using System;
using System.Collections.Generic;
using System.Windows;
using System.Windows.Input;
using System.Windows.Interop;
using Microsoft.Office.Interop.PowerPoint;
using Office = Microsoft.Office.Core;
using oPenEfficiency.Utils;

namespace oPenEfficiency.UI
{
    public partial class ApplyTextWindow : Window
    {
        private readonly PowerPointManager _manager;

        public ApplyTextWindow(PowerPointManager manager)
        {
            InitializeComponent();
            _manager = manager;
            var app = _manager?.GetApplication();
            if (app != null)
                new WindowInteropHelper(this).Owner = new IntPtr(app.HWND);
            TxtValue.Focus();

            CheckMasterButtonState();
        }

        private void CheckMasterButtonState()
        {
            try
            {
                var shapes = _manager.GetSelectedShapes();
                BtnMaster.IsEnabled = shapes != null &&
                    shapes.Count > 1 &&
                    GetMasterShape(shapes)?.HasTextFrame == Office.MsoTriState.msoTrue;
            }
            catch (Exception ex)
            {
                BtnMaster.IsEnabled = false;
                ExceptionLogger.Log(ex, "ApplyTextWindow.CheckMasterButtonState");
            }
        }

        private void Window_MouseDown(object sender, MouseButtonEventArgs e)
        {
            if (e.ChangedButton == MouseButton.Left)
                this.DragMove();
        }

        private void BtnCancel_Click(object sender, RoutedEventArgs e)
        {
            this.Close();
        }

        private void TxtValue_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.Key == Key.Enter)
            {
                BtnApply_Click(this, new RoutedEventArgs());
            }
        }

        private void BtnApply_Click(object sender, RoutedEventArgs e)
        {
            string text = TxtValue.Text;
            ApplyToSelection(text);
        }

        private void BtnMaster_Click(object sender, RoutedEventArgs e)
        {
            ApplyMasterText(null);
        }

        public bool ApplyMasterText(bool? useLastSelected)
        {
            try
            {
                var app = _manager.GetApplication();
                if (app.ActiveWindow?.Selection?.Type != PpSelectionType.ppSelectionShapes) return false;

                var shapes = app.ActiveWindow.Selection.ShapeRange;
                if (shapes.Count < 1) return false;

                var masterShape = GetMasterShape(shapes, useLastSelected);

                if (masterShape?.HasTextFrame == Office.MsoTriState.msoTrue)
                {
                    string masterText = masterShape.TextFrame.TextRange.Text;
                    ApplyToSelection(masterText);
                    TxtValue.Text = masterText; // Update UI to show what was applied
                    return true;
                }
            }
            catch (Exception ex)
            {
                ExceptionLogger.Log(ex, "ApplyTextWindow.ApplyMasterText");
            }
            return false;
        }

        private Shape GetMasterShape(ShapeRange shapes, bool? useLastSelected = null)
        {
            if (shapes == null || shapes.Count < 1) return null;
            // The first/last decision routes through the shared production decision point instead of
            // re-deciding the configured mode locally.
            bool useLast = useLastSelected ?? _manager.UsesLastSelectedAnchor();
            return useLast
                ? shapes[shapes.Count]
                : shapes[1];
        }

        private bool ApplyToSelection(string text)
        {
            var states = new List<TextState>();
            try
            {
                var app = _manager.GetApplication();
                if (app.ActiveWindow?.Selection == null) return false;

                var selection = app.ActiveWindow.Selection;
                if (selection.Type == PpSelectionType.ppSelectionShapes)
                {
                    foreach (Shape shape in selection.ShapeRange)
                    {
                        CollectTextStates(shape, states);
                    }
                }
                else if (selection.Type == PpSelectionType.ppSelectionText)
                {
                    states.Add(new TextState(selection.TextRange));
                }

                if (states.Count == 0) return false;
                try { app.StartNewUndoEntry(); }
                catch (Exception ex) { ExceptionLogger.Log(ex, "ApplyTextWindow.StartNewUndoEntry"); }

                foreach (var state in states) state.Range.Text = text;
                return true;
            }
            catch (Exception ex)
            {
                foreach (var state in states)
                {
                    try { state.Range.Text = state.Text; }
                    catch (Exception rollbackEx) { ExceptionLogger.Log(rollbackEx, "ApplyTextWindow.Rollback"); }
                }
                ExceptionLogger.Log(ex, "ApplyTextWindow.ApplyToSelection");
                return false;
            }
        }

        private static void CollectTextStates(Shape shape, List<TextState> states)
        {
            if (shape.HasTable == Office.MsoTriState.msoTrue)
            {
                bool hasSelectedCells = false;
                for (int row = 1; row <= shape.Table.Rows.Count && !hasSelectedCells; row++)
                {
                    for (int column = 1; column <= shape.Table.Columns.Count; column++)
                    {
                        if (shape.Table.Cell(row, column).Selected)
                        {
                            hasSelectedCells = true;
                            break;
                        }
                    }
                }

                for (int row = 1; row <= shape.Table.Rows.Count; row++)
                {
                    for (int column = 1; column <= shape.Table.Columns.Count; column++)
                    {
                        var cell = shape.Table.Cell(row, column);
                        if (hasSelectedCells && !cell.Selected) continue;

                        var cellShape = cell.Shape;
                        if (cellShape.HasTextFrame == Office.MsoTriState.msoTrue)
                            states.Add(new TextState(cellShape.TextFrame.TextRange));
                    }
                }
                return;
            }

            if (shape.HasTextFrame == Office.MsoTriState.msoTrue)
                states.Add(new TextState(shape.TextFrame.TextRange));
        }

        private sealed class TextState
        {
            public TextState(TextRange range)
            {
                Range = range;
                Text = range.Text;
            }

            public TextRange Range { get; }
            public string Text { get; }
        }
    }
}
