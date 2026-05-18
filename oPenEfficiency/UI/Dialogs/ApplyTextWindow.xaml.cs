using System;
using System.Linq;
using System.Windows;
using System.Windows.Input;
using Microsoft.Office.Interop.PowerPoint;
using Office = Microsoft.Office.Core;

namespace oPenEfficiency.UI
{
    public partial class ApplyTextWindow : Window
    {
        private readonly PowerPointManager _manager;

        public ApplyTextWindow(PowerPointManager manager)
        {
            InitializeComponent();
            _manager = manager;
            TxtValue.Focus();
            
            CheckMasterButtonState();
        }

        private void CheckMasterButtonState()
        {
            try
            {
                var shapes = _manager.GetSelectedShapes();
                if (shapes == null || shapes.Count <= 1)
                {
                    BtnMaster.IsEnabled = false;
                }
            }
            catch
            {
                BtnMaster.IsEnabled = false;
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
            try
            {
                var app = _manager.GetApplication();
                if (app.ActiveWindow?.Selection?.Type != PpSelectionType.ppSelectionShapes) return;

                var shapes = app.ActiveWindow.Selection.ShapeRange;
                if (shapes.Count < 1) return;

                // Determine master shape based on settings
                var config = JsonService.Load<oPenEfficiency.Models.SidebarConfig>("sidebar_config.json");
                Shape masterShape;

                if (config != null && config.MasterObjectMode == "Last Selected")
                    masterShape = shapes[shapes.Count];
                else
                    masterShape = shapes[1];

                if (masterShape.HasTextFrame == Office.MsoTriState.msoTrue)
                {
                    string masterText = masterShape.TextFrame.TextRange.Text;
                    ApplyToSelection(masterText);
                    TxtValue.Text = masterText; // Update UI to show what was applied
                }
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine("Error in ApplyText Master: " + ex.Message);
            }
        }

        private void ApplyToSelection(string text)
        {
            try
            {
                var app = _manager.GetApplication();
                if (app.ActiveWindow?.Selection == null) return;

                var selection = app.ActiveWindow.Selection;

                if (selection.Type == PpSelectionType.ppSelectionShapes)
                {
                    foreach (Shape shape in selection.ShapeRange)
                    {
                        if (shape.HasTextFrame == Office.MsoTriState.msoTrue)
                        {
                            shape.TextFrame.TextRange.Text = text;
                        }
                    }
                }
                else if (selection.Type == PpSelectionType.ppSelectionText)
                {
                    selection.TextRange.Text = text;
                }
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine("Error applying text: " + ex.Message);
            }
        }
    }
}
