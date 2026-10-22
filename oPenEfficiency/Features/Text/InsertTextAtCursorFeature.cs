using System;
using Microsoft.Office.Interop.PowerPoint;
using Office = Microsoft.Office.Core;
using oPenEfficiency.Utils;
using oPenEfficiency.Services.Attributes;
using WinForms = System.Windows.Forms;

namespace oPenEfficiency.Features
{
    /// <summary>
    /// Feature: Insert text at the current cursor position or at the end of the selected shape's text.
    /// </summary>
    [FeatureMetadata(
        Id = "BtnInsertText",
        Name = "Insert Text",
        Tooltip = "Insert text at cursor",
        IconData = "M19,13H13V19H11V13H5V11H11V5H13V11H19V13Z",
        Color = "#FBBF24",
        Description = "Prompts for text and inserts it at the current text cursor or at the end of one selected text-bearing shape.",
        DetailedHelpText = "### Insert Text\n\nEnter a text snippet, then insert it after the active text cursor or selection. Right-click to insert a line break, bullet, en dash, or em dash directly. With exactly one text-bearing shape selected, the text is appended to that shape. Cancel and empty input leave the presentation unchanged; successful insertion creates one PowerPoint Undo entry.",
        MinSelection = 0)]
    public static class InsertTextAtCursorFeature
    {
        /// <summary>
        /// Opens the text prompt used by the auto-discovered sidebar command.
        /// </summary>
        public static bool Execute(PowerPointManager manager)
        {
            if (!CanInsert(manager)) return false;

            var text = PromptForText(manager);
            return text != null && Execute(manager, text);
        }

        public static bool Execute(PowerPointManager manager, string text)
        {
            if (manager == null || string.IsNullOrEmpty(text)) return false;

            try
            {
                var app = manager.GetApplication();
                if (app.ActiveWindow == null) return false;

                var selection = app.ActiveWindow.Selection;
                try { app.StartNewUndoEntry(); }
                catch (Exception ex) { ExceptionLogger.Log(ex, "InsertTextAtCursorFeature.StartNewUndoEntry"); }

                if (selection.Type == PpSelectionType.ppSelectionText)
                {
                    selection.TextRange.InsertAfter(text);
                    return true;
                }
                else if (selection.Type == PpSelectionType.ppSelectionShapes && selection.ShapeRange.Count == 1)
                {
                    var shape = selection.ShapeRange[1];
                    if (shape.HasTextFrame == Office.MsoTriState.msoTrue)
                    {
                        shape.TextFrame.TextRange.InsertAfter(text);
                        return true;
                    }
                }
                return false;
            }
            catch (Exception ex)
            {
                ExceptionLogger.Log(ex, "InsertTextAtCursorFeature.Execute");
                return false;
            }
        }

        private static bool CanInsert(PowerPointManager manager)
        {
            if (manager == null) return false;
            try
            {
                var app = manager.GetApplication();
                if (app?.ActiveWindow?.Selection == null) return false;
                var selection = app.ActiveWindow.Selection;
                if (selection.Type == PpSelectionType.ppSelectionText)
                    return selection.TextRange != null;

                if (selection.Type != PpSelectionType.ppSelectionShapes ||
                    selection.ShapeRange.Count != 1)
                    return false;

                return selection.ShapeRange[1].HasTextFrame == Office.MsoTriState.msoTrue;
            }
            catch (Exception ex)
            {
                ExceptionLogger.Log(ex, "InsertTextAtCursorFeature.CanInsert");
                return false;
            }
        }

        private static string PromptForText(PowerPointManager manager)
        {
            using (var form = new WinForms.Form())
            using (var label = new WinForms.Label())
            using (var textBox = new WinForms.TextBox())
            using (var ok = new WinForms.Button())
            using (var cancel = new WinForms.Button())
            {
                form.Text = "Insert Text";
                form.FormBorderStyle = WinForms.FormBorderStyle.FixedDialog;
                form.StartPosition = WinForms.FormStartPosition.CenterParent;
                form.MinimizeBox = false;
                form.MaximizeBox = false;
                form.ShowInTaskbar = false;
                form.ClientSize = new System.Drawing.Size(400, 180);

                label.SetBounds(12, 12, 376, 20);
                label.Text = "Text to insert:";

                textBox.SetBounds(12, 38, 376, 88);
                textBox.Multiline = true;
                textBox.ScrollBars = WinForms.ScrollBars.Vertical;

                ok.SetBounds(220, 140, 80, 28);
                ok.Text = "Insert";
                ok.DialogResult = WinForms.DialogResult.OK;

                cancel.SetBounds(308, 140, 80, 28);
                cancel.Text = "Cancel";
                cancel.DialogResult = WinForms.DialogResult.Cancel;

                form.Controls.AddRange(new WinForms.Control[] { label, textBox, ok, cancel });
                form.AcceptButton = ok;
                form.CancelButton = cancel;

                var app = manager.GetApplication();
                var result = app != null
                    ? form.ShowDialog(new NativeWindowOwner(new IntPtr(app.HWND)))
                    : form.ShowDialog();
                return result == WinForms.DialogResult.OK ? textBox.Text : null;
            }
        }

        private sealed class NativeWindowOwner : WinForms.IWin32Window
        {
            public NativeWindowOwner(IntPtr handle)
            {
                Handle = handle;
            }

            public IntPtr Handle { get; }
        }
    }
}
