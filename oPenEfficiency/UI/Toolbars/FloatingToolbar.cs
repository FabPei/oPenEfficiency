using System;
using System.Drawing;
using System.Windows.Forms;
using Microsoft.Office.Core;

namespace oPenEfficiency.UI
{
    public class FloatingToolbar : Form
    {
        private ToolStrip _toolStrip;

        public FloatingToolbar()
        {
            this.FormBorderStyle = FormBorderStyle.None;
            this.ShowInTaskbar = false;
            this.StartPosition = FormStartPosition.Manual;
            this.TopMost = true;
            this.Size = new Size(1, 1); // Auto-size later
            this.AutoSize = true;
            this.AutoSizeMode = AutoSizeMode.GrowAndShrink;
            
            // To make it look like a floating menu/tooltip
            this.BackColor = Color.White;
            
            _toolStrip = new ToolStrip();
            _toolStrip.Dock = DockStyle.Fill;
            _toolStrip.GripStyle = ToolStripGripStyle.Hidden;
            _toolStrip.RenderMode = ToolStripRenderMode.System; // Or Professional
            _toolStrip.Padding = new Padding(2);
            
            this.Controls.Add(_toolStrip);

            // Close automatically when focus is lost
            this.Deactivate += (s, e) => 
            {
                // Small delay to allow button click to process first effectively
                this.Close(); 
            };
            
            // MouseEnter/Leave logic to keep it open? 
            // Standard ContextMenu behavior is usually simpler:
            // Click outside -> Close. Click item -> Action then Close (or stay open).
            // For now, close on Deactivate is robust.
        }

        protected override bool ProcessCmdKey(ref Message msg, Keys keyData)
        {
            if (keyData == Keys.Escape)
            {
                this.Close();
                return true;
            }
            return base.ProcessCmdKey(ref msg, keyData);
        }

        public void AddButton(string text, Image image, EventHandler onClick, bool isChecked = false, bool isBold = false)
        {
            var btn = new ToolStripButton(text, image, onClick);
            btn.Checked = isChecked;
            if (isBold) btn.Font = new Font(btn.Font, FontStyle.Bold);
            btn.DisplayStyle = ToolStripItemDisplayStyle.ImageAndText;
            if (image == null) btn.DisplayStyle = ToolStripItemDisplayStyle.Text;
            
            _toolStrip.Items.Add(btn);
        }

        public void AddSeparator()
        {
            _toolStrip.Items.Add(new ToolStripSeparator());
        }

        public void AddComboBox(string[] items, string currentText, EventHandler onSelectionChanged)
        {
            var combo = new ToolStripComboBox();
            combo.Items.AddRange(items);
            combo.Text = currentText;
            // Use selection change committed or text update? 
            // SelectedIndexChanged fires on text change too sometimes depending on mode.
            // Let's rely on Validated or explicit SelectionChangeCommitted equivalent (SelectedIndexChanged is usually fine here)
            combo.SelectedIndexChanged += (s, e) => onSelectionChanged?.Invoke(combo, EventArgs.Empty);
            
            // Fix width
            combo.Width = 50; 
            
            _toolStrip.Items.Add(combo);
        }

        public void AddSlider(int min, int max, int val, EventHandler onScroll)
        {
            var host = new ToolStripControlHost(new TrackBar 
            { 
                Minimum = min, Maximum = max, Value = val, 
                Width = 100, Height = 18, 
                TickStyle = TickStyle.None,
                AutoSize = false
            });
            var track = (TrackBar)host.Control;
            track.Scroll += onScroll;
            _toolStrip.Items.Add(host);
        }

        public void AddLabel(string text)
        {
            _toolStrip.Items.Add(new ToolStripLabel(text));
        }

        public void AddDropDownButton(string text, Image image, Action<ToolStripDropDownButton> populateItems)
        {
            var dropDown = new ToolStripDropDownButton(text, image);
            dropDown.DisplayStyle = ToolStripItemDisplayStyle.ImageAndText;
            if (image == null) dropDown.DisplayStyle = ToolStripItemDisplayStyle.Text;

            populateItems?.Invoke(dropDown);
            _toolStrip.Items.Add(dropDown);
        }

        public void SetVerticalLayout()
        {
            _toolStrip.LayoutStyle = ToolStripLayoutStyle.VerticalStackWithOverflow;
        }

        public void ShowAt(Point point)
        {
            this.Location = point;
            this.Show();
        }
    }
}
