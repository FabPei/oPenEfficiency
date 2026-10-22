using System.Windows;
using System.Windows.Controls;
using System.Windows.Input;

namespace oPenEfficiency.Excel.Features.ChartScanner
{
    /// <summary>Minimal modal text prompt (code-only WPF), used for entering axis values.</summary>
    internal sealed class PromptDialog : Window
    {
        private readonly TextBox _input;

        private PromptDialog(string title, string prompt, string initial)
        {
            Title = title;
            Width = 340;
            SizeToContent = SizeToContent.Height;
            WindowStartupLocation = WindowStartupLocation.CenterOwner;
            ResizeMode = ResizeMode.NoResize;
            Topmost = true;

            var panel = new StackPanel { Margin = new Thickness(16) };
            panel.Children.Add(new TextBlock { Text = prompt, Margin = new Thickness(0, 0, 0, 8), TextWrapping = TextWrapping.Wrap });

            _input = new TextBox { Text = initial ?? string.Empty };
            _input.SelectAll();
            panel.Children.Add(_input);

            var buttons = new StackPanel
            {
                Orientation = Orientation.Horizontal,
                HorizontalAlignment = HorizontalAlignment.Right,
                Margin = new Thickness(0, 12, 0, 0)
            };
            var ok = new Button { Content = "OK", Width = 80, IsDefault = true, Margin = new Thickness(0, 0, 8, 0) };
            var cancel = new Button { Content = "Cancel", Width = 80, IsCancel = true };
            ok.Click += (s, e) => { DialogResult = true; };
            buttons.Children.Add(ok);
            buttons.Children.Add(cancel);
            panel.Children.Add(buttons);

            Content = panel;
            Loaded += (s, e) => { _input.Focus(); _input.SelectAll(); };
        }

        /// <summary>Shows the prompt and returns the entered text, or null if cancelled.</summary>
        public static string Show(Window owner, string title, string prompt, string initial = "")
        {
            var dlg = new PromptDialog(title, prompt, initial);
            if (owner != null && owner.IsVisible) dlg.Owner = owner;
            return dlg.ShowDialog() == true ? dlg._input.Text : null;
        }
    }
}
