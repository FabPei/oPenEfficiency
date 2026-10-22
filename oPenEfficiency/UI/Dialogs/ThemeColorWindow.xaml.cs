using System;
using System.Windows;
using System.Windows.Input;

namespace oPenEfficiency.UI
{
    public partial class ThemeColorWindow : Window
    {
        public ThemeColorWindow(PowerPointManager manager, Features.PickColorFeature.ApplyTarget? target = null)
        {
            InitializeComponent();
            SourceInitialized += (sender, args) => FitToOwnerWorkArea();
            ColorControl.Initialize(manager);
            if (target.HasValue)
                Loaded += (s, e) => ColorControl.FocusTarget(target.Value);
        }

        public void FocusTarget(Features.PickColorFeature.ApplyTarget target)
        {
            ColorControl.FocusTarget(target);
            Activate();
        }

        private void FitToOwnerWorkArea()
        {
            var helper = new System.Windows.Interop.WindowInteropHelper(this);
            var screen = System.Windows.Forms.Screen.FromHandle(helper.Owner);
            var dpi = System.Windows.Media.VisualTreeHelper.GetDpi(this);
            MaxWidth = Math.Max(1, (screen.WorkingArea.Width / dpi.DpiScaleX) - 32);
            MaxHeight = Math.Max(1, (screen.WorkingArea.Height / dpi.DpiScaleY) - 32);
            Width = Math.Min(Width, MaxWidth);
            Height = Math.Min(Height, MaxHeight);
        }

        private void Window_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.Key == Key.Escape)
            {
                Close();
                e.Handled = true;
            }
        }

        private void Window_MouseDown(object sender, MouseButtonEventArgs e)
        {
            if (e.ChangedButton == MouseButton.Left && e.ButtonState == MouseButtonState.Pressed) DragMove();
        }

        private void BtnClose_Click(object sender, RoutedEventArgs e)
        {
            Close();
        }
    }
}
