using System;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Input;

namespace oPenEfficiency.UI
{
    public partial class SearchBarControl : UserControl
    {
        public event EventHandler<string> SearchChanged;
        public event EventHandler<string> SearchSubmitted;

        public SearchBarControl()
        {
            InitializeComponent();
        }

        public void SetSearchText(string text)
        {
            string value = text ?? string.Empty;
            if (string.Equals(TxtSearch.Text, value, StringComparison.Ordinal)) return;
            TxtSearch.Text = value;
            TxtSearch.CaretIndex = TxtSearch.Text.Length;
        }

        public void FocusSearch()
        {
            Dispatcher.BeginInvoke(new Action(() =>
            {
                TxtSearch.Focus();
                Keyboard.Focus(TxtSearch);
                TxtSearch.SelectAll();
            }), System.Windows.Threading.DispatcherPriority.Input);
        }

        private void TxtSearch_TextChanged(object sender, TextChangedEventArgs e)
        {
            if (TxtSearchPlaceholder != null)
                TxtSearchPlaceholder.Visibility = string.IsNullOrEmpty(TxtSearch.Text) ? Visibility.Visible : Visibility.Collapsed;
            
            SearchChanged?.Invoke(this, TxtSearch.Text);
        }

        private void TxtSearch_GotFocus(object sender, RoutedEventArgs e)
        {
            if (TxtSearchPlaceholder != null)
                TxtSearchPlaceholder.Visibility = Visibility.Collapsed;
        }

        private void TxtSearch_LostFocus(object sender, RoutedEventArgs e)
        {
            if (TxtSearchPlaceholder != null && string.IsNullOrEmpty(TxtSearch.Text))
                TxtSearchPlaceholder.Visibility = Visibility.Visible;
        }

        private void TxtSearch_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.Key == Key.Enter)
            {
                string query = TxtSearch.Text?.Trim();
                if (!string.IsNullOrEmpty(query)) SearchSubmitted?.Invoke(this, query);
                e.Handled = true;
            }
            else if (e.Key == Key.Escape)
            {
                SetSearchText(string.Empty);
                e.Handled = true;
            }
        }
    }
}
