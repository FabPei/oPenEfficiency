using System;
using System.Windows;
using System.Windows.Controls;

namespace oPenEfficiency.UI
{
    public partial class SearchBarControl : UserControl
    {
        public event EventHandler<string> SearchChanged;

        public SearchBarControl()
        {
            InitializeComponent();
        }

        public void SetSearchText(string text)
        {
            TxtSearch.Text = text;
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
    }
}