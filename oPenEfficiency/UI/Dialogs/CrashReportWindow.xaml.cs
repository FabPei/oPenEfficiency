using System;
using System.Diagnostics;
using System.Web;
using System.Windows;

namespace oPenEfficiency.UI.Dialogs
{
    public partial class CrashReportWindow : Window
    {
        private readonly string _context;
        private readonly Exception _exception;

        public CrashReportWindow(Exception ex, string context)
        {
            InitializeComponent();
            
            _context = context ?? "Unknown Feature";
            _exception = ex;

            ContextTextBlock.Text = $"Context: {_context}";
            ErrorMessageBlock.Text = ex.Message;
            StackTraceBox.Text = ex.StackTrace ?? "No stack trace available.";

            if (ex.InnerException != null)
            {
                StackTraceBox.Text += "\n\nInner Exception:\n" + ex.InnerException.Message + "\n" + ex.InnerException.StackTrace;
            }
        }

        private void CopyError_Click(object sender, RoutedEventArgs e)
        {
            string errorDetails = $"Context: {_context}\nError: {_exception.Message}\n\nStack Trace:\n{StackTraceBox.Text}";
            Clipboard.SetText(errorDetails);
            MessageBox.Show("Error details copied to clipboard.", "Copied", MessageBoxButton.OK, MessageBoxImage.Information);
        }

        private void ReportBug_Click(object sender, RoutedEventArgs e)
        {
            try
            {
                string title = Uri.EscapeDataString($"[Crash] {_context}: {_exception.GetType().Name}");
                string body = Uri.EscapeDataString($"**Describe the bug**\nWhat were you doing when this happened?\n\n**Error Details**\nContext: `{_context}`\nMessage: `{_exception.Message}`\n\n**Stack Trace**\n```csharp\n{StackTraceBox.Text}\n```");
                
                string url = $"https://github.com/FabPei/oPenEfficiency/issues/new?title={title}&body={body}";
                Process.Start(new ProcessStartInfo(url) { UseShellExecute = true });
            }
            catch (Exception)
            {
                MessageBox.Show("Could not open browser. Please go to github.com/FabPei/oPenEfficiency/issues to report this.", "Error", MessageBoxButton.OK, MessageBoxImage.Warning);
            }
        }

        private void Close_Click(object sender, RoutedEventArgs e)
        {
            this.Close();
        }
    }
}