using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Windows;
using Microsoft.Office.Interop.PowerPoint;
using oPenEfficiency.Services;

namespace oPenEfficiency.UI.Dialogs
{
    public partial class EditAssetDetailsDialog : Window
    {
        private string _filePath;
        public string AssetName => TxtName.Text;
        public string Tags => TxtTags.Text;

        public EditAssetDetailsDialog(string currentName, string currentTags, string filePath = null)
        {
            InitializeComponent();
            _filePath = filePath;
            TxtName.Text = currentName;
            TxtTags.Text = currentTags;
            TxtName.Focus();
            TxtName.SelectAll();

            if (string.IsNullOrEmpty(_filePath) || !File.Exists(_filePath) || !_filePath.EndsWith(".pptx", StringComparison.OrdinalIgnoreCase))
            {
                BtnAutoTag.Visibility = Visibility.Collapsed;
            }
        }

        private void BtnCancel_Click(object sender, RoutedEventArgs e)
        {
            DialogResult = false;
            Close();
        }

        private void BtnSave_Click(object sender, RoutedEventArgs e)
        {
            if (string.IsNullOrWhiteSpace(TxtName.Text))
            {
                MessageBox.Show("Please enter a name for the asset.", "Validation Error", MessageBoxButton.OK, MessageBoxImage.Warning);
                return;
            }
            DialogResult = true;
            Close();
        }

        private void BtnAutoTag_Click(object sender, RoutedEventArgs e)
        {
            if (string.IsNullOrEmpty(_filePath) || !File.Exists(_filePath)) return;

            try
            {
                var ppManager = new PowerPointManager(Globals.ThisAddIn.Application);
                var app = ppManager.GetApplication();
                
                Presentation pres = null;
                try
                {
                    pres = app.Presentations.Open(_filePath, 
                        Microsoft.Office.Core.MsoTriState.msoTrue, 
                        Microsoft.Office.Core.MsoTriState.msoFalse, 
                        Microsoft.Office.Core.MsoTriState.msoFalse);

                    var autoTags = new List<string>();
                    foreach (Slide slide in pres.Slides)
                    {
                        autoTags.AddRange(AutoTaggingService.IdentifyTags(slide));
                    }
                        
                    if (autoTags.Any())
                    {
                        var existingTags = TxtTags.Text.Split(new[] { ',' }, StringSplitOptions.RemoveEmptyEntries)
                                                      .Select(t => t.Trim())
                                                      .ToList();
                        
                        var combined = existingTags.Concat(autoTags).Distinct(StringComparer.OrdinalIgnoreCase);
                        TxtTags.Text = string.Join(", ", combined);
                    }
                    else
                    {
                        MessageBox.Show("No specific tags could be identified for this content.", "Auto-tagging", MessageBoxButton.OK, MessageBoxImage.Information);
                    }
                }
                finally
                {
                    pres?.Close();
                    if (pres != null) System.Runtime.InteropServices.Marshal.ReleaseComObject(pres);
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show("Error identifying tags: " + ex.Message, "Error", MessageBoxButton.OK, MessageBoxImage.Error);
            }
        }
    }
}
