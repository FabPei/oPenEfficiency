using System;
using System.Collections.ObjectModel;
using System.Linq;
using System.Windows;
using System.Windows.Input;
using Microsoft.Office.Interop.PowerPoint;
using Office = Microsoft.Office.Core;

namespace oPenEfficiency.UI
{
    public partial class ExcelLinkManagerWindow : Window
    {
        private PowerPointManager _pm;
        public ObservableCollection<ExcelLinkItem> Links { get; set; } = new ObservableCollection<ExcelLinkItem>();
        public bool UseRelativePaths { get; set; } = false;
        public bool AutoRefresh { get; set; } = false;

        public ExcelLinkManagerWindow(PowerPointManager pm)
        {
            InitializeComponent();
            _pm = pm;
            DataContext = this;
            LoadLinks();
        }

        private void LoadLinks()
        {
            Links.Clear();
            try {
                var pres = _pm.GetApplication().ActivePresentation;
                foreach (Slide slide in pres.Slides)
                {
                    foreach (Shape shape in slide.Shapes)
                    {
                        if (shape.Type == Office.MsoShapeType.msoLinkedOLEObject || shape.Type == Office.MsoShapeType.msoLinkedPicture)
                        {
                            Links.Add(new ExcelLinkItem { SlideNumber = slide.SlideNumber, ShapeName = shape.Name, SourcePath = shape.LinkFormat.SourceFullName, Status = "Linked", ShapeRef = shape });
                        }
                        else if (shape.HasChart == Office.MsoTriState.msoTrue && shape.Chart.ChartData != null)
                        {
                            try {
                                var wb = shape.Chart.ChartData.Workbook;
                                if (wb != null) {
                                    Links.Add(new ExcelLinkItem { SlideNumber = slide.SlideNumber, ShapeName = shape.Name, SourcePath = "Embedded Excel Chart", Status = "Active", ShapeRef = shape });
                                }
                            } catch { }
                        }
                    }
                }
            } catch { }
        }

        private void Window_MouseDown(object sender, MouseButtonEventArgs e) { if (e.ChangedButton == MouseButton.Left) DragMove(); }
        private void Close_Click(object sender, RoutedEventArgs e) => Close();

        private void Refresh_Click(object sender, RoutedEventArgs e)
        {
            foreach (var link in Links.Where(l => l.IsSelected))
            {
                try {
                    if (link.ShapeRef.Type == Office.MsoShapeType.msoLinkedOLEObject || link.ShapeRef.Type == Office.MsoShapeType.msoLinkedPicture)
                        link.ShapeRef.LinkFormat.Update();
                    else if (link.ShapeRef.HasChart == Office.MsoTriState.msoTrue)
                        link.ShapeRef.Chart.Refresh();
                    link.Status = "Refreshed";
                } catch { link.Status = "Error"; }
            }
        }

        private void BreakLink_Click(object sender, RoutedEventArgs e)
        {
            foreach (var link in Links.Where(l => l.IsSelected).ToList())
            {
                try {
                    if (link.ShapeRef.Type == Office.MsoShapeType.msoLinkedOLEObject || link.ShapeRef.Type == Office.MsoShapeType.msoLinkedPicture) {
                        link.ShapeRef.LinkFormat.BreakLink();
                        Links.Remove(link);
                    }
                } catch { }
            }
        }

        private void SplitScreen_Click(object sender, RoutedEventArgs e)
        {
            MessageBox.Show("Split Screen functionality requires deep OS-level window management (SetWindowPos P/Invoke) to arrange PowerPoint and Excel side-by-side. This is a prototype placeholder.", "Split Screen", MessageBoxButton.OK, MessageBoxImage.Information);
        }

        private void LinkAsImage_Click(object sender, RoutedEventArgs e)
        {
            MessageBox.Show("To link as image, copy a range in Excel, and paste it here using 'Paste Special -> Paste Link -> As Picture'. This UI will track msoLinkedPicture objects.", "Link as Image", MessageBoxButton.OK, MessageBoxImage.Information);
        }
    }

    public class ExcelLinkItem : System.ComponentModel.INotifyPropertyChanged
    {
        public bool IsSelected { get; set; }
        public int SlideNumber { get; set; }
        public string ShapeName { get; set; }
        public string SourcePath { get; set; }
        private string _status;
        public string Status { get => _status; set { _status = value; PropertyChanged?.Invoke(this, new System.ComponentModel.PropertyChangedEventArgs("Status")); } }
        public Shape ShapeRef { get; set; }

        public event System.ComponentModel.PropertyChangedEventHandler PropertyChanged;
    }
}
