using System;
using System.Globalization;
using System.Windows.Data;
using System.Windows.Media;
using oPenEfficiency.UI;

namespace oPenEfficiency
{
    public class FeatureIconConverter : IValueConverter
    {
        public object Convert(object value, Type targetType, object parameter, CultureInfo culture)
        {
            if (value is string id)
            {
                var info = FeatureLibrary.GetFeatureInfo(id);
                if (parameter?.ToString() == "Color")
                {
                    string hex = string.IsNullOrEmpty(info.Color) ? "#888888" : info.Color;
                    try { return (SolidColorBrush)new BrushConverter().ConvertFrom(hex); }
                    catch { return Brushes.Gray; }
                }
                return string.IsNullOrEmpty(info.IconData) ? "" : info.IconData;
            }
            return null;
        }

        public object ConvertBack(object value, Type targetType, object parameter, CultureInfo culture)
        {
            throw new NotImplementedException();
        }
    }
}
