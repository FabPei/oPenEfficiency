using System;
using System.Globalization;
using System.Windows.Data;
using oPenEfficiency.UI;

namespace oPenEfficiency
{
    public class FeatureIdToDescriptionConverter : IValueConverter
    {
        public object Convert(object value, Type targetType, object parameter, CultureInfo culture)
        {
            if (value is string id)
            {
                var info = FeatureLibrary.GetFeatureInfo(id);
                return string.IsNullOrEmpty(info.Description) ? null : info.Description;
            }
            return null;
        }

        public object ConvertBack(object value, Type targetType, object parameter, CultureInfo culture)
        {
            throw new NotImplementedException();
        }
    }
}
