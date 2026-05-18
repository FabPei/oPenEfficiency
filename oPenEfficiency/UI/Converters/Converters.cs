using System;
using System.Globalization;
using System.Windows;
using System.Windows.Data;

namespace oPenEfficiency.UI
{
    public class StringEqualsConverter : IValueConverter
    {
        public static readonly StringEqualsConverter Instance = new StringEqualsConverter();
        public object Convert(object value, Type targetType, object parameter, CultureInfo culture)
        {
            var a = value?.ToString();
            var b = parameter?.ToString();
            return string.Equals(a, b, StringComparison.OrdinalIgnoreCase);
        }
        public object ConvertBack(object value, Type targetType, object parameter, CultureInfo culture)
        {
            if (value is bool b && b)
                return parameter?.ToString();
            return Binding.DoNothing;
        }
    }

    public class StringEqualsToVisibilityConverter : IValueConverter
    {
        public static readonly StringEqualsToVisibilityConverter Instance = new StringEqualsToVisibilityConverter();
        public object Convert(object value, Type targetType, object parameter, CultureInfo culture)
        {
            var a = value?.ToString();
            var b = parameter?.ToString();
            return string.Equals(a, b, StringComparison.OrdinalIgnoreCase) ? Visibility.Visible : Visibility.Collapsed;
        }
        public object ConvertBack(object value, Type targetType, object parameter, CultureInfo culture)
        {
            return Binding.DoNothing;
        }
    }

    public class BooleanToVisibilityConverter : IValueConverter
    {
        public static readonly BooleanToVisibilityConverter Instance = new BooleanToVisibilityConverter();
        public object Convert(object value, Type targetType, object parameter, CultureInfo culture)
        {
            return (value is bool b && b) ? Visibility.Visible : Visibility.Collapsed;
        }
        public object ConvertBack(object value, Type targetType, object parameter, CultureInfo culture) => throw new NotImplementedException();
    }

    public class InverseBooleanToVisibilityConverter : IValueConverter
    {
        public static readonly InverseBooleanToVisibilityConverter Instance = new InverseBooleanToVisibilityConverter();
        public object Convert(object value, Type targetType, object parameter, CultureInfo culture)
        {
            return (value is bool b && b) ? Visibility.Collapsed : Visibility.Visible;
        }
        public object ConvertBack(object value, Type targetType, object parameter, CultureInfo culture) => throw new NotImplementedException();
    }

    /// <summary>
    /// Converts item spacing from points to a formatted string showing points, centimeters, and inches.
    /// Example: "15 pts | 0.53 cm | 0.21 in"
    /// </summary>
    public class ItemSpacingConverter : IValueConverter
    {
        public static readonly ItemSpacingConverter Instance = new ItemSpacingConverter();

        // 1 point = 1/72 inch = 0.0352778 cm
        private const double PointsToCm = 0.0352778;
        private const double PointsToInches = 1.0 / 72.0;

        public object Convert(object value, Type targetType, object parameter, CultureInfo culture)
        {
            if (value is double doubleValue)
            {
                return FormatSpacing(doubleValue);
            }
            if (value is int intValue)
            {
                return FormatSpacing(intValue);
            }
            if (value is string stringValue && double.TryParse(stringValue, NumberStyles.Any, CultureInfo.InvariantCulture, out double parsedValue))
            {
                return FormatSpacing(parsedValue);
            }
            return "15 pts | 0.53 cm | 0.21 in"; // Default value
        }

        public object ConvertBack(object value, Type targetType, object parameter, CultureInfo culture)
        {
            // For two-way binding, just return the original value
            return value;
        }

        private string FormatSpacing(double points)
        {
            double cm = points * PointsToCm;
            double inches = points * PointsToInches;
            return $"{points} pts | {cm:F2} cm | {inches:F2} in";
        }
    }
}
