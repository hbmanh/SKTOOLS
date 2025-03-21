using System;
using System.Globalization;
using System.Windows.Data;

namespace SKRevitAddins.Utils
{
    public class FractionToDoubleConverter : IValueConverter
    {
        public object Convert(object value, Type targetType, object parameter, CultureInfo culture)
        {
            if (value is double d)
            {
                // Chuyển về chuỗi thập phân
                return d.ToString(culture);
            }
            return value;
        }

        public object ConvertBack(object value, Type targetType, object parameter, CultureInfo culture)
        {
            var input = value as string;
            if (string.IsNullOrWhiteSpace(input)) return 0.0;

            if (input.Contains("/"))
            {
                var parts = input.Split('/');
                if (parts.Length == 2 &&
                    double.TryParse(parts[0].Trim(), NumberStyles.Any, culture, out double numerator) &&
                    double.TryParse(parts[1].Trim(), NumberStyles.Any, culture, out double denominator) &&
                    denominator != 0)
                {
                    return numerator / denominator;
                }
            }

            if (double.TryParse(input, NumberStyles.Any, culture, out double result))
                return result;

            return 0.0;
        }
    }
}
