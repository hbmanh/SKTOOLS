using System;
using System.Globalization;
using System.Windows.Data;

namespace SKRevitAddins.Commands.LayoutsToDWG
{
    public class EnumBooleanConverter : IValueConverter
    {
        public object Convert(object value, Type targetType, object parameter, CultureInfo culture)
            => value?.ToString().Equals(parameter?.ToString(), StringComparison.OrdinalIgnoreCase) ?? false;

        public object ConvertBack(object value, Type targetType, object parameter, CultureInfo culture)
            => (bool)value ? Enum.Parse(targetType, parameter.ToString()) : Binding.DoNothing;
    }
}
