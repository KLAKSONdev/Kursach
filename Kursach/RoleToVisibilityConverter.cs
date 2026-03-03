using System;
using System.Globalization;
using System.Windows;
using System.Windows.Data;

namespace Kursach
{
    public class RoleToVisibilityConverter : IValueConverter
    {
        public object Convert(object value, Type targetType, object parameter, CultureInfo culture)
        {
            if (value == null || parameter == null)
                return Visibility.Collapsed;

            string userRole = value.ToString();
            string[] allowedRoles = parameter.ToString().Split(',');

            foreach (string role in allowedRoles)
            {
                if (userRole == role.Trim())
                    return Visibility.Visible;
            }

            return Visibility.Collapsed;
        }

        public object ConvertBack(object value, Type targetType, object parameter, CultureInfo culture)
        {
            throw new NotImplementedException();
        }
    }
}