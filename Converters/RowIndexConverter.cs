using qaImageViewer.Service;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Data;

namespace qaImageViewer.Converters
{
    class RowIndexConverter: IValueConverter
    {
        public object Convert(object value, Type targetType, object parameter, System.Globalization.CultureInfo culture)
        {
            int index = System.Convert.ToInt32(parameter);
            LoggerService.LogError($"{(value as List<string>)[index]} at index {index} where param was {System.Convert.ToString(parameter)}");
            return (value as List<string>)[index];
        }

        public object ConvertBack(object value, Type targetType, object parameter, System.Globalization.CultureInfo culture)
        {
            LoggerService.LogError("Convertback called");

            return DependencyProperty.UnsetValue;
        }


    }
}
