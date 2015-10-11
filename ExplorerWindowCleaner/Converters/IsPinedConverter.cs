using System;
using System.Globalization;
using System.Windows.Data;
using System.Windows.Media.Imaging;

namespace ExplorerWindowCleaner.Converters
{
    class IsPinedConverter : IValueConverter
    {
        private readonly BitmapImage _pined;
        private readonly BitmapImage _empty;

        public IsPinedConverter()
        {
            _pined = new BitmapImage(new Uri(@"pack://application:,,,/Resources/pin.png"));


            _empty = new BitmapImage(new Uri(@"pack://application:,,,/Resources/empty.png"));
        }
        public object Convert(object value, Type targetType, object parameter, CultureInfo culture)
        {
            var val = (bool)value;
            return val ? _pined : _empty;
        }

        public object ConvertBack(object value, Type targetType, object parameter, CultureInfo culture)
        {
            return null;
        }
    }
}
