using System;
using System.Globalization;
using System.Windows.Data;
using System.Windows.Media.Imaging;

namespace ExplorerWindowCleaner.Converters
{
    class IsFavoritedConverter : IValueConverter
    {
        private readonly BitmapImage _favorited;
        private readonly BitmapImage _empty;

        public IsFavoritedConverter()
        {
            _favorited = new BitmapImage(new Uri(@"pack://application:,,,/Resources/favorite.png"));


            _empty = new BitmapImage(new Uri(@"pack://application:,,,/Resources/empty.png"));
        }
        public object Convert(object value, Type targetType, object parameter, CultureInfo culture)
        {
            var val = (bool)value;
            return val ? _favorited : _empty;
        }

        public object ConvertBack(object value, Type targetType, object parameter, CultureInfo culture)
        {
            return null;
        }
    }
}
