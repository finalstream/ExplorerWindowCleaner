using System;
using System.Collections.ObjectModel;
using System.ComponentModel;
using System.Diagnostics;
using System.Reflection;
using System.Runtime.InteropServices;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Data;
using System.Windows.Input;
using MahApps.Metro.Controls;
using MahApps.Metro.Controls.Dialogs;
using Newtonsoft.Json;

namespace ExplorerWindowCleaner 
{
    /// <summary>
    ///     MainWindow.xaml の相互作用ロジック
    /// </summary>
    public partial class MainWindow : MetroWindow, INotifyPropertyChanged
    {
        [DllImport("user32.dll", SetLastError = true)]
        private static extern void SwitchToThisWindow(IntPtr hWnd, bool fAltTab);

        private static string ClosedWindows = "ClosedWindows";
        private static string NowWindows = "NowWindows";
        private readonly ExplorerCleaner _ec;

        public MainWindow(ExplorerCleaner ec)
        {
            _ec = ec;
            InitializeComponent();

            DataContext = this;

            this.Closing += (sender, args) =>
            {
                args.Cancel = true;
                this.Visibility = Visibility.Hidden;
            };

            SwitchLabel = ClosedWindows;
        }

        public ObservableCollection<Explorer> NowExplorers
        {
            get { return _ec.Explorers; }
        }

        public ObservableCollection<Explorer> ClosedExplorers
        {
            get { return _ec.ClosedExplorers; }
        }

        #region NowWindowCount変更通知プロパティ

        private int _nowWindowCount;

        public int NowWindowCount
        {
            get { return _nowWindowCount; }
            set
            {
                if (_nowWindowCount == value) return;
                _nowWindowCount = value;
                OnPropertyChanged("NowWindowCount");
            }
        }

        #endregion

        #region MaxWindowCount変更通知プロパティ

        private int _maxWindowCount;

        public int MaxWindowCount
        {
            get { return _maxWindowCount; }
            set
            {
                if (_maxWindowCount == value) return;
                _maxWindowCount = value;
                OnPropertyChanged("MaxWindowCount");
            }
        }

        #endregion

        #region PinedCount変更通知プロパティ

        private int _pinedCount;

        public int PinedCount
        {
            get { return _pinedCount; }
            set
            {
                if (_pinedCount == value) return;
                _pinedCount = value;
                OnPropertyChanged("PinedCount");
            }
        }

        #endregion

        #region TotalClosedWindow変更通知プロパティ

        private int _totalClosedWindow;

        public int TotalClosedWindow
        {
            get { return _totalClosedWindow; }
            set
            {
                if (_totalClosedWindow == value) return;
                _totalClosedWindow = value;
                OnPropertyChanged("TotalClosedWindow");
            }
        }

        #endregion

        #region CurrentExplorer変更通知プロパティ

        private Explorer _currentExplorer;

        public Explorer CurrentExplorer
        {
            get { return _currentExplorer; }
            set
            {
                if (_currentExplorer == value) return;
                _currentExplorer = value;
                OnPropertyChanged("CurrentExplorer");
            }
        }

        #endregion

        #region IsShowClosed変更通知プロパティ

        private bool _isShowClosed;

        public bool IsShowClosed
        {
            get { return _isShowClosed; }
            set
            {
                if (_isShowClosed == value) return;
                _isShowClosed = value;
                OnPropertyChanged("IsShowClosed");
            }
        }

        #endregion

        #region SwitchLabel変更通知プロパティ

        private string _switchLabel;

        public string SwitchLabel
        {
            get { return _switchLabel; }
            set
            {
                if (_switchLabel == value) return;
                _switchLabel = value;
                OnPropertyChanged("SwitchLabel");
            }
        }

        #endregion

        #region INotifyPropertyChanged メンバ

        public event PropertyChangedEventHandler PropertyChanged;
        protected virtual void OnPropertyChanged(string name)
        {
            if (PropertyChanged == null) return;

            PropertyChanged(this, new PropertyChangedEventArgs(name));
        }
        #endregion

        private void DataGrid_OnDoubleClick(object sender, MouseButtonEventArgs e)
        {
            var item = (DataGridRow)sender;
            var explorer = (Explorer)item.DataContext;

            SwitchToThisWindow(new IntPtr(explorer.Handle), true);
        }

        private void Pin_OnMouseLeftButtonDown(object sender, MouseButtonEventArgs e)
        {
            var item = (DataGridRow) ((FrameworkElement)sender).Tag;
            var explorer = (Explorer)item.DataContext;

            explorer.SwitchPined();
            var cloneExplorer = JsonConvert.DeserializeObject<Explorer>(JsonConvert.SerializeObject(explorer));
            cloneExplorer.IsFavorited = true; // ピン留めされたときにお気に入りに登録する。
            _ec.AddOrUpdateClosedDictionary(cloneExplorer);
        }

        private void ClosedList_OnMouseDown(object sender, MouseButtonEventArgs e)
        {
            SwitchList();
        }

        private void SwitchList()
        {
            if (IsShowClosed)
            {
                SwitchLabel = ClosedWindows;
            }
            else
            {
                SwitchLabel = NowWindows;
            }
            IsShowClosed = !IsShowClosed;
        }

        private void ClosedDataGrid_OnDoubleClick(object sender, MouseButtonEventArgs e)
        {
            var item = (DataGridRow)sender;
            var explorer = (Explorer)item.DataContext;
            _ec.OpenExplorer(explorer);
        }


        private void OpenFavorited_OnMouseDown(object sender, MouseButtonEventArgs e)
        {
            _ec.OpenFavoritedExplorer();
        }

        private void Favorite_OnMouseLeftButtonDown(object sender, MouseButtonEventArgs e)
        {
            var item = (DataGridRow)((FrameworkElement)sender).Tag;
            var explorer = (Explorer)item.DataContext;

            explorer.SwitchFavorited();
        }

        private void Version_OnClick(object sender, RoutedEventArgs e)
        {
            var assembly = Assembly.GetExecutingAssembly();
            var exeInfo = FileVersionInfo.GetVersionInfo(assembly.Location);

            this.ShowMessageAsync(
                string.Format("{0} ver.{1}", exeInfo.ProductName, exeInfo.ProductVersion),
                exeInfo.LegalCopyright + "\n" + "https://github.com/finalstream/ExplorerWindowCleaner");
        }

        private void DataGrid_OnSorting(object sender, DataGridSortingEventArgs e)
        {
            if (e.Column.SortDirection == ListSortDirection.Descending)
            {
                // 降順の次はソートを無効にする
                e.Column.SortDirection = null;
                e.Handled = true; // イベントを処理済みにする。（デフォルトのソート機能を実行しない）
                // このままでは矢印アイコンが消えて降順になるだけなので、以下の処理をいれる。
                var view = CollectionViewSource.GetDefaultView(((DataGrid)sender).ItemsSource);
                view.SortDescriptions.Clear();
            }
        }

        private void CopyLocationMenuItem_OnClick(object sender, RoutedEventArgs e)
        {
            if (CurrentExplorer == null) return;
            Clipboard.SetText(CurrentExplorer.LocationPath);
        }


        private void OpenLocationMenuItem_OnClick(object sender, RoutedEventArgs e)
        {
            if (CurrentExplorer == null) return;
            _ec.OpenExplorer(CurrentExplorer);
        }

        private void AddFavoriteMenuItem_OnClick(object sender, RoutedEventArgs e)
        {
            if (CurrentExplorer == null) return;
            var cloneExplorer = JsonConvert.DeserializeObject<Explorer>(JsonConvert.SerializeObject(CurrentExplorer));
            cloneExplorer.IsFavorited = true; // ピン留めされたときにお気に入りに登録する。
            _ec.AddOrUpdateClosedDictionary(cloneExplorer);
        }

        private void NewWindow_OnClick(object sender, RoutedEventArgs e)
        {
            _ec.OpenExplorer("shell:MyComputerFolder");
        }
    }
}