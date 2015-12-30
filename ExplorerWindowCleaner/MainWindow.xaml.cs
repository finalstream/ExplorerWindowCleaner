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
        private readonly ExplorerWindowCleanerClient _ewClient;

        public MainWindow(ExplorerWindowCleanerClient ewClient)
        {
            _ewClient = ewClient;
            InitializeComponent();

            DataContext = this;

            this.Closing += (sender, args) =>
            {
                args.Cancel = true;
                this.Visibility = Visibility.Hidden;
            };

            SwitchViewLabel = ClosedWindows;
        }

        public ObservableCollection<Explorer> NowExplorers
        {
            get { return _ewClient.Explorers; }
        }

        public ObservableCollection<Explorer> ClosedExplorers
        {
            get { return _ewClient.ClosedExplorers; }
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

        #region IsShowApplication変更通知プロパティ

        private bool _IsShowApplication;

        public bool IsShowApplication
        {
            get { return _IsShowApplication; }
            set
            {
                if (_IsShowApplication == value) return;
                _IsShowApplication = value;
                OnPropertyChanged("IsShowApplication");
            }
        }

        #endregion

        #region SwitchViewLabel変更通知プロパティ

        private string _switchViewLabel;

        public string SwitchViewLabel
        {
            get { return _switchViewLabel; }
            set
            {
                if (_switchViewLabel == value) return;
                _switchViewLabel = value;
                OnPropertyChanged("SwitchViewLabel");
            }
        }

        #endregion

        #region SwitchShowApplicationLabel変更通知プロパティ

        private string _SwitchShowApplicationLabel;

        public string SwitchShowApplicationLabel
        {
            get { return _SwitchShowApplicationLabel; }
            set
            {
                if (_SwitchShowApplicationLabel == value) return;
                _SwitchShowApplicationLabel = value;
                OnPropertyChanged("SwitchShowApplicationLabel");
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
            var item = (DataGridRow)((FrameworkElement)sender).Tag;
            var explorer = (Explorer)item.DataContext;
            _ewClient.SwitchPin(explorer);
        }

        private void ClosedList_OnMouseDown(object sender, MouseButtonEventArgs e)
        {
            SwitchList();
        }

        private void SwitchList()
        {
            if (IsShowClosed)
            {
                SwitchViewLabel = ClosedWindows;
            }
            else
            {
                SwitchViewLabel = NowWindows;
            }
            IsShowClosed = !IsShowClosed;
        }

        private void ClosedDataGrid_OnDoubleClick(object sender, MouseButtonEventArgs e)
        {
            var item = (DataGridRow)sender;
            var explorer = (Explorer)item.DataContext;

            _ewClient.OpenExplorer(explorer);
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
            _ewClient.OpenExplorer(CurrentExplorer);
        }

        private void AddFavoriteMenuItem_OnClick(object sender, RoutedEventArgs e)
        {
            if (CurrentExplorer == null) return;

            _ewClient.AddFavorite(CurrentExplorer);
        }

        private void NewWindow_OnClick(object sender, RoutedEventArgs e)
        {
            _ewClient.OpenExplorer("shell:MyComputerFolder");
        }

        private void Close_OnMouseLeftButtonDown(object sender, MouseButtonEventArgs e)
        {
            var item = (DataGridRow)((FrameworkElement)sender).Tag;
            var explorer = (Explorer)item.DataContext;

            _ewClient.CloseExplorer(explorer);
        }

        private void DeleteLocationMenuItem_OnClick(object sender, RoutedEventArgs e)
        {
            if (CurrentExplorer == null) return;

            _ewClient.RemoveExplorer(CurrentExplorer);
        }

        private void OpenFavs_OnClick(object sender, RoutedEventArgs e)
        {
            _ewClient.OpenFavoritedExplorer();
        }

        private void ShowApplocation_OnMouseDown(object sender, MouseButtonEventArgs e)
        {

            IsShowApplication = _ewClient.SwitchShowApplication();
        }
    }
}