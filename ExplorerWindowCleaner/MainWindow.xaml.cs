using System;
using System.Collections.ObjectModel;
using System.ComponentModel;
using System.Diagnostics;
using System.Runtime.InteropServices;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Input;
using MahApps.Metro.Controls;

namespace ExplorerWindowCleaner 
{
    /// <summary>
    ///     MainWindow.xaml の相互作用ロジック
    /// </summary>
    public partial class MainWindow : MetroWindow, INotifyPropertyChanged
    {
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

        [DllImport("user32.dll", SetLastError = true)]
        private static extern void SwitchToThisWindow(IntPtr hWnd, bool fAltTab);

        private void MenuItem_OnClick(object sender, RoutedEventArgs e)
        {
            if (CurrentExplorer == null) return;
            Clipboard.SetText(CurrentExplorer.LocationPath);
        }

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
            _ec.UpdateClosedDictionary(explorer);
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
            Process.Start("EXPLORER.EXE", string.Format("/n,/root,\"{0}\"", explorer.LocationPath));
        }
    }
}