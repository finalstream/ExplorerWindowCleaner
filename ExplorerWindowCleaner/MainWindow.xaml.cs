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
        }

        public ObservableCollection<Explorer> Explorers
        {
            get { return _ec.Explorers; }
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


        [DllImport("user32.dll", SetLastError = true)]
        private static extern void SwitchToThisWindow(IntPtr hWnd, bool fAltTab);

        private void ListView_MouseDoubleClick(object sender, MouseButtonEventArgs e)
        {
            var item = (ListViewItem)sender;
            var explorer = (Explorer)item.DataContext;

            SwitchToThisWindow(new IntPtr(explorer.Handle), true);
            //Process.Start("EXPLORER.EXE", explorer.LocalPath);
        }

        private void MenuItem_OnClick(object sender, RoutedEventArgs e)
        {
            if (CurrentExplorer == null) return;
            Clipboard.SetText(CurrentExplorer.LocalPath);
        }

        #region INotifyPropertyChanged メンバ

        public event PropertyChangedEventHandler PropertyChanged;
        protected virtual void OnPropertyChanged(string name)
        {
            if (PropertyChanged == null) return;

            PropertyChanged(this, new PropertyChangedEventArgs(name));
        }
        #endregion
    }
}