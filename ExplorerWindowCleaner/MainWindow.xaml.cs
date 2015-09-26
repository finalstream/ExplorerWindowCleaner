using System;
using System.Collections.ObjectModel;
using System.Diagnostics;
using System.Runtime.InteropServices;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Input;

namespace ExplorerWindowCleaner
{
    /// <summary>
    ///     MainWindow.xaml の相互作用ロジック
    /// </summary>
    public partial class MainWindow : Window
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

        [DllImport("user32.dll", SetLastError = true)]
        private static extern void SwitchToThisWindow(IntPtr hWnd, bool fAltTab);

        private void ListView_MouseDoubleClick(object sender, MouseButtonEventArgs e)
        {
            var item = (ListViewItem)sender;
            var explorer = (Explorer)item.DataContext;

            SwitchToThisWindow(new IntPtr(explorer.Handle), true);
            //Process.Start("EXPLORER.EXE", explorer.LocalPath);
        }
    }
}