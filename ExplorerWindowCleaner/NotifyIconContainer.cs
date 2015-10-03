using System;
using System.ComponentModel;
using System.Windows;
using System.Windows.Forms;
using System.Windows.Threading;
using Application = System.Windows.Application;

namespace ExplorerWindowCleaner
{
    public partial class NotifyIconContainer : Component
    {
        private readonly ExplorerCleaner _explorerCleaner;
        private readonly MainWindow _mainWindow;

        public NotifyIconContainer(ExplorerCleaner explorerCleaner)
        {
            _explorerCleaner = explorerCleaner;
            
            _mainWindow = new MainWindow(_explorerCleaner);
            InitializeComponent();
            
            // コンテキストメニューのイベントを設定
            toolStripMenuItemOpen.Click += toolStripMenuItemOpen_Click;
            toolStripMenuItemExit.Click += toolStripMenuItemExit_Click;
            toolStripMenuItemAutoClose.Click += ToolStripMenuItemAutoCloseOnClick;
            toolStripMenuItemAutoClose.Checked = Properties.Settings.Default.IsAutoCloseUnused;

            _explorerCleaner.Updated += (sender, args) =>
            {
                notifyIcon.Text = string.Format("ExplorerWindowCleaner - {0} Windows", _explorerCleaner.WindowCount);
                _mainWindow.Dispatcher.Invoke(() =>
                {
                    toolStripMenuItemAutoClose.Text = _explorerCleaner.IsAutoCloseUnused
                        ? string.Format("Auto Close Unused expire:{0}",
                            _explorerCleaner.ExporeDateTime.ToString("yyyy-MM-dd HH:mm:ss"))
                        : "Auto Close Unused";
                });
                if (Properties.Settings.Default.IsNotifyCloseWindow && args.CloseWindowCount > 0)
                {
                    notifyIcon.ShowBalloonTip(3000, "Closed Window", string.Format("{0} Windows Closed.", args.CloseWindowCount), ToolTipIcon.Info);
                }
                _mainWindow.NowWindowCount = _explorerCleaner.WindowCount;
                _mainWindow.MaxWindowCount = _explorerCleaner.MaxWindowCount;
                _mainWindow.TotalClosedWindow = _explorerCleaner.TotalCloseWindowCount;
            };

            _explorerCleaner.Start();
        }

        private void ToolStripMenuItemAutoCloseOnClick(object sender, EventArgs eventArgs)
        {
            toolStripMenuItemAutoClose.Checked = !toolStripMenuItemAutoClose.Checked;
            _explorerCleaner.IsAutoCloseUnused = toolStripMenuItemAutoClose.Checked;
        }

        public NotifyIconContainer(IContainer container)
        {
            container.Add(this);

            InitializeComponent();
        }

        private void toolStripMenuItemOpen_Click(object sender, EventArgs e)
        {
            ShowWindowList();
        }

        private void toolStripMenuItemExit_Click(object sender, EventArgs e)
        {
            // 現在のアプリケーションを終了
            Application.Current.Shutdown();
        }

        private void notifyIcon_DoubleClick(object sender, EventArgs e)
        {
            ShowWindowList();
        }

        private void ShowWindowList()
        {
            _mainWindow.Show();
            _mainWindow.WindowState = WindowState.Normal;
            _mainWindow.Activate();
        }
    }
}