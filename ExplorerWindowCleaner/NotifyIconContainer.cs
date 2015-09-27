using System;
using System.ComponentModel;
using System.Windows;

namespace ExplorerWindowCleaner
{
    public partial class NotifyIconContainer : Component
    {
        private readonly ExplorerCleaner _explorerCleaner;
        private readonly MainWindow _mainWindow;

        public NotifyIconContainer(ExplorerCleaner explorerCleaner)
        {
            _explorerCleaner = explorerCleaner;
            _explorerCleaner.Start();
            _mainWindow = new MainWindow(_explorerCleaner);
            InitializeComponent();

            // コンテキストメニューのイベントを設定
            toolStripMenuItemOpen.Click += toolStripMenuItemOpen_Click;
            toolStripMenuItemExit.Click += toolStripMenuItemExit_Click;
            toolStripMenuItemAutoClose.Click += ToolStripMenuItemAutoCloseOnClick;
            toolStripMenuItemAutoClose.Checked = Properties.Settings.Default.IsAutoCloseUnused;
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
            _mainWindow.Show();
        }

        private void toolStripMenuItemExit_Click(object sender, EventArgs e)
        {
            // 現在のアプリケーションを終了
            Application.Current.Shutdown();
        }

        private void notifyIcon_DoubleClick(object sender, EventArgs e)
        {
            _mainWindow.Show();
            _mainWindow.WindowState = WindowState.Normal;
            _mainWindow.Activate();
        }
    }
}