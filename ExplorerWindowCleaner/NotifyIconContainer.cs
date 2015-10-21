using System;
using System.ComponentModel;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Windows;
using System.Windows.Forms;
using System.Windows.Threading;
using Application = System.Windows.Application;
using MessageBox = System.Windows.Forms.MessageBox;

namespace ExplorerWindowCleaner
{
    public partial class NotifyIconContainer : Component
    {
        private readonly ExplorerCleaner _explorerCleaner;
        private readonly MainWindow _mainWindow;

        private readonly string _ewclink = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.Startup),
                    "ExplorerWindowCleaner.lnk");

        public NotifyIconContainer(ExplorerCleaner explorerCleaner)
        {
            _explorerCleaner = explorerCleaner;
            
            _mainWindow = new MainWindow(_explorerCleaner);
            InitializeComponent();
            
            // コンテキストメニューの設定
            SetContextMenuStartUp();
            toolStripMenuItemOpen.Click += toolStripMenuItemOpen_Click;
            toolStripMenuItemExit.Click += toolStripMenuItemExit_Click;
            toolStripMenuItemAutoClose.Click += ToolStripMenuItemAutoCloseOnClick;
            toolStripMenuItemStartup.Click += ToolStripMenuItemStartupOnClick;
            toolStripMenuItemAutoClose.Checked = Properties.Settings.Default.IsAutoCloseUnused;

            _explorerCleaner.WindowClosed += (sender, args) =>
            {
                
                _mainWindow.Dispatcher.Invoke(() =>
                {
                    notifyIcon.Text = string.Format("ExplorerWindowCleaner - {0} Windows", _explorerCleaner.WindowCount);
                    toolStripMenuItemAutoClose.Text = _explorerCleaner.IsAutoCloseUnused
                        ? string.Format("Auto Close Unused expire:{0}",
                            _explorerCleaner.ExporeDateTime.ToString("yyyy-MM-dd HH:mm:ss"))
                        : "Auto Close Unused";
                    if (Properties.Settings.Default.IsNotifyCloseWindow && args.CloseWindowTitles.Count > 0)
                    {
                        notifyIcon.ShowBalloonTip(3000,
                            string.Format("{0} Windows Closed.", args.CloseWindowTitles.Count),
                            string.Format("{0}", string.Join("\n", args.CloseWindowTitles)), ToolTipIcon.Info);
                    }
                    _mainWindow.NowWindowCount = _explorerCleaner.WindowCount;
                    _mainWindow.MaxWindowCount = _explorerCleaner.MaxWindowCount;
                    _mainWindow.PinedCount = _explorerCleaner.PinedCount;
                    _mainWindow.TotalClosedWindow = _explorerCleaner.TotalCloseWindowCount;
                });
                
               
            };

            _explorerCleaner.Start();
        }

        private void ToolStripMenuItemStartupOnClick(object sender, EventArgs eventArgs)
        {
            toolStripMenuItemStartup.Checked = !toolStripMenuItemStartup.Checked;
            RegistStartup(toolStripMenuItemStartup.Checked);
        }

        private void RegistStartup(bool isRegist)
        {
            try
            {
                if (isRegist)
                {
                    CreateShortCut(
                        _ewclink,
                        Assembly.GetExecutingAssembly().Location,
                        "");
                }
                else
                {
                    File.Delete(_ewclink);
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(string.Format("fail regist startup. {0}", ex));
            }
        }

        /// <summary>
        /// ショートカットの作成
        /// </summary>
        /// <remarks>WSHを使用して、ショートカット(lnkファイル)を作成します。(遅延バインディング)</remarks>
        /// <param name="path">出力先のファイル名(*.lnk)</param>
        /// <param name="targetPath">対象のアセンブリ(*.exe)</param>
        /// <param name="description">説明</param>
        private static void CreateShortCut(String path, String targetPath, String description)
        {
            //using System.Reflection;

            // WSHオブジェクトを作成し、CreateShortcutメソッドを実行する
            var shellType = Type.GetTypeFromProgID("WScript.Shell");
            var shell = Activator.CreateInstance(shellType);
            var shortCut = shellType.InvokeMember("CreateShortcut", BindingFlags.InvokeMethod, null, shell, new object[] { path });

            var shortcutType = shell.GetType();
            // TargetPathプロパティをセットする
            shortcutType.InvokeMember("TargetPath", BindingFlags.SetProperty, null, shortCut, new object[] { targetPath });
            // Descriptionプロパティをセットする
            shortcutType.InvokeMember("Description", BindingFlags.SetProperty, null, shortCut, new object[] { description });
            // Saveメソッドを実行する
            shortcutType.InvokeMember("Save", BindingFlags.InvokeMethod, null, shortCut, null);

        }

        private void SetContextMenuStartUp()
        {
            toolStripMenuItemStartup.Checked = File.Exists(_ewclink);
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