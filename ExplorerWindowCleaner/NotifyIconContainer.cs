﻿using System;
using System.ComponentModel;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Windows;
using System.Windows.Forms;
using System.Windows.Threading;
using NLog;
using Application = System.Windows.Application;
using MessageBox = System.Windows.Forms.MessageBox;

namespace ExplorerWindowCleaner
{
    public partial class NotifyIconContainer : Component
    {
        private readonly Logger _log = LogManager.GetCurrentClassLogger();
        private readonly ExplorerWindowCleanerClient _ewClient;
        private readonly MainWindow _mainWindow;

        private readonly string _ewclink = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.Startup),
                    "ExplorerWindowCleaner.lnk");

        public NotifyIconContainer(ExplorerWindowCleanerClient client)
        {
            _ewClient = client;
            
            _mainWindow = new MainWindow(_ewClient);
            InitializeComponent();
            
            // コンテキストメニューの設定
            SetContextMenuStartUp();
            toolStripMenuItemOpen.Click += toolStripMenuItemOpen_Click;
            toolStripMenuItemExit.Click += toolStripMenuItemExit_Click;
            toolStripMenuItemAutoClose.Click += ToolStripMenuItemAutoCloseOnClick;
            toolStripMenuItemStartup.Click += ToolStripMenuItemStartupOnClick;
            toolStripMenuItemAutoClose.Checked = client.AppConfig.IsAutoCloseUnused;

            _ewClient.Cleaned += (sender, args) =>
            {
                
                _mainWindow.Dispatcher.Invoke(() =>
                {
                    notifyIcon.Text = string.Format("ExplorerWindowCleaner - {0} Windows", args.WindowCount);
                    toolStripMenuItemAutoClose.Text = _ewClient.AppConfig.IsAutoCloseUnused
                        ? string.Format("Auto Close Unused expire:{0}",
                            args.ExpireDateTime.ToString("yyyy-MM-dd HH:mm:ss"))
                        : "Auto Close Unused";
                    if (_ewClient.AppConfig.IsNotifyCloseWindow && args.CloseInfos.Count > 0)
                    {
                        notifyIcon.ShowBalloonTip(3000,
                            string.Format("{0} Windows Closed.", args.CloseInfos.Count),
                            string.Format("{0}", string.Join("\n", args.CloseInfos.Select(x => string.Format("{0} [{1}]", x.WindowTitle, x.CloseReason.ToString())))), ToolTipIcon.Info);
                    }
                    _mainWindow.NowWindowCount = args.WindowCount;
                    _mainWindow.MaxWindowCount = args.MaxWindowCount;
                    _mainWindow.PinedCount = args.PinedCount;
                    _mainWindow.TotalClosedWindow = args.TotalCloseWindowCount;
                    if (args.IsUpdated) _mainWindow.UpdateView();
                });
                
               
            };

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
            // TargetPathプロパティをセットする
            shortcutType.InvokeMember("WorkingDirectory", BindingFlags.SetProperty, null, shortCut, new object[] { Path.GetDirectoryName(targetPath) });
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
            _ewClient.AppConfig.IsAutoCloseUnused = toolStripMenuItemAutoClose.Checked;
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