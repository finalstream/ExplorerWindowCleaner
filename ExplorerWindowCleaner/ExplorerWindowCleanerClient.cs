using ExplorerWindowCleaner.Actions;
using ExplorerWindowCleaner.Properties;
using FinalstreamCommons.Windows;
using Firk.Core;
using Firk.Core.Actions;
using MahApps.Metro;
using Newtonsoft.Json;
using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.Diagnostics;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Runtime.InteropServices;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Controls.Primitives;
using System.Windows.Forms;
using static System.Net.Mime.MediaTypeNames;
using Application = System.Windows.Application;

namespace ExplorerWindowCleaner
{
    public class ExplorerWindowCleanerClient : AppClient<ExplorerWindowCleanerAppConfig>
    {
        public const string ClipboardFileName = "clipboard.json";
        public const string ShortcutFileName = "shortcut.json";
        private ActionExecuter<ExplorerWindowCleanerClientOperator> _actionExecuter;
        private Queue<ClipboardHistoryItem> _clipboardItemQueue;
        private ClipboardMonitor _clipboardMonitor;
        //private ContextMenuStrip _contextMenuClipboardHistories;
        private ContextMenuStrip _contextMenuShortcuts;
        private ExplorerCleaner _explorerCleaner;
        private GlobalMouseHook _globalMouseHook;
        private ShortcutItem[] _shortcuts;
        private System.Windows.Controls.ContextMenu _clipboardMenu;


        public ExplorerWindowCleanerClient(Assembly executingAssembly) : base(executingAssembly)
        {
        }

        public ObservableCollection<Explorer> Explorers
        {
            get { return _explorerCleaner.Explorers; }
        }

        public ObservableCollection<Explorer> ClosedExplorers
        {
            get { return _explorerCleaner.ClosedExplorers; }
        }

        public override string GetConfigPath()
        {
            return Path.Combine(CurrentDirectory, "ExplorerWindowCleanerConfig.json");
        }

        private sealed class HwndWrapper : System.Windows.Forms.IWin32Window
        {
            public IntPtr Handle { get; }
            public HwndWrapper(IntPtr handle) { Handle = handle; }
        }


        protected override void InitializeCore()
        {
            // get the theme from the current application
            var theme = ThemeManager.DetectAppStyle(Application.Current);

            // now set the Green accent and dark theme
            ThemeManager.ChangeAppStyle(Application.Current,
                ThemeManager.GetAccent(AppConfig.AccentColor),
                ThemeManager.GetAppTheme(AppConfig.AppTheme));

            _explorerCleaner = new ExplorerCleaner(AppConfig);
            _explorerCleaner.Cleaned += (sender, args) => OnCleaned(args);
            _actionExecuter =
                new ActionExecuter<ExplorerWindowCleanerClientOperator>(new ExplorerWindowCleanerClientOperator(this,
                    _explorerCleaner));
            DisposableCollection.Add(_actionExecuter);
            ResetBackgroundWorker(AppConfig.Interval, new BackgroundAction[] {new CleanerAction(this)});

            if (AppConfig.IsMouseHook)
            {
                RestoreClipboardHistories();

                MonitoringClipboard();

                CreateShortcutContextMenu();
                CreateClipboardContextMenu();

                _globalMouseHook = new GlobalMouseHook();
                _globalMouseHook.MouseHooked += (sender, args) =>
                {
                    if (args.MouseButton == MouseButtons.Left && args.IsDesktop)
                    {
                        //_contextMenuClipboardHistories.Close(ToolStripDropDownCloseReason.AppFocusChange);
                        Application.Current.Dispatcher.BeginInvoke(new Action(() =>
                        {
                            var p = args.Point; // screen座標（System.Drawing.Point）

                            // メニュー外クリックだけ閉じる
                            //if (!_clipboardMenuScreenRect.Contains(p))
                            Application.Current.Dispatcher.BeginInvoke(new Action(() =>
                            {
                                    _clipboardMenu.IsOpen = false;
                            }), System.Windows.Threading.DispatcherPriority.ApplicationIdle);

                            _contextMenuShortcuts.Close(ToolStripDropDownCloseReason.AppFocusChange);
                        }));
                        
                    }

                    if (args.IsDoubleClick)
                    {
                        if (args.MouseButton == MouseButtons.Right)
                        {

                            if (_windowHwnd != IntPtr.Zero)
                                SetForegroundWindow(_windowHwnd);
                            //_contextMenuClipboardHistories.Tag = _windowHwnd;
                            //var owner = new HwndWrapper(_windowHwnd);
                            //_contextMenuClipboardHistories.Show( args.Point);

                            Application.Current.Dispatcher.BeginInvoke(new Action(() =>
                            {
                                _clipboardMenu.Items.Clear();
                            var clipboardHistoriesItem = new System.Windows.Controls.MenuItem
                            {
                                Header = "Clipboard Histories",
                               IsEnabled = false
                            };
                             clipboardHistoriesItem.Click += (s, e) =>
                             {
                                 _clipboardMenu.IsOpen = false; // 閉じる → Closed で処理する
                             };
                            _clipboardMenu.Items.Add(clipboardHistoriesItem);
                            _clipboardMenu.Items.Add(new Separator());

                            foreach (var c in _clipboardItemQueue.Reverse())
                            {
                                var text = c.GetText();
                                var hwnd = this._windowHwnd;

                                var item = new System.Windows.Controls.MenuItem
                                {
                                    Header = c.ToString(),
                                    ToolTip = c.GetText()
                                };
                                item.Click += (s, e) =>
                                {
                                    _pendingText = ((System.Windows.Controls.MenuItem)s).ToolTip.ToString();
                                    _pendingHwnd = _windowHwnd;
                                    _pendingFromClick = true;

                                    

                                    _clipboardMenu.IsOpen = false; // 閉じる → Closed で処理する
                                };
                                _clipboardMenu.Items.Add(item);
                                _clipboardMenu.Items.Add(new Separator());
                            }
                            _clipboardMenu.Items.RemoveAt(_clipboardMenu.Items.Count - 1);
                                _clipboardMenu.MaxHeight = SystemParameters.WorkArea.Height * 0.8;
                                _clipboardMenu.IsOpen = true;
                            }));
                        }
                        else if (args.MouseButton == MouseButtons.Left && args.IsDesktop)
                        {
                            _contextMenuShortcuts.Show(args.Point);
                            Debug.WriteLine($"Show Shortcut Menu. {args.Point}");
                        }
                        Debug.WriteLine("DoubleClick Mouse.");
                    }
                };
            }
        }

        private void CreateShortcutContextMenu()
        {
            RestoreShortcut();
            _contextMenuShortcuts = new ContextMenuStrip();
            _contextMenuShortcuts.Opening += (sender, args) =>
            {
                _contextMenuShortcuts.Items.Clear();
                _contextMenuShortcuts.Items.Add("Explorer", null, (o, eventArgs) => _explorerCleaner.OpenExplorer(""));
                _contextMenuShortcuts.Items.Add(new ToolStripSeparator());

                var closedExplorers = _explorerCleaner.ClosedExplorers.OrderByDescending(x => x.IsFavorited)
                    .ThenByDescending(x => x.LastUpdateDateTime).Take(10);

                foreach (var closedExplorer in closedExplorers)
                {
                    var item = new ToolStripMenuItem(closedExplorer.LocationPath);
                    item.Image = closedExplorer.IsFavorited ? Resources.favorite : null;
                    item.Click += (o, eventArgs) => _explorerCleaner.OpenExplorer(closedExplorer);
                    _contextMenuShortcuts.Items.Add(item);
                }

                if (_shortcuts.Length > 0)
                {
                    _contextMenuShortcuts.Items.Add(new ToolStripSeparator());
                    foreach (var shortcutItem in _shortcuts)
                    {
                        _contextMenuShortcuts.Items.Add(shortcutItem.Name, null,
                            (o, eventArgs) => shortcutItem.Exec());
                    }
                }
                args.Cancel = false;
            };
            _contextMenuShortcuts.Closing += (sender, args) =>
            {
                // 任意のクリックで閉じるが、カーソルがコンテキストメニューにあるときは閉じない（イベントが消えるので）
                if (args.CloseReason == ToolStripDropDownCloseReason.AppFocusChange &&
                    _contextMenuShortcuts.ClientRectangle.Contains(
                        _contextMenuShortcuts.PointToClient(Cursor.Position))) args.Cancel = true;
            };
        }

        private IntPtr _pendingHwnd = IntPtr.Zero;
        private string _pendingText = null;
        private bool _pendingFromClick;
        private IntPtr _windowHwnd;
        private MainWindow _window;
        private bool _isClipboardMenuOpen;
        private Rectangle _clipboardMenuScreenRect;

        private void CreateClipboardContextMenu()
        {
            _clipboardMenu = new System.Windows.Controls.ContextMenu();
            _clipboardMenu.PlacementTarget = this._window;

            _clipboardMenu.Opened += (s, e) =>
            {
                _isClipboardMenuOpen = true;
            };

            _clipboardMenu.Closed += (s, e) =>
            {
                if (!_pendingFromClick) return;
                _pendingFromClick = false;

                SetClipboardTextWithRetry(_pendingText);
                Application.Current.Dispatcher.BeginInvoke(new Action(() =>
                {
                    SendCtrlV(_pendingHwnd);
                }), System.Windows.Threading.DispatcherPriority.ApplicationIdle);

                // 閉じた後の次ターンで実行（フォーカス戻り待ち）
                //Application.Current.Dispatcher.BeginInvoke(new Action(() =>
                //{
                //    SetClipboardTextWithRetry(text); // ← WPFなら System.Windows.Clipboard を推奨（後述）
                //    SendCtrlV(hwnd);
                //}), System.Windows.Threading.DispatcherPriority.Background);
                /*
                if (!_pendingFromClick || _pendingText is null) return;
                _pendingFromClick = false;

                var text = _pendingText;
                var hwnd = _pendingHwnd;
                _pendingText = null;
                _pendingHwnd = IntPtr.Zero;

                //_clipboardMenu.BeginInvoke(new Action(() =>
                //{
                    // クリップボードは競合するのでリトライ推奨
                    SetClipboardTextWithRetry(text);
                    SendCtrlV(hwnd);
                //}));
                */
            };

            /*
            _contextMenuClipboardHistories.Opening += (sender, args) =>
            {
                _contextMenuClipboardHistories.Items.Clear();
                var clipboardHistoriesItem = new ToolStripMenuItem("Clipboard Histories");
                clipboardHistoriesItem.Enabled = false;
                _contextMenuClipboardHistories.Items.Add(clipboardHistoriesItem);
                _contextMenuClipboardHistories.Items.Add(new ToolStripSeparator());

                foreach (var c in _clipboardItemQueue.Reverse())
                {
                    var text = c.GetText();
                    if (text.Length < 20) continue;
                    var hwnd = this._windowHwnd;

                    var item = new ToolStripMenuItem(c.ToString(), null, (o, eventArgs) =>
                    {

                        _pendingText = text;
                        _pendingHwnd = hwnd;
                        _pendingFromClick = true;
                        _contextMenuClipboardHistories.Close(ToolStripDropDownCloseReason.ItemClicked);
                        
                        Clipboard.SetText(text);
                        _contextMenuClipboardHistories.Close(ToolStripDropDownCloseReason.ItemClicked);

                        // メニューが閉じてフォーカスが戻ってから貼り付け
                        _contextMenuClipboardHistories.BeginInvoke(new Action(() =>
                        {
                            SendCtrlV(hwnd);
                        }));
                        
                    });
                    item.ToolTipText = c.GetText();
                    _contextMenuClipboardHistories.Items.Add(item);
                    _contextMenuClipboardHistories.Items.Add(new ToolStripSeparator());
                }
                _contextMenuClipboardHistories.Items.RemoveAt(_contextMenuClipboardHistories.Items.Count - 1);
            };
*/

            /*
            _contextMenuClipboardHistories.Closing += (sender, args) =>
            {
                // クリックで貼り付け待ちなら、理由に関係なく閉じるのを邪魔しない
                if (_pendingFromClick)
                    return;

                if (args.CloseReason == ToolStripDropDownCloseReason.AppFocusChange &&
                    _contextMenuClipboardHistories.ClientRectangle.Contains(
                        _contextMenuClipboardHistories.PointToClient(Cursor.Position)))
                    args.Cancel = true;
            };*/
        }

        private static void SetClipboardTextWithRetry(string text, int retry = 12)
        {
            if (text == null) text = "";

            int delayMs = 5;

            for (int i = 0; i < retry; i++)
            {
                try
                {
                    System.Windows.Clipboard.SetText(text);
                    return;
                }
                catch (System.Runtime.InteropServices.COMException ex) when ((uint)ex.HResult == 0x800401D0) // CLIPBRD_E_CANT_OPEN
                {
                    System.Threading.Thread.Sleep(delayMs);
                    delayMs = Math.Min(delayMs * 2, 80); // 5,10,20,40,80...
                }
                catch (System.Runtime.InteropServices.ExternalException)
                {
                    System.Threading.Thread.Sleep(delayMs);
                    delayMs = Math.Min(delayMs * 2, 80);
                }
            }

            // ここで例外投げて落とすのはやめる（必要ならログ出して諦める）
            // throw; ←しない
        }

        [DllImport("user32.dll", CharSet = CharSet.Auto)]
        static public extern IntPtr GetForegroundWindow();

        // import the function in your class
        [DllImport("User32.dll")]
        static extern int SetForegroundWindow(IntPtr point);

        private void SendCtrlV(IntPtr hWnd)
        {
            SetForegroundWindow(hWnd);

            SendKeys.SendWait("^{v}");

        }



        private readonly object _queueLock = new object();
        private void MonitoringClipboard()
        {
            _clipboardMonitor = new ClipboardMonitor();
            _clipboardMonitor.ClipboardChanged += (sender, args) =>
            {
                var item = new ClipboardHistoryItem(args.DataObject);
                var text = item.GetText();
                if (string.IsNullOrEmpty(text)) return;

                lock (_queueLock)
                {
                    // 同一テキストを除外してキューを作り直す
                    if (_clipboardItemQueue.Any(x => x.GetText() == text))
                    {
                        _clipboardItemQueue = new Queue<ClipboardHistoryItem>(
                            _clipboardItemQueue.Where(x => x.GetText() != text)
                        );
                    }

                    _clipboardItemQueue.Enqueue(item);
                    if (_clipboardItemQueue.Count == 501) _clipboardItemQueue.Dequeue();
                }

                SaveClipboardHistories();
            };
        }

        private void RestoreClipboardHistories()
        {
            _clipboardItemQueue = new Queue<ClipboardHistoryItem>();
            if (!File.Exists(ClipboardFileName)) return;
            var clipboardHistories = JsonConvert.DeserializeObject<string[]>(File.ReadAllText(ClipboardFileName));
            if (clipboardHistories == null) return;
            _clipboardItemQueue =
                new Queue<ClipboardHistoryItem>(clipboardHistories.Select(x => new ClipboardHistoryItem(x)));
        }

        private void RestoreShortcut()
        {
            _shortcuts = new ShortcutItem[] {};
            if (!File.Exists(ShortcutFileName)) return;
            var shortcuts = JsonConvert.DeserializeObject<dynamic[]>(File.ReadAllText(ShortcutFileName));
            if (shortcuts == null) return;
            _shortcuts = shortcuts.Select(x=> new ShortcutItem(x.name.ToString(), x.value.ToString())).ToArray();
        }

        protected override void FinalizeCore()
        {
            SaveClipboardHistories();
            _explorerCleaner.SaveExit();
        }

        private void SaveClipboardHistories()
        {
            var clipboardSerialized = JsonConvert.SerializeObject(_clipboardItemQueue.Select(x => x.GetText()),
                Formatting.Indented);
            File.WriteAllText(ClipboardFileName, clipboardSerialized);
        }

        public void SwitchPin(Explorer explorer)
        {
            _actionExecuter.Post(new SwitchPinAction(explorer));
        }

        public void OpenExplorer(Explorer explorer)
        {
            _explorerCleaner.OpenExplorer(explorer);
        }

        public void AddFavorite(Explorer explorer)
        {
            _actionExecuter.Post(new AddFavoriteAction(explorer));
        }

        public void OpenExplorer(string shellcommand)
        {
            _explorerCleaner.OpenExplorer(shellcommand);
        }

        public void CloseExplorer(Explorer explorer)
        {
            _actionExecuter.Post(new CloseExplorerAction(explorer));
        }

        public void RemoveExplorer(Explorer explorer)
        {
            _actionExecuter.Post(new RemoveExplorerAction(explorer));
        }

        public void OpenFavoritedExplorer()
        {
            _explorerCleaner.OpenFavoritedExplorer();
        }

        public bool SwitchShowApplication()
        {
            _explorerCleaner.IsShowApplication = !_explorerCleaner.IsShowApplication;
            Clean();
            return _explorerCleaner.IsShowApplication;
        }

        public void Clean()
        {
            _actionExecuter.Post(new CleanAction());
        }

        public override void Dispose()
        {
            base.Dispose();
            _clipboardMonitor.Dispose();
            _globalMouseHook.Dispose();
        }

        #region WindowClosedイベント

        // Event object
        public event EventHandler<CleanedEventArgs> Cleaned;

        protected virtual void OnCleaned(CleanedEventArgs args)
        {
            var handler = Cleaned;
            if (handler != null)
            {
                handler(this, args);
            }
        }

        // ExplorerWindowCleanerClient クラスにウィンドウハンドルを設定するメソッドを追加
        public void SetWindowHwnd(IntPtr hwnd)
        {
            // 必要に応じてフィールドを追加し、ウィンドウハンドルを保持
            this._windowHwnd = hwnd;
        }

        internal void SetWindow(MainWindow mainWindow)
        {
            this._window = mainWindow;
        }

        #endregion
    }
}