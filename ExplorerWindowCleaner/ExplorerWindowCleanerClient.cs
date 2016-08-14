using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Windows.Forms;
using ExplorerWindowCleaner.Actions;
using ExplorerWindowCleaner.Properties;
using FinalstreamCommons.Windows;
using Firk.Core;
using Firk.Core.Actions;
using MahApps.Metro;
using Newtonsoft.Json;
using Application = System.Windows.Application;

namespace ExplorerWindowCleaner
{
    public class ExplorerWindowCleanerClient : AppClient<ExplorerWindowCleanerAppConfig>
    {
        public const string ClipboardFileName = "clipboard.json";
        private ActionExecuter<ExplorerWindowCleanerClientOperator> _actionExecuter;
        private Queue<ClipboardHistoryItem> _clipboardItemQueue;
        private ClipboardMonitor _clipboardMonitor;
        private ContextMenuStrip _contextMenuClipboardHistories;
        private ContextMenuStrip _contextMenuShortcuts;
        private ExplorerCleaner _explorerCleaner;
        private GlobalMouseHook _globalMouseHook;

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
                _clipboardItemQueue = new Queue<ClipboardHistoryItem>();
                RestoreClipboardHistories();

                MonitoringClipboard();

                CreateShortcutContextMenu();
                CreateClipboardContextMenu();

                _globalMouseHook = new GlobalMouseHook();
                _globalMouseHook.MouseHooked += (sender, args) =>
                {
                    if (args.MouseButton == MouseButtons.Left)
                    {
                        _contextMenuClipboardHistories.Close(ToolStripDropDownCloseReason.AppFocusChange);
                        _contextMenuShortcuts.Close(ToolStripDropDownCloseReason.AppFocusChange);
                    }

                    if (args.IsDoubleClick)
                    {
                        if (args.MouseButton == MouseButtons.Right)
                        {
                            _contextMenuClipboardHistories.Show(args.Point);
                        }
                        else if (args.MouseButton == MouseButtons.Left && args.IsDesktop)
                        {
                            _contextMenuShortcuts.Show(args.Point);
                        }
                        Debug.WriteLine("DoubleClick Mouse.");
                    }
                };
            }
        }

        private void CreateShortcutContextMenu()
        {
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
                    item.Click += (o, eventArgs) => _explorerCleaner.OpenExplorer(closedExplorer.LocationPath);
                    _contextMenuShortcuts.Items.Add(item);
                }
            };
            _contextMenuShortcuts.Closing += (sender, args) =>
            {
                // 任意のクリックで閉じるが、カーソルがコンテキストメニューにあるときは閉じない（イベントが消えるので）
                if (args.CloseReason == ToolStripDropDownCloseReason.AppFocusChange &&
                    _contextMenuShortcuts.ClientRectangle.Contains(
                        _contextMenuShortcuts.PointToClient(Cursor.Position))) args.Cancel = true;
            };
        }

        private void CreateClipboardContextMenu()
        {
            _contextMenuClipboardHistories = new ContextMenuStrip();
            _contextMenuClipboardHistories.Opening += (sender, args) =>
            {
                _contextMenuClipboardHistories.Items.Clear();
                var clipboardHistoriesItem = new ToolStripMenuItem("Clipboard Histories");
                clipboardHistoriesItem.Enabled = false;
                _contextMenuClipboardHistories.Items.Add(clipboardHistoriesItem);
                _contextMenuClipboardHistories.Items.Add(new ToolStripSeparator());

                foreach (var c in _clipboardItemQueue.Reverse())
                {
                    var item = new ToolStripMenuItem(c.ToString(), null, (o, eventArgs) =>
                    {
                        Clipboard.SetText(c.GetText());
                        _contextMenuClipboardHistories.Close(ToolStripDropDownCloseReason.ItemClicked);
                    });
                    item.ToolTipText = c.GetText();
                    _contextMenuClipboardHistories.Items.Add(item);
                    _contextMenuClipboardHistories.Items.Add(new ToolStripSeparator());
                }
                _contextMenuClipboardHistories.Items.RemoveAt(_contextMenuClipboardHistories.Items.Count - 1);
            };

            _contextMenuClipboardHistories.Closing += (sender, args) =>
            {
                // 任意のクリックで閉じるが、カーソルがコンテキストメニューにあるときは閉じない（イベントが消えるので）
                if (args.CloseReason == ToolStripDropDownCloseReason.AppFocusChange &&
                    _contextMenuClipboardHistories.ClientRectangle.Contains(
                        _contextMenuClipboardHistories.PointToClient(Cursor.Position))) args.Cancel = true;
            };
        }

        private void MonitoringClipboard()
        {
            _clipboardMonitor = new ClipboardMonitor();
            _clipboardMonitor.ClipboardChanged += (sender, args) =>
            {
                var clipboardItem = new ClipboardHistoryItem(args.DataObject);
                if (!string.IsNullOrEmpty(clipboardItem.GetText())
                    && !_clipboardItemQueue.Select(x => x.GetText()).Contains(clipboardItem.GetText()))
                    _clipboardItemQueue.Enqueue(clipboardItem);
                if (_clipboardItemQueue.Count == 11) _clipboardItemQueue.Dequeue();
            };
        }

        private void RestoreClipboardHistories()
        {
            if (!File.Exists(ClipboardFileName)) return;
            var clipboardHistories = JsonConvert.DeserializeObject<string[]>(File.ReadAllText(ClipboardFileName));
            if (clipboardHistories == null) return;
            _clipboardItemQueue =
                new Queue<ClipboardHistoryItem>(clipboardHistories.Select(x => new ClipboardHistoryItem(x)));
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

        #endregion
    }
}