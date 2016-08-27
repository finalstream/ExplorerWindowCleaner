using System;
using System.Collections.Concurrent;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Runtime.InteropServices;
using System.Windows.Data;
using FinalstreamCommons.Extensions;
using Newtonsoft.Json;
using NLog;
using SHDocVw;

namespace ExplorerWindowCleaner
{
    public class ExplorerCleaner
    {
        private const string NowFileName = "now.json";
        private const string HistoryFileName = "history.json";
        private readonly ExplorerWindowCleanerAppConfig _appConfig;
        private readonly Dictionary<int, Explorer> _explorerDic;
        private readonly Logger _log = LogManager.GetCurrentClassLogger();
        private readonly HashSet<string> _pinedRestoreHashSet;
        private readonly ShellWindows _shellWindows;
        private readonly SpecialFolderManager _specialFolderManager;
        private ConcurrentDictionary<string, Explorer> _closedExplorerDic;
        private DateTime _expireDateTime;
        private string _lastSerializeHistory;
        private string _lastSerializeNow;
        private int _maxWindowCount;

        public enum ExplorerCloseReason
        {
            /// <summary>
            /// すでに終了済み
            /// </summary>
            Terminated,
            /// <summary>
            /// 重複
            /// </summary>
            Duplication,
            /// <summary>
            /// 期限切れ
            /// </summary>
            Expiration,
            /// <summary>
            /// ユーザ操作
            /// </summary>
            UserOperation
        }

        /// <summary>
        ///     復元用エクスプローラディクショナリ（前回のNow）
        /// </summary>
        private Dictionary<int, Explorer> _restoreExplorerDic;

        private int _totalCloseWindowCount;

        internal ExplorerCleaner(ExplorerWindowCleanerAppConfig appConfig)
        {
            _appConfig = appConfig;
            _explorerDic = new Dictionary<int, Explorer>();
            _closedExplorerDic = new ConcurrentDictionary<string, Explorer>();
            _restoreExplorerDic = new Dictionary<int, Explorer>();
            _pinedRestoreHashSet = new HashSet<string>();
            Explorers = new ObservableCollection<Explorer>();
            ClosedExplorers = new ObservableCollection<Explorer>();
            BindingOperations.EnableCollectionSynchronization(Explorers, new object());
            BindingOperations.EnableCollectionSynchronization(ClosedExplorers, new object());
            _shellWindows = new ShellWindowsClass();

            _specialFolderManager = new SpecialFolderManager();
            Restore();
        }

        public bool IsShowApplication { get; set; }

        public int WindowCount
        {
            get { return _explorerDic.Values.Count(x => x.IsExplorer); }
        }

        public int PinedCount
        {
            get { return _explorerDic.Values.Count(x => x.IsPined); }
        }

        public ObservableCollection<Explorer> Explorers { get; }
        public ObservableCollection<Explorer> ClosedExplorers { get; }

        [DllImport("user32.dll")]
        private static extern IntPtr GetForegroundWindow();

        private void Restore()
        {
            LoadNow();
            LoadHistory();
        }

        private void LoadNow()
        {
            if (!File.Exists(NowFileName)) return;
            var oldNows = JsonConvert.DeserializeObject<Explorer[]>(File.ReadAllText(NowFileName));
            if (oldNows == null) return;
            // 復元用ディクショナリ
            _restoreExplorerDic = oldNows.GroupBy(x => x.Handle)
                .ToDictionary(x => x.Key, x => x.First());
        }

        private void LoadHistory()
        {
            if (!File.Exists(HistoryFileName)) return;
            var histories = JsonConvert.DeserializeObject<Explorer[]>(File.ReadAllText(HistoryFileName));
            if (histories == null) return;
            _closedExplorerDic =
                new ConcurrentDictionary<string, Explorer>(
                    histories.Where(x => x.CloseCount > 0).ToDictionary(x => x.LocationKey, x => x));
        }

        /// <summary>
        ///     エクスプローラのウインドウをクリーンします。
        /// </summary>
        /// <returns>クローズしたウインドウの名前リスト</returns>
        public void Clean()
        {
            var isUpdated = false;
            var cleanedInfos = new List<ExplorerCleanedInfo>();

            var closedExplorerHandleList = _explorerDic.Keys.ToList();

            _log.Debug("ShellWindows[{0}]", _shellWindows.Count);

            var userApps = IsShowApplication
                ? Process.GetProcesses()
                    .Where(x => x.MainWindowHandle != IntPtr.Zero)
                    .Select(x => new UserApplication(x))
                    .Where(x => !x.IsExplorer && !x.IsUnknown)
                : Enumerable.Empty<UserApplication>();

            var ies = _shellWindows.Cast<InternetExplorer>().Concat(userApps);

            foreach (var ie in ies)
            {
                if (ie.FullName == null) continue;
                var handle = 0;
                try
                {
                    handle = ie.HWND;
                }
                catch (COMException comex)
                {
                    // IEのコンポーネントを使用しているアプリを拾った場合、COMExceptionが出る場合があるので出た場合はそのプロセスはスキップする。
                    _log.Debug(comex, "COM Error!! FullName:{0}", ie.FullName);
                }
                if (handle == 0) continue;
                if (!_explorerDic.Keys.Contains(handle))
                {
                    var explorer = new Explorer(ie);
                    if (_appConfig.IsKeepPin)
                    {
                        explorer.PinLocationChanged += (sender, pinexp) =>
                        {
                            // ピン留めでパスが変わったらピン留めのパスを開く（キープ）
                            OpenExplorer(pinexp, true);
                            // ピンキープのためにキーを保存
                            _pinedRestoreHashSet.Add(pinexp.LocationKey);
                        };

                        // ピン復元
                        if (_pinedRestoreHashSet.Contains(explorer.LocationKey))
                        {
                            explorer.IsPined = true;
                            _pinedRestoreHashSet.Remove(explorer.LocationKey);
                        }
                    }

                    RegistExplorer(explorer);
                }
                else
                {
                    closedExplorerHandleList.Remove(handle);

                    // よくわからないけどここでKeyNotFoundが出る場合があるのでTryGetする。
                    Explorer explorer;
                    if (_explorerDic.TryGetValue(handle, out explorer))
                    {
                        bool updated;
                        if (_appConfig.IsKeepPin)
                        {
                            updated = explorer.UpdateWithKeepPin(ie);
                        }
                        else
                        {
                            updated = explorer.Update(ie);
                        }
                        if (updated) isUpdated = true;
                    }
                }
            }

            // すでに終了されたものがあればディクショナリから削除
            foreach (var closedHandle in closedExplorerHandleList)
            {
                Explorer explorer;
                if (_explorerDic.TryGetValue(closedHandle, out explorer))
                {
                    CloseExplorer(explorer, ExplorerCloseReason.Terminated);
                    _explorerDic.Remove(closedHandle);
                }
            }

            var duplicateExplorers = _explorerDic.Where(x => x.Value.IsExplorer)
                .GroupBy(g => g.Value.LocationKey)
                .Select(g => new {Explorer = g, Count = g.Count()})
                .Where(x => x.Count > 1).ToArray(); // Count > 1はいらないけど、旧バージョンからの互換のためにしばらく残す

            // 同じパスのがあれば一番新しいもの以外は終了させる
            foreach (var duplicateExplorer in duplicateExplorers)
            {
                var closeTargets = duplicateExplorer.Explorer
                    .Where(
                        x =>
                            x.Value.LastUpdateDateTime !=
                            duplicateExplorer.Explorer.Max(m => m.Value.LastUpdateDateTime))
                    .Select(x => x.Value)
                    .ToArray();
                foreach (var closeTarget in closeTargets)
                {
                    var closeReason = ExplorerCloseReason.Duplication;
                    if (CloseExplorer(closeTarget, closeReason)) cleanedInfos.Add(new ExplorerCleanedInfo(closeTarget.LocationName, closeReason));
                }
            }

            // 期限切れのものがあれば閉じる
            if (_appConfig.IsAutoCloseUnused && _appConfig.ExpireInterval != TimeSpan.Zero)
            {
                _expireDateTime = DateTime.Now.Subtract(_appConfig.ExpireInterval);
                _log.Debug("Expire Datetime {0}", _expireDateTime);

                var explorers = _explorerDic.Values.Where(x => x.IsExplorer).ToArray();
                foreach (var expireExplorer in explorers.Where(x => x.LastUpdateDateTime < _expireDateTime))
                {
                    _log.Debug("Expire Explorer {0}", expireExplorer);
                    var closeReason = ExplorerCloseReason.Expiration;
                    if (CloseExplorer(expireExplorer, closeReason)) cleanedInfos.Add(new ExplorerCleanedInfo(expireExplorer.LocationName, closeReason));
                }
            }

            // 表示用データ更新
            var updatedView = UpdateView();
            if (updatedView) isUpdated = true;

            Save();

            if (WindowCount > _maxWindowCount) _maxWindowCount = WindowCount;
            _totalCloseWindowCount += cleanedInfos.Count;

            OnCleaned(cleanedInfos, isUpdated);
        }

        private void RegistExplorer(Explorer explorer)
        {
            if (_restoreExplorerDic.ContainsKey(explorer.Handle))
                explorer.Restore(_restoreExplorerDic[explorer.Handle]);
            _log.Debug("Regist Explorer : {0}", explorer);
            _explorerDic.Add(explorer.Handle, explorer);
        }

        private void Save()
        {
            SaveNow();
            SaveHistory();
        }

        private void SaveNow()
        {
            var nowSerialized = JsonConvert.SerializeObject(_explorerDic.Values.ToArray(), Formatting.Indented);
            if (nowSerialized == _lastSerializeNow) return; // 同じであれば何もしない
            File.WriteAllText(NowFileName, nowSerialized);
            _lastSerializeNow = nowSerialized;
        }

        private void SaveHistory()
        {
            var histories =
                _closedExplorerDic.Values.ToArray()
                    .Where(x => x.IsFavorited || x.IsExplorer) // お気に入りでないアプリは保存しない
                    .OrderByDescending(x => x.IsFavorited)
                    .ThenByDescending(x => x.LastUpdateDateTime).Take(_appConfig.ExportLimitNum);
            var historySerialized = JsonConvert.SerializeObject(histories, Formatting.Indented);
            if (historySerialized == _lastSerializeHistory) return; // 同じであれば何もしない
            File.WriteAllText(HistoryFileName, historySerialized);
            _lastSerializeHistory = historySerialized;
        }

        public void SaveExit()
        {
            // 終了時にNowWindowsをクローズドに登録して保存
            var nows = _explorerDic.Values.ToArray();
            foreach (var explorer in nows)
            {
                AddOrUpdateClosedDictionary(explorer);
            }
            SaveHistory();
        }

        public bool UpdateView()
        {
            // アプリ表示時以外はアプリは表示しないようにする。
            var nowUpdated = Explorers.DiffUpdate(
                _explorerDic.Values.Where(x => IsShowApplication || x.IsExplorer).ToArray(),
                new ExplorerEqualityComparer());
            var closeUpdated = ClosedExplorers.DiffUpdate(
                _closedExplorerDic.Values.Where(x => IsShowApplication || x.IsExplorer).ToArray(),
                new ExplorerEqualityComparer());
            return nowUpdated || closeUpdated;
        }

        public bool CloseExplorer(Explorer explorer, ExplorerCloseReason closeReason)
        {
            if (explorer.IsPined) return false; // ピン留めの場合閉じない。
            if (GetForegroundWindow() == (IntPtr) explorer.Handle) return false; // アクティブな場合閉じない。
            if (closeReason == ExplorerCloseReason.Expiration && _closedExplorerDic.Values.Where(x => x.IsFavorited).Select(x => x.LocationUrl).Contains(explorer.LocationUrl)) return false; // 期限切れで閉じるときはお気に入りのパスは閉じない。
            var handle = explorer.Exit();
            _explorerDic.Remove(handle);
            Explorers.Remove(explorer);
            if (explorer.IsExplorer) AddOrUpdateClosedDictionary(explorer); // エクスプローラだけ追加する

            return true;
        }

        public void AddOrUpdateClosedDictionary(Explorer explorer)
        {
            if (string.IsNullOrEmpty(explorer.LocationName) && string.IsNullOrEmpty(explorer.LocationUrl)) return; // 不明なemptyはすてる
            _closedExplorerDic.AddOrUpdate(explorer.LocationKey,
                s =>
                {
                    explorer.UpdateRegistDateTime();
                    return explorer;
                },
                (s, oldExplorer) =>
                {
                    oldExplorer.UpdateClosedInfo(explorer);
                    return oldExplorer;
                });
        }

        public void RemoveClosedDictionary(Explorer explorer)
        {
            Explorer removeExplorer;
            _closedExplorerDic.TryRemove(explorer.LocationKey, out removeExplorer);
            ClosedExplorers.Remove(removeExplorer);
        }

        public void OpenExplorer(Explorer explorer, bool isMinimized = false)
        {
            var path = !explorer.IsSpecialFolder
                ? explorer.LocationPath
                : _specialFolderManager.ConvertSpecialFolder(explorer.LocationName);
            OpenExplorer(path, isMinimized);
        }

        public void OpenExplorer(ShortcutItem shortcut, bool isMinimized = false)
        {
            var path = !shortcut.IsSpecialFolder
                ? shortcut.Value
                : _specialFolderManager.ConvertSpecialFolder(shortcut.Value);
            OpenExplorer(path, isMinimized);
        }

        public void OpenExplorer(string path, bool isMinimized = false)
        {
            var psi = new ProcessStartInfo("EXPLORER.EXE", string.Format("/n,\"{0}\"", path));
            if (isMinimized) psi.WindowStyle = ProcessWindowStyle.Minimized;
            _log.Debug("{0} {1} {2}", (object) psi.FileName, (object) psi.Arguments, (object) psi.WindowStyle);
            Process.Start(psi);
        }

        public void OpenFavoritedExplorer()
        {
            var noaliveFavoriteExplorers = _closedExplorerDic.Values.Where(
                x => x.IsFavorited && _explorerDic.Values.All(y => y.LocationKey != x.LocationKey));
            foreach (var noaliveFavoriteExplorer in noaliveFavoriteExplorers)
            {
                OpenExplorer(noaliveFavoriteExplorer);
            }
        }

        #region Cleanedイベント

        // Event object
        public event EventHandler<CleanedEventArgs> Cleaned;

        protected virtual void OnCleaned(ICollection<ExplorerCleanedInfo> closeInfos, bool isUpdated)
        {
            var handler = Cleaned;
            if (handler != null)
            {
                handler(this, new CleanedEventArgs(
                    closeInfos, WindowCount, _expireDateTime, _maxWindowCount, PinedCount, _totalCloseWindowCount,
                    isUpdated));
            }
        }

        #endregion
    }

    public class ExplorerEqualityComparer : IEqualityComparer<Explorer>
    {
        public bool Equals(Explorer x, Explorer y)
        {
            return x.LocationKey == y.LocationKey;
        }

        public int GetHashCode(Explorer obj)
        {
            return obj.GetHashCode();
        }
    }
}