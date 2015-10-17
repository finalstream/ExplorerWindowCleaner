using System;
using System.Collections.Concurrent;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Threading;
using System.Threading.Tasks;
using System.Windows.Data;
using System.Windows.Forms.VisualStyles;
using ExplorerWindowCleaner.Properties;
using Newtonsoft.Json;
using SHDocVw;

namespace ExplorerWindowCleaner
{
    public class ExplorerCleaner
    {
        private const string NowFileName = "now.json";
        private const string HistoryFileName = "history.json";
        private static int _seqNo = 0;
        private readonly Dictionary<int, Explorer> _explorerDic;
        /// <summary>
        /// 復元用エクスプローラディクショナリ（前回のNow）
        /// </summary>
        private Dictionary<int, Explorer> _restoreExplorerDic;
        private Dictionary<string, Explorer> _closedExplorerDic; 
        private readonly TimeSpan _interval;
        private readonly TimeSpan _expireInterval;
        private readonly int _exportLimitNum;
        private CancellationTokenSource _cancellationTokenSource;

        private string _lastSerializeNow;
        private string _lastSerializeHistory;
        
        public bool IsAutoCloseUnused { get; set; }
        public int WindowCount { get { return _explorerDic.Count; } }
        public int PinedCount { get { return _explorerDic.Values.Count(x => x.IsPined); }}
        public int MaxWindowCount { get; private set; }
        public int TotalCloseWindowCount { get; private set; }
        public DateTime ExporeDateTime { get; private set; }

        #region Updatedイベント

        // Event object
        public event EventHandler<UpdatedEventArgs> Updated;

        protected virtual void OnUpdated(int closeWindowCount)
        {
            var handler = this.Updated;
            if (handler != null)
            {
                handler(this, new UpdatedEventArgs(closeWindowCount));
            }
        }

        #endregion


        public ExplorerCleaner(TimeSpan interval, bool isAutoCloseUnused, TimeSpan expireInterval, int exportLimitNum)
        {
            _interval = interval;
            IsAutoCloseUnused = isAutoCloseUnused;
            _expireInterval = expireInterval;
            _exportLimitNum = exportLimitNum;
            _explorerDic = new Dictionary<int, Explorer>();
            _closedExplorerDic = new Dictionary<string, Explorer>();
            _restoreExplorerDic = new Dictionary<int, Explorer>();
            Explorers = new ObservableCollection<Explorer>();
            ClosedExplorers = new ObservableCollection<Explorer>();
            BindingOperations.EnableCollectionSynchronization(Explorers, new object());
            BindingOperations.EnableCollectionSynchronization(ClosedExplorers, new object());

            Restore();
            
        }

        private void Restore()
        {
            LoadNow();
            LoadHistory();

        }

        private void LoadNow()
        {
            if (!File.Exists(NowFileName)) return;
            var oldNows = JsonConvert.DeserializeObject<Explorer[]>(File.ReadAllText(NowFileName));
            // 復元用ディクショナリ
            _restoreExplorerDic = oldNows.GroupBy(x => x.Handle)
                .ToDictionary(x => x.Key, x => x.First());
        }

        private void LoadHistory()
        {
            if (!File.Exists(HistoryFileName)) return;
            var histories = JsonConvert.DeserializeObject<Explorer[]>(File.ReadAllText(HistoryFileName));
            _closedExplorerDic = histories.Where(x => x.CloseCount > 0).ToDictionary(x => x.LocationKey, x => x);
        }

        public ObservableCollection<Explorer> Explorers { get; private set; }

        public ObservableCollection<Explorer> ClosedExplorers { get; private set; }

        public void Start()
        {
            _cancellationTokenSource = new CancellationTokenSource();

            // タスクを開始する。
            Task.Factory.StartNew(() =>
            {
                while (true)
                {
                    if (_cancellationTokenSource.Token.IsCancellationRequested) break;

                    // メイン処理
                    var closeWindowCount = Clean();

                    OnUpdated(closeWindowCount);

                    Task.Delay(_interval).Wait();
                }
            },
                _cancellationTokenSource.Token,
                TaskCreationOptions.LongRunning,
                TaskScheduler.Default);
        }

        private int Clean()
        {
            int closeWindowCount = 0;
            ShellWindows shellWindows = new ShellWindowsClass();

            var closedExplorerHandleList = _explorerDic.Keys.ToList();

            foreach (InternetExplorer ie in shellWindows)
            {
                var filename = Path.GetFileNameWithoutExtension(ie.FullName).ToLower();

                if (filename.Equals("explorer"))
                {

                    if (!_explorerDic.Keys.Contains(ie.HWND))
                    {
                        var explorer = new Explorer(Interlocked.Increment(ref _seqNo), ie);
                        if (_restoreExplorerDic.ContainsKey(explorer.Handle)) explorer.Restore(_restoreExplorerDic[explorer.Handle]);
                        Console.WriteLine("add explorer : {0} {1} {2} {3}", explorer.SeqNo, explorer.LocationKey,
                            explorer.LocationPath, explorer.Instance.HWND);
                        _explorerDic.Add(explorer.Handle, explorer);
                    }
                    else
                    {
                        closedExplorerHandleList.Remove(ie.HWND);
                        _explorerDic[ie.HWND].Update(ie);
                    }
                }
            }

            // すでに終了されたものがあればディクショナリから削除
            foreach (var closedHandle in closedExplorerHandleList)
            {
                CloseExplorer(_explorerDic[closedHandle]);
                _explorerDic.Remove(closedHandle);
            }

            var duplicateExplorers = _explorerDic
                .GroupBy(g => g.Value.LocationKey)
                .Select(g => new {Explorer = g, Count = g.Count()})
                .Where(x => x.Count > 1);

            // 同じパスのがあれば一番新しいもの以外は終了させる
            foreach (var duplicateExplorer in duplicateExplorers)
            {
                var closeTargets = duplicateExplorer.Explorer
                    .Where(x => x.Value.LastUpdateDateTime != duplicateExplorer.Explorer.Max(m => m.Value.LastUpdateDateTime)).Select(x=>x.Value);
                foreach (var closeTarget in closeTargets)
                {
                    if (CloseExplorer(closeTarget)) closeWindowCount++;
                }
            }

            // 期限切れのものがあれば閉じる
            if (IsAutoCloseUnused && _expireInterval != TimeSpan.Zero)
            {
                ExporeDateTime = DateTime.Now.Subtract(_expireInterval);
                Console.WriteLine("expire datetime {0}", ExporeDateTime);

                var explorers = _explorerDic.Values.ToArray();
                foreach (var expireExplorer in explorers.Where(x => x.LastUpdateDateTime < ExporeDateTime))
                {
                    Console.WriteLine("expire explorer {0}", expireExplorer.Handle);
                    if (CloseExplorer(expireExplorer)) closeWindowCount++;
                }
            }

            // 表示更新
            UpdateView();

            Save();

            if (WindowCount > MaxWindowCount) MaxWindowCount = WindowCount;
            TotalCloseWindowCount += closeWindowCount;

            return closeWindowCount;
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
                    .OrderByDescending(x => x.IsFavorited)
                    .ThenByDescending(x => x.LastUpdateDateTime).Take(_exportLimitNum);
            var historySerialized = JsonConvert.SerializeObject(histories, Formatting.Indented);
            if (historySerialized == _lastSerializeHistory) return; // 同じであれば何もしない
            File.WriteAllText(HistoryFileName, historySerialized);
            _lastSerializeHistory = historySerialized;
        }

        private void UpdateView()
        {
            Explorers.Clear();
            foreach (var aliveExplorer in _explorerDic.Values.OrderByDescending(x=>x.IsPined).ThenByDescending(x => x.LastUpdateDateTime))
            {
                Explorers.Add(aliveExplorer);
            }

            ClosedExplorers.Clear();
            foreach (var closedExplorer in _closedExplorerDic.Values.OrderByDescending(x => x.IsFavorited).ThenByDescending(x=>x.LastUpdateDateTime))
            {
                ClosedExplorers.Add(closedExplorer);
            }
        }

        private bool CloseExplorer(Explorer explorer)
        {
            if (explorer.IsPined) return false;
            var handle = explorer.Exit();
            _explorerDic.Remove(handle);
            UpdateClosedDictionary(explorer);

            return true;
        }

        public void UpdateClosedDictionary(Explorer explorer)
        {
            if (_closedExplorerDic.ContainsKey(explorer.LocationKey))
            {
                // update
                _closedExplorerDic[explorer.LocationKey].UpdateClosedInfo(explorer);
            }
            else
            {
                // add
                explorer.UpdateRegistDateTime();
                _closedExplorerDic.Add(explorer.LocationKey, explorer);
            }
        }

        public void OpenExplorer(string locationPath)
        {
            Process.Start("EXPLORER.EXE", string.Format("/n,/root,\"{0}\"", locationPath));
        }

        public void OpenPinedExplorer()
        {
            var closedPinLocationPaths = _closedExplorerDic.Values.Where(
                x => x.IsFavorited && _explorerDic.Values.All(y => y.LocationKey != x.LocationKey)).Select(x=>x.LocationKey);
            foreach (var closedPinLocationPath in closedPinLocationPaths)
            {
                OpenExplorer(closedPinLocationPath);
            }
        }

        #region Dispose

        // Flag: Has Dispose already been called?
        private bool disposed;

        // Public implementation of Dispose pattern callable by consumers.
        public void Dispose()
        {
            Dispose(true);
            GC.SuppressFinalize(this);
        }

        // Protected implementation of Dispose pattern.
        protected virtual void Dispose(bool disposing)
        {
            if (disposed)
                return;

            if (disposing)
            {
                // Free any other managed objects here.
                //
                _cancellationTokenSource.Cancel();
            }

            // Free any unmanaged objects here.
            //
            disposed = true;
        }

        #endregion

        
    }
}