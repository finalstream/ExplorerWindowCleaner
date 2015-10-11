using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.IO;
using System.Linq;
using System.Threading;
using System.Threading.Tasks;
using System.Windows.Data;
using ExplorerWindowCleaner.Properties;
using SHDocVw;

namespace ExplorerWindowCleaner
{
    public class ExplorerCleaner
    {
        private static int _seqNo = 0;
        private readonly Dictionary<int, Explorer> _explorerDic;
        private readonly Dictionary<string, Explorer> _closedExplorerDic; 
        private readonly TimeSpan _interval;
        private readonly TimeSpan _expireInterval;
        private CancellationTokenSource _cancellationTokenSource;
        public bool IsAutoCloseUnused { get; set; }
        public int WindowCount { get { return _explorerDic.Count; } }
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


        public ExplorerCleaner(TimeSpan interval, bool isAutoCloseUnused, TimeSpan expireInterval)
        {
            _interval = interval;
            IsAutoCloseUnused = isAutoCloseUnused;
            _expireInterval = expireInterval;
            _explorerDic = new Dictionary<int, Explorer>();
            _closedExplorerDic = new Dictionary<string, Explorer>();
            Explorers = new ObservableCollection<Explorer>();
            ClosedExplorers = new ObservableCollection<Explorer>();
            BindingOperations.EnableCollectionSynchronization(Explorers, new object());
            BindingOperations.EnableCollectionSynchronization(ClosedExplorers, new object());
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

            if (WindowCount > MaxWindowCount) MaxWindowCount = WindowCount;
            TotalCloseWindowCount += closeWindowCount;

            return closeWindowCount;
        }

        private void UpdateView()
        {
            Explorers.Clear();
            foreach (var aliveExplorer in _explorerDic.Values.OrderByDescending(x => x.LastUpdateDateTime))
            {
                Explorers.Add(aliveExplorer);
            }

            ClosedExplorers.Clear();
            foreach (var closedExplorer in _closedExplorerDic.Values.OrderByDescending(x => x.IsPined).ThenByDescending(x=>x.CloseCount).ThenByDescending(x=>x.LastUpdateDateTime))
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
                _closedExplorerDic[explorer.LocationKey].UpdateCloseCount();
            }
            else
            {
                // add
                _closedExplorerDic.Add(explorer.LocationKey, explorer);
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