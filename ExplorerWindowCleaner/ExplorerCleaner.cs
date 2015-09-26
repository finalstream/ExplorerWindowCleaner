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
        private readonly TimeSpan _interval;
        private readonly TimeSpan _expireInterval;
        private CancellationTokenSource _cancellationTokenSource;
        public bool IsAutoCloseUnused { get; set; }

        public ExplorerCleaner(TimeSpan interval, bool isAutoCloseUnused, TimeSpan expireInterval)
        {
            _interval = interval;
            IsAutoCloseUnused = isAutoCloseUnused;
            _expireInterval = expireInterval;
            _explorerDic = new Dictionary<int, Explorer>();
            Explorers = new ObservableCollection<Explorer>();
            BindingOperations.EnableCollectionSynchronization(Explorers, new object());
        }

        public ObservableCollection<Explorer> Explorers { get; private set; }

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
                    Clean();

                    Task.Delay(_interval).Wait();
                }
            },
                _cancellationTokenSource.Token,
                TaskCreationOptions.LongRunning,
                TaskScheduler.Default);
        }

        private void Clean()
        {
            ShellWindows shellWindows = new ShellWindowsClass();

            var closedExplorerHandleList = _explorerDic.Keys.ToList();

            foreach (InternetExplorer ie in shellWindows)
            {
                var filename = Path.GetFileNameWithoutExtension(ie.FullName).ToLower();

                if (filename.Equals("explorer"))
                {

                    if (string.IsNullOrEmpty(ie.LocationURL)) continue;


                    if (!_explorerDic.Keys.Contains(ie.HWND))
                    {
                        var explorer = new Explorer(Interlocked.Increment(ref _seqNo), ie);
                        Console.WriteLine("add explorer : {0} {1} {2} {3}", explorer.SeqNo, explorer.Location,
                            explorer.LocalPath, explorer.Instance.HWND);
                        _explorerDic.Add(explorer.Handle, explorer);
                    }
                    else
                    {
                        closedExplorerHandleList.Remove(ie.HWND);
                        _explorerDic[ie.HWND].Update(ie.LocationURL);
                    }
                }
            }

            // すでに終了されたものがあればディクショナリから削除
            foreach (var closedHandle in closedExplorerHandleList)
            {
                _explorerDic.Remove(closedHandle);
            }

            var duplicateExplorers = _explorerDic
                .GroupBy(g => g.Value.Location)
                .Select(g => new {Explorer = g, Count = g.Count()})
                .Where(x => x.Count > 1);

            // 同じパスのがあれば一番新しいもの以外は終了させる
            foreach (var duplicateExplorer in duplicateExplorers)
            {
                var closeTargets = duplicateExplorer.Explorer
                    .Where(x => x.Value.SeqNo != duplicateExplorer.Explorer.Max(m => m.Value.SeqNo)).Select(x=>x.Value);
                foreach (var closeTarget in closeTargets)
                {
                    var handle = closeTarget.Exit();
                    _explorerDic.Remove(handle);
                }
            }

            // 期限切れのものがあれば閉じる
            if (IsAutoCloseUnused && _expireInterval != TimeSpan.Zero)
            {
                var expireDateTime = DateTime.Now.Subtract(_expireInterval);
                Console.WriteLine("expire datetime {0}", expireDateTime);
                
                foreach (var expireExplorer in _explorerDic.Values.Where(x => x.LastUpdateDateTime < expireDateTime))
                {
                    Console.WriteLine("expire explorer {0}", expireExplorer.Handle);
                    var handle = expireExplorer.Exit();
                    _explorerDic.Remove(handle);
                }
            }

            // 更新
            Explorers.Clear();
            foreach (var aliveExplorer in _explorerDic.Values.OrderBy(x=>x.SeqNo))
            {
                Explorers.Add(aliveExplorer);
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