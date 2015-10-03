using System.Windows;
using ExplorerWindowCleaner.Properties;

namespace ExplorerWindowCleaner
{
    /// <summary>
    ///     App.xaml の相互作用ロジック
    /// </summary>
    public partial class App : Application
    {
        private ExplorerCleaner _explorerCleaner;

        /// <summary>
        ///     タスクトレイに表示するアイコン
        /// </summary>
        private NotifyIconContainer notifyIcon;

        /// <summary>
        ///     System.Windows.Application.Startup イベント を発生させます。
        /// </summary>
        /// <param name="e">イベントデータ を格納している StartupEventArgs</param>
        protected override void OnStartup(StartupEventArgs e)
        {
            base.OnStartup(e);

            ShutdownMode = ShutdownMode.OnExplicitShutdown;
            _explorerCleaner = new ExplorerCleaner(
                Settings.Default.Interval,
                Settings.Default.IsAutoCloseUnused,
                Settings.Default.ExpireInterval);
            notifyIcon = new NotifyIconContainer(_explorerCleaner);
        }

        /// <summary>
        ///     System.Windows.Application.Exit イベント を発生させます。
        /// </summary>
        /// <param name="e">イベントデータ を格納している ExitEventArgs</param>
        protected override void OnExit(ExitEventArgs e)
        {
            base.OnExit(e);
            _explorerCleaner.Dispose();
            notifyIcon.Dispose();
        }

        
    }
}