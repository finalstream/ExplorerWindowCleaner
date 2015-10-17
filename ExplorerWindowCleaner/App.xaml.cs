using System.Windows;
using ExplorerWindowCleaner.Properties;
using MahApps.Metro;

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

            // get the theme from the current application
            var theme = ThemeManager.DetectAppStyle(Application.Current);

            // now set the Green accent and dark theme
            ThemeManager.ChangeAppStyle(Application.Current,
                                        ThemeManager.GetAccent(Settings.Default.AccentColor),
                                        ThemeManager.GetAppTheme(Settings.Default.AppTheme));

            ShutdownMode = ShutdownMode.OnExplicitShutdown;
            _explorerCleaner = new ExplorerCleaner(
                Settings.Default.Interval,
                Settings.Default.IsAutoCloseUnused,
                Settings.Default.ExpireInterval,
                Settings.Default.ExportLimitNum);
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