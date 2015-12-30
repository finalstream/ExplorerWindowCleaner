using System.Reflection;
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
        private ExplorerWindowCleanerClient _client;

        /// <summary>
        ///     System.Windows.Application.Startup イベント を発生させます。
        /// </summary>
        /// <param name="e">イベントデータ を格納している StartupEventArgs</param>
        protected override void OnStartup(StartupEventArgs e)
        {
            base.OnStartup(e);
            ShutdownMode = ShutdownMode.OnExplicitShutdown;

            _client = new ExplorerWindowCleanerClient(Assembly.GetExecutingAssembly());
            _client.Initialize();
            
        }

        /// <summary>
        ///     System.Windows.Application.Exit イベント を発生させます。
        /// </summary>
        /// <param name="e">イベントデータ を格納している ExitEventArgs</param>
        protected override void OnExit(ExitEventArgs e)
        {
            base.OnExit(e);
            _client.Finish();
            _client.Dispose();
        }
    }
}