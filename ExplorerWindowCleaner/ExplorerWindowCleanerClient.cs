using System;
using System.Collections.Generic;
using System.Linq;
using System.Reflection;
using System.Text;
using System.Threading.Tasks;
using System.Windows;
using ExplorerWindowCleaner.Actions;
using ExplorerWindowCleaner.Properties;
using Firk.Core;
using Firk.Core.Actions;
using MahApps.Metro;

namespace ExplorerWindowCleaner
{
    class ExplorerWindowCleanerClient : AppClient
    {

        private ExplorerCleaner _explorerCleaner;

        /// <summary>
        ///     タスクトレイに表示するアイコン
        /// </summary>
        private NotifyIconContainer _notifyIcon;

        public ExplorerWindowCleanerClient(Assembly executingAssembly) : base(executingAssembly)
        {

        }

        protected override void InitializeCore()
        {
            // get the theme from the current application
            var theme = ThemeManager.DetectAppStyle(Application.Current);

            // now set the Green accent and dark theme
            ThemeManager.ChangeAppStyle(Application.Current,
                                        ThemeManager.GetAccent(Settings.Default.AccentColor),
                                        ThemeManager.GetAppTheme(Settings.Default.AppTheme));

            _explorerCleaner = new ExplorerCleaner(
                Settings.Default.IsAutoCloseUnused,
                Settings.Default.ExpireInterval,
                Settings.Default.ExportLimitNum,
                Settings.Default.IsKeepPin);
            _notifyIcon = new NotifyIconContainer(_explorerCleaner);

            ResetBackgroundWorker(Settings.Default.Interval, new BackgroundAction[] { new CleanerAction(_explorerCleaner) });
        }

        protected override void FinalizeCore()
        {
            _explorerCleaner.SaveExit();
        }


        #region Dispose

        // Flag: Has Dispose already been called?
        private bool disposed = false;

        // Public implementation of Dispose pattern callable by consumers.
        public override void Dispose()
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
                base.Dispose();
                _notifyIcon.Dispose();
            }

            // Free any unmanaged objects here.
            //
            disposed = true;
        }

        #endregion

        
    }
}
