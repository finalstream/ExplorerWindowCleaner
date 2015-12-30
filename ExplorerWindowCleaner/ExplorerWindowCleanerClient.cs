using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
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
    public class ExplorerWindowCleanerClient : AppClient
    {

        private ActionExecuter<ExplorerWindowCleanerClientOperator> _actionExecuter;
        private ExplorerCleaner _explorerCleaner;

        #region WindowClosedイベント

        // Event object
        public event EventHandler<CleanedEventArgs> Cleaned;

        protected virtual void OnCleaned(CleanedEventArgs args)
        {
            var handler = this.Cleaned;
            if (handler != null)
            {
                handler(this, args);
            }
        }

        #endregion

        private ExplorerWindowCleanerAppConfig _appConfig;
        internal new ExplorerWindowCleanerAppConfig AppConfig
        {
            get
            {
                return _appConfig;
            }
            private set
            {
                _appConfig = value;
                base.AppConfig = value;
            }
        }

        public ExplorerWindowCleanerClient(Assembly executingAssembly) : base(executingAssembly)
        {
            
            
        }

        public ObservableCollection<Explorer> Explorers { get { return _explorerCleaner.Explorers; } }

        public ObservableCollection<Explorer> ClosedExplorers { get { return _explorerCleaner.ClosedExplorers; } }
        
        protected override void InitializeCore()
        {
            _appConfig = new ExplorerWindowCleanerAppConfig(ExecutingAssemblyInfo.FileVersion, Settings.Default);

            // get the theme from the current application
            var theme = ThemeManager.DetectAppStyle(Application.Current);

            // now set the Green accent and dark theme
            ThemeManager.ChangeAppStyle(Application.Current,
                                        ThemeManager.GetAccent(AppConfig.AccentColor),
                                        ThemeManager.GetAppTheme(AppConfig.AppTheme));

            _explorerCleaner = new ExplorerCleaner(AppConfig);
            _explorerCleaner.Cleaned += (sender, args) => OnCleaned(args);
            _actionExecuter = new ActionExecuter<ExplorerWindowCleanerClientOperator>(new ExplorerWindowCleanerClientOperator(this, _explorerCleaner));

            ResetBackgroundWorker(Settings.Default.Interval, new BackgroundAction[] { new CleanerAction(this) });
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
                
            }

            // Free any unmanaged objects here.
            //
            disposed = true;
        }

        #endregion

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
    }
}
