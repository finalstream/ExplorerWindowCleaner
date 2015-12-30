using System;
using System.Windows;
using ExplorerWindowCleaner.Properties;
using Firk.Core;

namespace ExplorerWindowCleaner
{
    internal class ExplorerWindowCleanerAppConfig : IAppConfig
    {

        private string _appVersion;

        public ExplorerWindowCleanerAppConfig(string appVersion, Settings settings)
        {
            _appVersion = appVersion;
            IsAutoCloseUnused = settings.IsAutoCloseUnused;
            AccentColor = settings.AccentColor;
            AppTheme = settings.AppTheme;
            IsKeepPin = settings.IsKeepPin;
            ExpireInterval = settings.ExpireInterval;
            ExportLimitNum = settings.ExportLimitNum;
        }

        public string AppVersion
        {
            get { return _appVersion; }
        }

        public int SchemaVersion
        {
            get { throw new System.NotImplementedException(); }
        }

        public Rect WindowBounds
        {
            get { throw new System.NotImplementedException(); }
        }

        public bool IsAutoCloseUnused { get; set; }
        public string AccentColor { get; set; }
        public string AppTheme { get; set; }
        public bool IsKeepPin { get; set; }
        public TimeSpan ExpireInterval { get; set; }
        public int ExportLimitNum { get; set; }

        public void UpdateSchemaVersion(int version)
        {
            throw new System.NotImplementedException();
        }
    }
}