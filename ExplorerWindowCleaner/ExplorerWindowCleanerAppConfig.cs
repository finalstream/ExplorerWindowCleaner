using System;
using System.Windows;
using ExplorerWindowCleaner.Properties;
using Firk.Core;

namespace ExplorerWindowCleaner
{
    public class ExplorerWindowCleanerAppConfig : AppConfig
    {

        public ExplorerWindowCleanerAppConfig()
        {
            Interval = TimeSpan.FromSeconds(10);
            IsAutoCloseUnused = true;
            AccentColor = "Cobalt";
            AppTheme = "BaseLight";
            IsKeepPin = true;
            ExpireInterval = TimeSpan.FromHours(5);
            ExportLimitNum = 30;
            IsNotifyCloseWindow = true;
            IsMouseHook = true;
        }

        public TimeSpan Interval { get; set; }
        public bool IsAutoCloseUnused { get; set; }
        public string AccentColor { get; set; }
        public string AppTheme { get; set; }
        public bool IsKeepPin { get; set; }
        public TimeSpan ExpireInterval { get; set; }
        public int ExportLimitNum { get; set; }
        public bool IsNotifyCloseWindow { get; set; }
        public bool IsMouseHook { get; set; }
        

        protected override void UpdateCore<T>(T config)
        {
            var appConfig = config as ExplorerWindowCleanerAppConfig;
            if (appConfig == null) return;
            Interval = appConfig.Interval;
            IsAutoCloseUnused = appConfig.IsAutoCloseUnused;
            AccentColor = appConfig.AccentColor;
            AppTheme = appConfig.AppTheme;
            IsKeepPin = appConfig.IsKeepPin;
            ExpireInterval = appConfig.ExpireInterval;
            ExportLimitNum = appConfig.ExportLimitNum;
            IsNotifyCloseWindow = appConfig.IsNotifyCloseWindow;
            IsMouseHook = appConfig.IsMouseHook;
        }
    }
}