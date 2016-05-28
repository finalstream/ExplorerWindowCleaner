using System;
using System.Collections.Generic;

namespace ExplorerWindowCleaner
{
    public class CleanedEventArgs : EventArgs
    {
        public ICollection<ExplorerCleanedInfo> CloseInfos { get; private set; }
        public int WindowCount { get; private set; }
        public DateTime ExpireDateTime { get; private set; }
        public int MaxWindowCount { get; private set; }
        public int PinedCount { get; private set; }
        public int TotalCloseWindowCount { get; private set; }
        public bool IsUpdated { get; private set; }

        public CleanedEventArgs(ICollection<ExplorerCleanedInfo> closeInfos, int windowCount, DateTime expireDateTime, int maxWindowCount, int pinedCount, int totalCloseWindowCount, bool isUpdated)
        {
            CloseInfos = closeInfos;
            WindowCount = windowCount;
            ExpireDateTime = expireDateTime;
            MaxWindowCount = maxWindowCount;
            PinedCount = pinedCount;
            TotalCloseWindowCount = totalCloseWindowCount;
            IsUpdated = isUpdated;
        }
    }
}