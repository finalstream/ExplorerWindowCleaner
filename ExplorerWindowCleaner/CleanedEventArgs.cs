using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;

namespace ExplorerWindowCleaner
{
    public class CleanedEventArgs : EventArgs
    {
        public ICollection<string> CloseWindowTitles { get; private set; }
        public int WindowCount { get; private set; }
        public DateTime ExpireDateTime { get; private set; }
        public int MaxWindowCount { get; private set; }
        public int PinedCount { get; private set; }
        public int TotalCloseWindowCount { get; private set; }

        public CleanedEventArgs(ICollection<string> closeWindowTitles, 
            int windowCount, 
            DateTime expireDateTime, 
            int maxWindowCount, 
            int pinedCount, 
            int totalCloseWindowCount)
        {
            CloseWindowTitles = closeWindowTitles;
            WindowCount = windowCount;
            ExpireDateTime = expireDateTime;
            MaxWindowCount = maxWindowCount;
            PinedCount = pinedCount;
            TotalCloseWindowCount = totalCloseWindowCount;
        }
    }
}