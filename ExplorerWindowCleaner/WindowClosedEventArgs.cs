using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;

namespace ExplorerWindowCleaner
{
    public class WindowClosedEventArgs : EventArgs
    {
        public ICollection<string> CloseWindowTitles { get; private set; }

        public WindowClosedEventArgs(ICollection<string> closeWindowTitles)
        {
            CloseWindowTitles = closeWindowTitles;
        }
    }
}