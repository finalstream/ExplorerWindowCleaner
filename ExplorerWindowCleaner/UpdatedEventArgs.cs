using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;

namespace ExplorerWindowCleaner
{
    public class UpdatedEventArgs : EventArgs
    {
        public ICollection<string> CloseWindowTitles { get; private set; }

        public UpdatedEventArgs(ICollection<string> closeWindowTitles)
        {
            CloseWindowTitles = closeWindowTitles;
        }
    }
}