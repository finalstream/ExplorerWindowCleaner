using System;

namespace ExplorerWindowCleaner
{
    public class UpdatedEventArgs : EventArgs
    {
        public int CloseWindowCount { get; private set; }

        public UpdatedEventArgs(int closeWindowCount)
        {
            CloseWindowCount = closeWindowCount;
        }
    }
}