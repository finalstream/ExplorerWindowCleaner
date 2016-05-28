using System.Windows.Forms;
using ExplorerWindowCleaner.Annotations;

namespace ExplorerWindowCleaner
{
    public class ExplorerCleanedInfo
    {
        public string WindowTitle { get; private set; }

        public ExplorerCleaner.ExplorerCloseReason CloseReason { get; private set; }
        
        public ExplorerCleanedInfo(string windowTitle, ExplorerCleaner.ExplorerCloseReason closeReason)
        {
            WindowTitle = windowTitle;
            CloseReason = closeReason;
        }
    }
}