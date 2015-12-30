using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Firk.Core.Actions;

namespace ExplorerWindowCleaner.Actions
{
    class CleanerAction : BackgroundAction
    {

        private ExplorerCleaner _explorerCleaner;

        public CleanerAction(ExplorerCleaner ec)
        {
            _explorerCleaner = ec;
        }

        protected override void InvokeCoreAsync()
        {
            _explorerCleaner.Clean();
        }
    }
}
