using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Firk.Core.Actions;

namespace ExplorerWindowCleaner.Actions
{
    class CleanAction : IGeneralAction<ExplorerWindowCleanerClientOperator>
    {
        public void Invoke(ExplorerWindowCleanerClientOperator _)
        {
            _.ExplorerCleaner.Clean();
        }
    }
}
