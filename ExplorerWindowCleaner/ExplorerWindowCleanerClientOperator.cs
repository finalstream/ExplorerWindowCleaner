using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ExplorerWindowCleaner
{
    public class ExplorerWindowCleanerClientOperator
    {
        public ExplorerWindowCleanerClient Client { get; private set; }
        public ExplorerCleaner ExplorerCleaner { get; private set; }

        public ExplorerWindowCleanerClientOperator(ExplorerWindowCleanerClient explorerWindowCleanerClient, ExplorerCleaner explorerCleaner)
        {
            Client = explorerWindowCleanerClient;
            ExplorerCleaner = explorerCleaner;
        }
    }
}
