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

        private ExplorerWindowCleanerClient _client;

        public CleanerAction(ExplorerWindowCleanerClient ec)
        {
            _client = ec;
        }

        protected override void InvokeCoreAsync()
        {
            _client.Clean();
        }
    }
}
