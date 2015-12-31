using System;
using Firk.Core.Actions;

namespace ExplorerWindowCleaner.Actions
{
    public class CloseExplorerAction : IGeneralAction<ExplorerWindowCleanerClientOperator>
    {
        private readonly Explorer _explorer;

        public CloseExplorerAction(Explorer explorer)
        {
            _explorer = explorer;
        }

        public void Invoke(ExplorerWindowCleanerClientOperator _)
        {
            _.ExplorerCleaner.CloseExplorer(_explorer);
        }
    }
}