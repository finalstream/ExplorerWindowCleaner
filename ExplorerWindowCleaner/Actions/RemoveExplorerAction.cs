using System;
using Firk.Core.Actions;

namespace ExplorerWindowCleaner.Actions
{
    public class RemoveExplorerAction : IGeneralAction<ExplorerWindowCleanerClientOperator>
    {
        private readonly Explorer _explorer;

        public RemoveExplorerAction(Explorer explorer)
        {
            _explorer = explorer;
        }

        public void Invoke(ExplorerWindowCleanerClientOperator _)
        {
            _.ExplorerCleaner.RemoveClosedDictionary(_explorer);
        }
    }
}