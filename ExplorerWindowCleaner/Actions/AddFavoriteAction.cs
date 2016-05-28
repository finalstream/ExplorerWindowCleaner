using System;
using Firk.Core.Actions;
using Newtonsoft.Json;

namespace ExplorerWindowCleaner.Actions
{
    public class AddFavoriteAction : IGeneralAction<ExplorerWindowCleanerClientOperator>
    {
        private readonly Explorer _explorer;

        public AddFavoriteAction(Explorer explorer)
        {
            _explorer = explorer;
        }

        public void Invoke(ExplorerWindowCleanerClientOperator _)
        {
            var cloneExplorer = JsonConvert.DeserializeObject<Explorer>(JsonConvert.SerializeObject(_explorer));
            cloneExplorer.IsFavorited = true; // ピン留めされたときにお気に入りに登録する。
            _.ExplorerCleaner.AddOrUpdateClosedDictionary(cloneExplorer);
        }
    }
}