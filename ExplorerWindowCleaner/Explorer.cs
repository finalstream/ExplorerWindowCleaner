using System;
using SHDocVw;

namespace ExplorerWindowCleaner
{
    public class Explorer
    {

        public Explorer(int seqNo, InternetExplorer instance)
        {
            SeqNo = seqNo;
            Handle = instance.HWND;
            LocationUrl = instance.LocationURL;
            LocationName = instance.LocationName;
            Instance = instance;
            LastUpdateDateTime = DateTime.Now;
        }

        public int SeqNo { get; private set; }
        public string LastUpdate { get { return LastUpdateDateTime.ToString("yyyy-MM-dd HH:mm:ss"); } }
        public DateTime LastUpdateDateTime { get; private set; }
        public int Handle { get; private set; }
        public string LocationUrl { get; private set; }
        public string LocationKey {get { return !string.IsNullOrEmpty(LocationUrl) ? LocationUrl : LocationName; }}
        public string LocationName { get; private set; }
        public string LocationPath { get { return !string.IsNullOrEmpty(LocationUrl) ? AppUtils.GetUNCPath(new Uri(LocationUrl).LocalPath) : LocationName; } }
        public InternetExplorer Instance { get; private set; }
        public string LocationInfo { get { return string.Format("{0} - {1}", LocationName, LocationPath); }}

        public void Update(InternetExplorer ie)
        {
            var locationkey = !string.IsNullOrEmpty(ie.LocationURL) ? ie.LocationURL : ie.LocationName;
            if (LocationKey == locationkey) return; // パスに変更がない場合は何もしない。
            LocationUrl = ie.LocationURL;
            LocationName = ie.LocationName;
            LastUpdateDateTime = DateTime.Now;
        }

        public override string ToString()
        {
            return string.Format("{0} {1}", LastUpdateDateTime, LocationPath);
        }

        public int Exit()
        {
            try
            {
                Instance.Quit();
            }
            catch (Exception) { /* 失敗したとしても何もしない */ }
            
            Console.WriteLine("exit explorer : {0}", Handle);
            return Handle;
        }
    }
}