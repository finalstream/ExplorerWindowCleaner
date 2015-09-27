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
            Location = instance.LocationURL;
            Instance = instance;
            LastUpdateDateTime = DateTime.Now;
        }

        public int SeqNo { get; private set; }
        public string LastUpdate { get { return LastUpdateDateTime.ToString("yyyy-MM-dd HH:mm:ss"); } }
        public DateTime LastUpdateDateTime { get; private set; }
        public int Handle { get; private set; }
        public string Location { get; private set; }
        public string LocalPath { get { return new Uri(Location).LocalPath; } }
        public InternetExplorer Instance { get; private set; }

        public void Update(string newlocationUrl)
        {
            if (Location == newlocationUrl) return; // パスに変更がない場合は何もしない。
            Location = newlocationUrl;
            LastUpdateDateTime = DateTime.Now;
        }

        public override string ToString()
        {
            return string.Format("{0} {1}", LastUpdateDateTime, LocalPath);
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