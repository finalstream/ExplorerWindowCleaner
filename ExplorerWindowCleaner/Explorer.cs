using System;
using System.ComponentModel;
using System.Runtime.CompilerServices;
using ExplorerWindowCleaner.Annotations;
using Newtonsoft.Json;
using SHDocVw;

namespace ExplorerWindowCleaner
{
    public class Explorer : INotifyPropertyChanged
    {
        [JsonConstructor]
        public Explorer(DateTime lastUpdateDateTime, int handle, string locationUrl, string locationName, bool isPined, int closeCount)
        {
            LastUpdateDateTime = lastUpdateDateTime;
            Handle = handle;
            LocationUrl = locationUrl;
            LocationName = locationName;
            IsPined = isPined;
            CloseCount = closeCount;
        }

        public Explorer(int seqNo, InternetExplorer instance)
        {
            SeqNo = seqNo;
            Handle = instance.HWND;
            LocationUrl = instance.LocationURL;
            LocationName = instance.LocationName;
            Instance = instance;
            LastUpdateDateTime = DateTime.Now;
            IsPined = false;
        }

        [JsonIgnore]
        public int SeqNo { get; private set; }
        [JsonIgnore]
        public string LastUpdate { get { return LastUpdateDateTime.ToString("yyyy-MM-dd HH:mm:ss"); } }
        public DateTime LastUpdateDateTime { get; private set; }
        public int Handle { get; private set; }
        public string LocationUrl { get; private set; }
        [JsonIgnore]
        public string LocationKey { get { return !string.IsNullOrEmpty(LocationUrl) ? LocationUrl : LocationName; }}
        public string LocationName { get; private set; }
        [JsonIgnore]
        public string LocationPath { get { return !string.IsNullOrEmpty(LocationUrl) ? AppUtils.GetUNCPath(new Uri(LocationUrl).LocalPath) : LocationName; } }
        [JsonIgnore]
        public InternetExplorer Instance { get; private set; }
        [JsonIgnore]
        public string LocationInfo { get { return string.Format("{0} - {1}", LocationName, LocationPath); }}
        public bool IsPined { get; set; }
        public int CloseCount { get; private set; }

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
                UpdateClosedInfo();
                Instance.Quit();
            }
            catch (Exception) { /* 失敗したとしても何もしない */ }
            
            Console.WriteLine("exit explorer : {0}", Handle);
            return Handle;
        }

        public void SwitchPined()
        {
            IsPined = !IsPined;
            OnPropertyChanged("IsPined");
        }

        public event PropertyChangedEventHandler PropertyChanged;

        [NotifyPropertyChangedInvocator]
        protected virtual void OnPropertyChanged([CallerMemberName] string propertyName = null)
        {
            var handler = PropertyChanged;
            if (handler != null) handler(this, new PropertyChangedEventArgs(propertyName));
        }

        public void UpdateClosedInfo()
        {
            LastUpdateDateTime = DateTime.Now;
            CloseCount++;
        }

        public void Restore(Explorer explorer)
        {
            LastUpdateDateTime = explorer.LastUpdateDateTime;
            IsPined = explorer.IsPined;
        }
    }
}