using System;
using System.ComponentModel;
using System.IO;
using System.Runtime.CompilerServices;
using ExplorerWindowCleaner.Annotations;
using Newtonsoft.Json;
using NLog;
using SHDocVw;

namespace ExplorerWindowCleaner
{
    public class Explorer : INotifyPropertyChanged
    {
        private readonly Logger _log = LogManager.GetCurrentClassLogger();

        #region PinLocationChangedイベント

        // Event object
        public event EventHandler<Explorer> PinLocationChanged;

        protected virtual void OnPinLocationChanged(Explorer pinExplorer)
        {
            var handler = this.PinLocationChanged;
            if (handler != null)
            {
                handler(this, pinExplorer);
            }
        }

        #endregion

        [JsonConstructor]
        public Explorer(DateTime registDateTime, DateTime lastUpdateDateTime, int handle, string locationUrl, string locationName, bool isPined, int closeCount, bool isFavorited, bool? isExplorer)
        {
            RegistDateTime = registDateTime;
            LastUpdateDateTime = lastUpdateDateTime;
            Handle = handle;
            LocationUrl = locationUrl;
            LocationName = locationName;
            IsPined = isPined;
            IsFavorited = isFavorited;
            CloseCount = closeCount;
            IsExplorer = isExplorer ?? true; // v1.2互換のため
        }

        public Explorer(InternetExplorer instance)
        {
            RegistDateTime = DateTime.Now;
            Handle = instance.HWND;
            LocationUrl = instance.LocationURL;
            LocationName = instance.LocationName;
            Instance = instance;
            LastUpdateDateTime = DateTime.Now;
            IsPined = false;

            var iename = Path.GetFileNameWithoutExtension(instance.FullName);
            IsExplorer = iename.ToLower().Equals("explorer");
        }

        [JsonIgnore]
        public string LastUpdate { get { return LastUpdateDateTime.ToString("yyyy-MM-dd HH:mm:ss"); } }
        public DateTime RegistDateTime { get; private set; }

        #region LastUpdateDateTime変更通知プロパティ

        private DateTime _LastUpdateDateTime;

        public DateTime LastUpdateDateTime
        {
            get { return _LastUpdateDateTime; }
            set
            {
                if (_LastUpdateDateTime == value) return;
                _LastUpdateDateTime = value;
                OnPropertyChanged();
                OnPropertyChanged("LastUpdate");
            }
        }

        #endregion

        public int Handle { get; private set; }

        #region LocationUrl変更通知プロパティ

        private string _LocationUrl;

        public string LocationUrl
        {
            get { return _LocationUrl; }
            set
            {
                if (_LocationUrl == value) return;
                _LocationUrl = value;
                OnPropertyChanged();
                OnPropertyChanged("LocationPath");
            }
        }

        #endregion

        [JsonIgnore]
        public string LocationKey { get { return !IsSpecialFolder ? LocationUrl : LocationName; } }

        #region LocationName変更通知プロパティ

        private string _LocationName;

        public string LocationName
        {
            get { return _LocationName; }
            set
            {
                if (_LocationName == value) return;
                _LocationName = value;
                OnPropertyChanged();
                OnPropertyChanged("LocationPath");
            }
        }

        #endregion
        [JsonIgnore]
        public string LocationPath { get { return !IsSpecialFolder ? AppUtils.GetUNCPath(new Uri(LocationUrl).LocalPath) : LocationName; } }
        
        [JsonIgnore]
        public InternetExplorer Instance { get; private set; }
        public bool IsExplorer { get; private set; }

        #region IsPined変更通知プロパティ

        private bool _IsPined;

        public bool IsPined
        {
            get { return _IsPined; }
            set
            {
                if (_IsPined == value) return;
                _IsPined = value;
                OnPropertyChanged();
            }
        }

        #endregion

        public bool IsFavorited { get; set; }
        [JsonIgnore]
        public bool IsSpecialFolder { get { return string.IsNullOrEmpty(LocationUrl); } }
        public int CloseCount { get; private set; }

        public bool Update(InternetExplorer ie)
        {
            var newlocationkey = GetLocationKey(ie);
            if (LocationKey == newlocationkey) return false; // パスに変更がない場合は何もしない。
            LocationUrl = ie.LocationURL;
            LocationName = ie.LocationName;
            LastUpdateDateTime = DateTime.Now;
            return true;
        }

        public bool UpdateWithKeepPin(InternetExplorer ie)
        {
            Explorer pinExplorer = null;
            var newlocationkey = GetLocationKey(ie);
            if (LocationKey == newlocationkey) return false; // パスに変更がない場合は何もしない。
            
            if (this.IsPined) pinExplorer = JsonConvert.DeserializeObject<Explorer>(JsonConvert.SerializeObject(this));

            Handle = ie.HWND;
            Instance = ie;
            LocationUrl = ie.LocationURL;
            LocationName = ie.LocationName;
            LastUpdateDateTime = DateTime.Now;

            if (pinExplorer != null)
            {
                IsPined = false;
                OnPinLocationChanged(pinExplorer);
            }
            return true;
        }

        private string GetLocationKey(InternetExplorer ie)
        {
            return !string.IsNullOrEmpty(ie.LocationURL) ? ie.LocationURL : ie.LocationName;
        }

        public override string ToString()
        {
            return string.Format("{0}-{1}-{2}", LocationKey, LocationName, LastUpdateDateTime);
        }

        public int Exit()
        {
            try
            {
                UpdateClosedInfo(this);
                Instance.Quit();
            }
            catch (Exception) { /* 失敗したとしても何もしない */ }
            
            _log.Debug("Exit Explorer : {0}", this);
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

        public void UpdateClosedInfo(Explorer explorer)
        {
            LastUpdateDateTime = DateTime.Now;
            CloseCount++;
            if (explorer.IsFavorited) IsFavorited = explorer.IsFavorited;
        }

        public void Restore(Explorer explorer)
        {
            if (LocationKey != explorer.LocationKey) return; // キーが違っていたら何もしない
            LastUpdateDateTime = explorer.LastUpdateDateTime;
            IsPined = explorer.IsPined;
        }

        public void SwitchFavorited()
        {
            IsFavorited = !IsFavorited;
            OnPropertyChanged("IsFavorited");
        }

        public void UpdateRegistDateTime()
        {
            RegistDateTime = DateTime.Now;
            LastUpdateDateTime = DateTime.Now;
        }
    }
}