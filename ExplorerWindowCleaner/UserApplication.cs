using System;
using System.Diagnostics;
using System.IO;
using FinalstreamCommons.Utils;
using SHDocVw;

namespace ExplorerWindowCleaner
{
    class UserApplication : InternetExplorer
    {

        private Process _process = null;
        private string _executablePath = null;
        private FileVersionInfo _fileVersionInfo = null;

        public UserApplication(Process process)
        {
            _process = process;
            _executablePath = ProcessUtils.GetExecutablePath(_process.Id);
            _fileVersionInfo = FileVersionInfo.GetVersionInfo(_executablePath);
        }

        public string ProcessName {get { return Path.GetFileNameWithoutExtension(_executablePath); }}

        public bool IsExplorer { get { return ProcessName.ToLower() == "explorer"; } }

        void IWebBrowser.GoBack()
        {
            throw new NotImplementedException();
        }

        void IWebBrowser2.GoForward()
        {
            throw new NotImplementedException();
        }

        void IWebBrowser2.GoHome()
        {
            throw new NotImplementedException();
        }

        void IWebBrowser2.GoSearch()
        {
            throw new NotImplementedException();
        }

        void IWebBrowser2.Navigate(string URL, ref object Flags, ref object TargetFrameName, ref object PostData, ref object Headers)
        {
            throw new NotImplementedException();
        }

        void IWebBrowser2.Refresh()
        {
            throw new NotImplementedException();
        }

        void IWebBrowser2.Refresh2(ref object Level)
        {
            throw new NotImplementedException();
        }

        void IWebBrowser2.Stop()
        {
            throw new NotImplementedException();
        }

        object IWebBrowser2.Application
        {
            get { throw new NotImplementedException(); }
        }

        object IWebBrowser2.Parent
        {
            get { throw new NotImplementedException(); }
        }

        object IWebBrowser2.Container
        {
            get { throw new NotImplementedException(); }
        }

        object IWebBrowser2.Document
        {
            get { throw new NotImplementedException(); }
        }

        bool IWebBrowser2.TopLevelContainer
        {
            get { throw new NotImplementedException(); }
        }

        string IWebBrowser2.Type
        {
            get { throw new NotImplementedException(); }
        }

        int IWebBrowser2.Left
        {
            get { throw new NotImplementedException(); }
            set { throw new NotImplementedException(); }
        }

        int IWebBrowser2.Top
        {
            get { throw new NotImplementedException(); }
            set { throw new NotImplementedException(); }
        }

        int IWebBrowser2.Width
        {
            get { throw new NotImplementedException(); }
            set { throw new NotImplementedException(); }
        }

        int IWebBrowser2.Height
        {
            get { throw new NotImplementedException(); }
            set { throw new NotImplementedException(); }
        }

        void IWebBrowser2.Quit()
        {
            throw new NotImplementedException();
        }

        void IWebBrowser2.ClientToWindow(ref int pcx, ref int pcy)
        {
            throw new NotImplementedException();
        }

        void IWebBrowser2.PutProperty(string Property, object vtValue)
        {
            throw new NotImplementedException();
        }

        object IWebBrowser2.GetProperty(string Property)
        {
            throw new NotImplementedException();
        }

        public void Navigate2(ref object URL, ref object Flags, ref object TargetFrameName, ref object PostData, ref object Headers)
        {
            throw new NotImplementedException();
        }

        public OLECMDF QueryStatusWB(OLECMDID cmdID)
        {
            throw new NotImplementedException();
        }

        public void ExecWB(OLECMDID cmdID, OLECMDEXECOPT cmdexecopt, ref object pvaIn, ref object pvaOut)
        {
            throw new NotImplementedException();
        }

        public void ShowBrowserBar(ref object pvaClsid, ref object pvarShow, ref object pvarSize)
        {
            throw new NotImplementedException();
        }

        void IWebBrowser2.GoBack()
        {
            throw new NotImplementedException();
        }

        void IWebBrowserApp.GoForward()
        {
            throw new NotImplementedException();
        }

        void IWebBrowserApp.GoHome()
        {
            throw new NotImplementedException();
        }

        void IWebBrowserApp.GoSearch()
        {
            throw new NotImplementedException();
        }

        void IWebBrowserApp.Navigate(string URL, ref object Flags, ref object TargetFrameName, ref object PostData, ref object Headers)
        {
            throw new NotImplementedException();
        }

        void IWebBrowserApp.Refresh()
        {
            throw new NotImplementedException();
        }

        void IWebBrowserApp.Refresh2(ref object Level)
        {
            throw new NotImplementedException();
        }

        void IWebBrowserApp.Stop()
        {
            throw new NotImplementedException();
        }

        object IWebBrowserApp.Application
        {
            get { throw new NotImplementedException(); }
        }

        object IWebBrowserApp.Parent
        {
            get { throw new NotImplementedException(); }
        }

        object IWebBrowserApp.Container
        {
            get { throw new NotImplementedException(); }
        }

        object IWebBrowserApp.Document
        {
            get { throw new NotImplementedException(); }
        }

        bool IWebBrowserApp.TopLevelContainer
        {
            get { throw new NotImplementedException(); }
        }

        string IWebBrowserApp.Type
        {
            get { throw new NotImplementedException(); }
        }

        int IWebBrowserApp.Left
        {
            get { throw new NotImplementedException(); }
            set { throw new NotImplementedException(); }
        }

        int IWebBrowserApp.Top
        {
            get { throw new NotImplementedException(); }
            set { throw new NotImplementedException(); }
        }

        int IWebBrowserApp.Width
        {
            get { throw new NotImplementedException(); }
            set { throw new NotImplementedException(); }
        }

        int IWebBrowserApp.Height
        {
            get { throw new NotImplementedException(); }
            set { throw new NotImplementedException(); }
        }

        void IWebBrowserApp.Quit()
        {
            throw new NotImplementedException();
        }

        void IWebBrowserApp.ClientToWindow(ref int pcx, ref int pcy)
        {
            throw new NotImplementedException();
        }

        void IWebBrowserApp.PutProperty(string Property, object vtValue)
        {
            throw new NotImplementedException();
        }

        object IWebBrowserApp.GetProperty(string Property)
        {
            throw new NotImplementedException();
        }

        void IWebBrowserApp.GoBack()
        {
            throw new NotImplementedException();
        }

        void IWebBrowser.GoForward()
        {
            throw new NotImplementedException();
        }

        void IWebBrowser.GoHome()
        {
            throw new NotImplementedException();
        }

        void IWebBrowser.GoSearch()
        {
            throw new NotImplementedException();
        }

        void IWebBrowser.Navigate(string URL, ref object Flags, ref object TargetFrameName, ref object PostData, ref object Headers)
        {
            throw new NotImplementedException();
        }

        void IWebBrowser.Refresh()
        {
            throw new NotImplementedException();
        }

        void IWebBrowser.Refresh2(ref object Level)
        {
            throw new NotImplementedException();
        }

        void IWebBrowser.Stop()
        {
            throw new NotImplementedException();
        }

        object IWebBrowser.Application
        {
            get { throw new NotImplementedException(); }
        }

        object IWebBrowser.Parent
        {
            get { throw new NotImplementedException(); }
        }

        object IWebBrowser.Container
        {
            get { throw new NotImplementedException(); }
        }

        object IWebBrowser.Document
        {
            get { throw new NotImplementedException(); }
        }

        bool IWebBrowser.TopLevelContainer
        {
            get { throw new NotImplementedException(); }
        }

        string IWebBrowser.Type
        {
            get { throw new NotImplementedException(); }
        }

        int IWebBrowser.Left
        {
            get { throw new NotImplementedException(); }
            set { throw new NotImplementedException(); }
        }

        int IWebBrowser.Top
        {
            get { throw new NotImplementedException(); }
            set { throw new NotImplementedException(); }
        }

        int IWebBrowser.Width
        {
            get { throw new NotImplementedException(); }
            set { throw new NotImplementedException(); }
        }

        int IWebBrowser.Height
        {
            get { throw new NotImplementedException(); }
            set { throw new NotImplementedException(); }
        }

        string IWebBrowser.LocationName
        {
            get { throw new NotImplementedException(); }
        }

        string IWebBrowser2.LocationURL
        {
            get { return new Uri(_executablePath).AbsoluteUri; }
        }

        bool IWebBrowser2.Busy
        {
            get { throw new NotImplementedException(); }
        }

        string IWebBrowser2.Name
        {
            get { throw new NotImplementedException(); }
        }

        int IWebBrowser2.HWND
        {
            get { return (int) _process.MainWindowHandle; }
        }

        string IWebBrowser2.FullName
        {
            get { return _executablePath; }
        }

        string IWebBrowser2.Path
        {
            get { throw new NotImplementedException(); }
        }

        bool IWebBrowser2.Visible
        {
            get { throw new NotImplementedException(); }
            set { throw new NotImplementedException(); }
        }

        bool IWebBrowser2.StatusBar
        {
            get { throw new NotImplementedException(); }
            set { throw new NotImplementedException(); }
        }

        string IWebBrowser2.StatusText
        {
            get { throw new NotImplementedException(); }
            set { throw new NotImplementedException(); }
        }

        int IWebBrowser2.ToolBar
        {
            get { throw new NotImplementedException(); }
            set { throw new NotImplementedException(); }
        }

        bool IWebBrowser2.MenuBar
        {
            get { throw new NotImplementedException(); }
            set { throw new NotImplementedException(); }
        }

        bool IWebBrowser2.FullScreen
        {
            get { throw new NotImplementedException(); }
            set { throw new NotImplementedException(); }
        }

        public tagREADYSTATE ReadyState
        {
            get { throw new NotImplementedException(); }
        }

        public bool Offline
        {
            get { throw new NotImplementedException(); }
            set { throw new NotImplementedException(); }
        }

        public bool Silent
        {
            get { throw new NotImplementedException(); }
            set { throw new NotImplementedException(); }
        }

        public bool RegisterAsBrowser
        {
            get { throw new NotImplementedException(); }
            set { throw new NotImplementedException(); }
        }

        public bool RegisterAsDropTarget
        {
            get { throw new NotImplementedException(); }
            set { throw new NotImplementedException(); }
        }

        public bool TheaterMode
        {
            get { throw new NotImplementedException(); }
            set { throw new NotImplementedException(); }
        }

        public bool AddressBar
        {
            get { throw new NotImplementedException(); }
            set { throw new NotImplementedException(); }
        }

        public bool Resizable
        {
            get { throw new NotImplementedException(); }
            set { throw new NotImplementedException(); }
        }

        string IWebBrowser2.LocationName
        {
            get { return _fileVersionInfo.FileDescription ?? Path.GetFileNameWithoutExtension(_fileVersionInfo.FileName); }
        }

        string IWebBrowserApp.LocationURL
        {
            get { throw new NotImplementedException(); }
        }

        bool IWebBrowserApp.Busy
        {
            get { throw new NotImplementedException(); }
        }

        string IWebBrowserApp.Name
        {
            get { throw new NotImplementedException(); }
        }

        int IWebBrowserApp.HWND
        {
            get { throw new NotImplementedException(); }
        }

        string IWebBrowserApp.FullName
        {
            get { throw new NotImplementedException(); }
        }

        string IWebBrowserApp.Path
        {
            get { throw new NotImplementedException(); }
        }

        bool IWebBrowserApp.Visible
        {
            get { throw new NotImplementedException(); }
            set { throw new NotImplementedException(); }
        }

        bool IWebBrowserApp.StatusBar
        {
            get { throw new NotImplementedException(); }
            set { throw new NotImplementedException(); }
        }

        string IWebBrowserApp.StatusText
        {
            get { throw new NotImplementedException(); }
            set { throw new NotImplementedException(); }
        }

        int IWebBrowserApp.ToolBar
        {
            get { throw new NotImplementedException(); }
            set { throw new NotImplementedException(); }
        }

        bool IWebBrowserApp.MenuBar
        {
            get { throw new NotImplementedException(); }
            set { throw new NotImplementedException(); }
        }

        bool IWebBrowserApp.FullScreen
        {
            get { throw new NotImplementedException(); }
            set { throw new NotImplementedException(); }
        }

        string IWebBrowserApp.LocationName
        {
            get { throw new NotImplementedException(); }
        }

        string IWebBrowser.LocationURL
        {
            get { throw new NotImplementedException(); }
        }

        bool IWebBrowser.Busy
        {
            get { throw new NotImplementedException(); }
        }

        public event DWebBrowserEvents2_StatusTextChangeEventHandler StatusTextChange;
        public event DWebBrowserEvents2_ProgressChangeEventHandler ProgressChange;
        public event DWebBrowserEvents2_CommandStateChangeEventHandler CommandStateChange;
        public event DWebBrowserEvents2_DownloadBeginEventHandler DownloadBegin;
        public event DWebBrowserEvents2_DownloadCompleteEventHandler DownloadComplete;
        public event DWebBrowserEvents2_TitleChangeEventHandler TitleChange;
        public event DWebBrowserEvents2_PropertyChangeEventHandler PropertyChange;
        public event DWebBrowserEvents2_BeforeNavigate2EventHandler BeforeNavigate2;
        public event DWebBrowserEvents2_NewWindow2EventHandler NewWindow2;
        public event DWebBrowserEvents2_NavigateComplete2EventHandler NavigateComplete2;
        public event DWebBrowserEvents2_DocumentCompleteEventHandler DocumentComplete;
        public event DWebBrowserEvents2_OnQuitEventHandler OnQuit;
        public event DWebBrowserEvents2_OnVisibleEventHandler OnVisible;
        public event DWebBrowserEvents2_OnToolBarEventHandler OnToolBar;
        public event DWebBrowserEvents2_OnMenuBarEventHandler OnMenuBar;
        public event DWebBrowserEvents2_OnStatusBarEventHandler OnStatusBar;
        public event DWebBrowserEvents2_OnFullScreenEventHandler OnFullScreen;
        public event DWebBrowserEvents2_OnTheaterModeEventHandler OnTheaterMode;
        public event DWebBrowserEvents2_WindowSetResizableEventHandler WindowSetResizable;
        public event DWebBrowserEvents2_WindowSetLeftEventHandler WindowSetLeft;
        public event DWebBrowserEvents2_WindowSetTopEventHandler WindowSetTop;
        public event DWebBrowserEvents2_WindowSetWidthEventHandler WindowSetWidth;
        public event DWebBrowserEvents2_WindowSetHeightEventHandler WindowSetHeight;
        public event DWebBrowserEvents2_WindowClosingEventHandler WindowClosing;
        public event DWebBrowserEvents2_ClientToHostWindowEventHandler ClientToHostWindow;
        public event DWebBrowserEvents2_SetSecureLockIconEventHandler SetSecureLockIcon;
        public event DWebBrowserEvents2_FileDownloadEventHandler FileDownload;
        public event DWebBrowserEvents2_NavigateErrorEventHandler NavigateError;
        public event DWebBrowserEvents2_PrintTemplateInstantiationEventHandler PrintTemplateInstantiation;
        public event DWebBrowserEvents2_PrintTemplateTeardownEventHandler PrintTemplateTeardown;
        public event DWebBrowserEvents2_UpdatePageStatusEventHandler UpdatePageStatus;
        public event DWebBrowserEvents2_PrivacyImpactedStateChangeEventHandler PrivacyImpactedStateChange;
        public event DWebBrowserEvents2_NewWindow3EventHandler NewWindow3;
        public event DWebBrowserEvents2_SetPhishingFilterStatusEventHandler SetPhishingFilterStatus;
        public event DWebBrowserEvents2_WindowStateChangedEventHandler WindowStateChanged;
        public event DWebBrowserEvents2_NewProcessEventHandler NewProcess;
        public event DWebBrowserEvents2_ThirdPartyUrlBlockedEventHandler ThirdPartyUrlBlocked;
        public event DWebBrowserEvents2_RedirectXDomainBlockedEventHandler RedirectXDomainBlocked;
        public event DWebBrowserEvents2_BeforeScriptExecuteEventHandler BeforeScriptExecute;
        public event DWebBrowserEvents2_WebWorkerStartedEventHandler WebWorkerStarted;
        public event DWebBrowserEvents2_WebWorkerFinsihedEventHandler WebWorkerFinsihed;
    }
}
