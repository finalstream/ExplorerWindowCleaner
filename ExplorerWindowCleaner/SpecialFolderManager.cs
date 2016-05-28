using System;
using System.Collections.Generic;
using System.Globalization;
using System.Linq;
using System.Runtime.InteropServices;
using System.Text;
using System.Threading.Tasks;

namespace ExplorerWindowCleaner
{
    public class SpecialFolderManager
    {
        /// <summary>
        /// 特殊フォルダディクショナリ
        /// </summary>
        /// <remarks>ref http://www.eightforums.com/tutorials/6050-shell-commands-windows-8-a.html </remarks>
        private readonly Dictionary<string, Dictionary<string, string>> _convDic = new Dictionary<string, Dictionary<string, string>>()
        {
            
            {
                "ja-JP", 
                new Dictionary<string, string>()
                {
                    {"デスクトップ", "Desktop"},
                    {"お気に入り", "Favorites"},
                    {"PC", "MyComputerFolder"},
                    {"コンピューター", "MyComputerFolder"},
                    {"ダウンロード", "Downloads"},
                    {"ライブラリ", "Libraries"},
                    {"ドキュメント", "DocumentsLibrary"},
                    {"ミュージック", "MusicLibrary"},
                    {"ピクチャ", "PicturesLibrary"},
                    {"ビデオ", "VideosLibrary"},
                    {"最近表示した場所", "::{22877A6D-37A1-461A-91B0-DBDA5AAEBC99}"},
                    {"最近使った項目", "Recent"},
                    {"ネットワーク", "NetworkPlacesFolder"},
                    {"ごみ箱", "RecycleBinFolder"},
                    {"コントロールパネル", "::{26EE0668-A00A-44D7-9371-BEB064C98683}"},
                    {"すべてのコントロール パネル項目", "ControlPanelFolder"},
                    {"プログラムと機能", "ChangeRemoveProgramsFolder"},
                    {"ネットワーク接続", "ConnectionsFolder"},
                } 
            
            }
        
        };

        private readonly string _localeName;

        public SpecialFolderManager()
        {
            _localeName = CultureInfo.InstalledUICulture.Name;
        }

        public string ConvertSpecialFolder(string specialFolderName)
        {
            if (!_convDic.ContainsKey(_localeName)) return "";

            var dic = _convDic[_localeName];
            if (dic.ContainsKey(specialFolderName))
            {
                return "shell:" + dic[specialFolderName];
            }
            return "";
        }


    }
}
