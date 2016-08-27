using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Newtonsoft.Json;

namespace ExplorerWindowCleaner
{
    public class ShortcutItem
    {
        private SpecialFolderManager _sfm;

        public string Name { get; private set; }

        public string Value { get; private set; }

        public bool IsSpecialFolder { get; private set; }

        public ShortcutItem(string name, string value)
        {
            _sfm = new SpecialFolderManager();
            IsSpecialFolder = _sfm.IsMatchSpecialFolder(value);
            Name = name;
            Value = value;
        }

        public void Exec(bool isMinimized = false)
        {
            ProcessStartInfo psi;
            if (IsSpecialFolder)
            {
                var path = IsSpecialFolder ? _sfm.ConvertSpecialFolder(Value) : Value;
                psi = new ProcessStartInfo("EXPLORER.EXE", string.Format("/n,\"{0}\"", path));
            }
            else
            {
                psi = new ProcessStartInfo(Value);
                psi.WorkingDirectory = Path.GetDirectoryName(Value);
            }
            if (isMinimized) psi.WindowStyle = ProcessWindowStyle.Minimized;
            Process.Start(psi);
        }



    }
}
