using System.Windows.Forms;

namespace ExplorerWindowCleaner
{
    internal class ClipboardHistoryItem
    {
        private readonly string _textData;

        public ClipboardHistoryItem(IDataObject dataObject)
        {
            _textData = GetText(dataObject);
        }

        public override string ToString()
        {
            return _textData.Length <= 100 ? _textData : _textData.Substring(0, 100) + " ...";
        }

        public string GetText()
        {
            return _textData;
        }

        private string GetText(IDataObject dataObject)
        {
            if(dataObject.GetDataPresent(DataFormats.UnicodeText))
            {
                var text = dataObject.GetData(DataFormats.UnicodeText);
                return text?.ToString() ?? "";

            }
            return "";
        }
    }
}