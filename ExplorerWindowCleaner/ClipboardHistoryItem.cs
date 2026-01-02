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

        public ClipboardHistoryItem(string text)
        {
            _textData = text.Trim();
        }

        public override string ToString()
        {
            return this.GetText().Length <= 50 ? this.GetText() : this.GetText().Substring(0, 50) + " ...";
        }

        public string GetText()
        {
            return _textData.Trim();
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