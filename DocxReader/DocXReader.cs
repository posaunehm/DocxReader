using Novacode;

namespace DocxReader
{
    public class DocXReader
    {
        private readonly DocX _doc;

        public DocXReader(string path)
        {
            _doc = DocX.Load(path);

            Properties = new DocumentPropertyStore(_doc.CoreProperties);

        }

        public DocumentPropertyStore Properties
        {
            get; private set;
        }

        public string Text {
            get
            {
                return _doc.Text;
            }
        }

    }
}
