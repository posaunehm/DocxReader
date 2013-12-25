using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using Novacode;

namespace DocxReader
{
    public class DocXReader
    {
        private DocX _doc;

        public DocXReader(string path)
        {
            _doc = DocX.Load(path);
        }

        public string Text {
            get
            {
                return _doc.Text;
            }
        }
    }
}
