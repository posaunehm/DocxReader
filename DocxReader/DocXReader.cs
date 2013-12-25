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

    public class DocumentPropertyStore
    {
        private Dictionary<string, string> _coreProperties;

        public DocumentPropertyStore(Dictionary<string, string> coreProperties)
        {
            _coreProperties = coreProperties;
        }

        //[0]: {[dc:title, Title]}
        //[1]: {[dc:subject, サブタイトル]}
        //[2]: {[dc:creator, posaunehm]}
        //[3]: {[cp:keywords, Keyword]}
        //[4]: {[dc:description, ]}
        //[5]: {[cp:lastModifiedBy, 前川博志]}
        //[6]: {[cp:revision, 5]}
        //[7]: {[dcterms:created, 2013-12-25T13:00:00Z]}
        //[8]: {[dcterms:modified, 2013-12-25T13:54:00Z]}
        //[9]: {[cp:category, 分類]}
        //[10]: {[cp:contentStatus, 状態]}

        public string Title { get { return _coreProperties["dc:title"]; } }
        public string Creator { get { return _coreProperties["dc:creator"]; } }
        public string SubTitle { get { return _coreProperties["dc:subject"]; } }
        public string Keywords { get { return _coreProperties["cp:keywords"]; }}
        public string Description { get { return _coreProperties["dc:description"]; } }
        public string Category { get { return _coreProperties["cp:category"]; } }
        public string Status { get { return _coreProperties["cp:contentStatus"]; } }
        public DateTime CreatedDate { get { return DateTime.Parse(_coreProperties["dcterms:created"]); } }
    }
}
