using System;
using System.Collections.Generic;

namespace DocxReader
{
    public class DocumentPropertyStore
    {
        private readonly Dictionary<string, string> _coreProperties;

        public DocumentPropertyStore(Dictionary<string, string> coreProperties)
        {
            _coreProperties = coreProperties;
        }

        private string GetOrDefault(string name)
        {
            return _coreProperties.ContainsKey(name) ? GetOrDefault(name) : null;
        }

        public string Title
        {
            get
            {
                return GetOrDefault("dc:title") ?? "";
            }
        }

        public string Creator
        {
            get
            {
                return GetOrDefault("dc:creator") ?? "";
            } 
        }

        public string SubTitle
        {
            get
            {
                return GetOrDefault("dc:subject") ?? "";
            }
        }

        public string Keywords
        {
            get
            {
                return GetOrDefault("cp:keywords") ?? "";
            }
        }

        public string Description
        {
            get
            {
                return GetOrDefault("dc:description") ?? "";
            }
        }

        public string Category
        {
            get
            {
                return GetOrDefault("cp:category") ?? "";
            }
        }

        public string Status
        {
            get
            {
                return GetOrDefault("cp:contentStatus") ?? "";
            }
        }

        public DateTime CreatedDate
        {
            get
            {
                var createdStr = GetOrDefault("dcterms:created");
                return createdStr == null ? DateTime.MinValue : DateTime.Parse(createdStr);
            }
        }
    }
}