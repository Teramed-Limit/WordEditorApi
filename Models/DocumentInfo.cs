using System;
using System.Collections.Generic;
using System.Linq;
using System.Threading.Tasks;

namespace WordEditorApi.Models
{
    public class DocumentInfo
    {
        public string Id { get; set; }
        public string FileName { get; set; }
        public string Url { get; set; } // OnlyOffice 能存取的 URL
    }
}