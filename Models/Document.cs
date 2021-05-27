using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace qaImageViewer.Models
{
    class DocumentColumn { 
        public ImportColumnMapping Mapping { get; set; }
        public object Value { get; set; }
    }


    class Document
    {
        public int Id { get; set; }

        public string ResultTableName { get; set; }
        public List<DocumentColumn> Columns { get; set; }
    }
}
