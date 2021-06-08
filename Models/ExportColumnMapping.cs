using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace qaImageViewer.Models
{
    class ExportColumnMapping
    {
        public int Id { get; set; }
        public int ProfileId { get; set; }
        public int ImportColumnMappingId { get; set; }
        public string ExcelColumnAlias { get; set; }
        public bool Match { get; set; }
    }
}
