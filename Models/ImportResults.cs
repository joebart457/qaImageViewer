using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace qaImageViewer.Models
{
    class ImportResults
    {
        public int Id { get; set; }
        public int TaskId { get; set; }
        public int ProfileId { get; set; }
        public string ResultTableName { get; set; }
        public string WorkbookName { get; set; }
        public string WorksheetName { get; set; }
        public DateTime EndTime { get; set; }
    }
}
