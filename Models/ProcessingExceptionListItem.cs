using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace qaImageViewer.Models
{
    class ProcessingExceptionListItem
    {
        public int Id { get; set; }
        public int TaskId { get; set; }
        public int ResultSetId { get; set; }
        public string ResultSetName { get; set; }
        public int RowIndex { get; set; }
        public string ErrorTrace { get; set; }
        public DateTime ErrorTime { get; set; }
        public string Type { get; set; }
    }
}
