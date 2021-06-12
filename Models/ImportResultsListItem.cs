using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace qaImageViewer.Models
{
    class ImportResultsListItem
    {
        public int Id { get; set; }
        public int ProfileId { get; set; }
        public string ProfileName { get; set; }
        public string ResultTableName { get; set; }
        public string WorkbookName { get; set; }
        public string WorksheetName { get; set; }
        public DateTime EndTime { get; set; }
        public int TaskId { get; set; }

        public override string ToString()
        {
            return $"{{TaskId:{TaskId}}}{ProfileName} - {WorkbookName}:{WorksheetName} ({EndTime.ToString()})";
        }
    }
}
