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
        public int ProfileId { get; set; }
        public string ResultTableName { get; set; }
        public DateTime EndTime { get; set; }
    }
}
