using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace qaImageViewer.Models
{
    class DataMapper
    {
        public int Id { get; set; }
        public int ProfileId { get; set; }
        public int AttributeId { get; set; }
        public string ExcelColumnCode { get; set; }
        public bool Ignore { get; set; }
    }
}
