using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace qaImageViewer.Models
{

    class ImportTableMapping
    {
        public int ProfileId;
        public List<ColumnMapping> ColumnMappings = new List<ColumnMapping>();
    }
}
