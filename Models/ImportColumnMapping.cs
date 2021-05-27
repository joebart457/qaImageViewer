using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace qaImageViewer.Models
{
    enum DBColumnType
    {
        TEXT,
        INTEGER,
        REAL,
        BOOLEAN,
        DATE
    }

    class ImportColumnMapping
    {
        public int Id { get; set; }
        public int ProfileId { get; set; }
        public string ColumnName { get; set; }
        public string ColumnAlias { get; set; }
        public string ExcelColumnAlias { get; set; }
        public DBColumnType ColumnType { get; set; }
    }
}
