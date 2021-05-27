using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Excel = Microsoft.Office.Interop.Excel;


namespace qaImageViewer.Models
{
    class ExcelWorksheetListItem
    {
        public string WorkbookPath { get; set; }
        public int UsedRowCount { get; set; }
        public int SheetIndex { get; set; }
        public string Name { get; set; }
        public List<List<string>> SheetData { get; set; }
        public override string ToString()
        {
            return Name;
        }
    }
}
