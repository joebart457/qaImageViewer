using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;


namespace qaImageViewer.Models
{
    class ExcelWorksheetListItem
    {
        public string Name { get; set; }
        public List<List<string>> SheetData { get; set; }
        public override string ToString()
        {
            return Name;
        }
    }
}
