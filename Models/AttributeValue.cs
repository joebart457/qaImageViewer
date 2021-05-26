using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace qaImageViewer.Models
{
    class AttributeValue
    {
        public int Id { get; set; }
        public Attribute Attr { get; set; }
        public string RawValue { get; set; }
    }
}
