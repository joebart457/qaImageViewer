using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace qaImageViewer.Models
{
    class MappingProfile
    {
        public int Id { get; set; }
        public string Name { get; set; }
        public bool Locked { get; set; }
        public ImportTableMapping ImportMapping { get; set; }

        public override string ToString()
        {
            return Name;
        }
    }
}
