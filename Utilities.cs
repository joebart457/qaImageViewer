using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace qaImageViewer
{
    public static class Utilities
    {

        public static void CheckNull(object obj)
        {
            if (obj is null) throw new Exception("Reference to object was null");
        }
        public class Double<_Ty>{

            public _Ty A { get; set; }
            public _Ty B { get; set; }
            public Double(_Ty a, _Ty b)
            {
                A = a;
                B = b;
            }
        }

    }
}
