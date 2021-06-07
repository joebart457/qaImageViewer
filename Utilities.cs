using qaImageViewer.Service;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Runtime.InteropServices;
using System.Text;
using System.Threading.Tasks;

namespace qaImageViewer
{
    public static class Utilities
    {

        [DllImport("user32.dll")]
        public static extern int GetWindowThreadProcessId(int hWnd, out int lpdwProcessId);

        public static void CheckNull(object obj)
        {
            if (obj is null) throw new Exception("Reference to object was null");
        }

        public static void ReleaseObject(object obj)
        {
            try
            {
                System.Runtime.InteropServices.Marshal.ReleaseComObject(obj);
                obj = null;
            }
            catch (Exception ex)
            {
                LoggerService.LogError(ex.ToString());
                obj = null;
            }
            finally
            {
                GC.Collect();
            }
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
