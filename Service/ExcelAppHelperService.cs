

using System;
using System.Collections.Generic;
using System.Runtime.InteropServices;
using System.Runtime.InteropServices.ComTypes;

namespace qaImageViewer
{

    public class ExcelAppHelperService
    {
        static public List<string> GetExcelColumnOptionsAsList()
        {
            List<string> excelColumnOptions = new List<string>();
            for (int i = 0; i < 3; i++)
            {
                for (int j = 0; j < 26; j++)
                {
                    excelColumnOptions.Add(new string(Convert.ToChar('A' + j), i + 1));
                }
            }

            return excelColumnOptions;
        }
    }

    /// <summary>
    /// Gets MS Excel running workbook instances via ROT
    /// </summary>
    public class MSExcelWorkbookRunningInstances
    {
        [DllImport("ole32.dll")]
        static extern int CreateBindCtx(uint reserved, out IBindCtx ppbc);

        [DllImport("ole32.dll")]
        public static extern void GetRunningObjectTable(int reserved, out IRunningObjectTable prot);

        public static IEnumerable<dynamic> Enum()
        {
            // get Running Object Table ...
            IRunningObjectTable Rot;
            GetRunningObjectTable(0, out Rot);

            // get enumerator for ROT entries
            IEnumMoniker monikerEnumerator = null;
            Rot.EnumRunning(out monikerEnumerator);

            IntPtr pNumFetched = new IntPtr();
            IMoniker[] monikers = new IMoniker[1];

            IBindCtx bindCtx;
            CreateBindCtx(0, out bindCtx);

            while (monikerEnumerator.Next(1, monikers, pNumFetched) == 0)
            {
                string applicationName = "";
                dynamic workBook = null;
                try
                {
                    Guid IUnknown = new Guid("{00000000-0000-0000-C000-000000000046}");
                    monikers[0].BindToObject(bindCtx, null, ref IUnknown, out workBook);
                    applicationName = workBook.Application.Name;
                }
                catch { }

                if (applicationName == "Microsoft Excel") yield return workBook;
            }
        }

    }
}