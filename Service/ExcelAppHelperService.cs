

using qaImageViewer.Service;
using System;
using System.Collections.Generic;
using System.Runtime.InteropServices;
using System.Runtime.InteropServices.ComTypes;

using Excel = Microsoft.Office.Interop.Excel;

namespace qaImageViewer
{

    public class ExcelAppHelperService
    {

        public static readonly string IGNORE_OPTION = "_IGNORE_";
        public static readonly string ROWID_OPTION = "_ROWID_";
        static public List<string> GetExcelColumnOptionsAsList(bool includeIgnoreOption = false, bool includeRowIdOption = false)
        {
            List<string> excelColumnOptions = new List<string>();
            if (includeIgnoreOption)
            {
                excelColumnOptions.Add(IGNORE_OPTION);
            }
            if (includeRowIdOption)
            {
                excelColumnOptions.Add(ROWID_OPTION);
            }
            for (int i = 0; i < 3; i++)
            {
                for (int j = 0; j < 26; j++)
                {
                    excelColumnOptions.Add(new string(Convert.ToChar('A' + j), i + 1));
                }
            }

            return excelColumnOptions;
        }

        static public List<List<string>> GetSheetData(IProgress<int> progress, Excel.Worksheet worksheet, int maxRows, int maxColumns)
        {
            try
            {
                int usedRowCount = worksheet.UsedRange.Rows.Count;
                int usedColumnCount = worksheet.UsedRange.Columns.Count;

                int maxColumnIndex = maxColumns < usedColumnCount ? maxColumns : usedColumnCount;
                int maxRowIndex = maxRows < usedRowCount ? maxRows : usedRowCount;
                List<List<string>> data = new List<List<string>>();
                for (int i = 1; i <= maxRowIndex; i++)
                {
                    List<string> rowValues = new List<string>();
                    for (int j = 1; j <= maxColumnIndex; j++)
                    {
                        Excel.Range intermediateValue = (Excel.Range)worksheet.Cells[i, j];
                        string valueString = "NULL";
                        if (intermediateValue is not null)
                        {
                            valueString = Convert.ToString(intermediateValue.Value);
                        } 
                        rowValues.Add(valueString);
                        progress.Report(i * maxColumnIndex + j);
                    }
                    data.Add(rowValues);
                }
                return data;
            }
            catch (Exception ex)
            {
                LoggerService.LogError(ex.ToString());
                return new List<List<String>>();
            }
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