using qaImageViewer.Repository;
using qaImageViewer.Service;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Excel = Microsoft.Office.Interop.Excel;

namespace qaImageViewer.Tasks
{
    class ExcelRunDateLabelMacrosTask
    {
        private string _dateColumn { get; set; }
        private string _targetColumn { get; set; }
        private string _fallback { get; set; }
        private ConnectionManager _connectionManager { get; set; }

        private Excel.Worksheet _activeSheet { get; set; }

        public ExcelRunDateLabelMacrosTask(ConnectionManager connectionManager, Excel.Worksheet activeSheet, 
            string dateColumn, string targetColumn, string fallback)
        {
            _connectionManager = connectionManager;
            _activeSheet = activeSheet;
            _dateColumn = dateColumn;
            _targetColumn = targetColumn;
            _fallback = fallback;
        }

        public void Execute(IProgress<int> progress, CallBack callback)
        {

            if (!(ValidatorService.ValidateSingleColumn(_dateColumn) && ValidatorService.ValidateSingleColumn(_targetColumn)))
            {
               throw new TaskException("Expect date column and target column to be a single column value.");
            }

            if (_activeSheet != null)
            {
                var xlRange = _activeSheet.UsedRange;
                int totalRows = _activeSheet.UsedRange.Rows.Count;

                for (int i = 1; i <= xlRange.Rows.Count; i++)
                {
                    Excel.Range rawDateColumn = ((dynamic)xlRange.Rows[i]).Columns[_dateColumn];

                    if (rawDateColumn != null)
                    {
                        if (rawDateColumn.Value is System.DateTime)
                        {
                            string label = "label";//DateRangeRepository.GetDateLabelsForDateTime(_connectionManager, rawDateColumn.Value, _fallback);
                            ((dynamic)xlRange.Rows[i]).Columns[_targetColumn] = label;
                        }
                        else
                        {
                            LoggerService.LogWarning($"Date column in row {i.ToString()} was not of type System.DateTime");
                        }
                    }
                    else
                    {
                        LoggerService.LogWarning($"Date column in row {i.ToString()} was null.");
                    }
                    progress.Report(i);
                    callback();
                }
            }
            else
            {
                throw new TaskException("Target sheet was null");
            }
        }
    }
}
