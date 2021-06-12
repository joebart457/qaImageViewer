using qaImageViewer.Models;
using qaImageViewer.Repository;
using qaImageViewer.Service;
using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Runtime.InteropServices;
using System.Text;
using System.Threading.Tasks;
using Excel = Microsoft.Office.Interop.Excel;

namespace qaImageViewer.Tasks
{
    class ExcelImportItemsTask: TaskInterface
    {
        public int TaskId { get; set; }
        private ConnectionManager _connectionManager { get; set; }
        private int _mappingProfileId { get; set; }
        private MappingProfile _fullProfile;
        private string _filename { get; set; }
        private int _sheetIndex { get; set; }
        private Excel.Worksheet _worksheet { get; set; }
        private int _batchSize { get; set; }

        public string GetTaskData()
        {
            return $"MappingProfileId => {_mappingProfileId}, "
                + $"FileName => {_filename}, SheetIndex => {_sheetIndex}, BatchSize => {_batchSize}";
        }


        public ExcelImportItemsTask(ConnectionManager connectionManager, string filename, int sheetIndex,
            int mappingProfileId, int batchSize)
        {
            _connectionManager = connectionManager;
            _filename = filename;
            _sheetIndex = sheetIndex;
            _mappingProfileId = mappingProfileId;
            _batchSize = batchSize;
        }

        private void ValidateParameters()
        {
            string invalidParameters = "";
            if (_connectionManager is null)
            {
                invalidParameters += "{ConnectionManager}";
            } 
            if (_fullProfile is null)
            {
                invalidParameters += "{MappingProfile}";
            } else
            {
                if (_fullProfile.ImportColumnMappings.Count == 0)
                {
                    invalidParameters += "{MappingProfile.ImportColumnMappings}";
                } else
                {
                    foreach (ImportColumnMappingListItem columnMapping in _fullProfile.ImportColumnMappings)
                    {
                        if (!ValidatorService.ValidateSingleColumnOrRowIdOption(columnMapping.ExcelColumnAlias))
                        {
                            invalidParameters += $"{{MappingProfile.ImportColumnMappings.${columnMapping.Id.ToString()}}}";
                        }
                    }
                }
            }
            if (_worksheet is null)
            {
                invalidParameters += "{Worksheet}";
            }

            if (invalidParameters.Length > 0)
            {
                throw new TaskException($"Invalid parameters in ExcelImportItemsTask: {invalidParameters}");
            }
        }
        private object ConvertToCorrectType(object obj, DBColumnType type)
        {
            switch (type)
            {
                case DBColumnType.INTEGER:
                    return Convert.ToInt32(obj);
                case DBColumnType.BOOLEAN:
                    return Convert.ToBoolean(obj);
                case DBColumnType.REAL:
                    return Convert.ToDouble(obj);
                case DBColumnType.TEXT:
                    return Convert.ToString(obj);
                case DBColumnType.DATE:
                    return Convert.ToDateTime(obj);
                default:
                    return obj;
            }
        }
        public void Execute(IProgress<int> progress, CallBack callback)
        {

            _fullProfile = MappingProfileRepository.GetFullMappingProfileById(_connectionManager, _mappingProfileId);

            Excel.Application xlApp = new Excel.Application();
            Excel.Workbook xlWorkBook = xlApp.Workbooks.Open(_filename, ReadOnly: true);
            _worksheet = _sheetIndex <= xlWorkBook.Worksheets.Count ? (Excel.Worksheet)xlWorkBook.Worksheets[_sheetIndex] : null;
            ValidateParameters();
            int resultSetId = 0;
            try
            {
                resultSetId = ResultSetRepository.CreateResultSet(_connectionManager, TaskId, _fullProfile, _filename, _worksheet.Name);
            } catch(Exception ex)
            {
                string msg = $"Caught in ExcelImportItemsTask {ex.ToString()}";
                LoggerService.LogError(msg);
                throw new TaskException(msg);
            }
            
                var xlRange = _worksheet.UsedRange;
                int totalRows = _worksheet.UsedRange.Rows.Count;

                List<Document> documents = new List<Document>();
            for (int i = 1; i <= xlRange.Rows.Count; i++)
            {

                Nullable<int> rowIndex = null;
                try
                {
                    Excel.Range rawRowData = (Excel.Range)xlRange.Rows[i];
                    rowIndex = rawRowData.Row;
                    Document docToAdd = new Document();
                    foreach (ImportColumnMappingListItem columnMapping in _fullProfile.ImportColumnMappings)
                    {
                        object value = null;
                        rowIndex = null;
                        if (columnMapping.ExcelColumnAlias == ExcelAppHelperService.ROWID_OPTION)
                        {
                            value = rawRowData.Row;
                        } else
                        {
                            Excel.Range range = (Excel.Range)rawRowData.Columns[columnMapping.ExcelColumnAlias];
                            value = range is null ? null : range.Value;
                        }

                        docToAdd.Columns.Add(new DocumentColumn
                        {
                            Mapping = ColumnMappingService.ConvertFromListItem(columnMapping),
                            Value = ConvertToCorrectType(value, columnMapping.ColumnType)
                        });

                    }
                    documents.Add(docToAdd);
                }
                catch (Exception ex)
                {
                    LoggerService.LogWarning(ex.ToString());
                    try
                    {
                        ProcessingExceptionRepository.InsertProcessingException(_connectionManager,
                            new ProcessingExceptionListItem
                            {
                                ResultSetId = resultSetId,
                                RowIndex = rowIndex == null ? i : rowIndex.GetValueOrDefault(), //If unable to grab row index,
                                                                                                //just make best guess with i
                                ErrorTrace = ex.ToString(),
                                Type = "IMPORT"
                            });
                    }
                    catch (Exception exception)
                    {
                        LoggerService.LogError(exception.ToString());
                        string msg = $"Caught in ExcelImportItemsTask {exception.ToString()}";
                        throw new TaskException(msg);
                    }
                }

                if (documents.Count == _batchSize || i + 1 >= xlRange.Rows.Count)
                {
                    try
                    {
                        ResultSetRepository.InsertIntoResultSet(_connectionManager, resultSetId, documents);
                        documents.Clear();
                    }
                    catch (Exception ex)
                    {
                        string msg = $"Caught in ExcelImportItemsTask {ex.ToString()}";
                        LoggerService.LogError(msg);
                        throw new TaskException(msg);
                    }
                }
                progress.Report(i);
                callback();
            }

            xlWorkBook.Close();
            if (xlApp is not null) {
                int id;
                // Find the Process Id
                Utilities.GetWindowThreadProcessId(xlApp.Hwnd, out id);
                Process excelProcess = Process.GetProcessById(id);
                xlApp.Quit();
                excelProcess.Kill();
            }
            Utilities.ReleaseObject(_worksheet);
            Utilities.ReleaseObject(xlWorkBook);
            Utilities.ReleaseObject(xlApp);
        }

  
    }
}
