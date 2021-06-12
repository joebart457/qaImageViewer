using qaImageViewer.Models;
using qaImageViewer.Repository;
using qaImageViewer.Service;
using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Text;
using System.Threading.Tasks;
using Excel = Microsoft.Office.Interop.Excel;


namespace qaImageViewer.Tasks
{
    enum ExportType {
        NewFile,
        Overlay
    }

    enum AttributeExportMode
    {
        None,
        First,
        CommaDelimited
    }

    class ExcelExportItemsTask: TaskInterface
    {

        private class Column_Value
        {
            public string ColumnString { get; set; }
            public object Value { get; set; }
            public DBColumnType Type { get; set; }
        }

        public int TaskId { get; set; }
        private ConnectionManager _connectionManager { get; set; }
        private ExportType _exportType { get; set; }
        private int _mappingProfileId { get; set; }
        private MappingProfile _fullProfile;
        private string _filename { get; set; }
        private int _sheetIndex { get; set; }
        private Excel.Worksheet _worksheet { get; set; }
        private int _resultSetId { get; set; }
        private AttributeExportMode _attributeExportMode { get; set; }
        private string _attributeExportTargetColumn { get; set; }
        private bool _trySave { get; set; }

        public string GetTaskData()
        {
            return $"ExportType => {_exportType}, MappingProfileId => {_mappingProfileId}, " 
                + $"FileName => {_filename}, SheetIndex => {_sheetIndex}, "
                + $"ResultSetId => {_resultSetId}, AttributeExportMode => {_attributeExportMode}, AttributeExportTargetColumn => {_attributeExportTargetColumn}, TrySave => {_trySave}";
        }

        public ExcelExportItemsTask(ConnectionManager connectionManager, ExportType exportType, string filename, int sheetIndex,
            int mappingProfileId, int resultSetId, AttributeExportMode attributeExportMode, string attributeExportTargetColumn, bool trySave)
        {
            _connectionManager = connectionManager;
            _exportType = exportType;
            _filename = filename;
            _sheetIndex = sheetIndex;
            _mappingProfileId = mappingProfileId;
            _resultSetId = resultSetId;
            _attributeExportMode = attributeExportMode;
            _attributeExportTargetColumn = attributeExportTargetColumn;
            _trySave = trySave;
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
            }
            if (!(_exportType == ExportType.NewFile || _exportType == ExportType.Overlay))
            {
                invalidParameters += "{ExportType}";
            }
            if (!(_attributeExportMode == AttributeExportMode.CommaDelimited 
                || _attributeExportMode == AttributeExportMode.First 
                || _attributeExportMode == AttributeExportMode.None))
            {
               invalidParameters += "{AttributeExportMode}";
            }

            if (!ValidatorService.ValidateSingleColumn(_attributeExportTargetColumn))
            {
                invalidParameters += "{AttributeExportTargetColumn}";
            }

            if (_fullProfile.ExportColumnMappings.Count == 0)
            {
                invalidParameters += "{MappingProfile.ImportColumnMappings}";
            }
            else
            {
                foreach (ExportColumnMappingListItem columnMapping in _fullProfile.ExportColumnMappings)
                {
                    if (!ValidatorService.ValidateSingleColumnOrRowIdOrIgnoreOption(columnMapping.ExcelColumnAlias))
                    {
                        invalidParameters += $"{{MappingProfile.ImportColumnMappings.${columnMapping.Id.ToString()}}}";
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

        private int FindMatchingRowIndex(List<ExportColumnMappingListItem> columnsToMatch, DocumentListItem doc)
        {
            if (_worksheet is not null)
            {
                var xlRange = _worksheet.UsedRange;

                List<DocumentColumn> docData = ResultSetRepository.GetFullRowDataAsKeyValuePairs(_connectionManager, doc);

                List<Column_Value> columnValues = new List<Column_Value>();
                foreach (ExportColumnMappingListItem column in columnsToMatch)
                {
                    DocumentColumn docDataToMatch = docData.Find(x => x.Mapping.Id == column.ImportColumnMappingId && column.ExcelColumnAlias != ExcelAppHelperService.IGNORE_OPTION);
                    if (docDataToMatch is null)
                    {
                        LoggerService.LogWarning($"item entry from result set {doc.ResultSetId.ToString()} with id {doc.Id.ToString()} did not contain column definition for export mapping with id {column.Id}");
                        // Insert into processing exception

                        return -1;
                    }
                    else
                    {
                        columnValues.Add(new Column_Value
                        {
                            ColumnString = column.ExcelColumnAlias,
                            Value = docDataToMatch.Value,
                            Type = docDataToMatch.Mapping.ColumnType
                        });
                    }
                }

                for (int i = 1; i <= xlRange.Rows.Count; i++)
                {
                    Excel.Range rawRowData = (Excel.Range)xlRange.Rows[i];
                    int rowIndex = rawRowData.Row;

                    bool match = true;

                    try
                    {

                        foreach (Column_Value colvalue in columnValues)
                        {
                            if (colvalue.ColumnString == ExcelAppHelperService.ROWID_OPTION)
                            {
                                if (rowIndex != Convert.ToInt32(colvalue.Value))
                                {
                                    match = false;
                                    break;
                                }
                            }
                            else if (colvalue.ColumnString == ExcelAppHelperService.IGNORE_OPTION)
                            {
                                continue;
                            }
                            else if (ValidatorService.ValidateSingleColumn(colvalue.ColumnString))
                            {
                                var cellData = (Excel.Range)rawRowData.Columns[colvalue.ColumnString];
                                

                                if (((object)cellData.Value).ToString() != colvalue.Value.ToString())
                                {
                                    match = false;
                                    break;
                                }
                            }
                            else
                            {
                                LoggerService.LogWarning($"item entry from result set {doc.ResultSetId.ToString()} with id {doc.Id.ToString()} : invalid target column value {colvalue.ColumnString}");
                                // Insert into processing_exception

                                return -1;
                            }
                        }

                        if (match)
                        {
                            return rowIndex;
                        }
                    } catch(Exception ex)
                    {
                        continue;
                    }

                }
                return -1;

            }
            throw new TaskException("target worksheet was null during processing");

        }


        private string GetAttributeDataString(DocumentListItem doc)
        {
            if (_attributeExportMode == AttributeExportMode.None) return "";

            List<AttributeListItem> attributes = AttributeRepository.GetAssignedAttributes(_connectionManager, doc.Id, _resultSetId);

            switch (_attributeExportMode)
            {
                case AttributeExportMode.CommaDelimited:
                    {
                        string result = "";
                        foreach(AttributeListItem attribute in attributes)
                        {
                            result += (result.Length > 0 ? "," : "") + attribute.Name;
                        }
                        return result;

                    }
                case AttributeExportMode.First:
                    {
                        if (attributes.Count > 0)
                        {
                            return attributes[0].Name;
                        }
                        return "";
                    }
                default:
                    return "";
            }
        }


        private void SaveDataToRow(int rowIndex, DocumentListItem doc)
        {
            if (_worksheet is not null)
            {
                var xlRange = _worksheet.UsedRange;
                var xlTargetRow = (Excel.Range)xlRange.Rows[rowIndex];

                if (xlTargetRow is null)
                {
                    LoggerService.LogWarning($"Target row index {rowIndex} was not valid");
                    //Insert into processing_exception

                    return;
                }

                List<DocumentColumn> docData = ResultSetRepository.GetFullRowDataAsKeyValuePairs(_connectionManager, doc);


                List<Column_Value> columnValues = new List<Column_Value>();
                foreach (ExportColumnMappingListItem column in _fullProfile.ExportColumnMappings)
                {
                    DocumentColumn docDataToMatch = docData.Find(x => x.Mapping.Id == column.ImportColumnMappingId && column.ExcelColumnAlias != ExcelAppHelperService.IGNORE_OPTION);
                    if (docDataToMatch is null)
                    {
                        LoggerService.LogWarning($"error exporting item. Item entry from result set {doc.ResultSetId.ToString()} with id {doc.Id.ToString()} did not contain column definition for export mapping with id {column.Id}");
                        // Insert into processing exception

                        return;
                    }
                    else
                    {
                        columnValues.Add(new Column_Value
                        {
                            ColumnString = column.ExcelColumnAlias,
                            Value = docDataToMatch.Value,
                            Type = docDataToMatch.Mapping.ColumnType
                        });
                    }
                }

                if (_attributeExportMode != AttributeExportMode.None)
                {

                    columnValues.Add(new Column_Value
                    {
                        ColumnString = _attributeExportTargetColumn,
                        Type = DBColumnType.TEXT,
                        Value = GetAttributeDataString(doc)
                    });
                }

                foreach(Column_Value colvalue in columnValues)
                {
                    if (colvalue.ColumnString == ExcelAppHelperService.IGNORE_OPTION)
                    {
                        continue;
                    } else
                    {
                        if (ValidatorService.ValidateSingleColumn(colvalue.ColumnString) && xlTargetRow.Columns[colvalue.ColumnString] is not null) {
                            ((Excel.Range)xlTargetRow.Columns[colvalue.ColumnString]).Value = colvalue.Value;
                        }
                        else
                        {
                            LoggerService.LogWarning($"error exporting item. Item entry from result set {doc.ResultSetId.ToString()} with id {doc.Id.ToString()} could not export to column with invalid format: {colvalue.ColumnString}");
                            // insert into processing exception
                        }
                    }
                }

                return;
            }
            throw new TaskException("target worksheet was null during processing");
        }


        public void InsertRowData(int rowIndex, DocumentListItem doc)
        {
            if (_worksheet is not null)
            {

                ((Excel.Range)_worksheet.Rows[rowIndex]).Insert();
                Excel.Range xlTargetRow = (Excel.Range)_worksheet.Rows[rowIndex];

                List<DocumentColumn> docData = ResultSetRepository.GetFullRowDataAsKeyValuePairs(_connectionManager, doc);


                List<Column_Value> columnValues = new List<Column_Value>();
                foreach (ExportColumnMappingListItem column in _fullProfile.ExportColumnMappings)
                {
                    DocumentColumn docDataToMatch = docData.Find(x => x.Mapping.Id == column.ImportColumnMappingId && column.ExcelColumnAlias != ExcelAppHelperService.IGNORE_OPTION);
                    if (docDataToMatch is null)
                    {
                        LoggerService.LogWarning($"error exporting item. Item entry from result set {doc.ResultSetId.ToString()} with id {doc.Id.ToString()} did not contain column definition for export mapping with id {column.Id}");
                        // Insert into processing exception

                        return;
                    }
                    else
                    {
                        columnValues.Add(new Column_Value
                        {
                            ColumnString = column.ExcelColumnAlias,
                            Value = docDataToMatch.Value,
                            Type = docDataToMatch.Mapping.ColumnType
                        });
                    }
                }

                foreach (Column_Value colvalue in columnValues)
                {
                    if (colvalue.ColumnString == ExcelAppHelperService.ROWID_OPTION)
                    {
                        continue;
                    }
                    if (colvalue.ColumnString == ExcelAppHelperService.IGNORE_OPTION)
                    {
                        continue;
                    }
                    else
                    {
                        if (ValidatorService.ValidateSingleColumn(colvalue.ColumnString) && xlTargetRow.Columns[colvalue.ColumnString] is not null)
                        {
                            ((Excel.Range)xlTargetRow.Columns[colvalue.ColumnString]).Value = colvalue.Value;
                        }
                        else
                        {
                            LoggerService.LogWarning($"error exporting item. Item entry from result set {doc.ResultSetId.ToString()} with id {doc.Id.ToString()} could not export to column with invalid format: {colvalue.ColumnString}");
                            // insert into processing exception
                        }
                    }
                }

                return;
            }
            throw new TaskException("target worksheet was null during processing");
        }

        public void Execute(IProgress<int> progress, CallBack callback)
        {

            _fullProfile = MappingProfileRepository.GetFullMappingProfileById(_connectionManager, _mappingProfileId);

            Excel.Application xlApp = new Excel.Application();
            Excel.Workbook xlWorkbook = null;

            if (_exportType == ExportType.Overlay)
            {
                xlApp.Visible = true;
                xlWorkbook = xlApp.Workbooks.Open(_filename);
                _worksheet = _sheetIndex <= xlWorkbook.Worksheets.Count ? (Excel.Worksheet)xlWorkbook.Worksheets[_sheetIndex] : null;
                
                ValidateParameters();

                List<DocumentListItem> results = ResultSetRepository.GetListItemsFromResultSet(_connectionManager, _resultSetId, new List<ColumnFilter>());

                var xlRange = _worksheet.UsedRange;

                var columnsToMatch = _fullProfile.ExportColumnMappings.FindAll(xportCol => xportCol.Match);

                for (int i = 0; i < results.Count; i++)
                {
                    int rowIndex = -1;
                    try
                    {
                        DocumentListItem doc = results[i];
                        rowIndex = FindMatchingRowIndex(columnsToMatch, doc);
                        if (rowIndex != 0 && rowIndex != -1)
                        {
                            SaveDataToRow(rowIndex, doc);
                        }
                    } catch (Exception ex)
                    {
                        LoggerService.LogWarning(ex.ToString());
                        ProcessingExceptionRepository.InsertProcessingException(_connectionManager, new ProcessingExceptionListItem
                        {
                            ErrorTrace = ex.ToString(),
                            RowIndex = rowIndex,
                            ResultSetId = _resultSetId,
                        });
                    }

                    progress.Report(i + 1);
                }

                if (_trySave) {
                    xlWorkbook.Save();
                    xlWorkbook.Close();
                }

            } else if (_exportType == ExportType.NewFile)
            {
                xlApp.SheetsInNewWorkbook = 1;
                xlApp.Visible = true;

                xlWorkbook = xlApp.Workbooks.Add(Missing.Value);
                _worksheet = (Excel.Worksheet)xlWorkbook.Sheets[1];

                ValidateParameters();

                List<DocumentListItem> results = ResultSetRepository.GetListItemsFromResultSet(_connectionManager, _resultSetId, new List<ColumnFilter>());

                for (int i = 0; i < results.Count; i++)
                {
                    try
                    {
                        DocumentListItem doc = results[i];
                        InsertRowData(i + 1, doc);
                    }
                    catch (Exception ex)
                    {
                        LoggerService.LogWarning(ex.ToString());
                        ProcessingExceptionRepository.InsertProcessingException(_connectionManager, new ProcessingExceptionListItem
                        {
                            ErrorTrace = ex.ToString(),
                            RowIndex = i+1,
                            ResultSetId = _resultSetId,
                            Type = "EXPORT"
                        });
                    }
                    progress.Report(i + 1);
                }

                if (_trySave)
                {
                    xlWorkbook.SaveAs2(_filename);
                    xlWorkbook.Close();
                }
            } else
            {
                throw new TaskException("invalid export type");
            }

            if (_trySave)
            {
                if (xlApp is not null)
                {
                    int id;
                    // Find the Process Id
                    Utilities.GetWindowThreadProcessId(xlApp.Hwnd, out id);
                    Process excelProcess = Process.GetProcessById(id);
                    xlApp.Quit();
                    excelProcess.Kill();
                }
                Utilities.ReleaseObject(_worksheet);
                Utilities.ReleaseObject(xlWorkbook);
                Utilities.ReleaseObject(xlApp);
            }
        }
    }
}
