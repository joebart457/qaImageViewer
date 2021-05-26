using qaImageViewer.Service;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Reflection;
using System.Text;
using System.Threading.Tasks;
using Excel = Microsoft.Office.Interop.Excel;

namespace qaImageViewer.Tasks
{
    class ExcelHideColumnsTask: TaskInterface
    {
        private Excel.Workbook _activeWB { get; set; }
        private Excel.Worksheet _activeSheet { get; set; }
        private List<string> _columnsToHide { get; set; }
        private bool _convertToXLS { get; set; }
        private bool _deleteOldCSV { get; set; }
        private bool _saveAsNewFile { get; set; }
        private string _outputFolder { get; set; }
        private string _outputPrefix { get; set; }

        public ExcelHideColumnsTask(Excel.Workbook activeWB, Excel.Worksheet activeSheet, 
            List<string> columnsToHide, bool convertToXLS, bool deleteOldCSV, bool saveAsNewFile, string outputFolder, string outputPrefix)
        {
            _activeWB = activeWB;
            _activeSheet = activeSheet;
            _columnsToHide = columnsToHide;
            _convertToXLS = convertToXLS;
            _deleteOldCSV = deleteOldCSV;
            _saveAsNewFile = saveAsNewFile;
            _outputFolder = outputFolder;
            _outputPrefix = outputPrefix;
        }

        public void Execute(CallBack callback)
        {
            if (_activeWB != null && _activeSheet != null)
            {
                foreach (string col in _columnsToHide)
                {
                    if (ValidatorService.ValidateColumnFormat(col))
                    {
                        _activeSheet.Range[col].EntireColumn.Hidden = !(bool)_activeSheet.Range[col].EntireColumn.Hidden;
                    } else
                    {
                        LoggerService.LogWarning($"Unable to hide columns; invalid format: {col}");
                    }
                }


                if (_convertToXLS && _activeWB.FullName.ToLower().EndsWith("csv"))
                {
                    string oldFilePath = _activeWB.FullName;
                    string filepath = System.IO.Path.ChangeExtension(_activeWB.FullName, ".xlsx");
                    LoggerService.Log("Converting file '" + _activeWB.FullName + "' to file: '" + filepath + "'...");
                    _activeWB.SaveAs(filepath, Excel.XlFileFormat.xlOpenXMLWorkbook, Missing.Value, Missing.Value, Missing.Value, Missing.Value,
                                Excel.XlSaveAsAccessMode.xlNoChange, Missing.Value, Missing.Value, Missing.Value, Missing.Value);
                    LoggerService.Log("...action completed.");

                    if (_deleteOldCSV)
                    {
                        if (System.IO.File.Exists(oldFilePath))
                        {
                            System.IO.File.Delete(oldFilePath);
                            LoggerService.Log("Successfully deleted file '" + oldFilePath + "'.");
                        }
                        else
                        {
                            LoggerService.LogWarning("Unable to delete file '" + oldFilePath + "'. File does not exist.");
                        }
                    }
                }



                if (_saveAsNewFile)
                {
                    if (System.IO.Directory.Exists(_outputFolder))
                    {
                        string fName = _outputFolder + "\\" + _outputPrefix + _activeWB.Name;
                        try
                        {
                            _activeWB.SaveAs(fName, Missing.Value, Missing.Value, Missing.Value, Missing.Value, Missing.Value,
                                Excel.XlSaveAsAccessMode.xlNoChange, Missing.Value, Missing.Value, Missing.Value, Missing.Value);
                            LoggerService.Log("Saved file '" + fName + "'.");
                        }
                        catch (Exception ex)
                        {
                            LoggerService.LogError("An error has occurred while trying to save the file '" + fName + "': " + ex.ToString());
                            throw new TaskException("Hmm. Something went wrong. Columns were hidden but unable to save the file. See logs for more details.");
                        }
                    }
                    else
                    {
                        LoggerService.LogError("Columns were hidden, but unable to save file. Directory '" + _outputFolder + "' is not a valid directory");
                        throw new TaskException("Columns were hidden, but unable to save file. Directory '" + _outputFolder + "' is not a valid directory.");
                    }
                }
            }
            else
            {
                throw new TaskException("Target sheet was null.");
            }
        }
    }
}
