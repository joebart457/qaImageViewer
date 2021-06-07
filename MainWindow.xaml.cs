using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.ComponentModel;
using System.Data;
using System.Diagnostics;
using System.Linq;
using System.Runtime.InteropServices;
using System.Text;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Data;
using System.Windows.Documents;
using System.Windows.Input;
using System.Windows.Media;
using System.Windows.Media.Imaging;
using System.Windows.Navigation;
using System.Windows.Shapes;
using qaImageViewer.Converters;
using qaImageViewer.Models;
using qaImageViewer.Repository;
using qaImageViewer.Service;
using qaImageViewer.Tasks;
using Excel = Microsoft.Office.Interop.Excel;

namespace qaImageViewer
{

    /// <summary>
    /// Interaction logic for MainWindow.xaml
    /// </summary>
    ///
    public partial class MainWindow : Window
    {
        private ConnectionManager _connectionManager = new ConnectionManager();
        private ProcessingReportWindow _processingReportWindow = null;
        public MainWindow()
        {
            InitializeComponent();
            SetupImportColumnMappingsViewColumns();
            SetupImportColumnMappingsEditColumns();
            SetupExportColumnMappingsEditColumns();
            SetupPreviousImportResultsDataGrid();
            PopulateImportProfilesComboBox();
            HideExcelPreviewStatusLabels();
            ResetExcelPreviewData();
            
        }


        private void PopulateImportProfilesComboBox()
        {
            ComboBox_ImportProfilesSelector.ItemsSource = MappingProfileRepository.GetMappingProfiles(_connectionManager);
        }

        private void PopulatePreviousImportResultsDataGrid()
        {
            DataGrid_PreviousImportResults.ItemsSource = ImportResultRepository.GetImportResultListItems(_connectionManager);
        }
        private void PopulateMappingProfilesImportViewComboBox()
        {
            ComboBox_MappingProfilesSelector.ItemsSource = MappingProfileRepository.GetMappingProfiles(_connectionManager);
        }

        private void SetupPreviousImportResultsDataGrid()
        {
            DataGrid_PreviousImportResults.Columns.Clear();
            DataGrid_PreviousImportResults.Columns.Add(new DataGridTextColumn { 
                Header = "Profile", 
                Binding = new Binding("ProfileName")
            });

            DataGrid_PreviousImportResults.Columns.Add(new DataGridTextColumn
            {
                Header = "Workbook",
                Binding = new Binding("WorkbookName")
            });

            DataGrid_PreviousImportResults.Columns.Add(new DataGridTextColumn
            {
                Header = "Worksheet",
                Binding = new Binding("WorksheetName")
            });

            DataGrid_PreviousImportResults.Columns.Add(new DataGridTextColumn
            {
                Header = "Date",
                Binding = new Binding("EndTime")
            }); 
        }

        private void SetupExportColumnMappingsEditColumns()
        {
            DataGrid_ExportColumnMappingsEdit.Columns.Clear();
            DataGrid_ExportColumnMappingsEdit.Columns.Add(new DataGridTextColumn { 
                Header = "Alias", 
                Binding = new Binding("ImportColumnMappingAlias"), 
                IsReadOnly = true
            });
            
            var excelColumnAliasComboBoxTemplate = new FrameworkElementFactory(typeof(ComboBox));
            excelColumnAliasComboBoxTemplate.SetValue(ComboBox.ItemsSourceProperty, ExcelAppHelperService.GetExcelColumnOptionsAsList(true));
            excelColumnAliasComboBoxTemplate.SetBinding(ComboBox.SelectedItemProperty, new Binding("ExcelColumnAlias"));
            excelColumnAliasComboBoxTemplate.AddHandler(
                ComboBox.SelectionChangedEvent,
                new SelectionChangedEventHandler((o, e) => {
                    if (DataGrid_ExportColumnMappingsEdit.SelectedItem is not null)
                    {
                        ((ExportColumnMappingListItem)DataGrid_ExportColumnMappingsEdit.SelectedItem).ExcelColumnAlias = e.AddedItems[0].ToString();
                        ExportColumnMappingRepository.UpdateColumnMapping(_connectionManager, ColumnMappingService.ConvertFromListItem((ExportColumnMappingListItem)DataGrid_ExportColumnMappingsEdit.SelectedItem));
                    }
                })
            );
            DataGrid_ExportColumnMappingsEdit.Columns.Add(
                new DataGridTemplateColumn()
                {
                    Header = "Map to Excel Column",
                    CellTemplate = new DataTemplate() { VisualTree = excelColumnAliasComboBoxTemplate },
                }
            );
        }

        private void SetupImportColumnMappingsViewColumns()
        {
            DataGrid_ImportColumnMappingsView.Columns.Clear();
            DataGrid_ImportColumnMappingsView.Columns.Add(new DataGridTextColumn { Header = "Alias", Binding = new Binding("ColumnAlias"), IsReadOnly = true });
            DataGrid_ImportColumnMappingsView.Columns.Add(new DataGridTextColumn { Header = "Map from Excel Column", Binding = new Binding("ExcelColumnAlias"), IsReadOnly = true });
            DataGrid_ImportColumnMappingsView.Columns.Add(new DataGridComboBoxColumn { Header = "Type", ItemsSource = Enum.GetValues(typeof(DBColumnType)), TextBinding = new Binding("ColumnType"), IsReadOnly = true });
        }


        private void SetupImportColumnMappingsEditColumns()
        {
            DataGrid_ImportColumnMappingsEdit.Columns.Clear();

            DataGrid_ImportColumnMappingsEdit.Columns.Add(new DataGridTextColumn { Header = "Alias", Binding = new Binding("ColumnAlias"), });
            // Add Column Type ComboBox
            var columnTypeComboBoxTemplate = new FrameworkElementFactory(typeof(ComboBox));
            columnTypeComboBoxTemplate.SetValue(ComboBox.ItemsSourceProperty, Enum.GetValues(typeof(DBColumnType)));
            columnTypeComboBoxTemplate.SetBinding(ComboBox.SelectedItemProperty, new Binding("ColumnType"));
            columnTypeComboBoxTemplate.AddHandler(
                ComboBox.SelectionChangedEvent,
                new SelectionChangedEventHandler((o, e) => {
                    if (DataGrid_ImportColumnMappingsEdit.SelectedItem is not null)
                    {
                        DBColumnType columnType = Enum.IsDefined(typeof(DBColumnType), e.AddedItems[0]) ? (DBColumnType)e.AddedItems[0] : DBColumnType.TEXT;
                        ((ImportColumnMappingListItem)DataGrid_ImportColumnMappingsEdit.SelectedItem).ColumnType = columnType;
                    }
                })
            );
            DataGrid_ImportColumnMappingsEdit.Columns.Add(
                new DataGridTemplateColumn()
                {
                    Header = "Type",
                    CellTemplate = new DataTemplate() { VisualTree = columnTypeComboBoxTemplate },
                }
            );

            // Add Excel Column Mapping ComboBox
            var excelColumnAliasComboBoxTemplate = new FrameworkElementFactory(typeof(ComboBox));
            excelColumnAliasComboBoxTemplate.SetValue(ComboBox.ItemsSourceProperty, ExcelAppHelperService.GetExcelColumnOptionsAsList(false, true));
            excelColumnAliasComboBoxTemplate.SetBinding(ComboBox.SelectedItemProperty, new Binding("ExcelColumnAlias"));
            excelColumnAliasComboBoxTemplate.AddHandler(
                ComboBox.SelectionChangedEvent,
                new SelectionChangedEventHandler((o, e) => {
                    if (DataGrid_ImportColumnMappingsEdit.SelectedItem is not null)
                    {
                        ((ImportColumnMappingListItem)DataGrid_ImportColumnMappingsEdit.SelectedItem).ExcelColumnAlias = e.AddedItems[0].ToString();
                    }
                })
            );
            DataGrid_ImportColumnMappingsEdit.Columns.Add(
                new DataGridTemplateColumn()
                {
                    Header = "Map to Excel Column",
                    CellTemplate = new DataTemplate() { VisualTree = excelColumnAliasComboBoxTemplate },
                }
            );

            // Add Save Button 
            var saveButtonTemplate = new FrameworkElementFactory(typeof(Button));
            saveButtonTemplate.SetValue(Button.ContentProperty, "Save");
            saveButtonTemplate.SetBinding(Button.VisibilityProperty, new Binding
            {
                Path = new PropertyPath("Changed"),
                Converter = new BooleanToVisibilityConverter()
            });

            saveButtonTemplate.AddHandler(
                Button.ClickEvent,
                new RoutedEventHandler((o, e) => {
                    ((ImportColumnMappingListItem)DataGrid_ImportColumnMappingsEdit.SelectedItem).Changed = false;
                    ImportColumnMappingRepository.UpdateColumnMapping(_connectionManager, 
                        ColumnMappingService.ConvertFromListItem((ImportColumnMappingListItem)DataGrid_ImportColumnMappingsEdit.SelectedItem));
                    PopulateExportColumnMappingsEditDataGrid((MappingProfile)ComboBox_ImportProfilesSelector.SelectedItem);
                })
            );
            DataGrid_ImportColumnMappingsEdit.Columns.Add(
                new DataGridTemplateColumn()
                {
                    Header = "",
                    CellTemplate = new DataTemplate() { VisualTree = saveButtonTemplate },
                }
            );

            // Add Delete Button 
            var deleteButtonTemplate = new FrameworkElementFactory(typeof(Button));
            deleteButtonTemplate.SetValue(Button.ContentProperty, "Delete");
            deleteButtonTemplate.AddHandler(
                Button.ClickEvent,
                new RoutedEventHandler((o, e) => {
                    ImportColumnMappingRepository.DeleteColumnMapping(_connectionManager, 
                        ColumnMappingService.ConvertFromListItem((ImportColumnMappingListItem)DataGrid_ImportColumnMappingsEdit.SelectedItem));
                    MappingProfile profile = (MappingProfile)ComboBox_ImportProfilesSelector.SelectedItem;
                    PopulateImportColumnMappingsEditViewDataGrid(profile);
                    PopulateExportColumnMappingsEditDataGrid(profile);

                })
            );
            DataGrid_ImportColumnMappingsEdit.Columns.Add(
                new DataGridTemplateColumn()
                {
                    Header = "",
                    CellTemplate = new DataTemplate() { VisualTree = deleteButtonTemplate },
                }
            );
        }

        private void Button_SaveImportProfile_Click(object sender, RoutedEventArgs e)
        {
            try
            {
                MappingProfile profile = new MappingProfile { Name = ComboBox_ImportProfilesSelector.Text, Locked = false };
                MappingProfileRepository.InsertMappingProfile(_connectionManager, profile, out _);
                PopulateImportProfilesComboBox();
                ComboBox_ImportProfilesSelector.Text = "--Select Profile--";
                Button_SaveImportProfile.IsEnabled = false;
            } catch (Exception ex)
            {
                MessageBox.Show(ex.ToString(), "Error creating profile");
            }
        }

        private void ComboBox_ImportProfilesSelector_SelectionChanged(object sender, SelectionChangedEventArgs e)
        {
            Button_SaveImportProfile.IsEnabled = false;
            if (ComboBox_ImportProfilesSelector.SelectedItem is not null)
            {
                MappingProfile profile = (MappingProfile)ComboBox_ImportProfilesSelector.SelectedItem;
                if (profile is null)
                {
                    Button_AddColumnMapping.IsEnabled = false;
                }
                else
                {
                    PopulateImportColumnMappingsEditViewDataGrid(profile);
                    PopulateExportColumnMappingsEditDataGrid(profile);
                }
            }
        }

        private void PopulateExportColumnMappingsEditDataGrid(MappingProfile profile)
        {
            try
            {
                if (profile is not null)
                {
                    DataGrid_ExportColumnMappingsEdit.ItemsSource = ExportColumnMappingRepository.GetColumnMappingListItemsByProfileId(_connectionManager, profile.Id);
                }
                else
                {
                    MessageBox.Show("Could not retrieve full export mapping profile. Please try a different one", "Oops");
                    return;
                }
            }
            catch (Exception)
            {
                MessageBox.Show("Could not retrieve export mapping profile. Please try a different one", "Error!");
                return;
            }
        }

        private void PopulateImportColumnMappingsEditViewDataGrid(MappingProfile profile)
        {
            try
            {
                MappingProfile fullProfile = MappingProfileRepository.GetFullMappingProfileById(_connectionManager, profile.Id);
                if (fullProfile is not null)
                {
                    if (fullProfile.Locked)
                    {
                        // If it is locked, just view
                        DataGrid_ImportColumnMappingsView.ItemsSource = fullProfile.ImportColumnMappings;
                        DataGrid_ImportColumnMappingsEdit.Visibility = Visibility.Hidden;
                        DataGrid_ImportColumnMappingsView.Visibility = Visibility.Visible;
                    }
                    else
                    {
                        // If it's not locked we can edit
                        DataGrid_ImportColumnMappingsEdit.ItemsSource = fullProfile.ImportColumnMappings;
                        DataGrid_ImportColumnMappingsView.Visibility = Visibility.Hidden;
                        DataGrid_ImportColumnMappingsEdit.Visibility = Visibility.Visible;
                    }
                }
                else
                {
                    MessageBox.Show("Could not retrieve full mapping profile. Please try a different one", "Oops");
                    return;
                }
            }
            catch (Exception)
            {
                MessageBox.Show("Could not retrieve mapping profile. Please try a different one", "Error!");
                return;
            }
        }

        private void ComboBox_ImportProfilesSelector_KeyUp(object sender, KeyEventArgs e)
        {
            if (ComboBox_ImportProfilesSelector.Text.Length > 0)
            {
                Button_SaveImportProfile.IsEnabled = true;
            } else
            {
                Button_SaveImportProfile.IsEnabled = false;
            }
            
        }

        private void Button_AddColumnMapping_Click(object sender, RoutedEventArgs e)
        {
            MappingProfile profile = (MappingProfile)ComboBox_ImportProfilesSelector.SelectedItem;
            if (profile is not null && !profile.Locked)
            {
                profile = MappingProfileRepository.GetFullMappingProfileById(_connectionManager, profile.Id);
                if (profile is not null)
                {
                    ImportColumnMapping newImportMapping = new ImportColumnMapping
                    {
                        ColumnAlias = "alias",
                        ColumnType = DBColumnType.TEXT,
                        ExcelColumnAlias = "A",
                        ProfileId = profile.Id,
                        ColumnName = MappingProfileHelperService.GetNextColumnName(profile)
                    };

                    int id = ImportColumnMappingRepository.InsertColumnMapping(_connectionManager, newImportMapping);

                    PopulateImportColumnMappingsEditViewDataGrid(profile);

                    ExportColumnMapping newExportMapping = new ExportColumnMapping
                    {
                        ImportColumnMappingId = id,
                        ExcelColumnAlias = "IGNORE",
                        ProfileId = profile.Id
                    };

                    ExportColumnMappingRepository.InsertColumnMapping(_connectionManager, newExportMapping);
                    PopulateExportColumnMappingsEditDataGrid(profile);
                }
            }
        }

        private void TabControl_SelectionChanged(object sender, SelectionChangedEventArgs e)
        {
            if (e.AddedItems.Count > 0)
            {
                TabItem tab = e.AddedItems[0] as TabItem;
                if (tab is not null)
                {
                    if (tab.Header.ToString() == "Import")
                    {
                        PopulateMappingProfilesImportViewComboBox();
                        PopulatePreviousImportResultsDataGrid();
                    }
                    
                }
            }
           
        }

        private void Button_SelectExcelTargetFile_Click(object sender, RoutedEventArgs e)
        {
            Microsoft.Win32.OpenFileDialog fileDialog = new Microsoft.Win32.OpenFileDialog();
            fileDialog.DefaultExt = ".xls|.xlsx";
            fileDialog.Filter = "(.xls)|*.xls|(.xlsx)|*.xlsx";
            Nullable<bool> isFileChosen = fileDialog.ShowDialog();
            if (isFileChosen == true)
            {

                LoadExcelPreview(fileDialog.FileName);

            }
        }

        private void HideExcelPreviewStatusLabels()
        {
            Label_UnableToLoadExcelPreview.Visibility = Visibility.Hidden;
            ProgressBar_LoadingExcelPreview.Visibility = Visibility.Hidden;
            Label_LoadingExcelPreview.Visibility = Visibility.Hidden;
            Label_ExcelPreviewDataLoaded.Visibility = Visibility.Hidden;
        }

        private void ResetExcelPreviewData()
        {
            ListBox_ExcelPreviewSheets.Items.Clear();
            DataGrid_ExcelPreview.Items.Clear();
            DataGrid_ExcelPreview.Columns.Clear();
        }

        private void ShowExcelPreviewInProgess()
        {
            Label_LoadingExcelPreview.Visibility = Visibility.Visible;
            ProgressBar_LoadingExcelPreview.Visibility = Visibility.Visible;
        }

        private void ShowDataLoadedSuccessfully()
        {
            HideExcelPreviewStatusLabels();
            Label_ExcelPreviewDataLoaded.Visibility = Visibility.Visible;
        }


        private async void LoadExcelPreview(string filename)
        {
            ResetExcelPreviewData();
            HideExcelPreviewStatusLabels();
            try
            {
               
                ProgressBar_LoadingExcelPreview.Value = 0;
                ShowExcelPreviewInProgess();
                IProgress<int> progress = new Progress<int>(value => {
                    ProgressBar_LoadingExcelPreview.Value = value;
                });

                ProgressBar_LoadingExcelPreview.IsIndeterminate = true;
               
                int maxPreviewRows = ConfigRepository.GetIntegerOption(_connectionManager, "Excel.Preview.Row.Count", 5);
                int maxPreviewColumns = ConfigRepository.GetIntegerOption(_connectionManager, "Excel.Preview.Column.Count", 5);

                
                Excel.Application xlApp = null;
                Excel.Workbook workbook = await Task.Run(() => { 
                    xlApp = new Excel.Application(); 
                    return xlApp.Workbooks.Open(filename, ReadOnly: true); 
                });

                int id;
                // Find the excel process id
                Utilities.GetWindowThreadProcessId(xlApp.Hwnd, out id);
                Process excelProcess = Process.GetProcessById(id);

                Label_ExcelWorkbookName.Content = filename;

                ProgressBar_LoadingExcelPreview.IsIndeterminate = false;

                ProgressBar_LoadingExcelPreview.Maximum = maxPreviewRows * maxPreviewColumns * workbook.Worksheets.Count;

                for (int i = 1; i <= workbook.Worksheets.Count; i++)
                {
                    ListBox_ExcelPreviewSheets.Items.Add(new ExcelWorksheetListItem
                    {
                        
                        Name = ((Excel.Worksheet)workbook.Worksheets[i]).Name,
                        UsedRowCount = ((Excel.Worksheet)workbook.Worksheets[i]).UsedRange.Rows.Count,
                        WorkbookPath = filename,
                        SheetIndex = i,
                        SheetData = await Task.Run(() => ExcelAppHelperService.GetSheetData(progress, (Excel.Worksheet)workbook.Worksheets[i], maxPreviewRows, maxPreviewColumns))
                    });
                }
                ProgressBar_LoadingExcelPreview.Value = ProgressBar_LoadingExcelPreview.Maximum;
                ShowDataLoadedSuccessfully();
                workbook.Close();
                if (xlApp is not null) xlApp.Quit();
                excelProcess.Kill();
                Utilities.ReleaseObject(workbook);
                Utilities.ReleaseObject(xlApp);

            } catch (Exception ex)
            {
                Label_UnableToLoadExcelPreview.Visibility = Visibility.Visible;
                LoggerService.LogError(ex.ToString());
                MessageBox.Show("Unable to load preview.", "Error");
            }
        }



        private void ListBox_ExcelPreviewSheets_SelectionChanged(object sender, SelectionChangedEventArgs e)
        {
            ExcelWorksheetListItem item = (ExcelWorksheetListItem)ListBox_ExcelPreviewSheets.SelectedItem;
            if (item is not null && item.SheetData.Count > 0 && item.SheetData[0].Count > 0)
            {
                HideExcelPreviewStatusLabels();
                ProgressBar_LoadingExcelPreview.Value = 0;
                ShowExcelPreviewInProgess();
                ProgressBar_LoadingExcelPreview.Maximum = item.SheetData.Count + item.SheetData[0].Count;

                DataGrid_ExcelPreview.Columns.Clear();
                DataGrid_ExcelPreview.Items.Clear();
                
                int columnCount = item.SheetData[0].Count;
                for (int i = 0; i < columnCount; i++)
                {
                    DataGrid_ExcelPreview.Columns.Add(new DataGridTextColumn { 
                         Binding = new Binding
                         {
                             Path = new PropertyPath(""),
                             Converter = new RowIndexConverter(),
                             ConverterParameter = i,
                         }
                    });
                    ProgressBar_LoadingExcelPreview.Value++;
                }

             

                for (int i = 0; i< item.SheetData.Count; i++)
                {
                    DataGrid_ExcelPreview.Items.Add(item.SheetData[i]);
                    ProgressBar_LoadingExcelPreview.Value++;
                }
                HideExcelPreviewStatusLabels();
            } else
            {
                Label_UnableToLoadExcelPreview.Visibility = Visibility.Visible;

            }
        }

        private async void Button_RunExcelImport_Click(object sender, RoutedEventArgs e)
        {
            ExcelWorksheetListItem ws = (ExcelWorksheetListItem)ListBox_ExcelPreviewSheets.SelectedItem;
            if (ws is null)
            {
                MessageBox.Show("Unable to import. Please select target sheet.");
                return;
            }

            MappingProfile profile = (MappingProfile)ComboBox_MappingProfilesSelector.SelectedItem;
            if (profile is null)
            {
                MessageBox.Show("Unable to import. Please select mapping profile.");
                return;
            }

            Button_RunExcelImport.IsEnabled = false;

            int batchSize = ConfigRepository.GetIntegerOption(_connectionManager, "Tasks.Excel.Import.BatchSize", 10);

            ExcelImportItemsTask excelImportItemsTask = new ExcelImportItemsTask(
                _connectionManager, 
                ws.WorkbookPath,
                ws.SheetIndex,
                profile.Id, 
                batchSize);

            ProgressBar_ExcelImportItemsTask.Value = 0;
            ProgressBar_ExcelImportItemsTask.Maximum = ws.UsedRowCount;
            IProgress<int> progress = new Progress<int>(value => {
                ProgressBar_ExcelImportItemsTask.Value = value;
            });
            try
            {
                await Task.Run(() => { excelImportItemsTask.Execute(progress, () => { }); });
                Button_RunExcelImport.IsEnabled = true;
            } catch (Exception ex)
            {
                LoggerService.LogError(ex.ToString());
                MessageBox.Show("Unable to complete import task. See logs for details", "Error");
                Button_RunExcelImport.IsEnabled = true;
            }
        }

        private void Button_RouteToResults_Click(object sender, RoutedEventArgs e)
        {
            ImportResultsListItem selectedResults = (ImportResultsListItem)DataGrid_PreviousImportResults.SelectedItem;
            if (selectedResults != null)
            {
                if (_processingReportWindow == null || _processingReportWindow.IsLoaded == false)
                {
                    _processingReportWindow = new ProcessingReportWindow(_connectionManager, selectedResults.Id);
                    _processingReportWindow.Show();
                } else
                {
                    _processingReportWindow.Close();
                    _processingReportWindow = new ProcessingReportWindow(_connectionManager, selectedResults.Id);
                    _processingReportWindow.Show();
                }
            }
        }
    }
}
