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
using qaImageViewer.Managers;
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
        private ConnectionManager _connectionManager = null;
        private ProcessingReportWindow _processingReportWindow = null;
        private ImageViewer _imageViewerWindow = null;

        private Brush _primaryBackground = new SolidColorBrush(Colors.White);
        private Brush _primary = new SolidColorBrush(Colors.Blue);
        private Brush _progressBarFill = new LinearGradientBrush(Colors.Purple, Colors.Pink, 0.0);
        private Brush _buttonBackground = new SolidColorBrush(Colors.Blue);
        private Brush _buttonBorder = new SolidColorBrush(Colors.LightBlue);
        private Brush _buttonForeground = new SolidColorBrush(Colors.White);

        public MainWindow()
        {
            InitializeComponent();
            InitializeColorScheme();
            try
            {
                _connectionManager = new ConnectionManager();

                SetupImportColumnMappingsViewColumns();
                SetupImportColumnMappingsEditColumns();
                SetupExportColumnMappingsEditColumns();
                SetupPreviousImportResultsDataGrid();
                SetupTaskViewDataGridColumns();
                PopulateImportProfilesComboBox();
                PopulateExportTypeComboBox();
                PopulateAttributeExportModeComboBox();
                PopulateAttributeExportTargetComboBox();
                HideExcelPreviewStatusLabels();
                ResetExcelPreviewData();
            } catch (Exception ex)
            {
                LoggerService.LogError(ex.ToString());
                throw ex;
            }
        }

        private void InitializeColorScheme()
        {
            Grid_MappingProfilesTab.Background = _primaryBackground;
            Grid_ImportTab.Background = _primaryBackground;
            Grid_ExportTab.Background = _primaryBackground;
            Grid_TasksTab.Background = _primaryBackground;

            GroupBox_AllTasks.BorderBrush = _primary;
            GroupBox_AttributeExportOptions.BorderBrush = _primary;
            GroupBox_ExportOptions.BorderBrush = _primary;
            GroupBox_Import.BorderBrush = _primary;
            GroupBox_ImportResults.BorderBrush = _primary;
            GroupBox_NewFileOptions.BorderBrush = _primary;
            GroupBox_OverlayOptions.BorderBrush = _primary;
            GroupBox_TaskStatus.BorderBrush = _primary;

            Rectange_ExportMappingBorder.Stroke = _primary;
            Rectange_ImportMappingBorder.Stroke = _primary;

            ProgressBar_ExcelImportItemsTask.Foreground = _progressBarFill;
            ProgressBar_ExportTaskStatus.Foreground = _progressBarFill;
            ProgressBar_LoadingExcelPreview.Foreground = _progressBarFill;

            Button_AddColumnMapping.Background = _buttonBackground;
            Button_ChooseExportFile.Background = _buttonBackground;
            Button_RouteToResults.Background = _buttonBackground;
            Button_RouteToReviewWindow.Background = _buttonBackground;
            Button_RunExcelImport.Background = _buttonBackground;
            //Button_SaveImportProfile.Background = _buttonBackground;
            Button_SelectExcelTargetFile.Background = _buttonBackground;
            Button_StartExport.Background = _buttonBackground;

            Button_AddColumnMapping.BorderBrush = _buttonBorder;
            Button_ChooseExportFile.BorderBrush = _buttonBorder;
            Button_RouteToResults.BorderBrush = _buttonBorder;
            Button_RouteToReviewWindow.BorderBrush = _buttonBorder;
            Button_RunExcelImport.BorderBrush = _buttonBorder;
            //Button_SaveImportProfile.BorderBrush = _buttonBorder;
            Button_SelectExcelTargetFile.BorderBrush = _buttonBorder;
            Button_StartExport.BorderBrush = _buttonBorder;

            Button_AddColumnMapping.Foreground = _buttonForeground;
            Button_ChooseExportFile.Foreground = _buttonForeground;
            Button_RouteToResults.Foreground = _buttonForeground;
            Button_RouteToReviewWindow.Foreground = _buttonForeground;
            Button_RunExcelImport.Foreground = _buttonForeground;
            //Button_SaveImportProfile.Foreground = _buttonForeground;
            Button_SelectExcelTargetFile.Foreground = _buttonForeground;
            Button_StartExport.Foreground = _buttonForeground;
        }

        private void PopulateExportResultSetTargetComboBox()
        {
            ComboBox_ExportResultSetTarget.ItemsSource = ImportResultRepository.GetImportResultListItems(_connectionManager);
        }

        private void PopulateAttributeExportTargetComboBox()
        {
            ComboBox_AttributeExportTarget.ItemsSource = ExcelAppHelperService.GetExcelColumnOptionsAsList(false, false);
            ComboBox_AttributeExportTarget.SelectedItem = "A";
        }

        private void PopulateAttributeExportModeComboBox()
        {
            ComboBox_AttributeExportMode.ItemsSource = Enum.GetValues(typeof(AttributeExportMode));
            ComboBox_AttributeExportMode.SelectedItem = AttributeExportMode.First;
        }

        private void PopulateExportTypeComboBox()
        {
            ComboBox_ExportType.ItemsSource = Enum.GetValues(typeof(ExportType));
            ComboBox_ExportType.SelectedItem = ExportType.NewFile;
        }
        private void PopulateImportProfilesComboBox()
        {
            ComboBox_ImportProfilesSelector.ItemsSource = MappingProfileRepository.GetMappingProfiles(_connectionManager);
            ComboBox_ImportProfilesSelector.Text = "--Select Profile--";
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
            excelColumnAliasComboBoxTemplate.SetValue(ComboBox.ItemsSourceProperty, ExcelAppHelperService.GetExcelColumnOptionsAsList(true, true));
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

            var matchCheckboxTemplate = new FrameworkElementFactory(typeof(CheckBox));
            matchCheckboxTemplate.SetBinding(CheckBox.IsCheckedProperty, new Binding("Match"));
            matchCheckboxTemplate.AddHandler(
                CheckBox.CheckedEvent,
                new RoutedEventHandler((o,e) => {
                    if (DataGrid_ExportColumnMappingsEdit.SelectedItem is not null)
                    {
                        ((ExportColumnMappingListItem)DataGrid_ExportColumnMappingsEdit.SelectedItem).Match = true;
                        ExportColumnMappingRepository.UpdateColumnMapping(_connectionManager, ColumnMappingService.ConvertFromListItem((ExportColumnMappingListItem)DataGrid_ExportColumnMappingsEdit.SelectedItem));
                    }
                })
            );
            DataGrid_ExportColumnMappingsEdit.Columns.Add(
                new DataGridTemplateColumn()
                {
                    Header = "Match",
                    CellTemplate = new DataTemplate() { VisualTree = matchCheckboxTemplate },
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
                        ExcelColumnAlias = ExcelAppHelperService.ROWID_OPTION,
                        ProfileId = profile.Id,
                        ColumnName = MappingProfileHelperService.GetNextColumnName(profile)
                    };

                    int id = ImportColumnMappingRepository.InsertColumnMapping(_connectionManager, newImportMapping);

                    PopulateImportColumnMappingsEditViewDataGrid(profile);

                    ExportColumnMapping newExportMapping = new ExportColumnMapping
                    {
                        ImportColumnMappingId = id,
                        ExcelColumnAlias = ExcelAppHelperService.IGNORE_OPTION,
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
                    if (tab.Header.ToString() == "Export")
                    {
                        PopulateExportResultSetTargetComboBox();
                    }
                    if (tab.Header.ToString() == "Tasks")
                    {
                        PopulateTaskViewDataGrid();
                    }
                    if (tab.Header.ToString() == "Mapping Profiles")
                    {
                        PopulateImportProfilesComboBox();
                        DataGrid_ImportColumnMappingsEdit.ItemsSource = null;
                        DataGrid_ImportColumnMappingsView.ItemsSource = null;
                        DataGrid_ExportColumnMappingsEdit.ItemsSource = null;
                    }


                }
            }
           
        }

        private string? ChooseExcelFile()
        {
            Microsoft.Win32.OpenFileDialog fileDialog = new Microsoft.Win32.OpenFileDialog();
            fileDialog.DefaultExt = ".xls|.xlsx";
            fileDialog.Filter = "(.xlsx)|*.xlsx|(.xls)|*.xls";
            Nullable<bool> isFileChosen = fileDialog.ShowDialog();
            if (isFileChosen == true)
            {
                return fileDialog.FileName;
            }
            return null;
        }

        private void Button_SelectExcelTargetFile_Click(object sender, RoutedEventArgs e)
        {
            string? fileName = ChooseExcelFile();
            if (fileName is not null)
            {
                LoadExcelPreview(fileName);
            }
        }

        private void HideExcelPreviewStatusLabels()
        {
            Label_UnableToLoadExcelPreview.Visibility = Visibility.Hidden;
            ProgressBar_LoadingExcelPreview.Visibility = Visibility.Hidden;
            Label_LoadingExcelPreview.Visibility = Visibility.Hidden;
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
                Label_ImportStatus.Content = $"importing row {value.ToString()} of {ProgressBar_ExcelImportItemsTask.Maximum.ToString()}";
            });
            try
            {
                await TaskManager.Launch(_connectionManager, excelImportItemsTask, progress);
                Button_RunExcelImport.IsEnabled = true;
                Label_ImportStatus.Content = "Done";
                PopulatePreviousImportResultsDataGrid();
            } catch (Exception ex)
            {
                MessageBox.Show(ex.ToString(), "Import Task Failed");
                Button_RunExcelImport.IsEnabled = true;
                Label_ImportStatus.Content = "Done";
            }
        }

        private void Button_RouteToResults_Click(object sender, RoutedEventArgs e)
        {
            ImportResultsListItem selectedResults = (ImportResultsListItem)DataGrid_PreviousImportResults.SelectedItem;
            if (selectedResults != null)
            {
                if (_processingReportWindow == null || _processingReportWindow.IsLoaded == false)
                {
                    _processingReportWindow = new ProcessingReportWindow(_connectionManager, selectedResults.TaskId);
                    _processingReportWindow.Show();
                } else
                {
                    _processingReportWindow.Close();
                    _processingReportWindow = new ProcessingReportWindow(_connectionManager, selectedResults.TaskId);
                    _processingReportWindow.Show();
                }
            }
        }

        private void DataGrid_PreviousImportResults_SelectionChanged(object sender, SelectionChangedEventArgs e)
        {
            Button_RouteToResults.IsEnabled = true;
            Button_RouteToReviewWindow.IsEnabled = true;
        }

        private void Button_RouteToReviewWindow_Click(object sender, RoutedEventArgs e)
        {
            ImportResultsListItem selectedResults = (ImportResultsListItem)DataGrid_PreviousImportResults.SelectedItem;
            if (selectedResults != null)
            {
                if (_imageViewerWindow == null || _imageViewerWindow.IsLoaded == false)
                {
                    _imageViewerWindow = new ImageViewer(_connectionManager, selectedResults.Id);
                    _imageViewerWindow.Show();
                }
                else
                {
                    _imageViewerWindow.Close();
                    _imageViewerWindow = new ImageViewer(_connectionManager, selectedResults.Id);
                    _imageViewerWindow.Show();
                }
            }
        }

        private void Button_ChooseExportFile_Click(object sender, RoutedEventArgs e)
        {
            string? fileName = ChooseExcelFile();
            if (fileName is not null)
            {
                Button_StartExport.IsEnabled = false;
                PopulateExportSheetNamesComboBox(fileName);
                Button_StartExport.IsEnabled = true;
            }
        }

        private async void PopulateExportSheetNamesComboBox(string filename)
        {
            Button_ChooseExportFile.IsEnabled = false;
            ComboBox_ExportSheetNames.Items.Clear();
            try
            {
                Excel.Application xlApp = null;
                Excel.Workbook workbook = await Task.Run(() =>
                {
                    xlApp = new Excel.Application();
                    return xlApp.Workbooks.Open(filename, ReadOnly: true);
                });

                int id;
                // Find the excel process id
                Utilities.GetWindowThreadProcessId(xlApp.Hwnd, out id);
                Process excelProcess = Process.GetProcessById(id);

                Label_ExportFileName.Content = filename;


                for (int i = 1; i <= workbook.Worksheets.Count; i++)
                {
                    ComboBox_ExportSheetNames.Items.Add(new ExcelWorksheetListItem
                    {

                        Name = ((Excel.Worksheet)workbook.Worksheets[i]).Name,
                        UsedRowCount = ((Excel.Worksheet)workbook.Worksheets[i]).UsedRange.Rows.Count,
                        WorkbookPath = filename,
                        SheetIndex = i,
                    });
                }

                workbook.Close();
                if (xlApp is not null) xlApp.Quit();
                excelProcess.Kill();
                Utilities.ReleaseObject(workbook);
                Utilities.ReleaseObject(xlApp);
            } catch (Exception ex)
            {
                LoggerService.LogError(ex.ToString());
                MessageBox.Show(ex.ToString());
            }
            Button_ChooseExportFile.IsEnabled = true;
        }

        private void ComboBox_AttributeExportMode_SelectionChanged(object sender, SelectionChangedEventArgs e)
        {
            if (ComboBox_AttributeExportMode.SelectedItem is null || 
                Enum.IsDefined(typeof(AttributeExportMode), ComboBox_AttributeExportMode.SelectedItem) 
                && (AttributeExportMode)ComboBox_AttributeExportMode.SelectedItem == AttributeExportMode.None)
            {
                ComboBox_AttributeExportTarget.IsEnabled = false;
            } else
            {
                ComboBox_AttributeExportTarget.IsEnabled = true;
            }
        }

        private void ComboBox_ExportType_SelectionChanged(object sender, SelectionChangedEventArgs e)
        {
            if (Enum.IsDefined(typeof(ExportType), ComboBox_ExportType.SelectedItem))
            {
                ExportType exportType = (ExportType)ComboBox_ExportType.SelectedItem;
                if (exportType == ExportType.NewFile)
                {
                    GroupBox_NewFileOptions.IsEnabled = true;
                    GroupBox_OverlayOptions.IsEnabled = false;
                    CheckBox_TrySave.IsChecked = true;
                } else
                {
                    GroupBox_NewFileOptions.IsEnabled = false;
                    GroupBox_OverlayOptions.IsEnabled = true;
                }
            } else
            {
                GroupBox_NewFileOptions.IsEnabled = false;
                GroupBox_OverlayOptions.IsEnabled = false;
            }
        }

        private async void Button_StartExport_Click(object sender, RoutedEventArgs e)
        {
            Button_StartExport.IsEnabled = false;
            try
            {
                if (!Enum.IsDefined(typeof(ExportType), ComboBox_ExportType.SelectedItem))
                {
                    MessageBox.Show("Please select a valid export type");
                    Button_StartExport.IsEnabled = true;
                    return;
                }


                ExportType exportType = (ExportType)ComboBox_ExportType.SelectedItem;
                string filename = exportType == ExportType.NewFile ? TextBox_ExportFileName.Text.Trim() : Label_ExportFileName.Content.ToString().Trim();

                if (filename.Length <= 0)
                {
                    MessageBox.Show("Please enter valid filename");
                    Button_StartExport.IsEnabled = true;
                    return;
                }

                ImportResultsListItem importResults = (ImportResultsListItem)ComboBox_ExportResultSetTarget.SelectedItem;
                if (importResults is null)
                {
                    MessageBox.Show("Please select result set to export");
                    Button_StartExport.IsEnabled = true;
                    return;
                }

                int sheetIndex = -1;
                if ((ExcelWorksheetListItem)ComboBox_ExportSheetNames.SelectedItem is not null)
                {
                    sheetIndex = ((ExcelWorksheetListItem)ComboBox_ExportSheetNames.SelectedItem).SheetIndex;
                }

                AttributeExportMode attributeExportMode = Enum.IsDefined(typeof(AttributeExportMode), ComboBox_AttributeExportMode.SelectedItem) ?
                    (AttributeExportMode)ComboBox_AttributeExportMode.SelectedItem : AttributeExportMode.None;


                if (attributeExportMode != AttributeExportMode.None && ComboBox_AttributeExportTarget.SelectedItem is null)
                {
                    MessageBox.Show("Please select attribute export column target");
                    return;
                }

                string attributeExportTarget = ComboBox_AttributeExportTarget.SelectedItem.ToString();

                bool? trySave = CheckBox_TrySave.IsChecked;

                int progressMaxCount = ResultSetRepository.GetResultSetSize(_connectionManager, importResults.Id);

                ExcelExportItemsTask exportTask = new ExcelExportItemsTask(
                    _connectionManager,
                    exportType,
                    filename,
                    sheetIndex,
                    importResults.ProfileId,
                    importResults.Id,
                    attributeExportMode,
                    attributeExportTarget,
                    trySave is null || trySave == false? false : true
                );

                ProgressBar_ExportTaskStatus.Maximum = progressMaxCount;
                ProgressBar_ExportTaskStatus.Value = 0;

                IProgress<int> progress = new Progress<int>(value =>
                {
                    ProgressBar_ExportTaskStatus.Value = value;
                    Label_ExportStatus.Content = $"exporting row {value.ToString()} of {ProgressBar_ExportTaskStatus.Maximum.ToString()}";
                });

                await TaskManager.Launch(_connectionManager, exportTask, progress);
            } catch (Exception ex)
            {
                MessageBox.Show(ex.ToString(), "Export failed");
            }

            Button_StartExport.IsEnabled = true;
            Label_ExportStatus.Content = "Idle";
        }

        private void PopulateTaskViewDataGrid()
        {
            DataGrid_TaskView.ItemsSource = TaskRepository.GetTasks(_connectionManager, "");
        }

        private void SetupTaskViewDataGridColumns()
        {
            DataGrid_TaskView.Columns.Clear();
            DataGrid_TaskView.Columns.Add(new DataGridTextColumn
            {
                Header = "Task Id",
                Binding = new Binding("Id"),
            });
            DataGrid_TaskView.Columns.Add(new DataGridTextColumn
            {
                Header = "Type",
                Binding = new Binding("Type"),
            });
            DataGrid_TaskView.Columns.Add(new DataGridTextColumn
            {
                Header = "Start Time",
                Binding = new Binding("StartTime"),
            });

            DataGrid_TaskView.Columns.Add(new DataGridTextColumn
            {
                Header = "Update Time",
                Binding = new Binding("UpdateTime"),
            });
            DataGrid_TaskView.Columns.Add(new DataGridTextColumn
            {
                Header = "Status",
                Binding = new Binding("Status"),
            });
            DataGrid_TaskView.Columns.Add(new DataGridTextColumn
            {
                Header = "Data",
                Binding = new Binding("Data"),
            });
            
        }

        private void DataGrid_TaskView_MouseDoubleClick(object sender, MouseButtonEventArgs e)
        {
            AppTask task = (AppTask)DataGrid_TaskView.SelectedItem;
            if (task is not null)
            {
                if (_processingReportWindow == null || _processingReportWindow.IsLoaded == false)
                {
                    _processingReportWindow = new ProcessingReportWindow(_connectionManager, task.Id);
                    _processingReportWindow.Show();
                }
                else
                {
                    _processingReportWindow.Close();
                    _processingReportWindow = new ProcessingReportWindow(_connectionManager, task.Id);
                    _processingReportWindow.Show();
                }
            }
        }
    }
}
