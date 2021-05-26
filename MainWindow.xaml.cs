using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.ComponentModel;
using System.Linq;
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

using qaImageViewer.Models;
using qaImageViewer.Repository;
using qaImageViewer.Service;

namespace qaImageViewer
{
    class XTEST { 
        public string Val { get; set; }

    }

    class XTEST_PARENT
    {
        public XTEST xTest = new XTEST { Val = "G" };

        public override string ToString()
        {
            return xTest.Val;
        }
    }

    /// <summary>
    /// Interaction logic for MainWindow.xaml
    /// </summary>
    ///
    public partial class MainWindow : Window
    {
        private ConnectionManager _connectionManager = new ConnectionManager();
        public MainWindow()
        {
            InitializeComponent();
            SetupImportColumnMappingsViewColumns();
            SetupImportColumnMappingsEditColumns();
            PopulateImportProfilesComboBox();
        }

        private void PopulateImportProfilesComboBox()
        {
            ComboBox_ImportProfilesSelector.ItemsSource = MappingProfileRepository.GetMappingProfiles(_connectionManager);
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
            DataGrid_ImportColumnMappingsEdit.Columns.Add(new DataGridCheckBoxColumn { Header = "Changed", Binding = new Binding("Changed"), IsReadOnly = true, });

            DataGrid_ImportColumnMappingsEdit.Columns.Add(new DataGridTextColumn { Header = "Alias", Binding = new Binding("ColumnAlias") });
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
                        ((ColumnMapping)DataGrid_ImportColumnMappingsEdit.SelectedItem).ColumnType = columnType;
                    }
                })
            );
            DataGrid_ImportColumnMappingsEdit.Columns.Add(
                new DataGridTemplateColumn()
                {
                    Header = "Type",
                    CellTemplate = new DataTemplate() { VisualTree = columnTypeComboBoxTemplate }
                }
            );

            // Add Excel Column Mapping ComboBox
            var excelColumnAliasComboBoxTemplate = new FrameworkElementFactory(typeof(ComboBox));
            excelColumnAliasComboBoxTemplate.SetValue(ComboBox.ItemsSourceProperty, ExcelAppHelperService.GetExcelColumnOptionsAsList());
            excelColumnAliasComboBoxTemplate.SetBinding(ComboBox.SelectedItemProperty, new Binding("ExcelColumnAlias"));
            excelColumnAliasComboBoxTemplate.AddHandler(
                ComboBox.SelectionChangedEvent,
                new SelectionChangedEventHandler((o, e) => {
                    if (DataGrid_ImportColumnMappingsEdit.SelectedItem is not null)
                    {
                        ((ColumnMapping)DataGrid_ImportColumnMappingsEdit.SelectedItem).ExcelColumnAlias = e.AddedItems[0].ToString();
                    }
                })
            );
            DataGrid_ImportColumnMappingsEdit.Columns.Add(
                new DataGridTemplateColumn()
                {
                    Header = "Map to Excel Column",
                    CellTemplate = new DataTemplate() { VisualTree = excelColumnAliasComboBoxTemplate }
                }
            );

            // Add Save Button 
            var saveButtonTemplate = new FrameworkElementFactory(typeof(Button));
            saveButtonTemplate.SetValue(Button.ContentProperty, "Save");
            saveButtonTemplate.AddHandler(
                Button.ClickEvent,
                new RoutedEventHandler((o, e) => {
                    ColumnMappingRepository.UpdateColumnMapping(_connectionManager, (ColumnMapping)DataGrid_ImportColumnMappingsEdit.SelectedItem);
                    //PopulateImportColumnMappingsEditViewDataGrid((MappingProfile)ComboBox_ImportProfilesSelector.SelectedItem);
                })
            );
            DataGrid_ImportColumnMappingsEdit.Columns.Add(
                new DataGridTemplateColumn()
                {
                    Header = "",
                    CellTemplate = new DataTemplate() { VisualTree = saveButtonTemplate }
                }
            );

            // Add Delete Button 
            var deleteButtonTemplate = new FrameworkElementFactory(typeof(Button));
            deleteButtonTemplate.SetValue(Button.ContentProperty, "Delete");
            deleteButtonTemplate.AddHandler(
                Button.ClickEvent,
                new RoutedEventHandler((o, e) => {
                    ColumnMappingRepository.DeleteColumnMapping(_connectionManager, (ColumnMapping)DataGrid_ImportColumnMappingsEdit.SelectedItem);
                    PopulateImportColumnMappingsEditViewDataGrid((MappingProfile)ComboBox_ImportProfilesSelector.SelectedItem);
                })
            );
            DataGrid_ImportColumnMappingsEdit.Columns.Add(
                new DataGridTemplateColumn()
                {
                    Header = "",
                    CellTemplate = new DataTemplate() { VisualTree = deleteButtonTemplate }
                }
            );


        }

        private void Button_SaveImportProfile_Click(object sender, RoutedEventArgs e)
        {
            try
            {
                MappingProfile profile = (MappingProfile)ComboBox_ImportProfilesSelector.SelectedItem;
                if (profile is null)
                {
                    MappingProfile newProfile;
                    profile = new MappingProfile { Name = ComboBox_ImportProfilesSelector.Text, Locked = false };
                    MappingProfileRepository.InsertMappingProfile(_connectionManager, profile, out newProfile);
                    PopulateImportProfilesComboBox();
                    ComboBox_ImportProfilesSelector.Text = "<profile name>";
                } 
                
                
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
                }
            }
        }


        private void PopulateImportColumnMappingsEditViewDataGrid(MappingProfile profile)
        {
            try
            {
                MappingProfile fullProfile = MappingProfileRepository.GetFullMappingProfileById(_connectionManager, profile.Id);
                if (fullProfile is not null && fullProfile.ImportMapping is not null)
                {
                    if (fullProfile.Locked)
                    {
                        // If it is locked, just view
                        DataGrid_ImportColumnMappingsView.ItemsSource = fullProfile.ImportMapping.ColumnMappings;
                        DataGrid_ImportColumnMappingsEdit.Visibility = Visibility.Hidden;
                        DataGrid_ImportColumnMappingsView.Visibility = Visibility.Visible;
                    }
                    else
                    {
                        // If it's not locked we can edit
                        DataGrid_ImportColumnMappingsEdit.ItemsSource = fullProfile.ImportMapping.ColumnMappings;
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
                    ColumnMapping newMapping = new ColumnMapping
                    {
                        ColumnAlias = "alias",
                        ColumnType = DBColumnType.TEXT,
                        ExcelColumnAlias = "A",
                        ProfileId = profile.Id,
                        ColumnName = MappingProfileHelperService.GetNextColumnName(profile)
                    };

                    ColumnMappingRepository.InsertColumnMapping(_connectionManager, newMapping);
                    PopulateImportColumnMappingsEditViewDataGrid(profile);
                }
            }
        }
    }
}
