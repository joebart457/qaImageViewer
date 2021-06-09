using qaImageViewer.Models;
using qaImageViewer.Repository;
using qaImageViewer.Service;
using System;
using System.Collections.Generic;
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
using System.Windows.Shapes;

namespace qaImageViewer
{
    /// <summary>
    /// Interaction logic for ImageViewer.xaml
    /// </summary>
    public partial class ImageViewer : Window
    {
        private int _resultSetId { get; set; }
        private ConnectionManager _connectionManager { get; set; }
        public ImageViewer(ConnectionManager cm, int resultSetId)
        {
            InitializeComponent();
            _connectionManager = cm;
            _resultSetId = resultSetId;

            SetupPropertyiewDataGridColumns();
            PopulateItemSelectionListBox();
            PopulateFilePathPropertyComboBox();
            SetupImageRotationComboBox();
            PopulateAttributesEditListBox();
        }


        private void SetupImageRotationComboBox()
        {
            ComboBox_ImageRotation.ItemsSource = Enum.GetValues(typeof(Rotation));
            ComboBox_ImageRotation.SelectedItem = Rotation.Rotate0;
        }
        private void PopulateItemSelectionListBox()
        {
            ListBox_ItemSelection.ItemsSource =
                 ResultSetRepository.GetListItemsFromResultSet(_connectionManager, _resultSetId, new List<ColumnFilter>());
        }

        private void SetupPropertyiewDataGridColumns()
        {
            DataGrid_PropertyView.Columns.Clear();
            DataGrid_PropertyView.Columns.Add(new DataGridTextColumn
            {
                Header = "Param",
                Binding = new Binding("Mapping.ColumnAlias"),
                IsReadOnly = true
            });

            DataGrid_PropertyView.Columns.Add(new DataGridTextColumn
            {
                Header = "Value",
                Binding = new Binding("Value"),
                IsReadOnly = true
            });
        }

        private void PopulatePropertyViewDataGrid()
        {
            DocumentListItem selected = (DocumentListItem)ListBox_ItemSelection.SelectedItem;

            if (selected == null) { DataGrid_PropertyView.ItemsSource = null; return; }
            DataGrid_PropertyView.ItemsSource = ResultSetRepository.GetFullRowDataAsKeyValuePairs(_connectionManager, selected);
        }

        private void PopulateFilePathPropertyComboBox()
        {  
            MappingProfile profile = MappingProfileRepository.GetFullMappingProfileForResultSet(_connectionManager, _resultSetId);
            if (profile is not null && profile.ImportColumnMappings is not null)
            {
                ComboBox_FilePathProperty.ItemsSource = profile.ImportColumnMappings;
            }
        }

        private void ListBox_ItemSelection_SelectionChanged(object sender, SelectionChangedEventArgs e)
        {
            PopulatePropertyViewDataGrid();
            PopulateAttributesEditListBox();
            LoadMainImage();
        }

        private void LoadMainImage()
        {
            DocumentListItem selected = (DocumentListItem)ListBox_ItemSelection.SelectedItem;
            if (selected is not null)
            {
                ImportColumnMappingListItem filePathProperty = (ImportColumnMappingListItem)ComboBox_FilePathProperty.SelectedItem;
                if (filePathProperty is not null)
                {
                    Rotation rotation = Enum.IsDefined(typeof(Rotation), ComboBox_ImageRotation.SelectedItem) ? (Rotation)ComboBox_ImageRotation.SelectedItem : Rotation.Rotate0;
                    try
                    {
                        Image_ViewCapture.Source =
                            ImageHelperService.GetImageSourceFromItemProperties(
                                ResultSetRepository.GetFullRowDataAsKeyValuePairs(_connectionManager, selected),
                                ColumnMappingService.ConvertFromListItem(filePathProperty),
                                rotation
                            );
                    }
                    catch (Exception ex)
                    {
                        MessageBox.Show(ex.ToString());
                    }
                }

            }
        }

        private void ComboBox_ImageRotation_SelectionChanged(object sender, SelectionChangedEventArgs e)
        {
            LoadMainImage();
        }

        private void ComboBox_FilePathProperty_SelectionChanged(object sender, SelectionChangedEventArgs e)
        {
            LoadMainImage();
        }

        private void PopulateAttributesEditListBox()
        {
            DocumentListItem selected = (DocumentListItem)ListBox_ItemSelection.SelectedItem;
            int selectionId = -1;
            if (selected is not null) selectionId = selected.Id;
            List<AttributeListItem> attributes = AttributeRepository.GetAllAttributeListItems(_connectionManager, selectionId, _resultSetId);
            ListBox_AttributesEdit.ItemsSource = attributes;
        }

        private void Button_AddAttribute_Click(object sender, RoutedEventArgs e)
        {
            AddAttributeDialog addAttributeDialog = new AddAttributeDialog(_connectionManager);
            addAttributeDialog.ShowDialog();
            if (addAttributeDialog.DialogResult == true)
            {
                PopulateAttributesEditListBox();
            }
        }

        private void ListBox_AttributesEdit_SelectionChanged(object sender, SelectionChangedEventArgs e)
        {
            DocumentListItem selected = (DocumentListItem)ListBox_ItemSelection.SelectedItem;
            if (selected is not null) {
                try
                {
                    List<AttributeListItem> attributesToAdd = new List<AttributeListItem>();
                    foreach (AttributeListItem item in ListBox_AttributesEdit.SelectedItems)
                    {
                        attributesToAdd.Add(item);
                    }
                    AttributeRepository.SaveAttributeAssignments(_connectionManager, selected.Id, _resultSetId, attributesToAdd);
                } catch (Exception ex)
                {
                    LoggerService.LogError(ex.ToString());
                    MessageBox.Show(ex.ToString());
                }
           }
        }
    }
}
