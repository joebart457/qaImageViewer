using qaImageViewer.Models;
using qaImageViewer.Repository;
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
    /// Interaction logic for ProcessingReportWindow.xaml
    /// </summary>
    public partial class ProcessingReportWindow : Window
    {

        private int _resultSetId { get; set; }
        private int _taskId { get; set; }
        private ConnectionManager _connectionManager { get; set; }
        public ProcessingReportWindow(ConnectionManager cm, int resultSetId, int taskId)
        {
            InitializeComponent();
            _connectionManager = cm;
            _resultSetId = resultSetId;
            _taskId = taskId;

            SetupProcessingExceptionDataGridColumns();
            SetupRowDataViewDataGridColumns();
            PopulateProcessingExceptionsDataGrid();
            PopulateRowSelectListBox();
        }


        private void SetupProcessingExceptionDataGridColumns()
        {
            DataGrid_ProcessingExceptions.Columns.Clear();
            DataGrid_ProcessingExceptions.Columns.Add(new DataGridTextColumn
            {
                Header = "Row Index",
                Binding = new Binding("RowIndex"),
                IsReadOnly = true
            });

            DataGrid_ProcessingExceptions.Columns.Add(new DataGridTextColumn
            {
                Header = "Error Trace",
                Binding = new Binding("ErrorTrace"),
                IsReadOnly = true
            });

            DataGrid_ProcessingExceptions.Columns.Add(new DataGridTextColumn
            {
                Header = "Error Time",
                Binding = new Binding("ErrorTime"),
                IsReadOnly = true
            });
        }

        private void PopulateProcessingExceptionsDataGrid()
        {
            DataGrid_ProcessingExceptions.ItemsSource = 
                ProcessingExceptionRepository.GetProcessingExceptionListItemsByTaskId(_connectionManager, _taskId);
        }

        private void PopulateRowSelectListBox()
        {
            ListBox_RowSelect.ItemsSource =
                 ResultSetRepository.GetListItemsFromResultSet(_connectionManager, _resultSetId, new List<ColumnFilter>());
        }
        
        private void SetupRowDataViewDataGridColumns()
        {
            DataGrid_RowDataView.Columns.Clear();
            DataGrid_RowDataView.Columns.Add(new DataGridTextColumn
            {
                Header = "Param",
                Binding = new Binding("Mapping.ColumnAlias"),
                IsReadOnly = true
            });

            DataGrid_RowDataView.Columns.Add(new DataGridTextColumn
            {
                Header = "Value",
                Binding = new Binding("Value"),
                IsReadOnly = true
            });

            DataGrid_RowDataView.Columns.Add(new DataGridTextColumn
            {
                Header = "Type",
                Binding = new Binding("Mapping.ColumnType"),
                IsReadOnly = true
            });
        }

        private void PopulateRowDataViewDataGrid(DocumentListItem doc)
        {
            if (doc == null) DataGrid_RowDataView.ItemsSource = null;
            DataGrid_RowDataView.ItemsSource = ResultSetRepository.GetFullRowDataAsKeyValuePairs(_connectionManager, doc);
        }

        private void ListBox_RowSelect_SelectionChanged(object sender, SelectionChangedEventArgs e)
        {
            DocumentListItem selected = (DocumentListItem)ListBox_RowSelect.SelectedItem;
            PopulateRowDataViewDataGrid(selected);
        }
    }
}
