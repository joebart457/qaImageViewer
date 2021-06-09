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
    /// Interaction logic for AddAttributeDialog.xaml
    /// </summary>
    public partial class AddAttributeDialog : Window
    {
        private ConnectionManager _connectionManager = null;
        public AddAttributeDialog(ConnectionManager connectionManager)
        {
            InitializeComponent();
            _connectionManager = connectionManager;
        }

        private void Button_Cancel_Click(object sender, RoutedEventArgs e)
        {
            this.Close();
        }

        private void Button_Add_Click(object sender, RoutedEventArgs e)
        {
            string attributeName = TextBox_AttributeName.Text.Trim();
            if (attributeName.Length > 0)
            {
                try
                {
                    AttributeRepository.InsertAttribute(_connectionManager, new Models.AttributeListItem { Name = attributeName });
                }
                catch (Exception ex)
                {
                    MessageBox.Show(ex.ToString(), "Unable to save attribute");
                }
                this.DialogResult = true;
                this.Close();
            }
        }

        private void TextBox_AttributeName_TextChanged(object sender, TextChangedEventArgs e)
        {
            if (Button_Add is not null) // Necessary for startup
            {
                if (TextBox_AttributeName.Text.Trim().Length > 0)
                {
                    Button_Add.IsEnabled = true;
                }
                else
                {
                    Button_Add.IsEnabled = false;
                }
            }
        }
    }
}
