using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace qaImageViewer.Models
{
    class ImportColumnMappingListItem : INotifyPropertyChanged
    {

        public event PropertyChangedEventHandler PropertyChanged;

        protected virtual void OnPropertyChanged(string propertyName = null)
        {
            var handler = PropertyChanged;
            if (handler != null) handler(this, new PropertyChangedEventArgs(propertyName));
        }


        private int _id { get; set; }
        private int _profileId { get; set; }
        private string _columnName { get; set; }
        private string _columnAlias { get; set; }
        private string _excelColumnAlias { get; set; }
        private DBColumnType _columnType { get; set; }
        private bool _changed { get; set; }
        
        public int Id { 
            get { return _id; }
            set { _id = value; Changed = true; OnPropertyChanged("Id"); }
        }
        public int ProfileId
        {
            get { return _profileId; }
            set { _profileId = value; Changed = true; OnPropertyChanged("ProfileId"); }
        }
        public string ColumnName
        {
            get { return _columnName; }
            set { _columnName = value; Changed = true; OnPropertyChanged("ColumnName"); }
        }
        public string ColumnAlias
        {
            get { return _columnAlias; }
            set { _columnAlias = value; Changed = true; OnPropertyChanged("ColumnAlias"); }
        }
        public string ExcelColumnAlias
        {
            get { return _excelColumnAlias; }
            set { _excelColumnAlias = value; Changed = true; OnPropertyChanged("ExcelColumnAlias"); }
        }
        public DBColumnType ColumnType
        {
            get { return _columnType; }
            set { _columnType = value; Changed = true; OnPropertyChanged("ColumnType"); }
        }

        public bool Changed
        {
            get { return _changed; }
            set { _changed = value; OnPropertyChanged("Changed"); }
        }

        public override string ToString()
        {
            return ColumnAlias;
        }
    }
}
