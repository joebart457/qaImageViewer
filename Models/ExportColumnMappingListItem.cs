using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace qaImageViewer.Models
{
    class ExportColumnMappingListItem: INotifyPropertyChanged
    {
        public event PropertyChangedEventHandler PropertyChanged;

        protected virtual void OnPropertyChanged(string propertyName = null)
        {
            var handler = PropertyChanged;
            if (handler != null) handler(this, new PropertyChangedEventArgs(propertyName));
        }
        private int _id { get; set; }
        private int _profileId { get; set; }
        private int _importColumnMappingId { get; set; }
        private string _excelColumnAlias { get; set; }
        private string _importColumnMappingAlias { get; set; }
        private bool _match { get; set; }
        public int Id
        {
            get { return _id; }
            set { _id = value; OnPropertyChanged("Id"); }
        }
        public int ProfileId
        {
            get { return _profileId; }
            set { _profileId = value; OnPropertyChanged("ProfileId"); }
        }
        public int ImportColumnMappingId
        {
            get { return _importColumnMappingId; }
            set { _importColumnMappingId = value; OnPropertyChanged("ImportColumnMappingId"); }
        }
        public string ExcelColumnAlias
        {
            get { return _excelColumnAlias; }
            set { _excelColumnAlias = value; OnPropertyChanged("ExcelColumnAlias"); }
        }

        public string ImportColumnMappingAlias
        {
            get { return _importColumnMappingAlias; }
            set { _importColumnMappingAlias = value; OnPropertyChanged("ImportColumnMappingAlias"); }
        }
        public bool Match
        {
            get { return _match; }
            set { _match = value; OnPropertyChanged("Match"); }
        }
    }
}
