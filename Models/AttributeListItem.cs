using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace qaImageViewer.Models
{
    class AttributeListItem: INotifyPropertyChanged
    {
        public event PropertyChangedEventHandler PropertyChanged;

        protected virtual void OnPropertyChanged(string propertyName = null)
        {
            var handler = PropertyChanged;
            if (handler != null) handler(this, new PropertyChangedEventArgs(propertyName));
        }
        public int Id { get; set; }
        public int ResultSetId { get; set; }
        public string Name { get; set; }
        private bool _isAssigned { get; set; }
        public bool IsAssigned
        {
            get { return _isAssigned; }
            set { _isAssigned = value; OnPropertyChanged("IsAssigned"); }
        }

        public override string ToString()
        {
            return Name;
        }
    }
}
