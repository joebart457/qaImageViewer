using qaImageViewer.Models;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Media;
using System.Windows.Media.Imaging;

namespace qaImageViewer.Service
{
    class ImageHelperService
    {

        public static BitmapImage GetImageSourceFromItemProperties(List<DocumentColumn> properties, ImportColumnMapping mapping, Rotation rotation)
        {
            if (properties is null) return null;
            DocumentColumn col = properties.Find(x => x.Mapping.ColumnName == mapping.ColumnName);
            if (col is null || col.Value is null) return null;
            if (!System.IO.File.Exists(col.Value.ToString())) return null;
            BitmapImage bmp = new BitmapImage();
            bmp.BeginInit();
            bmp.UriSource = new Uri(col.Value.ToString());
            switch (rotation) {
                case Rotation.Rotate90:
                    bmp.Rotation = Rotation.Rotate90;
                    break;
                case Rotation.Rotate180:
                    bmp.Rotation = Rotation.Rotate180;
                    break;
                case Rotation.Rotate270:
                    bmp.Rotation = Rotation.Rotate270;
                    break;
                default:
                    break;
            }
            bmp.EndInit();
            return bmp;
        } 
    }
}
