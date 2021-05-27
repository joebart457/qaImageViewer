using qaImageViewer.Models;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace qaImageViewer.Service
{
    class ColumnMappingService
    {
        static public List<ImportColumnMappingListItem> ConvertToListItems(List<ImportColumnMapping> mappings)
        {
            List<ImportColumnMappingListItem> listItems = new List<ImportColumnMappingListItem>();
            foreach(ImportColumnMapping mapping in mappings)
            {
                listItems.Add(new ImportColumnMappingListItem
                {
                    Id = mapping.Id,
                    ColumnAlias = mapping.ColumnAlias,
                    ColumnName = mapping.ColumnName,
                    ExcelColumnAlias = mapping.ExcelColumnAlias,
                    ColumnType = mapping.ColumnType,
                    ProfileId = mapping.ProfileId,
                    Changed = false
                }); ;
            }
            return listItems;
        }

        static public ImportColumnMapping ConvertFromListItem(ImportColumnMappingListItem listItem)
        {
            if (listItem is null) throw new Exception("failed to convert: listitem was null");
            return new ImportColumnMapping
            {
                Id = listItem.Id,
                ColumnAlias = listItem.ColumnAlias,
                ColumnName = listItem.ColumnName,
                ExcelColumnAlias = listItem.ExcelColumnAlias,
                ColumnType = listItem.ColumnType,
                ProfileId = listItem.ProfileId
            };
        }

        static public ExportColumnMapping ConvertFromListItem(ExportColumnMappingListItem listItem)
        {
            return new ExportColumnMapping
            {
                Id = listItem.Id,
                ExcelColumnAlias = listItem.ExcelColumnAlias,
                ProfileId = listItem.ProfileId,
                ImportColumnMappingId = listItem.ImportColumnMappingId
            };
        }

    }
}
