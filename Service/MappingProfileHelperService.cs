using qaImageViewer.Models;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace qaImageViewer.Service
{
    class MappingProfileHelperService
    {

        public static string GetNextColumnName(MappingProfile profile)
        {
            try
            {
                if (profile is not null)
                {
                    if (profile.ImportColumnMappings.Count == 0)
                    {
                        return "c1";
                    }
                    ImportColumnMapping mapping = ColumnMappingService.ConvertFromListItem(profile.ImportColumnMappings.OrderBy(x => x.ColumnName).LastOrDefault());
                    return $"c{GetNextColumnIndex(mapping.ColumnName)}";
                }
                throw new Exception("cannot get next column for malformed profile");
            } catch (Exception ex)
            {
                LoggerService.LogError(ex.ToString());
                throw new Exception("unable to retrieve next column for null mapping");
            }
        }

        private static int GetNextColumnIndex(string column)
        {
            if (column.StartsWith('c') && column.Length > 1)
            {
                return int.Parse(column.Substring(1)) + 1;
            }
            throw new Exception($"unrecognizable column name format {column}");
        }
    }
}
