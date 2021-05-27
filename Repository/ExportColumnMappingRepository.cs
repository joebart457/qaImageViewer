using qaImageViewer.Models;
using qaImageViewer.Service;
using System;
using System.Collections.Generic;
using System.Data.SQLite;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace qaImageViewer.Repository
{
    class ExportColumnMappingRepository
    {
        public static List<ExportColumnMapping> GetColumnMappingsByProfileId(ConnectionManager cm, int profileId)
        {
            try
            {
                Utilities.CheckNull(cm);
                List<ExportColumnMapping> mappings = new List<ExportColumnMapping>();
                var conn = cm.GetSQLConnection();
                var selectColumnMappingsCmd = conn.CreateCommand();

                selectColumnMappingsCmd.CommandText = @"SELECT id, profile_id, import_column_mapping_id, excel_column_alias FROM export_column_mapping WHERE profile_id = @ProfileId";
                selectColumnMappingsCmd.Parameters.Add(new SQLiteParameter("@ProfileId", profileId));

                var reader = selectColumnMappingsCmd.ExecuteReader();
                while (reader.Read())
                {
                    mappings.Add(new ExportColumnMapping
                    {
                        Id = reader.GetInt32(0),
                        ProfileId = reader.GetInt32(1),
                        ImportColumnMappingId = reader.GetInt32(2),
                        ExcelColumnAlias = reader.GetString(3),
                    });
                }

                return mappings;
            }
            catch (Exception ex)
            {
                LoggerService.LogError(ex.ToString());
                throw ex;
            }
        }

        public static List<ExportColumnMappingListItem> GetColumnMappingListItemsByProfileId(ConnectionManager cm, int profileId)
        {
            try
            {
                Utilities.CheckNull(cm);
                List<ExportColumnMappingListItem> mappings = new List<ExportColumnMappingListItem>();
                var conn = cm.GetSQLConnection();
                var selectColumnMappingsCmd = conn.CreateCommand();

                selectColumnMappingsCmd.CommandText = @"SELECT ecm.id, ecm.profile_id, ecm.import_column_mapping_id, ecm.excel_column_alias, icm.column_alias 
                                                        FROM export_column_mapping ecm, import_column_mapping icm 
                                                        WHERE ecm.profile_id = @ProfileId AND ecm.import_column_mapping_id = icm.id";
                selectColumnMappingsCmd.Parameters.Add(new SQLiteParameter("@ProfileId", profileId));

                var reader = selectColumnMappingsCmd.ExecuteReader();
                while (reader.Read())
                {
                    mappings.Add(new ExportColumnMappingListItem
                    {
                        Id = reader.GetInt32(0),
                        ProfileId = reader.GetInt32(1),
                        ImportColumnMappingId = reader.GetInt32(2),
                        ExcelColumnAlias = reader.GetString(3),
                        ImportColumnMappingAlias = reader.GetString(4),
                    });
                }

                return mappings;
            }
            catch (Exception ex)
            {
                LoggerService.LogError(ex.ToString());
                throw ex;
            }
        }

        public static void InsertColumnMapping(ConnectionManager cm, ExportColumnMapping columnMapping)
        {
            try
            {
                Utilities.CheckNull(cm);
                var conn = cm.GetSQLConnection();
                var insertColumnMappingCmd = conn.CreateCommand();

                insertColumnMappingCmd.CommandText = @"INSERT INTO export_column_mapping (profile_id, import_column_mapping_id, excel_column_alias) 
                                                      VALUES(@ProfileId, @ImportColumnMappingId, @ExcelColumnAlias)";
                insertColumnMappingCmd.Parameters.Add(new SQLiteParameter("@ProfileId", columnMapping.ProfileId));
                insertColumnMappingCmd.Parameters.Add(new SQLiteParameter("@ImportColumnMappingId", columnMapping.ImportColumnMappingId));
                insertColumnMappingCmd.Parameters.Add(new SQLiteParameter("@ExcelColumnAlias", columnMapping.ExcelColumnAlias));

                insertColumnMappingCmd.ExecuteNonQuery();
            }
            catch (Exception ex)
            {
                LoggerService.LogError(ex.ToString());
                throw ex;
            }
        }

        public static void UpdateColumnMapping(ConnectionManager cm, ExportColumnMapping columnMapping)
        {
            try
            {
                Utilities.CheckNull(cm);
                var conn = cm.GetSQLConnection();
                var updateColumnMappingCmd = conn.CreateCommand();

                updateColumnMappingCmd.CommandText = @"UPDATE export_column_mapping SET 
                                                         profile_id = @ProfileId,
                                                         import_column_mapping_id = @ImportColumnMappingId, 
                                                         excel_column_alias = @ExcelColumnAlias
                                                      WHERE id = @Id";
                updateColumnMappingCmd.Parameters.Add(new SQLiteParameter("@ProfileId", columnMapping.ProfileId));
                updateColumnMappingCmd.Parameters.Add(new SQLiteParameter("@ImportColumnMappingId", columnMapping.ImportColumnMappingId));
                updateColumnMappingCmd.Parameters.Add(new SQLiteParameter("@ExcelColumnAlias", columnMapping.ExcelColumnAlias));
                updateColumnMappingCmd.Parameters.Add(new SQLiteParameter("@Id", columnMapping.Id));

                updateColumnMappingCmd.ExecuteNonQuery();
            }
            catch (Exception ex)
            {
                LoggerService.LogError(ex.ToString());
                throw ex;
            }
        }


        public static void DeleteColumnMapping(ConnectionManager cm, ExportColumnMapping columnMapping)
        {
            try
            {
                Utilities.CheckNull(cm);
                var conn = cm.GetSQLConnection();
                var deleteColumnMappingCmd = conn.CreateCommand();

                deleteColumnMappingCmd.CommandText = @"DELETE FROM export_column_mapping WHERE id = @Id";
                deleteColumnMappingCmd.Parameters.Add(new SQLiteParameter("@Id", columnMapping.Id));

                deleteColumnMappingCmd.ExecuteNonQuery();
            }
            catch (Exception ex)
            {
                LoggerService.LogError(ex.ToString());
                throw ex;
            }
        }
    }
}
