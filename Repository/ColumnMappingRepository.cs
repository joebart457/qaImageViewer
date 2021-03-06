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
    class ColumnMappingRepository
    {
        public static List<ColumnMapping> GetColumnMappingsByProfileId(ConnectionManager cm, int profileId)
        {
            try
            {
                Utilities.CheckNull(cm);
                List<ColumnMapping> mappings = new List<ColumnMapping>();
                var conn = cm.GetSQLConnection();
                var selectColumnMappingsCmd = conn.CreateCommand();

                selectColumnMappingsCmd.CommandText = @"SELECT id, profile_id, column_name, column_alias, excel_column_alias, column_type FROM column_mapping WHERE profile_id = @ProfileId";
                selectColumnMappingsCmd.Parameters.Add(new SQLiteParameter("@ProfileId", profileId));

                var reader = selectColumnMappingsCmd.ExecuteReader();
                while (reader.Read())
                {
                    mappings.Add(new ColumnMapping
                    {
                        Id = reader.GetInt32(0),
                        ProfileId = reader.GetInt32(1),
                        ColumnName = reader.GetString(2),
                        ColumnAlias = reader.GetString(3),
                        ExcelColumnAlias = reader.GetString(4),
                        ColumnType = Enum.IsDefined(typeof(DBColumnType), reader.GetInt32(5)) ? (DBColumnType)reader.GetInt32(5) : DBColumnType.TEXT
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

        public static void InsertColumnMapping(ConnectionManager cm, ColumnMapping columnMapping)
        {
            try
            {
                Utilities.CheckNull(cm);
                var conn = cm.GetSQLConnection();
                var insertColumnMappingCmd = conn.CreateCommand();

                insertColumnMappingCmd.CommandText = @"INSERT INTO column_mapping (profile_id, column_name, column_alias, column_type, excel_column_alias) 
                                                      VALUES(@ProfileId, @ColumnName, @ColumnAlias, @ColumnType, @ExcelColumnAlias)";
                insertColumnMappingCmd.Parameters.Add(new SQLiteParameter("@ProfileId", columnMapping.ProfileId));
                insertColumnMappingCmd.Parameters.Add(new SQLiteParameter("@ColumnName", columnMapping.ColumnName));
                insertColumnMappingCmd.Parameters.Add(new SQLiteParameter("@ColumnAlias", columnMapping.ColumnAlias));
                insertColumnMappingCmd.Parameters.Add(new SQLiteParameter("@ColumnType", columnMapping.ColumnType));
                insertColumnMappingCmd.Parameters.Add(new SQLiteParameter("@ExcelColumnAlias", columnMapping.ExcelColumnAlias));

                insertColumnMappingCmd.ExecuteNonQuery();
            }
            catch (Exception ex)
            {
                LoggerService.LogError(ex.ToString());
                throw ex;
            }
        }

        public static void UpdateColumnMapping(ConnectionManager cm, ColumnMapping columnMapping)
        {
            try
            {
                Utilities.CheckNull(cm);
                var conn = cm.GetSQLConnection();
                var updateColumnMappingCmd = conn.CreateCommand();

                updateColumnMappingCmd.CommandText = @"UPDATE column_mapping SET 
                                                         profile_id = @ProfileId,
                                                         column_name = @ColumnName, 
                                                         column_alias = @ColumnAlias,
                                                         column_type = @ColumnType,
                                                         excel_column_alias = @ExcelColumnAlias
                                                      WHERE id = @Id";
                updateColumnMappingCmd.Parameters.Add(new SQLiteParameter("@ProfileId", columnMapping.ProfileId));
                updateColumnMappingCmd.Parameters.Add(new SQLiteParameter("@ColumnName", columnMapping.ColumnName));
                updateColumnMappingCmd.Parameters.Add(new SQLiteParameter("@ColumnAlias", columnMapping.ColumnAlias));
                updateColumnMappingCmd.Parameters.Add(new SQLiteParameter("@ColumnType", columnMapping.ColumnType));
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


        public static void DeleteColumnMapping(ConnectionManager cm, ColumnMapping columnMapping)
        {
            try
            {
                Utilities.CheckNull(cm);
                var conn = cm.GetSQLConnection();
                var deleteColumnMappingCmd = conn.CreateCommand();

                deleteColumnMappingCmd.CommandText = @"DELETE FROM column_mapping WHERE id = @Id";
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
