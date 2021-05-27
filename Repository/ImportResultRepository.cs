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
    class ImportResultRepository
    {
        public static ImportResults GetImportResult(ConnectionManager cm, int id)
        {
            try
            {
                Utilities.CheckNull(cm);

                List<ImportResults> results = new List<ImportResults>();
                var conn = cm.GetSQLConnection();
                var selectImportResultCmd = conn.CreateCommand();

                selectImportResultCmd.CommandText = @"SELECT id, profile_id, table_name, workbook_name, worksheet_name, datetime(end_time, 'unixepoch') FROM import_result WHERE id = @Id";
                selectImportResultCmd.Parameters.Add(new SQLiteParameter("@Id", id));

                var reader = selectImportResultCmd.ExecuteReader();
                while (reader.Read())
                {
                    results.Add(new ImportResults
                    {
                        Id = reader.GetInt32(0),
                        ProfileId = reader.GetInt32(1),
                        ResultTableName = reader.GetString(2),
                        WorkbookName = reader.GetString(3),
                        WorksheetName = reader.GetString(4),
                        EndTime = reader.GetDateTime(5),
                    });
                }

                return results.FirstOrDefault();
            }
            catch (Exception ex)
            {
                LoggerService.LogError(ex.ToString());
                throw ex;
            }
        }


        public static List<ImportResultsListItem> GetImportResultListItems(ConnectionManager cm)
        {
            try
            {
                Utilities.CheckNull(cm);

                List<ImportResultsListItem> results = new List<ImportResultsListItem>();
                var conn = cm.GetSQLConnection();
                var selectImportResultCmd = conn.CreateCommand();

                selectImportResultCmd.CommandText = @"SELECT ir.id, ir.profile_id, ir.table_name, ir.workbook_name,
                                                        ir.worksheet_name, datetime(end_time, 'unixepoch'), mp.name
                                                    FROM import_result ir, mapping_profile mp WHERE mp.id = ir.profile_id";

                var reader = selectImportResultCmd.ExecuteReader();
                while (reader.Read())
                {
                    results.Add(new ImportResultsListItem
                    {
                        Id = reader.GetInt32(0),
                        ProfileId = reader.GetInt32(1),
                        ResultTableName = reader.GetString(2),
                        WorkbookName = reader.GetString(3),
                        WorksheetName = reader.GetString(4),
                        EndTime = reader.GetDateTime(5),
                        ProfileName = reader.GetString(6)
                    });
                }

                return results;
            }
            catch (Exception ex)
            {
                LoggerService.LogError(ex.ToString());
                throw ex;
            }
        }

        public static void InsertImportResult(ConnectionManager cm, ImportResults importResults)
        {
            try
            {
                Utilities.CheckNull(cm);
                var conn = cm.GetSQLConnection();
                var insertImportResultCmd = conn.CreateCommand();

                insertImportResultCmd.CommandText = @"INSERT INTO import_result (profile_id, table_name, workbook_name, worksheet_name, end_time) 
                                                    VALUES (@ProfileId, @TableName, @WorkbookName, @WorksheetName, strftime('%s', 'now'))";
                insertImportResultCmd.Parameters.Add(new SQLiteParameter("@ProfileId", importResults.ProfileId));
                insertImportResultCmd.Parameters.Add(new SQLiteParameter("@TableName", importResults.ResultTableName));
                insertImportResultCmd.Parameters.Add(new SQLiteParameter("@WorkbookName", importResults.WorkbookName));
                insertImportResultCmd.Parameters.Add(new SQLiteParameter("@WorksheetName", importResults.WorksheetName));

                insertImportResultCmd.ExecuteNonQuery();
            }
            catch(Exception ex)
            {
                LoggerService.LogError(ex.ToString());
                throw ex;
            }
        }
    }
}
