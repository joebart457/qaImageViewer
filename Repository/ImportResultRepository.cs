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

                selectImportResultCmd.CommandText = @"SELECT id, profile_id, table_name, end_time FROM import_result WHERE id= @Id";
                selectImportResultCmd.Parameters.Add(new SQLiteParameter("@Id", id));

                var reader = selectImportResultCmd.ExecuteReader();
                while (reader.Read())
                {
                    results.Add(new ImportResults
                    {
                        Id = reader.GetInt32(0),
                        ProfileId = reader.GetInt32(1),
                        ResultTableName = reader.GetString(2),
                        EndTime = new DateTime(reader.GetInt64(3)),
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

        public static void InsertImportResult(ConnectionManager cm, ImportResults importResults)
        {
            try
            {
                Utilities.CheckNull(cm);
                var conn = cm.GetSQLConnection();
                var insertImportResultCmd = conn.CreateCommand();

                insertImportResultCmd.CommandText = @"INSERT INTO import_result (profile_id, table_name, end_time) 
                                                    VALUES (@ProfileId, @TableName, now())";
                insertImportResultCmd.Parameters.Add(new SQLiteParameter("@ProfileId", importResults.ProfileId));
                insertImportResultCmd.Parameters.Add(new SQLiteParameter("@TableName", importResults.ResultTableName));

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
