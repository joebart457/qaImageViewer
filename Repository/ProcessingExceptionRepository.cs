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
    class ProcessingExceptionRepository
    {
        public static List<ProcessingExceptionListItem> GetProcessingExceptionListItemsByResultSetId(ConnectionManager cm, int resultSetId)
        {
            try
            {
                Utilities.CheckNull(cm);
                List<ProcessingExceptionListItem> mappings = new List<ProcessingExceptionListItem>();
                var conn = cm.GetSQLConnection();
                var selectProcessingExceptionsCmd = conn.CreateCommand();

                selectProcessingExceptionsCmd.CommandText = @"SELECT id, result_set_id, row_index, error_trace, datetime(error_time, 'unixepoch'), type
                                                        FROM processing_exception
                                                        WHERE result_set_id = @ResultSetID";
                selectProcessingExceptionsCmd.Parameters.Add(new SQLiteParameter("@ResultSetId", resultSetId));

                var reader = selectProcessingExceptionsCmd.ExecuteReader();
                while (reader.Read())
                {

                    mappings.Add(new ProcessingExceptionListItem
                    {
                        Id = reader.GetInt32(0),
                        ResultSetId = reader.GetInt32(1),
                        RowIndex = reader.GetInt32(2),
                        ErrorTrace = reader.GetString(3),
                        ErrorTime = reader.GetDateTime(4),
                        Type = reader.GetString(5)
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

        public static List<ProcessingExceptionListItem> GetProcessingExceptionListItemsByTaskId(ConnectionManager cm, int taskId)
        {
            try
            {
                Utilities.CheckNull(cm);
                List<ProcessingExceptionListItem> mappings = new List<ProcessingExceptionListItem>();
                var conn = cm.GetSQLConnection();
                var selectProcessingExceptionsCmd = conn.CreateCommand();

                selectProcessingExceptionsCmd.CommandText = @"SELECT id, result_set_id, row_index, error_trace, datetime(error_time, 'unixepoch'), type, task_id
                                                        FROM processing_exception
                                                        WHERE task_id = @TaskId";
                selectProcessingExceptionsCmd.Parameters.Add(new SQLiteParameter("@TaskId", taskId));

                var reader = selectProcessingExceptionsCmd.ExecuteReader();
                while (reader.Read())
                {

                    mappings.Add(new ProcessingExceptionListItem
                    {
                        Id = reader.GetInt32(0),
                        ResultSetId = reader.GetInt32(1),
                        RowIndex = reader.GetInt32(2),
                        ErrorTrace = reader.GetString(3),
                        ErrorTime = reader.GetDateTime(4),
                        Type = reader.GetString(5),
                        TaskId = reader.GetInt32(6)
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

        public static void InsertProcessingException(ConnectionManager cm, ProcessingExceptionListItem processingException)
        {
            try
            {
                Utilities.CheckNull(cm);
                var conn = cm.GetSQLConnection();
                var insertProcessingExceptionCmd = conn.CreateCommand();

                insertProcessingExceptionCmd.CommandText = @"INSERT INTO processing_exception 
                                                                (result_set_id, row_index, error_trace, error_time, type, task_id) 
                                                            VALUES(@ResultSetId, @RowIndex, @ErrorTrace, strftime('%s', 'now'), @Type, @TaskId)";
                insertProcessingExceptionCmd.Parameters.Add(new SQLiteParameter("@ResultSetId", processingException.ResultSetId));
                insertProcessingExceptionCmd.Parameters.Add(new SQLiteParameter("@RowIndex", processingException.RowIndex));
                insertProcessingExceptionCmd.Parameters.Add(new SQLiteParameter("@ErrorTrace", processingException.ErrorTrace));
                insertProcessingExceptionCmd.Parameters.Add(new SQLiteParameter("@Type", processingException.Type));
                insertProcessingExceptionCmd.Parameters.Add(new SQLiteParameter("@TaskId", processingException.TaskId));

                insertProcessingExceptionCmd.ExecuteNonQuery();
            }
            catch (Exception ex)
            {
                LoggerService.LogError(ex.ToString());
                throw ex;
            }
        }
    }
}
