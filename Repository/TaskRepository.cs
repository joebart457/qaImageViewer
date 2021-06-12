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
    class TaskRepository
    {
        public static List<AppTask> GetTasks(ConnectionManager cm, string typeFilter)
        {
            try
            {
                Utilities.CheckNull(cm);
                List<AppTask> mappings = new List<AppTask>();
                var conn = cm.GetSQLConnection();
                var selectTasksCmd = conn.CreateCommand();

                selectTasksCmd.CommandText = @"SELECT id, type, datetime(start_time, 'unixepoch'), datetime(update_time, 'unixepoch'), data, status
                                                        FROM app_task
                                                        WHERE type like @TypeFilter";
                selectTasksCmd.Parameters.Add(new SQLiteParameter("@TypeFilter", $"{typeFilter}%"));

                var reader = selectTasksCmd.ExecuteReader();
                while (reader.Read())
                {

                    mappings.Add(new AppTask
                    {
                        Id = reader.GetInt32(0),
                        Type = reader.GetString(1),
                        StartTime = reader.GetDateTime(2),
                        UpdateTime = reader.GetDateTime(3),
                        Data = reader.GetString(4),
                        Status = Enum.IsDefined(typeof(AppTaskStatus), reader.GetInt32(5)) ? (AppTaskStatus)reader.GetInt32(5) : AppTaskStatus.None
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

        public static int InsertTask(ConnectionManager cm, AppTask task)
        {
            try
            {
                Utilities.CheckNull(cm);
                var conn = cm.GetSQLConnection();
                var insertTaskCmd = conn.CreateCommand();

                insertTaskCmd.CommandText = @"INSERT INTO app_task 
                                                                (type, start_time, update_time, data, status) 
                                                            VALUES(@Type, strftime('%s', 'now'), strftime('%s', 'now'), @Data, @Status)";
                insertTaskCmd.Parameters.Add(new SQLiteParameter("@Type", task.Type));
                insertTaskCmd.Parameters.Add(new SQLiteParameter("@Data", task.Data));
                insertTaskCmd.Parameters.Add(new SQLiteParameter("@Status", task.Status));

                insertTaskCmd.ExecuteNonQuery();
                return (int)conn.LastInsertRowId;
            }
            catch (Exception ex)
            {
                LoggerService.LogError(ex.ToString());
                throw ex;
            }
        }


        public static void UpdateTask(ConnectionManager cm, AppTask task)
        {
            try
            {
                Utilities.CheckNull(cm);
                var conn = cm.GetSQLConnection();
                var insertTaskCmd = conn.CreateCommand();

                insertTaskCmd.CommandText = @"UPDATE app_task SET
                                                                type = @Type,
                                                                update_time = strftime('%s', 'now'),
                                                                data = @Data,
                                                                status = @Status
                                                           WHERE id = @Id";
                insertTaskCmd.Parameters.Add(new SQLiteParameter("@Id", task.Id));
                insertTaskCmd.Parameters.Add(new SQLiteParameter("@Type", task.Type));
                insertTaskCmd.Parameters.Add(new SQLiteParameter("@Data", task.Data));
                insertTaskCmd.Parameters.Add(new SQLiteParameter("@Status", task.Status));

                insertTaskCmd.ExecuteNonQuery();
            }
            catch (Exception ex)
            {
                LoggerService.LogError(ex.ToString());
                throw ex;
            }
        }
    }
}
