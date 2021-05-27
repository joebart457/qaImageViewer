using qaImageViewer.Service;
using System;
using System.Collections.Generic;
using System.Data;
using System.Data.SQLite;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace qaImageViewer.Repository
{
    class ConfigRepository
    {
        public static bool GetBooleanOption(ConnectionManager cm, string param, bool onFail)
        {
            try
            {
                Utilities.CheckNull(cm);
                var conn = cm.GetSQLConnection();
                var selectParamCmd = conn.CreateCommand();

                selectParamCmd.CommandText = @"SELECT param, value FROM config WHERE param like @ParameterName";
                selectParamCmd.Parameters.Add(new SQLiteParameter("@ParameterName", "%"+param+"%"));

                bool result = onFail;
                var reader = selectParamCmd.ExecuteReader();
                while (reader.Read())
                {
                    result = bool.Parse(reader.GetString(1));
                }

                return result;
            }
            catch (Exception ex)
            {
                LoggerService.LogError(ex.ToString());
                return onFail;
            }
        }

        public static string GetStringOption(ConnectionManager cm, string param, string onFail)
        {
            try
            {
                Utilities.CheckNull(cm);
                var conn = cm.GetSQLConnection();
                var selectParamCmd = conn.CreateCommand();

                selectParamCmd.CommandText = @"SELECT param, value FROM config WHERE param like @ParameterName";
                selectParamCmd.Parameters.Add(new SQLiteParameter("@ParameterName", "%" + param + "%"));

                string result = onFail;
                var reader = selectParamCmd.ExecuteReader();
                while (reader.Read())
                {
                    result = reader.GetString(1);
                }

                return result;
            }
            catch (Exception ex)
            {
                LoggerService.LogError(ex.ToString());
                return onFail;
            }
        }

        public static int GetIntegerOption(ConnectionManager cm, string param, int onFail)
        {
            try
            {
                Utilities.CheckNull(cm);
                var conn = cm.GetSQLConnection();
                var selectParamCmd = conn.CreateCommand();

                selectParamCmd.CommandText = @"SELECT param, value FROM config WHERE param like @ParameterName";
                selectParamCmd.Parameters.Add(new SQLiteParameter("@ParameterName", "%" + param + "%"));

                int result = onFail;
                var reader = selectParamCmd.ExecuteReader();
                while (reader.Read())
                {
                    result = int.Parse(reader.GetString(1));
                }

                return result;
            }
            catch (Exception ex)
            {
                LoggerService.LogError(ex.ToString());
                return onFail;
            }
        }


        public static void SetParameterValue(ConnectionManager cm, string param, string value)
        {
            try
            {
                Utilities.CheckNull(cm);
                var conn = cm.GetSQLConnection();
                var updateParamCmd = conn.CreateCommand();

                updateParamCmd.CommandText = @"UPDATE config SET value=@Value WHERE param=@ParameterName";
                updateParamCmd.Parameters.Add(new SQLiteParameter("@Value", value));
                updateParamCmd.Parameters.Add(new SQLiteParameter("@ParameterName", param));

                updateParamCmd.ExecuteNonQuery();
            }
            catch (Exception ex)
            {
                LoggerService.LogError(ex.ToString());
                throw ex;
            }
        }

        public static void DeleteParameter(ConnectionManager cm, int id)
        {
            try
            {
                Utilities.CheckNull(cm);
                var conn = cm.GetSQLConnection();
                var deleteParamCmd = conn.CreateCommand();

                deleteParamCmd.CommandText = @"DELETE FROM config WHERE id=@Id";
                deleteParamCmd.Parameters.Add(new SQLiteParameter("@Id", id));

                deleteParamCmd.ExecuteNonQuery();
            }
            catch (Exception ex)
            {
                LoggerService.LogError(ex.ToString());
                throw ex;
            }
        }

        public static DataTable GetConfigurationAsDatatable(ConnectionManager cm)
        {
            try
            {
                Utilities.CheckNull(cm);

                DataTable configDataTable = new DataTable();
                string selectConfigSql = @"select id, param, value from config";

                using (SQLiteCommand command = new SQLiteCommand(selectConfigSql, cm.GetSQLConnection()))
                {
                    configDataTable = new DataTable();
                    configDataTable.Load(command.ExecuteReader());
                }
                return configDataTable;
            } catch (Exception ex)
            {
                LoggerService.LogError(ex.ToString());
                throw ex;
            }
        }

        public static void UpdateConfigurationByDatatable(ConnectionManager cm, DataTable configTable) 
        {
            try
            {
                Utilities.CheckNull(cm);
                using (SQLiteDataAdapter sQLiteDataAdapter = new SQLiteDataAdapter(@"select id, param, value from config", cm.GetSQLConnection()))
                {
                    SQLiteCommandBuilder commandBuilder = new SQLiteCommandBuilder(sQLiteDataAdapter);
                    sQLiteDataAdapter.Update(configTable);
                }
            } catch (Exception ex)
            {
                LoggerService.LogError(ex.ToString());
                throw ex;
            }
        }

    }
}
