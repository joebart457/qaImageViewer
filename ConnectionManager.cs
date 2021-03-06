using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Data.SQLite;
using System.IO;
using qaImageViewer.Service;

namespace qaImageViewer
{

    public class ConnectionManager
    {
        private string _dbFileName { get; } = "data.db";
        private SQLiteConnection _sqlConn { get; set; }

        private List<string> RELEASES = new List<string>
        {
            "1.0.0",
        };

        public ConnectionManager() { CreateConnection(); }
        ~ConnectionManager() { }

        public SQLiteConnection GetSQLConnection()
        {
            if (_sqlConn != null)
            {
                return _sqlConn;
            }
            throw new Exception("Fatal! db was unitialized when requested.");
        }
        private void CreateConnection()
        {
            if (File.Exists(_dbFileName))
            {
                try
                {
                    _sqlConn = new SQLiteConnection($"Data Source={_dbFileName};");
                    _sqlConn.Open();
                    var cmd = _sqlConn.CreateCommand();
                    cmd.CommandText = "select value from config where param='Db.Version'";
                    object version = cmd.ExecuteScalar();
                    if (version == null || !(version is string))
                    {
                        throw new Exception("Fatal! Invalid db data.");
                    } else
                    {
                        string versionStr = version.ToString();
                        int index = RELEASES.FindIndex(x => x == versionStr);
                        if (index == -1)
                        {
                            throw new Exception("Invalid db version. Please update manually. Or recreate db.");
                        } else
                        {
                            for (int i = index + 1; i < RELEASES.Count(); i++)
                            {
                                RunUpgrade(RELEASES.ElementAt(i));
                            }
                        }
                    }
                }
                catch (Exception ex)
                {
                    LoggerService.LogError(ex.ToString());
                    throw ex;
                }
            }
            else
            {
                LoggerService.Log("Existing db not found. Creating...");
                SetupNewDatabase();
                LoggerService.Log("Finished new db creation.");
            }
        }

        private void UpgradeDatabase() { }

        private void SetupNewDatabase()
        {
            try
            {
                _sqlConn = new SQLiteConnection($"Data Source={_dbFileName};");
                _sqlConn.Open();

                var createConfigTblCmd = _sqlConn.CreateCommand();
                createConfigTblCmd.CommandText = @"CREATE TABLE config (
                	id	INTEGER NOT NULL,
                	param	TEXT NOT NULL UNIQUE,
                	value	TEXT,
                	PRIMARY KEY(id AUTOINCREMENT)
                )";

                createConfigTblCmd.ExecuteNonQuery();

                var createMappingProfileTblCmd = _sqlConn.CreateCommand();
                createMappingProfileTblCmd.CommandText = @"CREATE TABLE mapping_profile (
                	id	INTEGER NOT NULL,
                	name	TEXT NOT NULL,
                    locked BOOLEAN NOT NULL DEFAULT FALSE,
                	PRIMARY KEY(id AUTOINCREMENT)
                );";

                createMappingProfileTblCmd.ExecuteNonQuery();

                var createImportColumnMappingTblCmd = _sqlConn.CreateCommand();
                createImportColumnMappingTblCmd.CommandText = @"CREATE TABLE import_column_mapping (
                	id	INTEGER NOT NULL,
                	column_alias	TEXT,
                    column_name TEXT,
                    excel_column_alias TEXT,
                    column_type INTEGER NOT NULL,
                    profile_id INTEGER NOT NULL,
                	PRIMARY KEY(id AUTOINCREMENT)
                );";

                createImportColumnMappingTblCmd.ExecuteNonQuery();


                var createExportColumnMappingTblCmd = _sqlConn.CreateCommand();
                createExportColumnMappingTblCmd.CommandText = @"CREATE TABLE export_column_mapping (
                	id	INTEGER NOT NULL,
                    profile_id INTEGER NOT NULL,
                	import_column_mapping_id	INTEGER NOT NULL,
                    excel_column_alias TEXT,
                    match BOOLEAN DEFAULT FALSE,
                	PRIMARY KEY(id AUTOINCREMENT),
                    FOREIGN KEY('import_column_mapping_id') REFERENCES import_column_mapping(id) ON DELETE CASCADE
                ); ";

                createExportColumnMappingTblCmd.ExecuteNonQuery();

                var createProcessingExceptionTblCmd = _sqlConn.CreateCommand();
                createProcessingExceptionTblCmd.CommandText = @"CREATE TABLE processing_exception (
                	id	INTEGER NOT NULL,
                    task_id INTEGER,
                    type TEXT,
                    result_set_id INTEGER,
                    row_index INTEGER,
                	error_trace	TEXT,
                    error_time INTEGER,
                	PRIMARY KEY(id AUTOINCREMENT)
                ); ";

                createProcessingExceptionTblCmd.ExecuteNonQuery();

                var createImportResultTblCmd = _sqlConn.CreateCommand();
                createImportResultTblCmd.CommandText = @"CREATE TABLE import_result (
                	id	INTEGER NOT NULL,
                    task_id INTEGER NOT NULL,
                    profile_id INTEGER NOT NULL,
                	table_name	TEXT NOT NULL,
                    workbook_name TEXT NOT NULL,
                    worksheet_name TEXT NOT NULL,
                    end_time INTEGER NOT NULL,
                	PRIMARY KEY(id AUTOINCREMENT)
                ); ";

                createImportResultTblCmd.ExecuteNonQuery();

                var createAttributeTblCmd = _sqlConn.CreateCommand();
                createAttributeTblCmd.CommandText = @"CREATE TABLE attribute (
                	id	INTEGER NOT NULL,
                	name	TEXT NOT NULL,
                	PRIMARY KEY(id AUTOINCREMENT)
                );";

                createAttributeTblCmd.ExecuteNonQuery();

                var createItemAttributeTblCmd = _sqlConn.CreateCommand();
                createItemAttributeTblCmd.CommandText = @"CREATE TABLE item_attribute (
                	id	INTEGER NOT NULL,
                    item_id INTEGER NOT NULL,
                    result_set_id NOT NULL,
                    attribute_id INTEGER NOT NULL,
                	PRIMARY KEY(id AUTOINCREMENT),
                    FOREIGN KEY('attribute_id') REFERENCES attribute(id) ON DELETE CASCADE
                );";

                createItemAttributeTblCmd.ExecuteNonQuery();


                var createAppTaskTblCmd = _sqlConn.CreateCommand();
                createAppTaskTblCmd.CommandText = @"CREATE TABLE app_task (
                	id	INTEGER NOT NULL,
                    type        TEXT NOT NULL,
                    start_time INTEGER NOT NULL,
                    update_time INTEGER NOT NULL,
                    data TEXT,
                	status INTEGER NOT NULL,
                	PRIMARY KEY(id AUTOINCREMENT)
                );";

                createAppTaskTblCmd.ExecuteNonQuery();

                SetApplicationSettings();
            }
            catch (Exception ex)
            {
                throw ex;
            }
        }

        private void SetApplicationSettings()
        {
            try
            {
                var insertAppParamsCmd = _sqlConn.CreateCommand();
                insertAppParamsCmd.CommandText = @"INSERT INTO config (param, value) 
                    VALUES  ('Db.Version', '1.0.0'),
                            ('Logger.DoLogDebug', 'false'),
                            ('Logger.DoLogWarning', 'true'),
                            ('Logger.DoLogError', 'true')";
                insertAppParamsCmd.ExecuteNonQuery();

            } catch (Exception ex)
            {
                LoggerService.LogError(ex.ToString());
                throw ex;
            }
        }

        private void RunUpgrade(string upgradeVersion)
        {
            LoggerService.Log("Running upgrade: " + upgradeVersion);
            if (upgradeVersion == "1.0.0")
            {
                // Initial Release
            }


        }
    }
}
