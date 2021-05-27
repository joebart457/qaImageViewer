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
    class ResultSetRepository
    {
        private static string GetNextResultTableName(ConnectionManager cm)
        {
            // Needs refactor
            Utilities.CheckNull(cm);
            var conn = cm.GetSQLConnection();
            var getSeqCmd = conn.CreateCommand();


            getSeqCmd.CommandText = @"select seq from sqlite_sequence where name = 'import_result'";
            var result = getSeqCmd.ExecuteScalar();
            return $"result_set_{Convert.ToInt32(result) + 1}";
        }

        private static string GetTypeString(DBColumnType columnType)
        {
            switch (columnType)
            {
                case DBColumnType.TEXT:
                    return "TEXT";
                case DBColumnType.BOOLEAN:
                    return "BOOLEAN";
                case DBColumnType.INTEGER:
                    return "INTEGER";
                case DBColumnType.DATE:
                    return "INTEGER";
                case DBColumnType.REAL:
                    return "REAL";
                default:
                    return "TEXT";
            }
        }

        private static string CreateColumnDefinitions(List<ImportColumnMappingListItem> mappings)
        {
            string result = "";
            foreach(ImportColumnMappingListItem map in mappings)
            {
                result += $"{map.ColumnName} {GetTypeString(map.ColumnType)}, ";
            }
            return result;
        }

        public static int CreateResultSet(ConnectionManager cm, MappingProfile mapping, string workbookName, string worksheetName)
        {
            try
            { 
                Utilities.CheckNull(cm);
                var conn = cm.GetSQLConnection();
                var createResultSetCmd = conn.CreateCommand();

                string nextResultTableName = GetNextResultTableName(cm);

                // needs refactor
                ImportResultRepository.InsertImportResult(cm, new ImportResults
                {
                    ResultTableName = nextResultTableName,
                    ProfileId = mapping.Id,
                    WorkbookName = workbookName,
                    WorksheetName = worksheetName
                });


                createResultSetCmd.CommandText = $"CREATE TABLE {nextResultTableName}" 
                    + "(id	INTEGER NOT NULL,"
                	+ CreateColumnDefinitions(mapping.ImportColumnMappings)
                	+ "PRIMARY KEY(id AUTOINCREMENT));";
                createResultSetCmd.ExecuteNonQuery();
                int resultSetId = Convert.ToInt32(conn.LastInsertRowId);

                // Lock mapping profile so more columns cannot be added
                MappingProfile mappingProfile = MappingProfileRepository.GetMappingProfileById(cm, mapping.Id);
                mappingProfile.Locked = true;
                MappingProfileRepository.UpdateMappingProfile(cm, mappingProfile);

                return resultSetId;

            } catch (Exception ex)
            {
                LoggerService.LogError(ex.ToString());
                throw ex;
            }
        }

        private static string GetValueString(DBColumnType columnType, object value)
        {
            switch (columnType) {
                case DBColumnType.INTEGER:
                case DBColumnType.REAL:
                case DBColumnType.BOOLEAN:
                    return value.ToString();
                case DBColumnType.TEXT:
                    return $"\'{value.ToString()}'";
                default:
                    return $"\'{value.ToString()}'";
            }
        }

        public static void InsertIntoResultSet(ConnectionManager cm, int importResultId, List<Document> documents)
        {
            try
            {
                Utilities.CheckNull(cm);

                var conn = cm.GetSQLConnection();


                ImportResults res = ImportResultRepository.GetImportResult(cm, importResultId);
                if (res is null)
                {
                    throw new Exception($"unable to find result set with id {importResultId}");
                }

                string resultTableName = res.ResultTableName;

                MappingProfile profile = MappingProfileRepository.GetFullMappingProfileById(cm, res.ProfileId);


                var insertIntoResultSetCmd = conn.CreateCommand();

                // Build query 
                string columnsToSelect = ""; // Don't want id here because it is autoincrement anyway
                string columnValuesToInsert = "";

                if (profile is not null)
                {
                    foreach (ImportColumnMappingListItem mapping in profile.ImportColumnMappings)
                    {
                        columnsToSelect += columnsToSelect.Length > 0? $", {mapping.ColumnName}" : mapping.ColumnName;
                    }

                    for (int i = 0; i < documents.Count; i++)
                    {
                        string columnValuesToInsertForRow = "";
                        foreach (ImportColumnMappingListItem mapping in profile.ImportColumnMappings)
                        {
                            DocumentColumn column = documents[i].Columns.Find(d => d.Mapping.Id == mapping.Id); // for now just try to match on id
                            object parameterValue = column is null ? null : column.Value;
                            string parameterName = $"@{mapping.ColumnName + i.ToString()}";
                            columnValuesToInsertForRow += columnValuesToInsertForRow.Length > 0 ? $", {parameterName}" : parameterName;
                            insertIntoResultSetCmd.Parameters.Add(new SQLiteParameter(parameterName, parameterValue));
                        }
                        columnValuesToInsert += columnValuesToInsert.Length == 0 ? $"({columnValuesToInsertForRow})" : $", ({columnValuesToInsertForRow})";
                    }
                }
                else
                {
                    throw new Exception("expected profile to have valid import_mapping");
                }

                
                insertIntoResultSetCmd.CommandText = $"INSERT INTO {resultTableName} ({columnsToSelect}) VALUES {columnValuesToInsert}";
                LoggerService.LogError($"INSERT INTO {resultTableName} ({columnsToSelect}) VALUES {columnValuesToInsert}");

                insertIntoResultSetCmd.ExecuteNonQuery();

            }
            catch (Exception ex)
            {
                LoggerService.LogError(ex.ToString());
                throw ex;
            }
        }


        public static void InsertIntoResultSet(ConnectionManager cm, int importResultId, Document document)
        {
            try
            {
                Utilities.CheckNull(cm);

                var conn = cm.GetSQLConnection();


                ImportResults res = ImportResultRepository.GetImportResult(cm, importResultId);
                if (res is null)
                {
                    throw new Exception($"unable to find result set with id {importResultId}");
                }

                string resultTableName = res.ResultTableName;

                MappingProfile profile = MappingProfileRepository.GetFullMappingProfileById(cm, res.ProfileId);


                var insertIntoResultSetCmd = conn.CreateCommand();

                // Build query 
                string columnsToSelect = ""; // No id on insert because it is autoincrement
                string columnValuesToInsert = "";

                if (profile is not null)
                {
                    foreach (ImportColumnMappingListItem mapping in profile.ImportColumnMappings)
                    {
                        columnsToSelect += columnsToSelect.Length > 0 ? $", {mapping.ColumnName}" : mapping.ColumnName;
                        DocumentColumn column = document.Columns.Find(d => d.Mapping.Id == mapping.Id); // for now just try to match on id
                        object parameterValue = column is null ? null : column.Value;
                        string parameterName = mapping.ColumnName;
                        columnValuesToInsert += columnValuesToInsert.Length > 0 ? $", {parameterName}" : parameterName;
                        insertIntoResultSetCmd.Parameters.Add(new SQLiteParameter(parameterName, parameterValue));
                    }
                }
                else
                {
                    throw new Exception("expected profile to have valid import_mapping");
                }


                insertIntoResultSetCmd.CommandText = $"INSERT INTO {resultTableName} ({columnsToSelect}) VALUES ({columnValuesToInsert})";


                insertIntoResultSetCmd.ExecuteNonQuery();

            } catch (Exception ex)
            {
                LoggerService.LogError(ex.ToString());
                throw ex;
            }
        }

        private static string BuildSelectClauseString(MappingProfile profile)
        {
            // Build query
            string columnsToSelect = "id";

            if (profile is not null)
            {
                foreach (ImportColumnMappingListItem mapping in profile.ImportColumnMappings)
                {
                    columnsToSelect += $", {mapping.ColumnName}";
                }
                return columnsToSelect;
            }
            else
            {
                throw new Exception("expected profile to have valid import_mapping");
            }
        }

        private static SQLiteCommand BuildResultSetQuery(ConnectionManager cm, string resultTableName, MappingProfile profile, List<ColumnFilter> filters)
        {
            Utilities.CheckNull(cm);

            var conn = cm.GetSQLConnection();
            var selectResultSetCmd = conn.CreateCommand();

            // Build query
            string columnsToSelect = BuildSelectClauseString(profile);

            string whereClause = "";


            if (filters.Count > 0)
            {
                whereClause += "WHERE ";
                for (int i = 0; i< filters.Count; i++)
                {
                    string filterParameterName = $"@{filters[i].Mapping.ColumnName}Filter";
                    whereClause += (i > 0 ? " AND " : "") + $"{filters[i].Mapping.ColumnName} like {filterParameterName}";
                    selectResultSetCmd.Parameters.Add(new SQLiteParameter(filterParameterName, $"{filters[i].Filter}%"));
                }
            }

            selectResultSetCmd.CommandText = $"SELECT {columnsToSelect} FROM {resultTableName} {whereClause}";
            return selectResultSetCmd;
        }

        public static List<Document> GetResultSet(ConnectionManager cm, int importResultId, List<ColumnFilter> filters) 
        {
            try
            {
                Utilities.CheckNull(cm);

                ImportResults res = ImportResultRepository.GetImportResult(cm, importResultId);
                if (res is null)
                {
                    throw new Exception($"unable to find result set with id {importResultId}");
                }

                string resultTableName = res.ResultTableName;

                MappingProfile profile = MappingProfileRepository.GetFullMappingProfileById(cm, res.ProfileId);


                List<Document> docResults = new List<Document>();
                var selectEntriesCmd = BuildResultSetQuery(cm, resultTableName, profile, filters);

                var reader = selectEntriesCmd.ExecuteReader();
                while (reader.Read())
                {
                    Document docToAdd = new Document
                    {
                        Id = reader.GetInt32(0),
                        ResultTableName = resultTableName,
                        Columns = new List<DocumentColumn>()
                    };

                    for (int i = 0; i < profile.ImportColumnMappings.Count; i++)
                    {
                        docToAdd.Columns.Add(new DocumentColumn
                        {
                            Mapping = ColumnMappingService.ConvertFromListItem(profile.ImportColumnMappings[i]),
                            Value = reader.GetValue(i + 1)
                        });
                        
                    }
                    docResults.Add(docToAdd);
                }

                return docResults;
            }
            catch (Exception ex)
            {
                LoggerService.LogError(ex.ToString());
                throw ex;
            }
        }

        
    }
}
