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
    class AttributeRepository
    {


        public static List<AttributeListItem> GetAssignedAttributes(ConnectionManager cm, int itemId, int resultSetId)
        {
            try
            {
                Utilities.CheckNull(cm);

                List<AttributeListItem> results = new List<AttributeListItem>();
                var conn = cm.GetSQLConnection();
                var selectAttributesCmd = conn.CreateCommand();

                selectAttributesCmd.CommandText = @"SELECT a.id, a.name, 
                                                        CASE WHEN ia.id is NULL THEN False ELSE True End as 'Assigned'
                                                    FROM attribute a, item_attribute ia
                                                    WHERE ia.attribute_id = a.id
                                                        AND ia.item_id=@ItemId
                                                        AND ia.result_set_id=@ResultSetId";
                selectAttributesCmd.Parameters.Add(new SQLiteParameter("@ItemId", itemId));
                selectAttributesCmd.Parameters.Add(new SQLiteParameter("@ResultSetId", resultSetId));

                var reader = selectAttributesCmd.ExecuteReader();
                while (reader.Read())
                {
                    results.Add(new AttributeListItem
                    {
                        Id = reader.GetInt32(0),
                        Name = reader.GetString(1),
                        IsAssigned = reader.GetBoolean(2),
                        ResultSetId = resultSetId
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
        public static List<AttributeListItem> GetAllAttributeListItems(ConnectionManager cm, int itemId, int resultSetId)
        {
            try
            {
                Utilities.CheckNull(cm);

                List<AttributeListItem> results = new List<AttributeListItem>();
                var conn = cm.GetSQLConnection();
                var selectAttributesCmd = conn.CreateCommand();

                selectAttributesCmd.CommandText = @"SELECT a.id, a.name, 
                                                        CASE WHEN ia.id is NULL THEN False ELSE True End as 'Assigned'
                                                    FROM attribute a
                                                    LEFT JOIN item_attribute ia
                                                    ON ia.attribute_id = a.id
                                                        AND ia.item_id=@ItemId
                                                        AND ia.result_set_id=@ResultSetId";
                selectAttributesCmd.Parameters.Add(new SQLiteParameter("@ItemId", itemId));
                selectAttributesCmd.Parameters.Add(new SQLiteParameter("@ResultSetId", resultSetId));

                var reader = selectAttributesCmd.ExecuteReader();
                while (reader.Read())
                {
                    results.Add(new AttributeListItem
                    {
                        Id = reader.GetInt32(0),
                        Name = reader.GetString(1),
                        IsAssigned = reader.GetBoolean(2),
                        ResultSetId = resultSetId
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

        public static void InsertAttribute(ConnectionManager cm, AttributeListItem attribute)
        {
            try
            {
                Utilities.CheckNull(cm);

                var conn = cm.GetSQLConnection();
                var insertAttributeCmd = conn.CreateCommand();

                insertAttributeCmd.CommandText = @"INSERT INTO attribute (name) VALUES (@Name)";
                insertAttributeCmd.Parameters.Add(new SQLiteParameter("@Name", attribute.Name));

                insertAttributeCmd.ExecuteNonQuery();
                
            }
            catch (Exception ex)
            {
                LoggerService.LogError(ex.ToString());
                throw ex;
            }
        }


        public static void SaveAttributeAssignments(ConnectionManager cm, int itemId, int resultSetId, List<AttributeListItem> attributes)
        {
            try
            {
                Utilities.CheckNull(cm);

                var conn = cm.GetSQLConnection();

                DeleteAttributesForItem(cm, itemId, resultSetId);

                if (attributes.Count == 0) return;

                string valueClause = "";
                for (int i = 0; i < attributes.Count; i++)
                {
                    valueClause += (valueClause.Length > 0 ? ", " : " ") + $" ({itemId.ToString()}, {resultSetId.ToString()}, {attributes[i].Id.ToString()})";
                }

                var insertAttributeCmd = conn.CreateCommand();

                insertAttributeCmd.CommandText = $"INSERT INTO item_attribute (item_id, result_set_id, attribute_id) VALUES {valueClause}";

                LoggerService.LogError($"INSERT INTO item_attribute (item_id, result_set_id, attribute_id) VALUES {valueClause}");

                insertAttributeCmd.ExecuteNonQuery();
            }
            catch (Exception ex)
            {
                LoggerService.LogError(ex.ToString());
                throw ex;
            }
        }

        public static void DeleteAttribute(ConnectionManager cm, int attributeId)
        {
            try
            {
                Utilities.CheckNull(cm);

                var conn = cm.GetSQLConnection();
                var deleteAttributeCommand = conn.CreateCommand();

                deleteAttributeCommand.CommandText = @"DELETE FROM attribute WHERE id=@Id";
                deleteAttributeCommand.Parameters.Add(new SQLiteParameter("@Id", attributeId));

                deleteAttributeCommand.ExecuteNonQuery();
            }
            catch (Exception ex)
            {
                LoggerService.LogError(ex.ToString());
                throw ex;
            }
        }

        public static void DeleteAttributesForItem(ConnectionManager cm, int itemId, int resultSetId)
        {
            try
            {
                Utilities.CheckNull(cm);

                var conn = cm.GetSQLConnection();
                var deleteAttributeCommand = conn.CreateCommand();

                deleteAttributeCommand.CommandText = @"DELETE FROM item_attribute WHERE item_id=@ItemId AND result_set_id=@ResultSetId";
                deleteAttributeCommand.Parameters.Add(new SQLiteParameter("@ItemId", itemId));
                deleteAttributeCommand.Parameters.Add(new SQLiteParameter("@ResultSetId", resultSetId));

                deleteAttributeCommand.ExecuteNonQuery();
            }
            catch (Exception ex)
            {
                LoggerService.LogError(ex.ToString());
                throw ex;
            }
        }
    }
}
