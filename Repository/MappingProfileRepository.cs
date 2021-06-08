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
    class MappingProfileRepository
    {

        public static List<MappingProfile> GetMappingProfiles(ConnectionManager cm)
        {
            try
            {
                Utilities.CheckNull(cm);
                List<MappingProfile> profiles = new List<MappingProfile>();
                var conn = cm.GetSQLConnection();
                var selectMappingProfileCmd = conn.CreateCommand();

                selectMappingProfileCmd.CommandText = @"SELECT id, name, locked FROM mapping_profile";

                var reader = selectMappingProfileCmd.ExecuteReader();
                while (reader.Read())
                {
                    profiles.Add(new MappingProfile
                    {
                        Id = reader.GetInt32(0),
                        Name = reader.GetString(1),
                        Locked = reader.GetBoolean(2)
                    });
                }

                return profiles;
            }
            catch (Exception ex)
            {
                LoggerService.LogError(ex.ToString());
                throw ex;
            }
        }
        public static MappingProfile GetMappingProfileById(ConnectionManager cm, int id)
        {
            try
            {
                Utilities.CheckNull(cm);
                List<MappingProfile> profiles = new List<MappingProfile>();
                var conn = cm.GetSQLConnection();
                var selectMappingProfileCmd = conn.CreateCommand();

                selectMappingProfileCmd.CommandText = @"SELECT id, name, locked FROM mapping_profile WHERE id = @Id";
                selectMappingProfileCmd.Parameters.Add(new SQLiteParameter("@Id", id));

                var reader = selectMappingProfileCmd.ExecuteReader();
                while (reader.Read())
                {
                    profiles.Add(new MappingProfile
                    {
                        Id = reader.GetInt32(0),
                        Name = reader.GetString(1),
                        Locked = reader.GetBoolean(2)
                    });
                }

                return profiles.FirstOrDefault();
            }
            catch (Exception ex)
            {
                LoggerService.LogError(ex.ToString());
                throw ex;
            }
        }

        public static MappingProfile GetFullMappingProfileById(ConnectionManager cm, int id)
        {
            try
            {
                Utilities.CheckNull(cm);
                MappingProfile profile = GetMappingProfileById(cm, id);
                if (profile is null)
                {
                    throw new Exception($"unable to find mapping profile with id {id}");
                }

                profile.ImportColumnMappings = ImportColumnMappingRepository.GetColumnMappingListItemsByProfileId(cm, id);
                profile.ExportColumnMappings = ExportColumnMappingRepository.GetColumnMappingListItemsByProfileId(cm, id);

                return profile;
            }
            catch (Exception ex)
            {
                LoggerService.LogError(ex.ToString());
                throw ex;
            }
        }

        public static MappingProfile GetFullMappingProfileForResultSet(ConnectionManager cm, int resultSetId)
        {
            try
            {
                Utilities.CheckNull(cm);
                ImportResults results = ImportResultRepository.GetImportResult(cm, resultSetId);
                if (results is null) return null;
                return GetFullMappingProfileById(cm, results.ProfileId);
            }
            catch (Exception ex)
            {
                LoggerService.LogError(ex.ToString());
                throw ex;
            }
        }



        public static void UpdateMappingProfile(ConnectionManager cm, MappingProfile profile)
        {
            try
            {
                Utilities.CheckNull(cm);

                var conn = cm.GetSQLConnection();
                var updateMappingProfileCmd = conn.CreateCommand();

                updateMappingProfileCmd.CommandText = @"UPDATE mapping_profile set name = @Name, locked = @Locked WHERE id = @Id";
                updateMappingProfileCmd.Parameters.Add(new SQLiteParameter("@Id", profile.Id));
                updateMappingProfileCmd.Parameters.Add(new SQLiteParameter("@Name", profile.Name));
                updateMappingProfileCmd.Parameters.Add(new SQLiteParameter("@Locked", profile.Locked));

                updateMappingProfileCmd.ExecuteNonQuery();
            }
            catch (Exception ex)
            {
                LoggerService.LogError(ex.ToString());
                throw ex;
            }
        }

        public static void InsertMappingProfile(ConnectionManager cm, MappingProfile profile, out MappingProfile insertedProfile)
        {
            try
            {
                Utilities.CheckNull(cm);

                var conn = cm.GetSQLConnection();
                var insertMappingProfileCmd = conn.CreateCommand();

                insertMappingProfileCmd.CommandText = @"INSERT INTO mapping_profile (name) VALUES (@Name)";
                insertMappingProfileCmd.Parameters.Add(new SQLiteParameter("@Name", profile.Name));

                insertMappingProfileCmd.ExecuteNonQuery();
                profile.Id = Convert.ToInt32(conn.LastInsertRowId);
                insertedProfile = profile;
            }
            catch (Exception ex)
            {
                LoggerService.LogError(ex.ToString());
                throw ex;
            }
        }
    }
}
