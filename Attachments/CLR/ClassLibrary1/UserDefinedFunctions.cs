using Microsoft.SqlServer.Server;
using System;
using System.Data.SqlClient;
using System.Data.SqlTypes;
using System.IO;
using System.IO.Compression;
using System.Linq;
using System.Text;

namespace CLR
{
    public partial class UserDefinedFunctions
    {

        [SqlFunction(IsDeterministic = true,
                         IsPrecise = true,
                         DataAccess = DataAccessKind.Read,
                         SystemDataAccess = SystemDataAccessKind.Read)]
        public static SqlString WriteToFile(SqlString path, SqlString efolderid, SqlString fileName)
        {
            try
            {
                if (!path.IsNull && !efolderid.IsNull && !fileName.IsNull)
                {
                    var dir = Path.GetDirectoryName(path.Value);
                    if (!Directory.Exists(dir))
                        Directory.CreateDirectory(dir);
                    string filename = Convert.ToString(fileName);
                    string folderid = Convert.ToString(efolderid);
                    string filepath = Convert.ToString(path);
                    SaveAttachmentToFile(filename, folderid, filepath);
                    return "Wrote file";
                }
                else
                    return "No data passed to method!";
            }
            catch (IOException e)
            {
                return "Make sure the assembly has external access!\n" + e.ToString();
            }
            catch (Exception ex)
            {
                return ex.ToString();
            }
        }

        private const int guidLength = 38 * 2;
        public static byte[] GetAttachment(string file, string efolderid)
        {
            string queryString = string.Format("SELECT eContents FROM eAttachment WHERE eKey = '0\t{0}\t{1}'",
                      efolderid, file);

            using (SqlConnection connection = new SqlConnection("context connection=true"))
            {
                connection.Open();

                using (SqlCommand selectAttachment = new SqlCommand(
                    queryString,
                    connection))
                {
                    using (SqlDataReader reader = selectAttachment.ExecuteReader())
                    {
                        if (!reader.Read())
                            return new byte[0];
                        if (reader[0] == System.DBNull.Value)
                            return new byte[0];
                        byte[] data = (byte[])reader[0];
                        if (data.Length == 0)
                            return new byte[0];
                        String base64String;
                        if (data.Length > 1 && data[1] == 00)
                            base64String = Encoding.Unicode.GetString(data);
                        else
                            base64String = Encoding.ASCII.GetString(data);
                        // Cuts off the GUID, and takes care of any trailing 00 bytes.
                        String truncatedString = base64String.Substring(38).TrimEnd('\0');
                        return Convert.FromBase64String(truncatedString);
                    }

                }
            }
        }

        /// <summary>
        /// Saves the specified attachment to a file on disk.
        /// </summary>
        /// <param name="type">The type of attachment.</param>
        /// <param name="file">The attachment filename.</param>
        /// <param name="folderid">The owner of the attachment (folder ID for folder attachments, map name for map attachments, procedure name for procedure attachments).</param>
        /// <param name="fileName">Name of the file to save to.</param>
        public static void SaveAttachmentToFile(string file, string folderid, string fileName)
        {
            byte[] data = GetAttachment(file, folderid);
            if (data == null)
                throw new ArgumentNullException("Attachment has no data, it may have been deleted");
            using (FileStream writer = new FileStream(fileName, FileMode.Create))
            {
                writer.Write(data, 0, data.Length);
            }
        }

    }
}