using AGOServer.Components;
using OpenText.Livelink.Service.DocMan;
using System;
using System.Collections.Generic;
using System.Data;
using System.Data.SqlClient;
using System.Drawing.Printing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Web.Http.Results;
using System.Web.UI.WebControls;

namespace AGOServer
{
    public static class AGODataAccess
    {
        static string EmailProperties_TableName = "OTEmailProperties";
        static string EmailConversationIDs_TableName = "OTEmailConversationIDs";
        static string Emailfiling_EmailFiles_TableName = "Emailfiling_EmailFiles";
        static string DTree_TableName = "DTree";
        static string EmailBody_TableName = "EmailBrowser_EmailBody";

        private static NLog.Logger logger = NLog.LogManager.GetCurrentClassLogger();

        public static string GetConnectionString()
        {
            string dbDataSource = Properties.Settings.Default.DB_DataSource;
            string dbInitialCatalog = Properties.Settings.Default.DB_InitialCatalog;

            string dbLoginId = "";
            string dbPassword = "";
            if (Properties.Settings.Default.UseSecureCredentials)
            {
                dbLoginId = SecureInfo.getSensitiveInfo(Properties.Settings.Default.SecureCSDBUsername_Filename);
                dbPassword = SecureInfo.getSensitiveInfo(Properties.Settings.Default.SecureCSDBPassword_Filename);
            }
            else
            {
                dbLoginId = Properties.Settings.Default.DB_Username;
                dbPassword = Properties.Settings.Default.DB_Password;
            }

            string connectionString = @"Data Source=" + dbDataSource + ";Initial Catalog=" + dbInitialCatalog + "; User ID=" + dbLoginId + ";Password=" + dbPassword + "";
            return connectionString;
        }

        public static List<FolderInfo> ListFoldersInfosFromCSDB(long NodeID)
        {
            DataTable table = new DataTable();
            using (var conn = new SqlConnection(GetConnectionString()))
            {
                using (var cmd = conn.CreateCommand())
                {
                    String SQL = "SELECT * FROM " + DTree_TableName + " where ParentID = @NodeID ";
                    SQL += @" and Deleted=0 and ( SubType=0 or SubType=751 ) and Catalog!=2 ORDER BY Name";
                    SqlDataAdapter adapter = new SqlDataAdapter();
                    cmd.CommandText = SQL;
                    cmd.Parameters.AddWithValue("@NodeID", NodeID);
                    // Add named parameters here to the command if needed...
                    adapter.SelectCommand = cmd;
                    adapter.Fill(table);
                }
                conn.Close();
            }

            List<FolderInfo> folderInfos = new List<FolderInfo>();
            foreach (DataRow row in table.Rows)
            {
                FolderInfo folderInfo = new FolderInfo();
                if (!row.IsNull("DataID"))
                {
                    folderInfo.NodeID = (long)row["DataID"];
                }
                if (!row.IsNull("Name"))
                {
                    folderInfo.Name = (string)row["Name"];
                }
                if (!row.IsNull("ParentID"))
                {
                    folderInfo.ParentNodeID = (long)row["ParentID"];
                }
                if (!row.IsNull("ChildCount"))
                {
                    folderInfo.ChildCount = (long)row["ChildCount"];
                }

                folderInfos.Add(folderInfo);
            }
            return folderInfos;
        }
        public static List<EmailInfo> ListEmailInfosFromCSDB(long NodeID)
        {
            DataTable table = new DataTable();
            using (var conn = new SqlConnection(GetConnectionString()))
            {
                using (var cmd = conn.CreateCommand())
                {
                    String SQL = "SELECT * FROM " + DTree_TableName + " where ParentID = @NodeID ";
                    SQL += @" and Deleted=0 and SubType = 749 and Catalog!=2 ORDER BY Name";//749 for emails
                    SqlDataAdapter adapter = new SqlDataAdapter();
                    cmd.CommandText = SQL;
                    cmd.Parameters.AddWithValue("@NodeID", NodeID);
                    // Add named parameters here to the command if needed...
                    adapter.SelectCommand = cmd;
                    adapter.Fill(table);
                }
                conn.Close();
            }

            List<EmailInfo> emailInfos = new List<EmailInfo>();
            foreach (DataRow row in table.Rows)
            {
                try
                {
                    EmailInfo emailInfo = new EmailInfo();
                    if (!row.IsNull("DataID"))
                    {
                        emailInfo.NodeID = (long)row["DataID"];
                    }
                    if (!row.IsNull("Name"))
                    {
                        emailInfo.Name = (string)row["Name"];
                    }
                    if (!row.IsNull("ParentID"))
                    {
                        emailInfo.ParentNodeID = (long)row["ParentID"];
                    }
                    emailInfos.Add(emailInfo);
                }
                catch (Exception ex)
                {
                    logger.Error(ex.Message + ex.StackTrace);
                }
            }
            return emailInfos;
        }

        public static List<EmailInfoVer2> GetEmails(bool showAsConversation, long userID, long folderID, int pageNumber, int pageSize, string sortColumn, string sortOrder, out int totalPage, out int totalCount)
        {
            List<EmailInfoVer2> emailList = new List<EmailInfoVer2>();
            using (SqlConnection conn = new SqlConnection(GetConnectionString()))
            {
                using (SqlCommand cmd = new SqlCommand())
                {
                    conn.Open();
                    cmd.Connection = conn;
                    cmd.CommandType = CommandType.StoredProcedure;
                    cmd.CommandText = "EmailBrowser_GetEmails";

                    cmd.Parameters.Clear();
                    cmd.Parameters.Add("@showAsConversation", SqlDbType.SmallInt).Value = showAsConversation;
                    cmd.Parameters.Add("@userID", SqlDbType.BigInt).Value = userID;
                    cmd.Parameters.Add("@folderID", SqlDbType.BigInt).Value = folderID;
                    cmd.Parameters.Add("@pageNumber", SqlDbType.SmallInt).Value = pageNumber;
                    cmd.Parameters.Add("@pageSize", SqlDbType.SmallInt).Value = pageSize;
                    cmd.Parameters.Add("@sortColumn", SqlDbType.VarChar, 100).Value = sortColumn;
                    cmd.Parameters.Add("@sortOrder", SqlDbType.VarChar, 100).Value = sortOrder;
                    cmd.Parameters.Add("@totalPage", SqlDbType.Int).Direction = ParameterDirection.Output;
                    cmd.Parameters.Add("@totalCount", SqlDbType.Int).Direction = ParameterDirection.Output;

                    using (SqlDataReader reader = cmd.ExecuteReader())
                    {
                        while (reader.Read())
                        {
                            EmailInfoVer2 email = new EmailInfoVer2();
                            email.NodeID = (long)reader["DataID"];
                            email.ParentNodeID = folderID; // The input folder where query search the emails.
                            email.Name = (string)reader["Name"];
                            email.EmailSubject = reader["EmailSubject"] == DBNull.Value ? "" : (string)reader["EmailSubject"];
                            email.EmailTo = reader["EmailTo"] == DBNull.Value ? "" : (string)reader["EmailTo"];
                            email.EmailCC = reader["EmailCC"] == DBNull.Value ? "" : (string)reader["EmailCC"];
                            email.EmailFrom = reader["EmailFrom"] == DBNull.Value ? "" : (string)reader["EmailFrom"];
                            if (reader["EmailSentDate"] == DBNull.Value)
                            {
                                email.SentDate = null;
                            }
                            else
                            {
                                email.SentDate = (DateTime)reader["EmailSentDate"];
                            }
                            if (reader["EmailReceivedDate"] == DBNull.Value)
                            {
                                email.ReceivedDate = null;
                            }
                            else
                            {
                                email.ReceivedDate = (DateTime)reader["EmailReceivedDate"];
                            }
                            email.ConversationId = reader["ConversationID"] == DBNull.Value ? "" : (string)reader["ConversationID"];
                            email.HasAttachments = reader["HasAttachments"] == DBNull.Value ? 0 : (int)reader["HasAttachments"];
                            email.FileSize = reader["FileSize"] == DBNull.Value ? 0 : (long)reader["FileSize"];
                            email.Summary = ""; // Store Procedure do not provide the value yet.
                            email.ClientLocalEmailID = 0; // Store Procedure do not provide the value yet.
                            emailList.Add(email);
                        }
                    }

                    totalPage = Convert.ToInt32(cmd.Parameters["@totalPage"].Value);
                    totalCount = Convert.ToInt32(cmd.Parameters["@totalCount"].Value);
                }
            }

            return emailList;
        }

        public static string GetEmailHTMLBody(long nodeID)
        {
            using (DataTable table = new DataTable())
            {
                using (var conn = new SqlConnection(GetConnectionString()))
                {
                    using (var cmd = conn.CreateCommand())
                    {
                        String SQL = "SELECT top 1 * FROM " + EmailBody_TableName + " where NodeID = @NodeID ";

                        SqlDataAdapter adapter = new SqlDataAdapter();
                        cmd.CommandText = SQL;
                        cmd.Parameters.AddWithValue("@NodeID", nodeID);
                        // Add named parameters here to the command if needed...
                        adapter.SelectCommand = cmd;
                        adapter.Fill(table);
                    }
                    conn.Close();
                }
                if (table.Rows.Count > 0)
                {
                    DataRow row = table.Rows[0];
                    if (!row.IsNull("HtmlBody"))
                    {
                        return (string)row["HtmlBody"];
                    }
                }
            }
            return "";
        }

        public static void StoreOrUpdateEmailHTMLBody(long nodeID, string htmlBody)
        {
            using (SqlConnection dbConnection = new SqlConnection(GetConnectionString()))
            {
                dbConnection.Open();
                SqlTransaction transaction = dbConnection.BeginTransaction();

                try
                {
                    using (var cmd = dbConnection.CreateCommand())
                    {
                        cmd.Transaction = transaction;

                        // Check if the record exists
                        cmd.CommandText = $"SELECT COUNT(*) FROM {EmailBody_TableName} WHERE NodeId = @nodeId";
                        cmd.Parameters.AddWithValue("@nodeId", nodeID);
                        int existingCount = (int)cmd.ExecuteScalar();

                        if (existingCount > 0)
                        {
                            // Update the existing record
                            cmd.CommandText = $"UPDATE {EmailBody_TableName} SET HtmlBody = @htmlBody WHERE NodeId = @nodeId";
                            cmd.Parameters.AddWithValue("@htmlBody", htmlBody);
                            cmd.ExecuteNonQuery();
                        }
                        else
                        {
                            // Insert a new record
                            cmd.CommandText = $@"INSERT INTO {EmailBody_TableName}
                                        (NodeId, HtmlBody) 
                                        VALUES(@nodeId, @htmlBody)";
                            cmd.Parameters.AddWithValue("@htmlBody", htmlBody);
                            cmd.ExecuteNonQuery();
                        }
                    }

                    transaction.Commit();
                }
                catch (Exception ex)
                {
                    transaction.Rollback();
                    logger.Info(ex.Message + ex.StackTrace);
                }
                finally
                {
                    dbConnection.Close();
                }
            }
        }

        public static List<EmailInfo> GetConversations(long nodeID, long userID)
        {
            List<EmailInfo> emailList = new List<EmailInfo>();
            using (SqlConnection conn = new SqlConnection(GetConnectionString()))
            {
                using (SqlCommand cmd = new SqlCommand())
                {
                    conn.Open();
                    cmd.Connection = conn;
                    cmd.CommandType = CommandType.StoredProcedure;
                    cmd.CommandText = "EmailBrowser_GetConversations";

                    cmd.Parameters.Clear();
                    cmd.Parameters.Add("@nodeID", SqlDbType.BigInt).Value = nodeID;
                    cmd.Parameters.Add("@userID", SqlDbType.BigInt).Value = userID;

                    using (SqlDataReader reader = cmd.ExecuteReader())
                    {
                        while (reader.Read())
                        {
                            EmailInfo email = new EmailInfo();
                            email.NodeID = (long)reader["DataID"];
                            email.ParentNodeID = (long)reader["ParentID"]; // The input folder where query search the emails.
                            email.Name = (string)reader["Name"];
                            email.EmailSubject = reader["EmailSubject"] == DBNull.Value ? "" : (string)reader["EmailSubject"];
                            email.EmailTo = reader["EmailTo"] == DBNull.Value ? "" : (string)reader["EmailTo"];
                            email.EmailCC = reader["EmailCC"] == DBNull.Value ? "" : (string)reader["EmailCC"];
                            email.EmailFrom = reader["EmailFrom"] == DBNull.Value ? "" : (string)reader["EmailFrom"];
                            if (reader["EmailSentDate"] == DBNull.Value)
                            {
                                email.SentDate = null;
                            }
                            else
                            {
                                email.SentDate = (DateTime)reader["EmailSentDate"];
                            }
                            if (reader["EmailReceivedDate"] == DBNull.Value)
                            {
                                email.ReceivedDate = null;
                            }
                            else
                            {
                                email.ReceivedDate = (DateTime)reader["EmailReceivedDate"];
                            }
                            email.ConversationId = reader["ConversationID"] == DBNull.Value ? "" : (string)reader["ConversationID"];
                            email.HasAttachments = reader["HasAttachments"] == DBNull.Value ? 0 : Convert.ToInt32(reader["HasAttachments"]);
                            email.FileSize = reader["FileSize"] == DBNull.Value ? 0 : (long)reader["FileSize"];
                            email.PermSeeContents = Convert.ToInt32(reader["See"]) > 0 ? true : false;
                            email.Summary = ""; // Store Procedure do not provide the value yet.
                            email.ClientLocalEmailID = 0; // Store Procedure do not provide the value yet.
                            emailList.Add(email);
                        }
                    }
                }
            }

            return emailList;
        }

        public static void GetEmailInfoFromCSDBVer2(ref EmailSearchInfoVer2 emailInfo)
        {
            if (emailInfo == null || emailInfo.NodeID <= 0)
            {
                return;
            }
            using (DataTable table = new DataTable())
            {
                using (var conn = new SqlConnection(GetConnectionString()))
                {
                    using (var cmd = conn.CreateCommand())
                    {
                        String SQL = "SELECT top 1 * FROM " + EmailProperties_TableName + " where NodeID = @NodeID ";

                        SqlDataAdapter adapter = new SqlDataAdapter();
                        cmd.CommandText = SQL;
                        cmd.Parameters.AddWithValue("@NodeID", emailInfo.NodeID);
                        // Add named parameters here to the command if needed...
                        adapter.SelectCommand = cmd;
                        adapter.Fill(table);
                    }
                    conn.Close();
                }
                if (table.Rows.Count > 0)
                {
                    DataRow row = table.Rows[0];
                    if (!row.IsNull("OTEmailSubject"))
                    {
                        emailInfo.EmailSubject = (string)row["OTEmailSubject"];
                    }
                    if (!row.IsNull("OTEmailTo"))
                    {
                        emailInfo.EmailTo = (string)row["OTEmailTo"];
                    }
                    if (!row.IsNull("OTEmailCC"))
                    {
                        emailInfo.EmailCC = (string)row["OTEmailCC"];
                    }
                    if (!row.IsNull("OTEmailFrom"))
                    {
                        emailInfo.EmailFrom = (string)row["OTEmailFrom"];
                    }
                    if (!row.IsNull("OTEmailSentDate"))
                    {
                        emailInfo.SentDate = (DateTime)row["OTEmailSentDate"];
                    }
                    if (!row.IsNull("OTEmailReceivedDate"))
                    {
                        emailInfo.ReceivedDate = (DateTime)row["OTEmailReceivedDate"];
                    }
                    if (!row.IsNull("HasAttachments"))
                    {
                        emailInfo.HasAttachments = Convert.ToInt32(row["HasAttachments"]);
                    }
                }
            }
        }

        public static void GetEmailInfoFromCSDB(ref EmailInfo emailInfo)
        {
            if (emailInfo == null || emailInfo.NodeID <= 0)
            {
                return;
            }
            using (DataTable table = new DataTable())
            {
                using (var conn = new SqlConnection(GetConnectionString()))
                {
                    using (var cmd = conn.CreateCommand())
                    {
                        String SQL = "SELECT top 1 * FROM " + EmailProperties_TableName + " where NodeID = @NodeID ";

                        SqlDataAdapter adapter = new SqlDataAdapter();
                        cmd.CommandText = SQL;
                        cmd.Parameters.AddWithValue("@NodeID", emailInfo.NodeID);
                        // Add named parameters here to the command if needed...
                        adapter.SelectCommand = cmd;
                        adapter.Fill(table);
                    }
                    conn.Close();
                }
                if (table.Rows.Count > 0)
                {
                    DataRow row = table.Rows[0];
                    if (!row.IsNull("OTEmailSubject"))
                    {
                        emailInfo.EmailSubject = (string)row["OTEmailSubject"];
                    }
                    if (!row.IsNull("OTEmailTo"))
                    {
                        emailInfo.EmailTo = (string)row["OTEmailTo"];
                    }
                    if (!row.IsNull("OTEmailCC"))
                    {
                        emailInfo.EmailCC = (string)row["OTEmailCC"];
                    }
                    if (!row.IsNull("OTEmailFrom"))
                    {
                        emailInfo.EmailFrom = (string)row["OTEmailFrom"];
                    }
                    if (!row.IsNull("OTEmailSentDate"))
                    {
                        emailInfo.SentDate = (DateTime)row["OTEmailSentDate"];
                    }
                    if (!row.IsNull("OTEmailReceivedDate"))
                    {
                        emailInfo.ReceivedDate = (DateTime)row["OTEmailReceivedDate"];
                    }
                    if (!row.IsNull("HasAttachments"))
                    {
                        emailInfo.HasAttachments = Convert.ToInt32(row["HasAttachments"]);
                    }
                }
            }
        }

        public static string GetEmailConversationIDFromCSDB(long NodeID)
        {
            string conversationID = "";
            using (DataTable table = new DataTable())
            {
                using (var conn = new SqlConnection(GetConnectionString()))
                {
                    using (var cmd = conn.CreateCommand())
                    {
                        String SQL = "SELECT top 1 * FROM " + EmailConversationIDs_TableName + " where NodeID = @NodeID ";

                        SqlDataAdapter adapter = new SqlDataAdapter();
                        cmd.CommandText = SQL;
                        cmd.Parameters.AddWithValue("@NodeID", NodeID);
                        // Add named parameters here to the command if needed...
                        adapter.SelectCommand = cmd;
                        adapter.Fill(table);
                    }
                    conn.Close();
                }
                if (table.Rows.Count > 0)
                {
                    DataRow row = table.Rows[0];
                    if (!row.IsNull("ConversationID"))
                    {
                        conversationID = (string)row["ConversationID"];
                    }
                }
            }
            return conversationID;
        }

        public static long GetLocalEmailIDfromEmailFilling(long NodeID)
        {
            long localEmailID = 0;
            using (DataTable table = new DataTable())
            {
                using (var conn = new SqlConnection(GetConnectionString()))
                {
                    using (var cmd = conn.CreateCommand())
                    {
                        String SQL = "SELECT top 1 * FROM " + Emailfiling_EmailFiles_TableName + " where CSNodeID = @NodeID ";
                        SqlDataAdapter adapter = new SqlDataAdapter();
                        cmd.CommandText = SQL;
                        cmd.Parameters.AddWithValue("@NodeID", NodeID);

                        adapter.SelectCommand = cmd;
                        adapter.Fill(table);
                    }
                    conn.Close();
                }
                if (table.Rows.Count > 0)
                {
                    DataRow row = table.Rows[0];
                    if (!row.IsNull("ClientLocalEmailID"))
                    {
                        localEmailID = (int)row["ClientLocalEmailID"];
                    }
                }
            }

            return localEmailID;
        }

        public static string GetEmailConversationIdStoredProcedure(long nodeId)
        {
            string conversationId = "";
            using (DataTable table = new DataTable())
            {
                using (var conn = new SqlConnection(GetConnectionString()))
                {
                    using (var cmd = conn.CreateCommand())
                    {
                        cmd.CommandText = "GetConversationId";
                        cmd.CommandType = CommandType.StoredProcedure;

                        // Set up parameters.
                        cmd.Parameters.AddWithValue("@nodeId", nodeId);
                        cmd.Parameters.Add("@conversationId", SqlDbType.NVarChar, 4000).Direction = ParameterDirection.Output;

                        conn.Open();
                        cmd.ExecuteNonQuery();

                        conversationId = Convert.ToString(cmd.Parameters["@conversationId"].Value);
                        conn.Close();
                    }
                }
            }

            return conversationId;
        }

        public static List<long> FindEmailIDWithConversationIDFromCSDB(string ConversationID, int maxCount)
        {
            DataTable table = new DataTable();
            using (var conn = new SqlConnection(GetConnectionString()))
            {
                using (var cmd = conn.CreateCommand())
                {
                    String SQL = "SELECT TOP " + maxCount + " NodeID FROM " + EmailConversationIDs_TableName + " where ConversationID = @ConversationID ";
                    SqlDataAdapter adapter = new SqlDataAdapter();
                    cmd.CommandText = SQL;
                    cmd.Parameters.AddWithValue("@ConversationID", ConversationID);
                    // Add named parameters here to the command if needed...
                    adapter.SelectCommand = cmd;
                    adapter.Fill(table);
                }
                conn.Close();
            }

            var emailIDs = new List<long>();
            foreach (DataRow row in table.Rows)
            {
                try
                {
                    if (!row.IsNull("NodeID"))
                    {
                        long NodeID = (long)row["NodeID"];
                        emailIDs.Add(NodeID);
                    }
                }
                catch (Exception ex)
                {
                    logger.Error(ex.Message + ex.StackTrace);
                }
            }
            return emailIDs;
        }

        public static List<EmailInfo> FindEmailConversations(string conversationID, int maxCount, string sortedBy, string sortDirection)
        {
            List<EmailInfo> emailList = new List<EmailInfo>();
            try
            {
                using (SqlConnection conn = new SqlConnection(GetConnectionString()))
                {
                    using (SqlCommand cmd = new SqlCommand())
                    {
                        conn.Open();
                        cmd.Connection = conn;
                        cmd.CommandType = CommandType.Text;

                        var stringBuilder = new StringBuilder();
                        stringBuilder.Append(" select top " + maxCount + " prop.*, dtree.Name, dtree.ParentID, dvers.DataSize from " + EmailConversationIDs_TableName + " conv left join OTEmailProperties prop on conv.NodeID = prop.NodeID left join DTree dtree on dtree.DataID = prop.NodeID left join DVersData dvers on dvers.DocID = prop.NodeID where conv.ConversationID = @conversationID ");

                        if (sortDirection == "asc")
                        {
                            stringBuilder.Append(string.Format(" order by {0} asc ", sortedBy));
                        }
                        else
                        {
                            stringBuilder.Append(string.Format(" order by {0} desc ", sortedBy));
                        }

                        cmd.CommandText = stringBuilder.ToString();

                        cmd.Parameters.Clear();
                        cmd.Parameters.Add("@conversationID", SqlDbType.NVarChar).Value = conversationID;

                        using (SqlDataReader reader = cmd.ExecuteReader())
                        {
                            while (reader.Read())
                            {
                                EmailInfo email = new EmailInfo();
                                email.NodeID = (long)reader["NodeID"];
                                email.ParentNodeID = reader["ParentID"] == DBNull.Value ? (long)0 : (long)reader["ParentID"];
                                email.Name = reader["Name"] == DBNull.Value ? "" : (string)reader["Name"];
                                email.EmailSubject = reader["OTEmailSubject"] == DBNull.Value ? "" : (string)reader["OTEmailSubject"];
                                email.EmailTo = reader["OTEmailTo"] == DBNull.Value ? "" : (string)reader["OTEmailTo"];
                                email.EmailCC = reader["OTEmailCC"] == DBNull.Value ? "" : (string)reader["OTEmailCC"];
                                email.EmailFrom = reader["OTEmailFrom"] == DBNull.Value ? "" : (string)reader["OTEmailFrom"];
                                if (reader["OTEmailSentDate"] == DBNull.Value)
                                {
                                    email.SentDate = null;
                                }
                                else
                                {
                                    email.SentDate = (DateTime)reader["OTEmailSentDate"];
                                }
                                if (reader["OTEmailReceivedDate"] == DBNull.Value)
                                {
                                    email.ReceivedDate = null;
                                }
                                else
                                {
                                    email.ReceivedDate = (DateTime)reader["OTEmailReceivedDate"];
                                }
                                email.ConversationId = "";
                                email.HasAttachments = reader["HasAttachments"] == DBNull.Value ? 0 : Convert.ToInt32(reader["HasAttachments"]);
                                email.FileSize = reader["DataSize"] == DBNull.Value ? 0 : (long)reader["DataSize"];
                                email.Summary = ""; // Store Procedure do not provide the value yet.
                                email.ClientLocalEmailID = 0; // Store Procedure do not provide the value yet.
                                emailList.Add(email);
                            }
                        }
                    }
                }
            }
            catch (Exception ex)
            {

            }

            return emailList;
        }

        public static List<long> GetFolderIDAfterEnterprise(long dataID)
        {
            DataTable table = new DataTable();
            using (var conn = new SqlConnection(GetConnectionString()))
            {
                using (var cmd = conn.CreateCommand())
                {
                    String SQL = "select *FROM DTree i, DTree p WHERE i.ParentID = p.DataID and i.ParentID=@dataID";
                    SqlDataAdapter adapter = new SqlDataAdapter();
                    cmd.CommandText = SQL;
                    cmd.Parameters.AddWithValue("@dataID", dataID);
                    // Add named parameters here to the command if needed...
                    adapter.SelectCommand = cmd;
                    adapter.Fill(table);
                }
                conn.Close();
            }

            var dataIDs = new List<long>();
            foreach (DataRow row in table.Rows)
            {
                try
                {
                    if (!row.IsNull("DataID"))
                    {
                        long NodeID = (long)row["DataID"];
                        dataIDs.Add(NodeID);
                    }
                }
                catch (Exception ex)
                {
                    logger.Error(ex.Message + ex.StackTrace);
                }
            }
            return dataIDs;
        }

        public static List<FolderSearchInfo> ListFoldersByNameFromCSDB(string FolderName, long UserNodeID)
        {
            DataTable table = new DataTable();
            using (var conn = new SqlConnection(GetConnectionString()))
            {
                using (var cmd = conn.CreateCommand())
                {
                    String SQL = @"SELECT * FROM (SELECT DISTINCT dt.DataID, dt.Name, dt.ParentID, dt.ChildCount, dt.GPermissions, dt.SPermissions, dt.SubType," +
                        "(select dbo.fn_llpath(dt.DataID)) as FullPath FROM DTree dt " +
                        "LEFT JOIN DTreeACL DTA on DTA.DataID = dt.DataID " +
                        "LEFT OUTER JOIN KUAFRightsListNew ku ON ku.RLRightID = DTA.RightID " +
                        "where ku.RLID = " + UserNodeID + " and dt.Name LIKE '%" + FolderName + "%' " +
                        "and dt.Deleted=0 and dt.ParentID != 2000 and dt.GPermissions != 128 and Catalog != 128 " +
                        "and (dt.SubType = 0 or dt.SubType=751) and dt.ChildCount > 0 and (SELECT CASE When DTA.See > 1 THEN 1 ELSE 0 END) > 0) " +
                        "AS DATA Order BY len(FullPath) DESC";

                    SqlDataAdapter adapter = new SqlDataAdapter();
                    cmd.CommandText = SQL;
                    cmd.Parameters.AddWithValue("@folderName", FolderName);
                    cmd.Parameters.AddWithValue("@userNodeId", UserNodeID);
                    // Add named parameters here to the command if needed...
                    adapter.SelectCommand = cmd;
                    adapter.Fill(table);
                }
                conn.Close();
            }

            List<FolderSearchInfo> folderInfos = new List<FolderSearchInfo>();
            foreach (DataRow row in table.Rows)
            {
                FolderSearchInfo folderInfo = new FolderSearchInfo();
                if (!row.IsNull("DataID"))
                {
                    folderInfo.NodeID = (long)row["DataID"];
                }
                if (!row.IsNull("Name"))
                {
                    folderInfo.Name = (string)row["Name"];
                }
                if (!row.IsNull("ParentID"))
                {
                    folderInfo.ParentNodeID = (long)row["ParentID"];
                }
                if (!row.IsNull("ChildCount"))
                {
                    folderInfo.ChildCount = (long)row["ChildCount"];
                }
                if (!row.IsNull("FullPath"))
                {
                    folderInfo.FullPath = (string)row["FullPath"];
                }

                folderInfos.Add(folderInfo);
            }
            logger.Info($"folderInfos {folderInfos}");
            return folderInfos;
        }

        public static String GetFolderPathFromCSDB(long NodeID)
        {
            var path = "";
            DataTable table = new DataTable();
            using (var conn = new SqlConnection(GetConnectionString()))
            {
                using (var cmd = conn.CreateCommand())
                {
                    string SQL = "select dbo.fn_llpath(@NodeId) As PATH";

                    cmd.CommandText = SQL;
                    cmd.Parameters.AddWithValue("@NodeId", NodeID);
                    SqlDataAdapter adapter = new SqlDataAdapter();
                    // Add named parameters here to the command if needed...
                    adapter.SelectCommand = cmd;
                    adapter.Fill(table);
                }
                Object paths = table.Rows[0][0];
                path = paths.ToString();
                conn.Close();
            }
            return path;
        }

        public static EmailSearchResult SearchAttachmentFileNameFromCSDB(long userID, string attachmentName, long pageSize, long pageNumber, string OTEmailFrom, string OTEmailSubject, string OTEmailTo, string OTEmailSentDate_From, string OTEmailSentDate_To, string SortColumn, string SortOrder)
        {
            DataTable table = new DataTable();
            using (var conn = new SqlConnection(GetConnectionString()))
            {
                using (var cmd = conn.CreateCommand())
                {
                    String SQL = @"WITH CTE AS (
                    SELECT 
                        dt.DataID, 
                        dt.Name, 
                        COUNT(em.ID) AS attachments,
                        ep.OTEmailSubject, 
                        ep.OTEmailTO, 
                        ep.OTEmailFrom, 
                        ep.OTEmailReceivedDate, 
                        ep.OTEmailSentDate
                    FROM Dtree AS dt 
                    LEFT JOIN EmailBrowser_ExtractedMailAttachments AS ea ON dt.DataID = ea.CSEmailID
                    LEFT JOIN OTEmailProperties AS ep ON ep.NodeID = dt.DataID
                    LEFT JOIN DTreeACL AS dta ON dta.DataID = dt.DataID
                    LEFT OUTER JOIN KUAFRightsListNew AS ku ON ku.RLRightID = dta.RightID
                    LEFT JOIN (
                        SELECT CSEmailID, COUNT(*) AS ID
                        FROM EmailBrowser_ExtractedMailAttachments
                        GROUP BY CSEmailID
                    ) AS em ON em.CSEmailID = dt.DataID
                    WHERE 
                        ea.FileName LIKE '%'+ @AttachmentName +'%'
		                AND ep.OTEmailFrom Like '%'+ @OTEmailFrom +'%'
		                AND ep.OTEmailTo LIKE '%'+ @OTEmailTo +'%'
		                AND ep.OTEmailSubject LIKE '%'+ @OTEmailSubject +'%'
		                AND ep.OTEmailSentDate between @OTEmailSentDate_From and @OTEmailSentDate_To
		                AND ku.RLID = @userID --userid
		                AND dta.See > 1    
                    GROUP BY 
                        dt.DataID, 
                        ea.FileName, 
                        dt.Name, 
                        ep.OTEmailSubject, 
                        ep.OTEmailTO, 
                        ep.OTEmailFrom, 
                        ep.OTEmailReceivedDate, 
                        ep.OTEmailSentDate
                ),
                TotalCount AS (
                    SELECT COUNT(*) AS TotalCount FROM CTE
                )
                SELECT 
                    CTE.*, 
                    TotalCount.TotalCount,
                    CEILING(CAST(TotalCount.TotalCount AS FLOAT) / @PageSize) AS TotalPages
                FROM CTE
                CROSS JOIN TotalCount
                ORDER BY 
	                OTEmailSubject ASC
                OFFSET (@PageNumber - 1) * @PageSize ROWS
                FETCH NEXT @PageSize ROWS ONLY";
                    //cmd.CommandText = SQL;

                    conn.Open();
                    cmd.Connection = conn;
                    cmd.CommandType = CommandType.StoredProcedure;
                    cmd.CommandText = "EmailBrowser_SearchAttachmentName";


                    // Add named parameters here to the command if needed...
                    cmd.Parameters.AddWithValue("@AttachmentName", SqlDbType.VarChar).Value = attachmentName;
                    cmd.Parameters.AddWithValue("@userID", SqlDbType.BigInt).Value = userID;
                    cmd.Parameters.AddWithValue("@pageSize", SqlDbType.BigInt).Value = pageSize;
                    cmd.Parameters.AddWithValue("@pageNumber", SqlDbType.BigInt).Value = pageNumber;
                    cmd.Parameters.AddWithValue("@OTEmailFrom", SqlDbType.VarChar).Value = OTEmailFrom;
                    cmd.Parameters.AddWithValue("@OTEmailSubject", SqlDbType.VarChar).Value = OTEmailSubject;
                    cmd.Parameters.AddWithValue("@OTEmailTo", SqlDbType.VarChar).Value = OTEmailTo;
                    //cmd.Parameters.AddWithValue("@SortColumn", SqlDbType.VarChar).Value = SortColumn;
                    //cmd.Parameters.AddWithValue("@SortOrder", SqlDbType.VarChar).Value = SortOrder;
                    cmd.Parameters.AddWithValue("@OTEmailSentDate_START", SqlDbType.VarChar).Value = OTEmailSentDate_From;
                    cmd.Parameters.AddWithValue("@OTEmailSentDate_END", SqlDbType.VarChar).Value = OTEmailSentDate_To;
                    //adapter.SelectCommand = cmd;
                    SqlDataAdapter adapter = new SqlDataAdapter(cmd);

                    adapter.Fill(table);
                }
                conn.Close();

                EmailSearchResult results = new EmailSearchResult();
                foreach (DataRow row in table.Rows)
                {
                    EmailInfo emailInfo = new EmailInfo();
                    if (!row.IsNull("DataID"))
                    {
                        emailInfo.NodeID = (long)row["DataID"];
                    }
                    if (!row.IsNull("Name"))
                    {
                        emailInfo.Name = (string)row["Name"];
                    }
                    if (!row.IsNull("ParentID"))
                    {
                        emailInfo.ParentNodeID = (long)row["ParentID"];
                    }
                    if (!row.IsNull("EmailSubject"))
                    {
                        emailInfo.EmailSubject = (string)row["EmailSubject"];
                    }
                    if (!row.IsNull("EmailTo"))
                    {
                        emailInfo.EmailTo = (string)row["EmailTo"];
                    }
                    if (!row.IsNull("EmailFrom"))
                    {
                        emailInfo.EmailFrom = (string)row["EmailFrom"];
                    }
                    if (!row.IsNull("OTEmailSentDate"))
                    {
                        emailInfo.SentDate = (DateTime)row["OTEmailSentDate"];
                    }
                    if (!row.IsNull("OTEmailReceivedDate"))
                    {
                        emailInfo.ReceivedDate = (DateTime)row["OTEmailReceivedDate"];
                    }
                    if (!row.IsNull("TotalCount"))
                    {
                        results.ActualCount = (int)row["TotalCount"];
                    }
                    if (!row.IsNull("TotalPages"))
                    {
                        results.ListHead = (int)row["TotalPages"];
                    }

                    results.EmailInfos.Add(emailInfo);
                }

                return results;
            }
        }

        #region attachments

        internal static List<AttachmentInfo> ListAttachmentsFromCSDB(long emailID)
        {
            DataTable table = new DataTable();
            using (var conn = new SqlConnection(GetConnectionString()))
            {
                using (var cmd = conn.CreateCommand())
                {
                    String SQL =
                        @" SELECT DISTINCT * 
                          FROM EmailBrowser_ExtractedEmailAttachments
                          WHERE CSEmailID = @CSEmailID 
                          AND Deleted = 0 ";

                    /*UNION SELECT * FROM EmailBrowser_ExtractedMailAttachments 
                    WHERE CSEmailID = @CSEmailID AND FileType = 'Email'";*/

                    SqlDataAdapter adapter = new SqlDataAdapter();
                    cmd.CommandText = SQL;
                    cmd.Parameters.AddWithValue("@CSEmailID", emailID);
                    // Add named parameters here to the command if needed...
                    adapter.SelectCommand = cmd;
                    adapter.Fill(table);
                }
                conn.Close();
            }

            List<AttachmentInfo> infos = new List<AttachmentInfo>();
            foreach (DataRow row in table.Rows)
            {
                try
                {
                    AttachmentInfo info = new AttachmentInfo();
                    if (!row.IsNull("CSEmailID"))
                    {
                        info.CSEmailID = (long)row["CSEmailID"];
                    }
                    if (!row.IsNull("FileHash"))
                    {
                        info.FileHash = (byte[])row["FileHash"];
                    }
                    if (!row.IsNull("CSID"))
                    {
                        info.CSID = (long)row["CSID"];
                    }
                    if (!row.IsNull("FileName"))
                    {
                        info.FileName = (string)row["FileName"];
                    }
                    if (!row.IsNull("FileSize"))
                    {
                        info.FileSize = (long)row["FileSize"];
                    }
                    if (!row.IsNull("FileType"))
                    {
                        info.FileType = (string)row["FileType"];
                    }

                    infos.Add(info);
                }
                catch (Exception ex)
                {
                    logger.Error(ex.Message + ex.StackTrace);
                }
            }
            return infos;
        }

        internal static bool WasAttachmentsExtracted(long emailID)
        {
            bool wasExtracted = false;
            using (SqlConnection dbConnection = new SqlConnection(GetConnectionString()))
            {
                try
                {
                    dbConnection.Open();
                    using (var cmd = dbConnection.CreateCommand())
                    {
                        String SQL =
                            @" SELECT IsExtracted FROM EmailBrowser_ExtractedEmails
                            WHERE CSID=@CSID";
                        cmd.CommandText = SQL;
                        cmd.Parameters.AddWithValue("@CSID", emailID);
                        var result = cmd.ExecuteScalar();
                        if (result != null)
                        {
                            wasExtracted = (bool)result;
                        }
                    }
                }
                catch (Exception ex)
                {
                    logger.Info(ex.Message + ex.StackTrace);
                }
                dbConnection.Close();
            }
            return wasExtracted;
        }
        internal static void UpdateEmailAsExtracted(long emailID)
        {
            using (SqlConnection dbConnection = new SqlConnection(GetConnectionString()))
            {
                try
                {
                    dbConnection.Open();
                    using (var cmd = dbConnection.CreateCommand())
                    {
                        String SQL =
                            @" UPDATE EmailBrowser_ExtractedEmails
                            SET IsExtracted=1, IsExtractingNow=0, LastExtractedDate=@LastExtractedDate 
                            WHERE CSID=@CSID";
                        cmd.CommandText = SQL;
                        cmd.Parameters.AddWithValue("@CSID", emailID);
                        cmd.Parameters.AddWithValue("@LastExtractedDate", DateTime.Now);
                        cmd.ExecuteNonQuery();
                    }
                }
                catch (Exception ex)
                {
                    logger.Info(ex.Message + ex.StackTrace);
                }
                dbConnection.Close();
            }
        }

        internal static void RegisterEmailForExtraction(ExtractedEmailInfo extractedEmailInfo)
        {
            using (SqlConnection dbConnection = new SqlConnection(GetConnectionString()))
            {
                try
                {
                    dbConnection.Open();
                    using (var cmd = dbConnection.CreateCommand())
                    {
                        String SQL =
                            @" INSERT INTO EmailBrowser_ExtractedEmails
                            (CSID, CSName, CSModifyDate, CSParentID, CSCachingFolderNodeID, IsExtractingNow) 
                            VALUES(@CSID, @CSName, @CSModifyDate, @CSParentID, @CSCachingFolderNodeID, @IsExtractingNow)  ";
                        cmd.CommandText = SQL;
                        cmd.Parameters.AddWithValue("@CSID", extractedEmailInfo.CSID);
                        cmd.Parameters.AddWithValue("@CSName", extractedEmailInfo.CSName);
                        cmd.Parameters.AddWithValue("@CSModifyDate", extractedEmailInfo.CSModifyDate);
                        cmd.Parameters.AddWithValue("@CSParentID", extractedEmailInfo.CSParentID);
                        cmd.Parameters.AddWithValue("@CSCachingFolderNodeID", extractedEmailInfo.CSCachingFolderNodeID);
                        cmd.Parameters.AddWithValue("@IsExtractingNow", extractedEmailInfo.IsExtractingNow);
                        cmd.ExecuteNonQuery();
                    }
                }
                catch (Exception ex)
                {
                    logger.Info(ex.Message + ex.StackTrace);
                }
                dbConnection.Close();
            }
        }
        internal static void RegisterEmailAttachment(AttachmentInfo attachmentInfo)
        {
            using (SqlConnection dbConnection = new SqlConnection(GetConnectionString()))
            {
                try
                {
                    dbConnection.Open();
                    using (var cmd = dbConnection.CreateCommand())
                    {
                        String SQL =
                            @" INSERT INTO EmailBrowser_ExtractedEmailAttachments
                            (CSEmailID, FileHash, CSID, FileName, FileType, FileSize) 
                            VALUES(@CSEmailID, @FileHash, @CSID, @FileName, @FileType, @FileSize) ";
                        cmd.CommandText = SQL;
                        cmd.Parameters.AddWithValue("@CSEmailID", attachmentInfo.CSEmailID);
                        cmd.Parameters.AddWithValue("@FileHash", attachmentInfo.FileHash);
                        cmd.Parameters.AddWithValue("@CSID", attachmentInfo.CSID);
                        cmd.Parameters.AddWithValue("@FileName", attachmentInfo.FileName);
                        cmd.Parameters.AddWithValue("@FileType", attachmentInfo.FileType);
                        cmd.Parameters.AddWithValue("@FileSize", attachmentInfo.FileSize);
                        // Add named parameters here to the command if needed...
                        cmd.ExecuteNonQuery();
                    }
                }
                catch (Exception ex)
                {
                    logger.Info(ex.Message + ex.StackTrace);
                }
                dbConnection.Close();
            }
        }
        internal static ExtractedEmailInfo GetExtractedEmailInfo(long emailID)
        {
            ExtractedEmailInfo extractedEmailInfo = null;
            using (DataTable table = new DataTable())
            {
                using (var conn = new SqlConnection(GetConnectionString()))
                {
                    using (var cmd = conn.CreateCommand())
                    {
                        String SQL = "SELECT top 1 * FROM EmailBrowser_ExtractedEmails where CSID = @NodeID ";

                        SqlDataAdapter adapter = new SqlDataAdapter();
                        cmd.CommandText = SQL;
                        cmd.Parameters.AddWithValue("@NodeID", emailID);
                        // Add named parameters here to the command if needed...
                        adapter.SelectCommand = cmd;
                        adapter.Fill(table);
                    }
                    conn.Close();
                }
                if (table.Rows.Count > 0)
                {
                    extractedEmailInfo = new ExtractedEmailInfo();
                    DataRow row = table.Rows[0];
                    if (!row.IsNull("CSID"))
                    {
                        extractedEmailInfo.CSID = (long)row["CSID"];
                    }
                    if (!row.IsNull("CSName"))
                    {
                        extractedEmailInfo.CSName = (string)row["CSName"];
                    }
                    if (!row.IsNull("CSVersionNum"))
                    {
                        extractedEmailInfo.CSVersionNum = (long)row["CSVersionNum"];
                    }
                    if (!row.IsNull("CSModifyDate"))
                    {
                        extractedEmailInfo.CSModifyDate = (DateTime)row["CSModifyDate"];
                    }
                    if (!row.IsNull("CSParentID"))
                    {
                        extractedEmailInfo.CSParentID = (long)row["CSParentID"];
                    }
                    if (!row.IsNull("CSCachingFolderNodeID"))
                    {
                        extractedEmailInfo.CSCachingFolderNodeID = (long)row["CSCachingFolderNodeID"];
                    }
                    if (!row.IsNull("IsExtracted"))
                    {
                        extractedEmailInfo.IsExtracted = (bool)(row["IsExtracted"]);
                    }
                    if (!row.IsNull("IsExtractingNow"))
                    {
                        extractedEmailInfo.IsExtractingNow = (bool)(row["IsExtractingNow"]);
                    }
                    if (!row.IsNull("LastExtractedDate"))
                    {
                        extractedEmailInfo.LastExtractedDate = (DateTime)row["LastExtractedDate"];
                    }
                }
            }
            return extractedEmailInfo;
        }
        #endregion


        public static bool IsFromMigratedFolder(string nodeId)
        {
            bool isMigratedEmail = false;

            try
            {
                DataTable table = new DataTable();
                using (var conn = new SqlConnection(GetConnectionString()))
                {
                    using (var cmd = conn.CreateCommand())
                    {
                        String SQL = "SELECT top 1 * FROM DTreeAncestors where DataID = @NodeID AND AncestorID=@AncestorID ";

                        SqlDataAdapter adapter = new SqlDataAdapter();
                        cmd.CommandText = SQL;
                        cmd.Parameters.AddWithValue("@NodeID", int.Parse(nodeId));
                        cmd.Parameters.AddWithValue("@AncestorID", int.Parse(Properties.Settings.Default.ArchiveFolderID.ToString()));

                        adapter.SelectCommand = cmd;
                        adapter.Fill(table);

                        if (table.Rows.Count > 0)
                        {
                            isMigratedEmail = true;
                        }

                    }
                    conn.Close();
                }
            }
            catch (Exception ex)
            {
                logger.Error(ex.Message + "\r\n" + ex.StackTrace.ToString());
            }

            return isMigratedEmail;
        }

        internal static List<FolderInfo> GetNavigationPath(long folderID)
        {
            List<FolderInfo> infos = new List<FolderInfo>();
            using (DataTable table = new DataTable())
            {
                using (var conn = new SqlConnection(GetConnectionString()))
                {
                    using (var cmd = conn.CreateCommand())
                    {
                        String SQL = "select dt.ParentID, dt.DataID, dt.Name, dt.ChildCount as 'ChildCount' from DTreeAncestors dta left join DTree dt on dt.DataID=dta.AncestorID where dta.DataID=@NodeID and (dt.SubType=0 OR dt.SubType=141)";
                        SqlDataAdapter adapter = new SqlDataAdapter();
                        cmd.CommandText = SQL;
                        cmd.Parameters.AddWithValue("@NodeID", folderID);
                        // Add named parameters here to the command if needed...
                        adapter.SelectCommand = cmd;
                        adapter.Fill(table);
                    }
                    conn.Close();
                }
                foreach (DataRow row in table.Rows)
                {
                    var folderInfo = new FolderInfo()
                    {
                        ParentNodeID = long.Parse(row["ParentID"].ToString()),
                        NodeID = long.Parse(row["DataID"].ToString()),
                        Name = row["Name"].ToString(),
                        ChildCount = long.Parse(row["ChildCount"].ToString()),
                        NodeType = "Folder"
                    };

                    if (folderInfo.NodeID == 2000)
                    {
                        folderInfo.NodeType = "EnterpriseWS";
                    }
                    infos.Add(folderInfo);
                }
            }
            return infos;
        }

        internal static string GetUserGroups(long userid)
        {
            string userGroupStr = "";
            try
            {
                using (DataTable table = new DataTable())
                {
                    using (var conn = new SqlConnection(GetConnectionString()))
                    {
                        using (var cmd = conn.CreateCommand())
                        {
                            string sql = @"SELECT KRL.RLRightID FROM KUAFRightsListNew KRL WHERE RLID=@UserID AND RLProxyType=0";
                            SqlDataAdapter adapter = new SqlDataAdapter();
                            cmd.CommandText = sql;
                            cmd.Parameters.AddWithValue("@UserID", userid);
                            // Add named parameters here to the command if needed...
                            adapter.SelectCommand = cmd;
                            adapter.Fill(table);
                        }
                        conn.Close();
                        foreach (DataRow row in table.Rows)
                        {
                            userGroupStr += $"{row["RLRightID"].ToString()},";
                        }
                    }
                }
            }
            catch (Exception ex)
            {
            }
            return userGroupStr.Trim().Trim(',');
        }
        internal static string GetUserGroups(string username)
        {
            string userGroupStr = "";
            try
            {
                using (DataTable table = new DataTable())
                {
                    using (var conn = new SqlConnection(GetConnectionString()))
                    {
                        using (var cmd = conn.CreateCommand())
                        {
                            string sql = @"SELECT KRL.RLRightID FROM KUAFRightsListNew KRL WHERE RLID IN (SELECT ID FROM KUAF WHERE Name=@UserName) AND RLProxyType=0";
                            SqlDataAdapter adapter = new SqlDataAdapter();
                            cmd.CommandText = sql;
                            cmd.Parameters.AddWithValue("@UserName", username);
                            // Add named parameters here to the command if needed...
                            adapter.SelectCommand = cmd;
                            adapter.Fill(table);
                        }
                        conn.Close();
                        foreach (DataRow row in table.Rows)
                        {
                            userGroupStr += $"{row["RLRightID"].ToString()},";
                        }
                    }
                }
            }
            catch (Exception ex)
            {
            }
            return userGroupStr.Trim().Trim(',');
        }

        internal static ListPaginatedEmailsResult GetPaginatedEmails(long folderID, int pageNo, int pageSize, string sortedBy, string sortDir, string userGroups)
        {
            ListPaginatedEmailsResult result = new ListPaginatedEmailsResult();
            List<EmailInfo> emails = new List<EmailInfo>();
            string sql = "";
            logger.Info($"folder ID : {folderID}");
            logger.Info($"pageSize : {pageSize}");
            logger.Info($"minIndex : {userGroups}");
            try
            {
                if (!string.IsNullOrEmpty(userGroups))
                {
                    using (DataTable table = new DataTable())
                    {
                        using (var conn = new SqlConnection(GetConnectionString()))
                        {
                            using (var cmd = conn.CreateCommand())
                            {
                                long minIndex = ((pageNo - 1) * pageSize);
                                string orderByColumn = "";
                                if (string.IsNullOrEmpty(sortDir))
                                {
                                    sortDir = "asc";
                                }

                                if (sortDir.ToLower() == "asc")
                                {
                                    switch (sortedBy)
                                    {
                                        case "OTEmailTo":
                                            orderByColumn = $"OTEmailTo asc";
                                            break;
                                        case "OTEmailFrom":
                                            orderByColumn = $"OTEmailFrom asc";
                                            break;
                                        case "OTEmailSubject":
                                            orderByColumn = $"OTEmailSubject asc";
                                            break;
                                        case "OTEmailReceivedDate":
                                            orderByColumn = $"OTEmailReceivedDate asc";
                                            break;
                                        case "OTEmailSentDate":
                                            orderByColumn = $"OTEmailSentDate asc";
                                            break;
                                        default:
                                            orderByColumn = "OTEmailSentDate asc";
                                            break;
                                    }
                                }
                                else
                                {
                                    switch (sortedBy)
                                    {
                                        case "OTEmailTo":
                                            orderByColumn = $"OTEmailTo desc";
                                            break;
                                        case "OTEmailFrom":
                                            orderByColumn = $"OTEmailFrom desc";
                                            break;
                                        case "OTEmailSubject":
                                            orderByColumn = $"OTEmailSubject desc";
                                            break;
                                        case "OTEmailReceivedDate":
                                            orderByColumn = $"OTEmailReceivedDate desc";
                                            break;
                                        case "OTEmailSentDate":
                                            orderByColumn = $"OTEmailSentDate desc";
                                            break;
                                        default:
                                            orderByColumn = "OTEmailSentDate desc";
                                            break;
                                    }
                                }

                                sql = string.Format(@"with Nodes(DataID, See) AS (SELECT acl.DataID, max(acl.See) FROM DTreeACL acl WHERE acl.ParentID=@FolderID and acl.RightID IN ({0}) and acl.See > 0 group by acl.DataID),
MsgNodes(DataID, See) AS (SELECT Nodes.DataID, Nodes.See FROM Nodes WHERE DataID NOT IN (SELECT DataID FROM DDeletedItemsNodes)),
EmailProperties(RowID, DataID, ConversationID, OTEmailFrom, OTEmailTo, HasAttachments, OTEmailSubject, OTEmailSentDate, See, TotalCount) AS (SELECT RowID=ROW_NUMBER() OVER (ORDER BY {1}), ote.NodeID, otec.ConversationID, ote.OTEmailFrom, ote.OTEmailTo, ote.HasAttachments, ote.OTEmailSubject, ote.OTEmailSentDate, msg.See, COUNT(*) OVER() FROM OTEmailProperties ote LEFT JOIN OTEmailConversationIDs otec on otec.NodeID=ote.NodeID INNER JOIN MsgNodes msg on msg.DataID=ote.NodeId WHERE ote.NodeID IN (SELECT DataID FROM MsgNodes)),
PaginatedResult(RowID, DataID, ConversationID, OTEmailFrom, OTEmailTo, HasAttachments, OTEmailSubject, OTEmailSentDate, See, TotalCount) AS (SELECT TOP (@PageSize) RowID, DataID, ConversationID, OTEmailFrom, OTEmailTo, HasAttachments, OTEmailSubject, OTEmailSentDate, See, TotalCount FROM EmailProperties WHERE RowID > @Min)
SELECT DataID, ConversationID, OTEmailFrom, OTEmailTo, HasAttachments, OTEmailSubject, OTEmailSentDate, See, TotalCount FROM PaginatedResult", userGroups, orderByColumn);
                                SqlDataAdapter adapter = new SqlDataAdapter();
                                cmd.CommandText = sql;
                                cmd.Parameters.AddWithValue("@FolderID", folderID);
                                cmd.Parameters.AddWithValue("@PageSize", pageSize);
                                cmd.Parameters.AddWithValue("@Min", minIndex);
                                cmd.CommandTimeout = 600; // 10 Minutes timeout
                                adapter.SelectCommand = cmd;
                                adapter.Fill(table);
                            }
                            logger.Info($"folder ID : {folderID}");
                            logger.Info($"pageSize : {pageSize}");
                            logger.Info($"minIndex : {userGroups}");
                            conn.Close();
                            long totalItemCount = 0;
                            if (table.Rows.Count > 0)
                            {
                                totalItemCount = long.Parse(table.Rows[0]["TotalCount"].ToString());
                            }
                            foreach (DataRow row in table.Rows)
                            {
                                long nodeParentID = folderID;
                                long nodeID = long.Parse(row["DataID"].ToString());
                                string convId = row["ConversationID"].ToString();
                                string from = row["OTEmailFrom"].ToString();
                                string to = row["OTEmailTo"].ToString();
                                string subject = row["OTEmailSubject"].ToString();
                                int hasAttachments = int.Parse(row["HasAttachments"].ToString());
                                DateTime sentDateVal;
                                int see = int.Parse(row["See"].ToString());
                                DateTime? sentDate;
                                bool hasValidSentDate = DateTime.TryParse(row["OTEmailSentDate"].ToString(), out sentDateVal);

                                if (!hasValidSentDate)
                                {
                                    sentDate = null;
                                }
                                else
                                {
                                    sentDate = sentDateVal;
                                }

                                var emailInfo = new EmailInfo()
                                {
                                    ParentNodeID = folderID,
                                    NodeID = nodeID,
                                    ConversationId = convId,
                                    EmailFrom = from,
                                    EmailTo = to,
                                    EmailSubject = subject,
                                    HasAttachments = hasAttachments,
                                    SentDate = sentDate,
                                    PermSeeContents = see >= 2 ? true : false,
                                };
                                logger.Info($"email info : {emailInfo}");


                                emails.Add(emailInfo);
                            }
                            result.PageNumber = pageNo;
                            result.PageSize = pageSize;
                            result.TotalCount = (int)totalItemCount;
                            result.NumberOfPages = (int)Math.Ceiling(((decimal)totalItemCount / pageSize));
                            result.EmailInfos = emails;
                            result.SortedBy = sortedBy;
                            result.SortDirection = sortDir;
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                logger.Error($"Something went wrong when trying to retrieve the list of emails ({sql}): {ex.Message}\t{ex.StackTrace}");
            }
            return result;
        }

        internal static List<EmailInfo> GetEmailSentDates(string nodeIds)
        {
            List<EmailInfo> emailInfos = new List<EmailInfo>();
            string sql = "";

            try
            {
                var nodeidList = nodeIds.Split(',');
                foreach (var nodeid in nodeidList)
                {
                    long nodeID;

                    if (!long.TryParse(nodeid, out nodeID))
                    {
                        return emailInfos;
                    }
                }

                using (DataTable table = new DataTable())
                {
                    using (var conn = new SqlConnection(GetConnectionString()))
                    {
                        using (var cmd = conn.CreateCommand())
                        {
                            sql = string.Format(@"select NodeID, OTEmailSentDate FROM OTEmailProperties WHERE NodeID IN ({0})", nodeIds);
                            SqlDataAdapter adapter = new SqlDataAdapter();
                            cmd.CommandText = sql;
                            adapter.SelectCommand = cmd;
                            adapter.Fill(table);
                        }
                        conn.Close();

                        foreach (DataRow row in table.Rows)
                        {
                            long nodeID = long.Parse(row["NodeID"].ToString());
                            DateTime sentDate;

                            DateTime.TryParse(row["OTEmailSentDate"].ToString(), out sentDate);

                            var emailInfo = new EmailInfo()
                            {
                                ParentNodeID = -1,
                                NodeID = nodeID,
                                ConversationId = "",
                                EmailFrom = "",
                                EmailTo = "",
                                EmailSubject = "",
                                HasAttachments = -1,
                                SentDate = sentDate
                            };
                            emailInfos.Add(emailInfo);
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                logger.Error($"Something went wrong when trying to retrieve the list of email sentdates ({sql}): {ex.Message}\t{ex.StackTrace}");
            }

            return emailInfos;
        }
    }
}