using HC10AutomationFramework.Logs;
using System;
using System.Data.SqlClient;

namespace HC10AutomationFramework.Helpers
{
    public static class DataHelperExtensions
    {
        public static SqlConnection DBConnect(this SqlConnection sqlConnection, string connectionString)
        {
            try
            {
                sqlConnection = new SqlConnection(connectionString);
                sqlConnection.Open();
                return sqlConnection;
            }

            catch(Exception e)
            {
                LogClass.AppendLogs(e);
                
            }
            return null;
        }

        public static void DBClose(this SqlConnection sqlConnection) 
        {
            try
            {
                sqlConnection.Close();
            }

            catch (Exception e)
            {
                LogClass.AppendLogs(e);
            }
        }

        public static void ExecuteQuery(this SqlConnection sqlConnection, string queryString)

        {
            
        }
    }
}
