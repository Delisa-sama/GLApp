using System;
using System.Data.SqlClient;

namespace GLApp
{
    static class UserData
    {
        static public SqlConnection Connection = null;
        static public string uname = "";
        static public string pwd = "";
        static public bool SuccLogin = false;
        static public Int16 UserType = 0;
        static public Int32 ID = -1;
        static public void OpenConnection()
        {
            if (Connection.State != System.Data.ConnectionState.Open)
            {
                Connection.Open();
            }
        }
        static public void CloseConnection()
        {
            if (Connection.State == System.Data.ConnectionState.Open)
            {
                Connection.Close();
            }
        }
        static public void SetUD(string Username, string Password)
        {
            uname = Username;
            pwd = Password;
        }
        
        static readonly public string ConnectionString = @"Data Source=" + Environment.MachineName + @"\sqlcore;Initial Catalog=GLPR;Connect Timeout=3;";
        static public string GetCS()
        {
            return ConnectionString + "User ID=" + uname + ";Password=" + pwd + ";";
        }
    }
}
