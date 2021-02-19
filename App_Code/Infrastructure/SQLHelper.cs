using System;
using System.Collections;
using System.Collections.Generic;
using System.Data;
using System.Diagnostics;
using System.Linq;
using System.Text;
using System.Data.SqlClient;
using System.Xml;
using System.IO;
using System.Web;
using System.Configuration;
using FilesToPDFConvertor.Infrastructure;

namespace FilesToPDFConvertor
{
    public class SQLHelper
    {
        #region "Get Connection string"

        //New Added
        public static string SMTP_SERVER = "";
        public static string NETWORK_USER = string.Empty;
        public static string NETWORK_PSWRD = string.Empty;

        public static string PORT = string.Empty;
        public static string REPORT_SERVER = string.Empty;
        public static string REPORT_USER = string.Empty;
        public static string REPORT_PSWRD = string.Empty;
        public static string REPORT_FOLDER = string.Empty;
        public static string REPORT_DOMAIN = string.Empty;

        public static string REPORT_VIEWER_URL = string.Empty;
        public static string AD_DOMAIN = string.Empty;
        public static string AD_USER = string.Empty;

        public static string AD_PSWRD = string.Empty;
        public static string AD_DOMAIN_JLLR = string.Empty;
        public static string AD_USER_JLLR = string.Empty;

        public static string AD_PSWRD_JLLR = string.Empty;
        public static string GetConnString()
        {

            string connectionString = "";
            //System.Configuration.AppSettingsReader settingsReader = new AppSettingsReader();
            // Get the key from config file
            string key = CryptorEngine.Decrypt(ConfigurationManager.AppSettings["ConnectionString"].ToString(), true);// @"Data Source=WIN-P99G4H9GTJL;Initial Catalog=MIS_JF;Persist Security Info=True;User ID=sa;Password=P@ssw0rd";// (string)settingsReader.GetValue("connectionString", typeof(String));
            connectionString = key;

            /*StreamReader objStreamReader = default(StreamReader);
            string[] config = new string[6];
            string[] fileName = new string[6];

            string basepath = System.AppDomain.CurrentDomain.BaseDirectory;
            objStreamReader = File.OpenText(basepath + "bin\\\\config.txt");

            int i = 1;
            while (!(objStreamReader.Peek() == -1) & i < 4)
            {
                fileName[i] = objStreamReader.ReadLine();
                i = i + 1;
            }
            objStreamReader.Close();

            objStreamReader = File.OpenText(fileName[1]);

            int k = 1;
            while (!(objStreamReader.Peek() == -1) & k < 10)
            {
                if (k <= 4)
                {
                    config[k] = CryptorEngine.Decrypt(objStreamReader.ReadLine(), true);
                }
                else if (k == 5)
                {
                    SMTP_SERVER = CryptorEngine.Decrypt(objStreamReader.ReadLine(), true);
                }
                else if (k == 6)
                {
                    NETWORK_USER = CryptorEngine.Decrypt(objStreamReader.ReadLine(), true);
                }
                else if (k == 7)
                {
                    NETWORK_PSWRD = CryptorEngine.Decrypt(objStreamReader.ReadLine(), true);
                }
                else if (k == 9)
                {
                    PORT = CryptorEngine.Decrypt(objStreamReader.ReadLine(),true);
                }

                k = k + 1;
            }

            objStreamReader.Close();


            //connectionString = "server=srcrm01;User ID=sa;Password=P#ssw0rd;database=SmartProcurement_Test;pooling=false;"
            connectionString = config[1] + ";" + config[2] + ";" + config[3] + ";" + config[4] + ";pooling=false;";

            objStreamReader = File.OpenText(fileName[2]);

            k = 1;
            while (!(objStreamReader.Peek() == -1) & k < 7)
            {
                if (k == 1)
                {
                    REPORT_SERVER = CryptorEngine.Decrypt(objStreamReader.ReadLine(), true);
                }
                else if (k == 2)
                {
                    REPORT_USER = CryptorEngine.Decrypt(objStreamReader.ReadLine(), true);
                }
                else if (k == 3)
                {
                    REPORT_PSWRD = CryptorEngine.Decrypt(objStreamReader.ReadLine(), true);
                }
                else if (k == 4)
                {
                    REPORT_FOLDER = CryptorEngine.Decrypt(objStreamReader.ReadLine(), true);
                }
                else if (k == 5)
                {
                    REPORT_DOMAIN = CryptorEngine.Decrypt(objStreamReader.ReadLine(), true);
                }
                else if (k == 6)
                {
                    REPORT_VIEWER_URL = CryptorEngine.Decrypt(objStreamReader.ReadLine(), true);
                }

                k = k + 1;
            }

            objStreamReader.Close();

            objStreamReader = File.OpenText(fileName[3]);

            k = 1;
            while (!(objStreamReader.Peek() == -1) & k < 4)
            {
                if (k == 1)
                {
                    AD_DOMAIN = CryptorEngine.Decrypt(objStreamReader.ReadLine(),true);
                }
                else if (k == 2)
                {
                    AD_USER = CryptorEngine.Decrypt(objStreamReader.ReadLine(), true);
                }
                else if (k == 3)
                {
                    AD_PSWRD = CryptorEngine.Decrypt(objStreamReader.ReadLine(), true);
                }

                k = k + 1;
            }

            objStreamReader.Close();*/

            return connectionString;
        }

        public static string GetDBName()
        {
            return CryptorEngine.Decrypt(ConfigurationManager.AppSettings["DBName"].ToString(), true);
        }

        public static string[] getADstring()
        {
            try
            {
                StreamReader objStreamReader = default(StreamReader);
                string[] config = new string[6];
                string[] fileName = new string[6];
                Int32 k;// = default(Int16);
                string connectionString = string.Empty;

                string basepath = System.AppDomain.CurrentDomain.BaseDirectory;
                objStreamReader = File.OpenText(basepath + "bin\\\\config.txt");

                int i = 1;
                while (!(objStreamReader.Peek() == -1) & i < 4)
                {
                    fileName[i] = objStreamReader.ReadLine();
                    i = i + 1;
                }
                objStreamReader.Close();

                objStreamReader = File.OpenText(fileName[3]);

                k = 1;
                while (!(objStreamReader.Peek() == -1) & k < 7)
                {
                    if (k == 1)
                    {
                        AD_DOMAIN = "LDAP://" + CryptorEngine.Decrypt(objStreamReader.ReadLine(), true) + ".com";
                    }
                    else if (k == 2)
                    {
                        AD_USER = CryptorEngine.Decrypt(objStreamReader.ReadLine(), true);
                    }
                    else if (k == 3)
                    {
                        AD_PSWRD = CryptorEngine.Decrypt(objStreamReader.ReadLine(), true);
                    }
                    else if (k == 4)
                    {
                        AD_DOMAIN_JLLR = "LDAP://" + CryptorEngine.Decrypt(objStreamReader.ReadLine(), true) + ".com";
                    }
                    else if (k == 5)
                    {
                        AD_USER_JLLR = CryptorEngine.Decrypt(objStreamReader.ReadLine(), true);
                    }
                    else if (k == 6)
                    {
                        AD_PSWRD_JLLR = CryptorEngine.Decrypt(objStreamReader.ReadLine(), true);
                    }

                    k = k + 1;
                }

                objStreamReader.Close();

                string[] arr = new string[7];
                arr[0] = AD_DOMAIN;
                arr[1] = AD_USER;
                arr[2] = AD_PSWRD;

                arr[3] = AD_DOMAIN_JLLR;
                arr[4] = AD_USER_JLLR;
                arr[5] = AD_PSWRD_JLLR;

                return arr;
            }
            catch (Exception ex)
            {
                throw new ApplicationException(ex.Message);
            }
        }

        #endregion

        #region "Get Bool Value"

        public static bool GetBooleanValue(string val)
        {
            if (val == "1")
            {
                return true;
            }
            else
            {
                return false;
            }
        }

        #endregion

        #region "Get Bool Value"

        public static bool GetBooleanValue(int val)
        {
            if (val == 1)
            {
                return true;
            }
            else
            {
                return false;
            }
        }

        #endregion

        #region "Get Value from Boolean"

        public static int GetValue(bool val)
        {
            if (val)
            {
                return 1;
            }
            else
            {
                return 0;
            }
        }

        #endregion

        #region "private utility methods & constructors"

        private SQLHelper()
        {
        }

        private static void AttachParameters(SqlCommand command, SqlParameter[] commandParameters)
        {
            if (command == null)
            {
                throw new ArgumentNullException("command");
            }
            if (commandParameters != null)
            {
                foreach (SqlParameter p in commandParameters)
                {
                    if (p != null)
                    {
                        if ((p.Direction == ParameterDirection.InputOutput || p.Direction == ParameterDirection.Input) && (p.Value == null))
                        {
                            p.Value = DBNull.Value;
                        }
                        command.Parameters.Add(p);
                    }
                }
            }
        }

        private static void AssignParameterValues(SqlParameter[] commandParameters, DataRow dataRow)
        {
            if ((commandParameters == null) || (dataRow == null))
            {
                return;
            }
            int i = 0;
            foreach (SqlParameter commandParameter in commandParameters)
            {
                if (commandParameter.ParameterName == null || commandParameter.ParameterName.Length <= 1)
                {
                    throw new Exception(string.Format("Please provide a valid parameter name on the parameter #{0}, the ParameterName property has the following value: '{1}'.", i, commandParameter.ParameterName));
                }
                if (dataRow.Table.Columns.IndexOf(commandParameter.ParameterName.Substring(1)) != -1)
                {
                    commandParameter.Value = dataRow[commandParameter.ParameterName.Substring(1)];
                }
                i += 1;
            }
        }

        private static void AssignParameterValues(SqlParameter[] commandParameters, object[] parameterValues)
        {
            if ((commandParameters == null) || (parameterValues == null))
            {
                return;
            }
            if (commandParameters.Length != parameterValues.Length)
            {
                throw new ArgumentException("Parameter count does not match Parameter Value count.");
            }
            int i = 0;
            int j = commandParameters.Length;
            while (i < j)
            {
                if (parameterValues[i] is IDbDataParameter)
                {
                    IDbDataParameter paramInstance = (IDbDataParameter)parameterValues[i];
                    if (paramInstance.Value == null)
                    {
                        commandParameters[i].Value = DBNull.Value;
                    }
                    else
                    {
                        commandParameters[i].Value = paramInstance.Value;
                    }
                }
                else if (parameterValues[i] == null)
                {
                    commandParameters[i].Value = DBNull.Value;
                }
                else
                {
                    commandParameters[i].Value = parameterValues[i];
                }
                i += 1;
            }
        }

        private static void PrepareCommand(SqlCommand command, SqlConnection connection, SqlTransaction transaction, CommandType commandType, string commandText, SqlParameter[] commandParameters, ref bool mustCloseConnection)
        {
            if (command == null)
            {
                throw new ArgumentNullException("command");
            }
            if (commandText == null || commandText.Length == 0)
            {
                throw new ArgumentNullException("commandText");
            }
            if (connection.State != ConnectionState.Open)
            {
                mustCloseConnection = true;
                connection.Open();
            }
            else
            {
                mustCloseConnection = false;
            }
            command.Connection = connection;
            command.CommandText = commandText;
            if (transaction != null)
            {
                if (transaction.Connection == null)
                {
                    throw new ArgumentException("The transaction was rollbacked or commited, please provide an open transaction.", "transaction");
                }
                command.Transaction = transaction;
            }
            command.CommandType = commandType;
            if (commandParameters != null)
            {
                AttachParameters(command, commandParameters);
            }
            return;
        }

        #endregion

        #region "ExecuteNonQuery"

        public static int ExecuteNonQuery(string connectionString, CommandType commandType, string commandText)
        {
            return ExecuteNonQuery(connectionString, commandType, commandText, (SqlParameter[])null);
        }

        public static int ExecuteNonQuery(string connectionString, CommandType commandType, string commandText, params SqlParameter[] commandParameters)
        {
            if (connectionString == null || connectionString.Length == 0)
            {
                throw new ArgumentNullException("connectionString");
            }
            using (SqlConnection connection = new SqlConnection(connectionString))
            {
                connection.Open();
                return ExecuteNonQuery(connection, commandType, commandText, commandParameters);
            }
        }

        public static int ExecuteNonQuery(string connectionString, CommandType commandType, string commandText, string moduleDatabase, params SqlParameter[] commandParameters)
        {
            if (connectionString == null || connectionString.Length == 0)
            {
                throw new ArgumentNullException("connectionString");
            }
            using (SqlConnection connection = new SqlConnection(connectionString))
            {
                connection.Open();
                connection.ChangeDatabase(moduleDatabase);
                return ExecuteNonQuery(connection, commandType, commandText, commandParameters);
            }
        }

        public static int ExecuteNonQuery(SqlConnection connection, CommandType commandType, string commandText)
        {
            return ExecuteNonQuery(connection, commandType, commandText, (SqlParameter[])null);
        }

        public static int ExecuteNonQuery(SqlConnection connection, CommandType commandType, string commandText, params SqlParameter[] commandParameters)
        {
            if (connection == null)
            {
                throw new ArgumentNullException("connection");
            }
            SqlCommand cmd = new SqlCommand();
            bool mustCloseConnection = false;
            PrepareCommand(cmd, connection, (SqlTransaction)null, commandType, commandText, commandParameters, ref mustCloseConnection);
            cmd.CommandTimeout = 300;
            int retval = cmd.ExecuteNonQuery();
            cmd.Parameters.Clear();
            if (mustCloseConnection)
            {
                connection.Close();
            }
            return retval;
        }

        public static int ExecuteNonQuery(SqlTransaction transaction, CommandType commandType, string commandText)
        {
            return ExecuteNonQuery(transaction, commandType, commandText, (SqlParameter[])null);
        }

        public static int ExecuteNonQuery(SqlTransaction transaction, CommandType commandType, string commandText, params SqlParameter[] commandParameters)
        {
            if (transaction == null)
            {
                throw new ArgumentNullException("transaction");
            }
            if (transaction != null && transaction.Connection == null)
            {
                throw new ArgumentException("The transaction was rollbacked or commited, please provide an open transaction.", "transaction");
            }
            SqlCommand cmd = new SqlCommand();
            bool mustCloseConnection = false;
            PrepareCommand(cmd, transaction.Connection, transaction, commandType, commandText, commandParameters, ref mustCloseConnection);
            int retval = cmd.ExecuteNonQuery();
            cmd.Parameters.Clear();
            return retval;
        }

        #endregion

        #region "ExecuteDataset"

        public static DataSet ExecuteDataset(string connectionString, CommandType commandType, string commandText)
        {
            return ExecuteDataset(connectionString, commandType, commandText, (SqlParameter[])null);
        }

        public static DataSet ExecuteDataset(string connectionString, CommandType commandType, string commandText, params SqlParameter[] commandParameters)
        {
            if (connectionString == null || connectionString.Length == 0)
            {
                throw new ArgumentNullException("connectionString");
            }
            using (SqlConnection connection = new SqlConnection(connectionString))
            {
                connection.Open();
                return ExecuteDataset(connection, commandType, commandText, commandParameters);
            }
        }

        public static DataSet ExecuteDataset(string connectionString, CommandType commandType, string commandText, string moduleDatabase, params SqlParameter[] commandParameters)
        {
            if (connectionString == null || connectionString.Length == 0)
            {
                throw new ArgumentNullException("connectionString");
            }
            using (SqlConnection connection = new SqlConnection(connectionString))
            {
                connection.Open();
                connection.ChangeDatabase(moduleDatabase);
                return ExecuteDataset(connection, commandType, commandText, commandParameters);
            }
        }

        public static DataSet ExecuteDataset(SqlConnection connection, CommandType commandType, string commandText)
        {
            return ExecuteDataset(connection, commandType, commandText, (SqlParameter[])null);
        }

        public static DataSet ExecuteDataset(SqlConnection connection, CommandType commandType, string commandText, params SqlParameter[] commandParameters)
        {
            if (connection == null)
            {
                throw new ArgumentNullException("connection");
            }
            SqlCommand cmd = new SqlCommand();
            bool mustCloseConnection = false;
            PrepareCommand(cmd, connection, (SqlTransaction)null, commandType, commandText, commandParameters, ref mustCloseConnection);
            cmd.CommandTimeout = 600;
            using (SqlDataAdapter da = new SqlDataAdapter(cmd))
            {
                DataSet ds = new DataSet();
                da.Fill(ds);
                cmd.Parameters.Clear();
                if (mustCloseConnection)
                {
                    connection.Close();
                }
                return ds;
            }
        }

        public static DataSet ExecuteDataset(SqlTransaction transaction, CommandType commandType, string commandText)
        {
            return ExecuteDataset(transaction, commandType, commandText, (SqlParameter[])null);
        }

        public static DataSet ExecuteDataset(SqlTransaction transaction, CommandType commandType, string commandText, params SqlParameter[] commandParameters)
        {
            if (transaction == null)
            {
                throw new ArgumentNullException("transaction");
            }
            if (transaction != null && transaction.Connection == null)
            {
                throw new ArgumentException("The transaction was rollbacked or commited, please provide an open transaction.", "transaction");
            }
            SqlCommand cmd = new SqlCommand();
            bool mustCloseConnection = false;
            PrepareCommand(cmd, transaction.Connection, transaction, commandType, commandText, commandParameters, ref mustCloseConnection);
            using (SqlDataAdapter da = new SqlDataAdapter(cmd))
            {
                DataSet ds = new DataSet();
                da.Fill(ds);
                cmd.Parameters.Clear();
                return ds;
            }
        }

        #endregion

        #region "ExecuteReader"

        private enum SqlConnectionOwnership
        {
            Internal,
            External
        }

        private static SqlDataReader ExecuteReader(SqlConnection connection, SqlTransaction transaction, CommandType commandType, string commandText, SqlParameter[] commandParameters, SqlConnectionOwnership connectionOwnership)
        {
            if (connection == null)
            {
                throw new ArgumentNullException("connection");
            }
            bool mustCloseConnection = false;
            SqlCommand cmd = new SqlCommand();
            try
            {
                PrepareCommand(cmd, connection, transaction, commandType, commandText, commandParameters, ref mustCloseConnection);
                SqlDataReader dataReader = default(SqlDataReader);
                if (connectionOwnership == SqlConnectionOwnership.External)
                {
                    dataReader = cmd.ExecuteReader();
                }
                else
                {
                    dataReader = cmd.ExecuteReader(CommandBehavior.CloseConnection);
                }
                bool canClear = true;
                foreach (SqlParameter commandParameter in cmd.Parameters)
                {
                    if (commandParameter.Direction != ParameterDirection.Input)
                    {
                        canClear = false;
                    }
                }
                if (canClear)
                {
                    cmd.Parameters.Clear();
                }
                return dataReader;
            }
            catch
            {
                if (mustCloseConnection)
                {
                    connection.Close();
                }
                throw;
            }
        }

        public static SqlDataReader ExecuteReader(string connectionString, CommandType commandType, string commandText)
        {
            return ExecuteReader(connectionString, commandType, commandText, (SqlParameter[])null);
        }

        public static SqlDataReader ExecuteReader(string connectionString, CommandType commandType, string commandText, params SqlParameter[] commandParameters)
        {
            if (connectionString == null || connectionString.Length == 0)
            {
                throw new ArgumentNullException("connectionString");
            }
            SqlConnection connection = null;
            try
            {
                connection = new SqlConnection(connectionString);
                connection.Open();
                return ExecuteReader(connection, null, commandType, commandText, commandParameters, SqlConnectionOwnership.Internal);
            }
            catch
            {
                if (connection != null)
                {
                    connection.Close();
                }
                throw;
            }
        }

        public static SqlDataReader ExecuteReader(string connectionString, CommandType commandType, string commandText, string moduleDatabase, params SqlParameter[] commandParameters)
        {
            if (connectionString == null || connectionString.Length == 0)
            {
                throw new ArgumentNullException("connectionString");
            }
            SqlConnection connection = null;
            try
            {
                connection = new SqlConnection(connectionString);
                connection.Open();
                connection.ChangeDatabase(moduleDatabase);
                return ExecuteReader(connection, null, commandType, commandText, commandParameters, SqlConnectionOwnership.Internal);
            }
            catch
            {
                if (connection != null)
                {
                    connection.Close();
                }
                throw;
            }
        }

        public static SqlDataReader ExecuteReader(SqlConnection connection, CommandType commandType, string commandText)
        {
            return ExecuteReader(connection, commandType, commandText, (SqlParameter[])null);
        }

        public static SqlDataReader ExecuteReader(SqlConnection connection, CommandType commandType, string commandText, params SqlParameter[] commandParameters)
        {
            return ExecuteReader(connection, (SqlTransaction)null, commandType, commandText, commandParameters, SqlConnectionOwnership.External);
        }

        public static SqlDataReader ExecuteReader(SqlTransaction transaction, CommandType commandType, string commandText)
        {
            return ExecuteReader(transaction, commandType, commandText, (SqlParameter[])null);
        }

        public static SqlDataReader ExecuteReader(SqlTransaction transaction, CommandType commandType, string commandText, params SqlParameter[] commandParameters)
        {
            if (transaction == null)
            {
                throw new ArgumentNullException("transaction");
            }
            if (transaction != null && transaction.Connection == null)
            {
                throw new ArgumentException("The transaction was rollbacked or commited, please provide an open transaction.", "transaction");
            }
            return ExecuteReader(transaction.Connection, transaction, commandType, commandText, commandParameters, SqlConnectionOwnership.External);
        }

        #endregion

        #region "ExecuteScalar"

        public static object ExecuteScalar(string connectionString, CommandType commandType, string commandText)
        {
            return ExecuteScalar(connectionString, commandType, commandText, (SqlParameter[])null);
        }

        public static object ExecuteScalar(string connectionString, CommandType commandType, string commandText, params SqlParameter[] commandParameters)
        {
            if (connectionString == null || connectionString.Length == 0)
            {
                throw new ArgumentNullException("connectionString");
            }
            using (SqlConnection connection = new SqlConnection(connectionString))
            {
                connection.Open();
                return ExecuteScalar(connection, commandType, commandText, commandParameters);
            }
        }

        public static object ExecuteScalar(string connectionString, CommandType commandType, string commandText, string moduleDatabase, params SqlParameter[] commandParameters)
        {
            if (connectionString == null || connectionString.Length == 0)
            {
                throw new ArgumentNullException("connectionString");
            }
            using (SqlConnection connection = new SqlConnection(connectionString))
            {
                connection.Open();
                connection.ChangeDatabase(moduleDatabase);
                return ExecuteScalar(connection, commandType, commandText, commandParameters);
            }
        }

        public static object ExecuteScalar(SqlConnection connection, CommandType commandType, string commandText)
        {
            return ExecuteScalar(connection, commandType, commandText, (SqlParameter[])null);
        }

        public static object ExecuteScalar(SqlConnection connection, CommandType commandType, string commandText, params SqlParameter[] commandParameters)
        {
            if (connection == null)
            {
                throw new ArgumentNullException("connection");
            }
            SqlCommand cmd = new SqlCommand();
            bool mustCloseConnection = false;
            PrepareCommand(cmd, connection, (SqlTransaction)null, commandType, commandText, commandParameters, ref mustCloseConnection);
            cmd.CommandTimeout = 100000;
            object retval = cmd.ExecuteScalar();
            cmd.Parameters.Clear();
            if (mustCloseConnection)
            {
                connection.Close();
            }
            return retval;
        }

        public static object ExecuteScalar(SqlTransaction transaction, CommandType commandType, string commandText)
        {
            return ExecuteScalar(transaction, commandType, commandText, (SqlParameter[])null);
        }

        public static object ExecuteScalar(SqlTransaction transaction, CommandType commandType, string commandText, params SqlParameter[] commandParameters)
        {
            if (transaction == null)
            {
                throw new ArgumentNullException("transaction");
            }
            if (transaction != null && transaction.Connection == null)
            {
                throw new ArgumentException("The transaction was rollbacked or commited, please provide an open transaction.", "transaction");
            }
            SqlCommand cmd = new SqlCommand();
            cmd.CommandTimeout = 100000;
            bool mustCloseConnection = false;
            PrepareCommand(cmd, transaction.Connection, transaction, commandType, commandText, commandParameters, ref mustCloseConnection);
            object retval = cmd.ExecuteScalar();
            cmd.Parameters.Clear();
            return retval;
        }

        #endregion

        #region "ExecuteXmlReader"

        public static System.Xml.XmlReader ExecuteXmlReader(SqlConnection connection, CommandType commandType, string commandText)
        {
            return ExecuteXmlReader(connection, commandType, commandText, (SqlParameter[])null);
        }

        public static XmlReader ExecuteXmlReader(SqlConnection connection, CommandType commandType, string commandText, params SqlParameter[] commandParameters)
        {
            if (connection == null)
            {
                throw new ArgumentNullException("connection");
            }
            bool mustCloseConnection = false;
            SqlCommand cmd = new SqlCommand();
            try
            {
                PrepareCommand(cmd, connection, (SqlTransaction)null, commandType, commandText, commandParameters, ref mustCloseConnection);
                XmlReader retval = cmd.ExecuteXmlReader();
                cmd.Parameters.Clear();
                return retval;
            }
            catch
            {
                if (mustCloseConnection)
                {
                    connection.Close();
                }
                throw;
            }
        }

        public static XmlReader ExecuteXmlReader(SqlTransaction transaction, CommandType commandType, string commandText)
        {
            return ExecuteXmlReader(transaction, commandType, commandText, (SqlParameter[])null);
        }

        public static XmlReader ExecuteXmlReader(SqlTransaction transaction, CommandType commandType, string commandText, params SqlParameter[] commandParameters)
        {
            if (transaction == null)
            {
                throw new ArgumentNullException("transaction");
            }
            if (transaction != null && transaction.Connection == null)
            {
                throw new ArgumentException("The transaction was rollbacked or commited, please provide an open transaction.", "transaction");
            }
            SqlCommand cmd = new SqlCommand();
            bool mustCloseConnection = false;
            PrepareCommand(cmd, transaction.Connection, transaction, commandType, commandText, commandParameters, ref mustCloseConnection);
            XmlReader retval = cmd.ExecuteXmlReader();
            cmd.Parameters.Clear();
            return retval;
        }

        #endregion

        #region "FillDataset"

        public static void FillDataset(string connectionString, CommandType commandType, string commandText, DataSet dataSet, string[] tableNames)
        {
            if (connectionString == null || connectionString.Length == 0)
            {
                throw new ArgumentNullException("connectionString");
            }
            if (dataSet == null)
            {
                throw new ArgumentNullException("dataSet");
            }
            using (SqlConnection connection = new SqlConnection(connectionString))
            {
                connection.Open();
                FillDataset(connection, commandType, commandText, dataSet, tableNames);
            }
        }

        public static void FillDataset(string connectionString, CommandType commandType, string commandText, DataSet dataSet, string[] tableNames, params SqlParameter[] commandParameters)
        {
            if (connectionString == null || connectionString.Length == 0)
            {
                throw new ArgumentNullException("connectionString");
            }
            if (dataSet == null)
            {
                throw new ArgumentNullException("dataSet");
            }
            using (SqlConnection connection = new SqlConnection(connectionString))
            {
                connection.Open();
                FillDataset(connection, commandType, commandText, dataSet, tableNames, commandParameters);
            }
        }

        public static void FillDataset(SqlConnection connection, CommandType commandType, string commandText, DataSet dataSet, string[] tableNames)
        {
            FillDataset(connection, commandType, commandText, dataSet, tableNames, null);
        }

        public static void FillDataset(SqlConnection connection, CommandType commandType, string commandText, DataSet dataSet, string[] tableNames, params SqlParameter[] commandParameters)
        {
            FillDataset(connection, null, commandType, commandText, dataSet, tableNames, commandParameters);
        }

        public static void FillDataset(SqlTransaction transaction, CommandType commandType, string commandText, DataSet dataSet, string[] tableNames)
        {
            FillDataset(transaction, commandType, commandText, dataSet, tableNames, null);
        }

        public static void FillDataset(SqlTransaction transaction, CommandType commandType, string commandText, DataSet dataSet, string[] tableNames, params SqlParameter[] commandParameters)
        {
            FillDataset(transaction.Connection, transaction, commandType, commandText, dataSet, tableNames, commandParameters);
        }

        private static void FillDataset(SqlConnection connection, SqlTransaction transaction, CommandType commandType, string commandText, DataSet dataSet, string[] tableNames, params SqlParameter[] commandParameters)
        {
            if (connection == null)
            {
                throw new ArgumentNullException("connection");
            }
            if (dataSet == null)
            {
                throw new ArgumentNullException("dataSet");
            }
            SqlCommand command = new SqlCommand();
            bool mustCloseConnection = false;
            PrepareCommand(command, connection, transaction, commandType, commandText, commandParameters, ref mustCloseConnection);
            using (SqlDataAdapter dataAdapter = new SqlDataAdapter(command))
            {
                if (tableNames != null && tableNames.Length > 0)
                {
                    string tableName = "Table";
                    for (int index = 0; index <= tableNames.Length - 1; index++)
                    {
                        if (tableNames[index] == null || tableNames[index].Length == 0)
                        {
                            throw new ArgumentException("The tableNames parameter must contain a list of tables, a value was provided as null or empty string.", "tableNames");
                        }
                        dataAdapter.TableMappings.Add(tableName, tableNames[index]);
                        tableName += (index + 1).ToString();
                    }
                }
                dataAdapter.Fill(dataSet);
                command.Parameters.Clear();
            }
            if (mustCloseConnection)
            {
                connection.Close();
            }
        }

        #endregion

        #region "UpdateDataset"

        public static void UpdateDataset(SqlCommand insertCommand, SqlCommand deleteCommand, SqlCommand updateCommand, DataSet dataSet, string tableName)
        {
            if (insertCommand == null)
            {
                throw new ArgumentNullException("insertCommand");
            }
            if (deleteCommand == null)
            {
                throw new ArgumentNullException("deleteCommand");
            }
            if (updateCommand == null)
            {
                throw new ArgumentNullException("updateCommand");
            }
            if (tableName == null || tableName.Length == 0)
            {
                throw new ArgumentNullException("tableName");
            }
            using (SqlDataAdapter dataAdapter = new SqlDataAdapter())
            {
                dataAdapter.UpdateCommand = updateCommand;
                dataAdapter.InsertCommand = insertCommand;
                dataAdapter.DeleteCommand = deleteCommand;
                dataAdapter.Update(dataSet, tableName);
                dataSet.AcceptChanges();
            }
        }

        #endregion

        #region "Check Null"

        public static object CheckDateNull(object readerValue)
        {
            if ((object.ReferenceEquals(readerValue, DBNull.Value)) || (readerValue == null))
            {
                return null;
            }
            else
            {
                return readerValue;
            }
        }

        public static int CheckIntNull(object readerValue)
        {
            if ((object.ReferenceEquals(readerValue, DBNull.Value)) || (readerValue == null))
            {
                return -1;
            }
            else
            {
                return Convert.ToInt32(readerValue);
            }
        }

        public static string CheckStringNull(object readerValue)
        {
            if ((object.ReferenceEquals(readerValue, DBNull.Value)) || (string.IsNullOrEmpty(readerValue.ToString())))
            {
                return "";
            }
            else
            {
                return Convert.ToString(readerValue);
            }
        }

        public static float CheckFloatNull(object readerValue)
        {
            if ((object.ReferenceEquals(readerValue, DBNull.Value)) || (readerValue == null))
            {
                return -1f;
            }
            else
            {
                return Convert.ToSingle(readerValue);
            }
        }

        public static bool CheckBooleanNull(object readerValue)
        {
            if ((object.ReferenceEquals(readerValue, DBNull.Value)) || (readerValue == null))
            {
                return false;
            }
            else
            {
                return Convert.ToBoolean(readerValue);
            }
        }

        #endregion

    }

    internal sealed class SqlValueConverter
    {

        private SqlValueConverter()
        {
        }

        #region "getters based on ordinals"

        public static String GetString(IDataReader reader, int ordinal)
        {
            if (reader.IsDBNull(ordinal))
            {
                return String.Empty;
            }
            return reader.GetString(ordinal);
        }

        public static int GetInt32(IDataReader reader, int ordinal)
        {
            if (reader.IsDBNull(ordinal))
            {
                return 0;
            }
            return Convert.ToInt32(reader.GetValue(ordinal));
        }

        public static long GetInt64(IDataReader reader, int ordinal)
        {
            if (reader.IsDBNull(ordinal))
            {
                return 0;
            }
            return Convert.ToInt64(reader.GetValue(ordinal));
        }

        public static DateTime GetDateTime(IDataReader reader, int ordinal)
        {
            if (reader.IsDBNull(ordinal))
            {
                return DateTime.MinValue;
            }
            return Convert.ToDateTime(reader.GetValue(ordinal));
        }

        #endregion

    }
}

//=======================================================
//Service provided by Telerik (www.telerik.com)
//Conversion powered by NRefactory.
//Twitter: @telerik
//Facebook: facebook.com/telerik
//=======================================================
