using Azure.Core;
using Microsoft.Data.SqlClient;
using Microsoft.Data.SqlClient.Server;
using System.Data;
using System.Runtime.InteropServices;

namespace Microsoft.ServiceHealthHub.Core
{
    public class MSSqlStructuredParameter
    {
        public string TypeName { get; set; }
        public IEnumerable<SqlDataRecord> Value { get; set; }
    }

    /// <summary>
    /// Connects to the Microsoft SQL Server database and executes T-SQL queries
    /// </summary>
    [Guid("3DE3EF45-7F09-4D61-A990-DD8A35C75B3D")]
    public class MSSQLServerConnector : IMSDatabaseConnector
    {
        #region Internal fields
        /// <summary>The _connectionString field represents the string used to open a SQL Server database.</summary>
        private string _connectionString = "";

        /// <summary>The _commandTimeout field represents the wait time in seconds before terminating the attempt to execute a command and generating an error.</summary>
        private int _commandTimeout = 120;
        #endregion

        #region Public properties
        /// <summary>The ConnectionString property represents the string used to open a SQL Server database.</summary>
        public string ConnectionString
        {
            get
            {
                return _connectionString;
            }
            set
            {
                if (String.IsNullOrEmpty(value))
                {
                    throw new System.ArgumentException("Connection string cannot be null or empty.");
                }

                _connectionString = value;
            }
        }

        /// <summary>The CommandTimeout property represents the wait time in seconds before terminating the attempt to execute a command and generating an error.</summary>
        public int CommandTimeout
        {
            get
            {
                return _commandTimeout;
            }
            set
            {
                _commandTimeout = value;
            }
        }
        #endregion

        #region Constructors
        /// <summary>
        /// Creates an instance of the mssqLServerConnector class
        /// </summary>
        /// <param name="connectionString">Microsoft SQL Server database connection string.</param>
        /// <returns></returns>
        /// <owner>adraskovic@outlook.com</owner>
        public MSSQLServerConnector(string connectionString)
        {
            if (String.IsNullOrEmpty(connectionString))
            {
                throw new ArgumentException("Connection string cannot be null or empty.");
            }

            _connectionString = connectionString;
        }

        /// <summary>
        /// Creates an instance of the mssqLServerConnector class
        /// </summary>
        /// <param name="server">Microsoft SQL Server name</param>
        /// <param name="database">Microsoft SQL Server database name</param>
        /// <param name="username">SQL Server user name. If not specified, the connection will be made in the IntegratedSecurity mode.</param>
        /// <param name="password">SQL Server user password. If not specified, the connection will be made in the IntegratedSecurity mode.</param>
        /// <param name="port">Microsoft SQL Server instance port. Default is 1433.</param>
        /// <returns></returns>
        /// <owner>adraskovic@outlook.com</owner>
        public MSSQLServerConnector(string server, string database, string username = "", string password = "", int port = 1433)
        {
            if (port == 0) { port = 1433; }

            SqlConnectionStringBuilder sqlConnStringBuilder = new SqlConnectionStringBuilder
            {
                DataSource = server + "," + port.ToString(),
                InitialCatalog = database
            };

            if (String.IsNullOrEmpty(username))
            {
                sqlConnStringBuilder.IntegratedSecurity = true;
            }
            else
            {
                sqlConnStringBuilder.UserID = username;
                sqlConnStringBuilder.Password = password;
            }
            _connectionString = sqlConnStringBuilder.ConnectionString;
        }
        #endregion

        #region Internal and private helper methods
        /// <summary>
        /// Returns the connection string with masked password. For internal use.
        /// </summary>
        /// <returns>SQL Server connection string with masked password</returns>
        /// <owner>adraskovic@outlook.com</owner>
        private string GetConnectionStringWithoutPassword()
        {
            SqlConnectionStringBuilder noPassConStrSB = new SqlConnectionStringBuilder(ConnectionString);
            if (!(string.IsNullOrWhiteSpace(noPassConStrSB.UserID) || noPassConStrSB.IntegratedSecurity))
                noPassConStrSB.Password = "******";

            return noPassConStrSB.ConnectionString;
        }

        /// <summary>
        /// Converts a SQL Server database dataset to the list of dictionaries. For internal use.
        /// </summary>
        /// <param name="sqlDataSet">Input dataset for the conversion</param>
        /// <returns>List of dictionaries containing the row data</returns>
        /// <owner>adraskovic@outlook.com</owner>
        private List<MSPropertyBag> ConvertDatasetToDictionary(DataSet sqlDataSet)
        {
            List<MSPropertyBag> result = new List<MSPropertyBag>();

            if (sqlDataSet != null)
            {
                DataTable dt = sqlDataSet.Tables[0];
                if (dt.Rows.Count > 0)
                {
                    foreach (DataRow dr in dt.Rows)
                    {
                        MSPropertyBag row = new MSPropertyBag();
                        foreach (DataColumn dc in dt.Columns)
                        {
                            row.Add(dc.ColumnName, dr[dc.ColumnName]);
                        }

                        result.Add(row);
                    }
                }
            }

            return result;
        }
        #endregion

        #region Public methods
        /// <summary>
        /// Get a dataset from the SQL Server database based on the provided query
        /// </summary>
        /// <param name="command"></param>
        /// <param name="parameters"></param>
        /// <returns>List of dictionaries containing the row data</returns>
        /// <owner>adraskovic@outlook.com</owner>
        public List<MSPropertyBag> GetDataSet(string command, MSPropertyBag parameters = null)
        {
            if (String.IsNullOrEmpty(command))
            {
                throw new ArgumentException("Command cannot be null or empty.");
            }

            using (SqlConnection sqlConn = new SqlConnection())
            {
                sqlConn.ConnectionString = @_connectionString;

                SqlConnectionStringBuilder sqlConnString = new SqlConnectionStringBuilder(_connectionString);
                if (string.IsNullOrEmpty(sqlConnString.UserID))
                {
                    var credential = GlobalConfiguration.Instance.DefaultAzureCredential;
                    var token = credential.GetToken(new TokenRequestContext(new[] { "https://database.windows.net/.default" }));
                    sqlConn.AccessToken = token.Token;
                }

                sqlConn.Open();

                using (SqlCommand sqlCmd = new SqlCommand())
                {
                    sqlCmd.CommandText = @command;
                    sqlCmd.Connection = sqlConn;
                    sqlCmd.CommandTimeout = _commandTimeout;

                    if (parameters != null)
                        foreach (string key in parameters.Keys)
                        {
                            if (parameters[key] is MSSqlStructuredParameter stParam)
                            {
                                SqlParameter p = sqlCmd.Parameters.Add(new SqlParameter(key, stParam?.Value));
                                p.SqlDbType = SqlDbType.Structured;
                                p.TypeName = stParam?.TypeName;
                            }
                            else
                                sqlCmd.Parameters.Add(new SqlParameter(key, parameters[key]));
                        }
                            
                    using (SqlDataAdapter sqlAdapter = new SqlDataAdapter())
                    {
                        sqlAdapter.SelectCommand = sqlCmd;
                        DataSet sqlDataSet = new DataSet();
                        try
                        {
                            sqlAdapter.Fill(sqlDataSet);
                        }
                        catch (Exception ex)
                        {
                            string paramList = string.Empty;

                            if (parameters != null)
                                foreach (string key in parameters.Keys)
                                {
                                    object parameterValue = parameters[key];
                                    paramList = string.Format("[{0} @{1} = '{2}'], ",
                                        parameterValue == null ? "<null>" : parameterValue.GetType().FullName,
                                        key,
                                        parameterValue ?? "<null>");
                                }

                            paramList = paramList.Trim().TrimEnd(new char[] { ',' });

                            /* MSTPLogger.LogEvent(
                                MSTPLoggingLevel.Critical,
                                "Change Management Core",
                                "Database",
                                "mssqpQD7",
                                string.Format(
                                    "Couldn't retrieve the data from the database. Connection string: {0}.\r\nSQL command: {1}.\r\nParameters: {2}.\r\nException: {3}.\r\nStack trace: {4}",
                                    GetConnectionStringWithoutPassword(),
                                    @command,
                                    paramList,
                                    ex,
                                    ex.StackTrace
                                    )); */
                            throw;
                        }

                        return ConvertDatasetToDictionary(sqlDataSet);
                    }
                }
            }
        }

        /// <summary>
        /// Executes a sequence of the SQL Server database commands
        /// </summary>
        /// <param name="command"></param>
        /// <returns></returns>
        /// <owner>adraskovic@outlook.com</owner>
        public void ExecuteQuerySequence(string command)
        {
            if (String.IsNullOrEmpty(command))
            {
                throw new System.ArgumentException("Command cannot be null or empty.");
            }

            string[] sqlStatements = @command.Split(new string[] { "GO\r\n", "GO ", "GO\t" }, System.StringSplitOptions.RemoveEmptyEntries);

            using (SqlConnection sqlConn = new SqlConnection())
            {
                sqlConn.ConnectionString = @_connectionString;

                SqlConnectionStringBuilder sqlConnString = new SqlConnectionStringBuilder(_connectionString);
                if (string.IsNullOrEmpty(sqlConnString.UserID))
                {
                    var credential = GlobalConfiguration.Instance.DefaultAzureCredential;
                    var token = credential.GetToken(new TokenRequestContext(new[] { "https://database.windows.net/.default" }));
                    sqlConn.AccessToken = token.Token;
                }
                
                sqlConn.Open();

                foreach (string sqlStatement in sqlStatements)
                {
                    if (sqlStatement.ToUpper() != "GO")
                    {
                        using (SqlCommand sqlCmd = new SqlCommand())
                        {
                            sqlCmd.Connection = sqlConn;
                            sqlCmd.CommandText = sqlStatement.Trim().TrimEnd(new char[] { 'G', 'O' });

                            if (!string.IsNullOrWhiteSpace(sqlCmd.CommandText))
                            {
                                sqlCmd.CommandTimeout = _commandTimeout;
                                try
                                {
                                    sqlCmd.ExecuteNonQuery();
                                }
                                catch (Exception ex)
                                {
                                    /* MSTPLogger.LogEvent(
                                        MSTPLoggingLevel.Critical,
                                        "Change Management Core",
                                        "Database",
                                        "mssqAo98",
                                        string.Format(
                                            "Couldn't execute the SQL query. Connection string: {0}.\r\nSQL command: {1}.\r\nException: {2}.\r\nStack trace: {3}",
                                            GetConnectionStringWithoutPassword(),
                                            @command,
                                            ex,
                                            ex.StackTrace
                                            )); */
                                    throw;
                                }
                            }
                        }
                    }
                }

                sqlConn.Close();
            }
        }

        /// <summary>
        /// Executes a single SQL Server database statement and passes provided parameters
        /// </summary>
        /// <param name="command">SQL Server command</param>
        /// <param name="parameters">A dictionery containing SQL Server command parameters</param>
        /// <owner>adraskovic@outlook.com</owner>
        public int ExecuteQuery(string command, MSPropertyBag parameters = null)
        {
            if (String.IsNullOrEmpty(command))
            {
                throw new ArgumentException("Command cannot be null or empty.");
            }

            int result = -1;

            using (SqlConnection sqlConn = new SqlConnection())
            {
                sqlConn.ConnectionString = @_connectionString;
                
                SqlConnectionStringBuilder sqlConnString = new SqlConnectionStringBuilder(_connectionString);
                if (string.IsNullOrEmpty(sqlConnString.UserID))
                {
                    var credential = GlobalConfiguration.Instance.DefaultAzureCredential;
                    var token = credential.GetToken(new TokenRequestContext(new[] { "https://database.windows.net/.default" }));
                    sqlConn.AccessToken = token.Token;
                }

                sqlConn.Open();

                using (SqlCommand sqlCmd = new SqlCommand(command))
                {
                    sqlCmd.Connection = sqlConn;

                    if (parameters != null)
                        foreach (string key in parameters.Keys)
                            sqlCmd.Parameters.Add(new SqlParameter(key, parameters[key]));

                    sqlCmd.CommandTimeout = _commandTimeout;
                    try
                    {
                        result = sqlCmd.ExecuteNonQuery();
                    }
                    catch (Exception ex)
                    {
                        string paramList = string.Empty;

                        if (parameters != null)
                            foreach (string key in parameters.Keys)
                            {
                                object parameterValue = parameters[key];
                                paramList = string.Format("[{0} @{1} = '{2}'], ",
                                    parameterValue == null ? "<null>" : parameterValue.GetType().FullName,
                                    key,
                                    parameterValue ?? "<null>");
                            }

                        paramList = paramList.Trim().TrimEnd(new char[] { ',' });

                        /* MSTPLogger.LogEvent(
                            MSTPLoggingLevel.Critical,
                            "Change Management Core",
                            "Database",
                            "mssq19Ka",
                            string.Format(
                                "Couldn't execute SQL query. Connection string: {0}.\r\nSQL command: {1}.\r\nParameters: {2}.\r\nException: {3}.\r\nStack trace: {4}",
                                GetConnectionStringWithoutPassword(),
                                @command,
                                paramList,
                                ex,
                                ex.StackTrace
                                )); */
                        throw;
                    }
                }

                sqlConn.Close();
            }

            return result;
        }

        /// <summary>
        /// Executes a single SQL Server database stored procedure and passes provided parameters
        /// </summary>
        /// <param name="command">SQL Server command</param>
        /// <param name="parameters">A dictionery containing SQL Server command parameters</param>
        /// <returns>Stored proceduree return code</returns>
        /// <owner>adraskovic@outlook.com</owner>
        public int ExecuteStoredProcedure(string storedProcedureName, MSPropertyBag parameters = null)
        {
            if (string.IsNullOrEmpty(storedProcedureName))
            {
                throw new System.ArgumentException("Command cannot be null or empty.");
            }

            int result = -1;

            using (SqlConnection sqlConn = new SqlConnection())
            {
                sqlConn.ConnectionString = @_connectionString;

                SqlConnectionStringBuilder sqlConnString = new SqlConnectionStringBuilder(_connectionString);
                if (string.IsNullOrEmpty(sqlConnString.UserID))
                {
                    var credential = GlobalConfiguration.Instance.DefaultAzureCredential;
                    var token = credential.GetToken(new TokenRequestContext(new[] { "https://database.windows.net/.default" }));
                    sqlConn.AccessToken = token.Token;
                }

                sqlConn.Open();

                using (SqlCommand sqlCmd = new SqlCommand(storedProcedureName))
                {
                    sqlCmd.CommandType = CommandType.StoredProcedure;
                    sqlCmd.Connection = sqlConn;

                    if (parameters != null)
                        foreach (string key in parameters.Keys)
                            sqlCmd.Parameters.Add(new SqlParameter(key, parameters[key]));

                    sqlCmd.Parameters.Add("@RETURN_VALUE", SqlDbType.Int).Direction = ParameterDirection.ReturnValue;
                    sqlCmd.CommandTimeout = _commandTimeout;
                    try
                    {
                        sqlCmd.ExecuteNonQuery();
                        result = sqlCmd.Parameters["@RETURN_VALUE"] == null ? 0 : (int)sqlCmd.Parameters["@RETURN_VALUE"].Value;
                    }
                    catch (Exception ex)
                    {
                        string paramList = string.Empty;

                        if (parameters != null)
                            foreach (string key in parameters.Keys)
                            {
                                object parameterValue = parameters[key];
                                paramList = string.Format("[{0} @{1} = '{2}'], ",
                                    parameterValue == null ? "<null>" : parameterValue.GetType().FullName,
                                    key,
                                    parameterValue ?? "<null>");
                            }

                        paramList = paramList.Trim().TrimEnd(new char[] { ',' });

                        /* MSTPLogger.LogEvent(
                            MSTPLoggingLevel.Critical,
                            "Change Management Core",
                            "Database",
                            "mssq0KlM",
                            string.Format(
                                "Couldn't execute SQL stored procedure. Connection string: {0}.\r\nStored procedure: {1}.\r\nParameters: {2}.\r\nException: {3}.\r\nStack trace: {4}",
                                GetConnectionStringWithoutPassword(),
                                storedProcedureName,
                                paramList,
                                ex,
                                ex.StackTrace
                                )); */
                        throw;
                    }
                }

                sqlConn.Close();
            }

            return result;
        }
        #endregion
    }
}
