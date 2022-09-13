using Oracle.ManagedDataAccess.Client;
using System;
using System.Collections.Generic;
using System.Data;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace BFDataCrawler.DataProviders
{
    class OracleProvider
    {
        //private static string host = "10.4.17.62";
        //private static int port = 3306;
        //private static string database = "Misumi_CheckJig";
        //private static string username = "root";
        //private static string password = "admin";
        //private static string connStr = string.Format("Server={0};Database={1};Port={2};Uid={3};Pwd={4};SslMode=none;CharSet=utf8;respect binary flags=false;", host, database, port, username, password);
        private static string connStr = System.Configuration.ConfigurationManager.AppSettings["BWDataCrawler_Connection"];
        private static OracleProvider instance;
        public static OracleProvider Instance
        {
            get
            {
                if (instance == null) instance = new OracleProvider(); return instance;
            }
            set
            {
                instance = value;
            }
        }

        public DataTable ExecuteQuery(string query, object[] parameter = null)
        {
            DataTable data = new DataTable();

            using (OracleConnection connection = new OracleConnection(connStr))
            {
                connection.Open();

                OracleCommand command = new OracleCommand(query, connection);

                if (parameter != null)
                {
                    string[] listPara = query.Split(' ');
                    int i = 0;
                    foreach (string item in listPara)
                    {
                        if (item.Contains('@'))
                        {
                            command.Parameters.Add(item, parameter[i]);
                            i++;
                        }
                    }
                }

                OracleDataAdapter adapter = new OracleDataAdapter(command);

                adapter.Fill(data);

                connection.Close();
            }

            return data;
        }

        public int ExecuteNonQuery(string query, object[] parameter = null)
        {
            int data = 0;

            using (OracleConnection connection = new OracleConnection(connStr))
            {
                connection.Open();

                OracleCommand command = new OracleCommand(query, connection);

                if (parameter != null)
                {
                    string[] listPara = query.Split(' ');
                    int i = 0;
                    foreach (string item in listPara)
                    {
                        if (item.Contains('@'))
                        {
                            command.Parameters.Add(item, parameter[i]);
                            i++;
                        }
                    }
                }

                data = command.ExecuteNonQuery();

                connection.Close();
            }

            return data;
        }


        public object ExecuteScalar(string query, object[] parameter = null)
        {
            object data = 0;

            using (OracleConnection connection = new OracleConnection(connStr))
            {
                connection.Open();

                OracleCommand command = new OracleCommand(query, connection);

                if (parameter != null)
                {
                    string[] listPara = query.Split(' ');
                    int i = 0;
                    foreach (string item in listPara)
                    {
                        if (item.Contains('@'))
                        {
                            command.Parameters.Add(item, parameter[i]);
                            i++;
                        }
                    }
                }

                data = command.ExecuteScalar();

                connection.Close();
            }

            return data;
        }



        public DataSet ExecuteQueryDS(string spName, params object[] pr)
        {
            using (OracleConnection conn = new OracleConnection(connStr))
            {
                if (conn.State == ConnectionState.Closed) conn.Open();
                OracleDataAdapter da = new OracleDataAdapter(spName, conn);
                da.SelectCommand.CommandType = CommandType.StoredProcedure;
                OracleCommandBuilder.DeriveParameters(da.SelectCommand);
                if (da.SelectCommand.Parameters.Count != pr.Length)
                {
                    throw new Exception("So luong tham so khong dung");
                }
                int i = 0;
                foreach (OracleParameter p in da.SelectCommand.Parameters)
                {
                    if (p.Direction == ParameterDirection.Input || p.Direction == ParameterDirection.InputOutput)
                    {
                        p.Value = pr[i++];
                    }
                }
                DataSet ds = new DataSet();
                da.Fill(ds);
                da.Dispose();
                return ds;
            }
        }

        public DataTable ExecuteQueryDT(string spName, params object[] pr)
        {
            try
            {
                DataSet ds = ExecuteQueryDS(spName, pr);
                if (ds.Tables.Count > 0)
                {
                    return ds.Tables[0];
                }
                else return null;
            }
            catch (Exception ex)
            {
                throw ex;
            }
        }

        public DataRow ExecuteQueryDR(string spName, params object[] pr)
        {
            try
            {
                DataTable dt = ExecuteQueryDT(spName, pr);
                if (dt != null && dt.Rows.Count > 0)
                    return dt.Rows[0];
                else
                    return null;
            }
            catch (Exception ex)
            {

                throw ex;
            }
        }

        public int ExecuteNonQuerySP(string spName, params object[] pr)
        {
            int result = 0;
            using (OracleConnection conn = new OracleConnection(connStr))
            {
                if (conn.State == ConnectionState.Closed)
                    conn.Open();

                OracleCommand cmd = new OracleCommand(spName, conn);
                cmd.CommandType = CommandType.StoredProcedure;
                OracleCommandBuilder.DeriveParameters(cmd);
                if (cmd.Parameters.Count != pr.Length)
                {
                    throw new Exception("So luong tham so khong dung");
                }
                int i = 0;
                foreach (OracleParameter p in cmd.Parameters)
                {
                    if (p.Direction == ParameterDirection.Input || p.Direction == ParameterDirection.InputOutput)
                    {
                        p.Value = pr[i];
                        i++;
                    }
                }

                cmd.ExecuteNonQuery();
                return result = Convert.ToInt32(cmd.Parameters[pr.Length - 1].Value);
            }
        }
    }
}
