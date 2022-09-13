
using System;
using System.Collections.Generic;
using System.Data;
using System.Data.SqlClient;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace BDDataCrawler.DataProvider
{
    class SQLProvider
    {
        private static string connStr = System.Configuration.ConfigurationManager.AppSettings["DailyReportF4Connection"];
        private static SQLProvider instance;
        public static SQLProvider Instance
        {
            get
            {
                if (instance == null) instance = new SQLProvider(); return instance;
            }
            set
            {
                instance = value;
            }
        }

        public void SetConnection(string connectionStr)
        {
            connStr = connectionStr;
        }

        public DataSet ExecuteQueryDS(string spName, params object[] pr)
        {
            using (SqlConnection conn = new SqlConnection(connStr))
            {
                if (conn.State == ConnectionState.Closed) conn.Open();
                SqlDataAdapter da = new SqlDataAdapter(spName, conn);
                da.SelectCommand.CommandType = CommandType.StoredProcedure;
                SqlCommandBuilder.DeriveParameters(da.SelectCommand);
                if (da.SelectCommand.Parameters.Count - 1 != pr.Length)
                {
                    throw new Exception("So luong tham so khong dung");
                }
                int i = 0;
                foreach (SqlParameter p in da.SelectCommand.Parameters)
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
        public int ExecuteNonQuery(string spName, params object[] pr)
        {
            int result = 0;
            using (SqlConnection conn = new SqlConnection(connStr))
            {
                SqlCommand cmd = new SqlCommand(spName, conn);
                cmd.CommandType = CommandType.StoredProcedure;
                SqlCommandBuilder.DeriveParameters(cmd);
                if (cmd.Parameters.Count - 1 != pr.Length)
                {
                    throw new Exception("So luong tham so khong dung");
                }
                int i = 0;
                foreach (SqlParameter p in cmd.Parameters)
                {
                    if (p.Direction == ParameterDirection.Input || p.Direction == ParameterDirection.InputOutput)
                    {
                        p.Value = pr[i++];
                    }
                }

                result = cmd.ExecuteNonQuery();
                return result;
            }
        }

        public DataTable ExecuteQuery(string query, object[] parameter = null)
        {
            DataTable data = new DataTable();

            using (SqlConnection connection = new SqlConnection(connStr))
            {
                connection.Open();

                SqlCommand command = new SqlCommand(query, connection);

                if (parameter != null)
                {
                    string[] listPara = query.Split(' ');
                    int i = 0;
                    foreach (string item in listPara)
                    {
                        if (item.Contains('@'))
                        {
                            command.Parameters.AddWithValue(item, parameter[i]);
                            i++;
                        }
                    }
                }

                SqlDataAdapter adapter = new SqlDataAdapter(command);

                adapter.Fill(data);

                connection.Close();
            }

            return data;
        }

        public int ExecuteNonQuerySP(string query, object[] parameter = null)
        {
            int data = 0;

            using (SqlConnection connection = new SqlConnection(connStr))
            {
                connection.Open();

                SqlCommand command = new SqlCommand(query, connection);

                if (parameter != null)
                {
                    string[] listPara = query.Split(' ');
                    int i = 0;
                    foreach (string item in listPara)
                    {
                        if (item.Contains('@'))
                        {
                            command.Parameters.AddWithValue(item, parameter[i]);
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

            using (SqlConnection connection = new SqlConnection(connStr))
            {
                connection.Open();

                SqlCommand command = new SqlCommand(query, connection);

                if (parameter != null)
                {
                    string[] listPara = query.Split(' ');
                    int i = 0;
                    foreach (string item in listPara)
                    {
                        if (item.Contains('@'))
                        {
                            command.Parameters.AddWithValue(item, parameter[i]);
                            i++;
                        }
                    }
                }

                data = command.ExecuteScalar();

                connection.Close();
            }

            return data;
        }
    }
}
