using BDDataCrawler.DataProvider;
using BDDataCrawler.Provider;
using BFDataCrawler.Model;
using System;
using System.Collections.Generic;
using System.Data;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace BFDataCrawler.DAL
{
    class MachineTypeDAL
    {
        private static MachineTypeDAL instance;
        public static MachineTypeDAL Instance
        {
            get
            {
                if (instance == null)
                    instance = new MachineTypeDAL();
                return instance;
            }
            set
            {
                instance = value;
            }
        }

        //public Task<List<MachineTypeModel>> GetMachineTypesAsync()
        //{
        //    string query = "SELECT * FROM MachineType";
        //    List<MachineTypeModel> lst = new List<MachineTypeModel>();
        //    try
        //    {
        //        return Task.Run(()=> {
        //            DataTable dt = MySQLProvider.Instance.ExecuteQuery(query);
        //            foreach (DataRow row in dt.Rows)
        //            {
        //                lst.Add(new MachineTypeModel()
        //                {
        //                    MaterialCode = row["MaterialCode"].ToString(),
        //                    Name = row["Name"].ToString(),
        //                    Size = Convert.ToInt32(row["Size"]),
        //                    LineAB = row["LineAB"].ToString(),
        //                });
        //            }
        //            return lst;
        //        }); 

        //    }
        //    catch (Exception ex)
        //    {

        //        throw ex;
        //    }
        //}

        public List<MachineTypeModel> GetMachineTypes()
        {
            string query = "SELECT * FROM MachineType";
            List<MachineTypeModel> lst = new List<MachineTypeModel>();
            try
            {
                DataTable dt = MySQLProvider.Instance.ExecuteQuery(query);
                foreach (DataRow row in dt.Rows)
                {
                    lst.Add(new MachineTypeModel()
                    {
                        MaterialCode = row["MaterialCode"].ToString(),
                        Name = row["Name"].ToString(),
                        Size = Convert.ToInt32(row["Size"]),
                        LineAB = row["LineAB"].ToString(),
                        LineName = row["LineName"].ToString()
                    });
                }
                return lst;

            }
            catch (Exception ex)
            {

                throw ex;
            }
        }

        public int EditFlagCheckBWOrder_CNCHob9(int value)
        {
            int result = 0;
            try
            {
                result = MySQLProvider.Instance.ExecuteNonQuery(string.Format("UPDATE CNCHob9 SET BWCNCHob9_HasBWOrders = {0}",value));
                return result;
            }
            catch (Exception ex)
            {

                throw ex;
            }
        }
        public int EditFlagCheckBWOrder_CNCHob10(int value)
        {
            int result = 0;
            try
            {
                result = MySQLProvider.Instance.ExecuteNonQuery(string.Format("UPDATE CNCHob10 SET BWCNCHob10_HasBWOrders = {0}",value));
                return result;
            }
            catch (Exception ex)
            {

                throw ex;
            }
        }

        public int GetFlagCheckBWOrder_CNCHob9()
        {
            try
            {                
                return Convert.ToInt32(MySQLProvider.Instance.ExecuteScalar(string.Format("SELECT BWCNCHob9_HasBWOrders FROM CNCHob9")));
            }
            catch (Exception ex)
            {

                throw ex;
            }
        }
        public int GetFlagCheckBWOrder_CNCHob10()
        {
            try
            {
                return Convert.ToInt32(MySQLProvider.Instance.ExecuteScalar(string.Format("SELECT BWCNCHob10_HasBWOrders FROM CNCHob10")));
            }
            catch (Exception ex)
            {

                throw ex;
            }
        }

        
    }
 
}
