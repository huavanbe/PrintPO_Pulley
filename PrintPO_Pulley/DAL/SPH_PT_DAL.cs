using BDDataCrawler.DataProvider;
using BDDataCrawler.Provider;
using System;
using System.Collections.Generic;
using System.Data;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace BFDataCrawler.DAL
{
    class SPH_PT_DAL
    {
        private static SPH_PT_DAL _instance;

        public static SPH_PT_DAL Instance {
            get {
                if (_instance == null) {
                    _instance = new SPH_PT_DAL();
                }

                return _instance;
            }
            set => _instance = value;
        }
        public DataTable GetSPH_PT()
        {
            try
            {
                return SQLProvider.Instance.ExecuteQuery(string.Format("SELECT * FROM F4_BASEINPUT WHERE [Date] = CONVERT(varchar(10),GETDATE(),112) AND ([Title] = 'SPH' OR [Title] = 'PT')"));
            }
            catch (Exception ex)
            {

                throw ex;
            }
        }

        public DataTable GetSumHourByShift(string shift,int hour)
        {
            try
            {
                string query = string.Format(@"SELECT SUM(Hour) totalHour,Shift
                                                FROM WorkingTime
                                                WHERE RowNum <
                                                (SELECT RowNum FROM `WorkingTime` WHERE MinHour = {0} AND Shift = '{1}') AND Shift = '{1}'
                                                GROUP BY Shift", hour,shift);
               
                return MySQLProvider.Instance.ExecuteQuery(query);
            }
            catch (Exception ex)
            {
                throw ex;
            }
        }
    }
}
