using BFDataCrawler.DAL;
using System;
using System.Collections.Generic;
using System.Data;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace BFDataCrawler.BLL
{
    class SPH_PT_BLL
    {
        private static SPH_PT_BLL _instance;

        public static SPH_PT_BLL Instance
        {
            get
            {
                if (_instance == null)
                {
                    _instance = new SPH_PT_BLL();
                }

                return _instance;
            }
            set => _instance = value;
        }

        public DataTable GetSPH_PT()
        {
            try
            {
                return SPH_PT_DAL.Instance.GetSPH_PT();
            }
            catch (Exception ex)
            {

                throw ex;
            }
        }

        public DataTable GetSumHourByShift(string shift, int hour)
        {
            try
            {
                return SPH_PT_DAL.Instance.GetSumHourByShift(shift,hour);
            }
            catch (Exception ex)
            {
                throw ex;
            }
        }
    }
}
