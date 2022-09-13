using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace BFDataCrawler.VMSubFunctions
{
    class BWCNCSubFunctions
    {
        private static BWCNCSubFunctions instance;
        public static BWCNCSubFunctions Instance
        {
            get {
                if (instance == null)
                    instance = new BWCNCSubFunctions();
                return instance;
            }
            set { instance = value; }
        }


        #region <<BW-FUNCTIONS>>
        public void CanGetOrdersBW()
        {

        }
        public void GetOrdersBW()
        {

        }
        #endregion <<BW-FUNCTIONS>>

        #region <<CNC-FUNCTIONS>>

        #endregion <<CNC-FUNCTIONS>>


    }
}
