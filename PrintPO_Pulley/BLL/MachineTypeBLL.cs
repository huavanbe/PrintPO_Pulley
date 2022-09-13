using BFDataCrawler.DAL;
using BFDataCrawler.Model;
using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace BFDataCrawler.BLL
{
    class MachineTypeBLL
    {
        private static MachineTypeBLL instance;
        public static MachineTypeBLL Instance
        {
            get
            {
                if (instance == null)
                    instance = new MachineTypeBLL();
                return instance;
            }
            set
            {
                instance = value;
            }
        }

        public Task<List<MachineTypeModel>> GetMachineTypes()
        {
            try
            {            
                return Task.Run(()=>  MachineTypeDAL.Instance.GetMachineTypes());
            }
            catch (Exception ex)
            {

                throw ex;
            }
        }

        public int EditFlagCheckBWOrder_CNCHob9(int value)
        {            
            try
            {                
                return MachineTypeDAL.Instance.EditFlagCheckBWOrder_CNCHob9(value);
            }
            catch (Exception ex)
            {

                throw ex;
            }
        }
        public int EditFlagCheckBWOrder_CNCHob10(int value)
        {            
            try
            {                
                return MachineTypeDAL.Instance.EditFlagCheckBWOrder_CNCHob10(value);
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
                return MachineTypeDAL.Instance.GetFlagCheckBWOrder_CNCHob9();
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
                return MachineTypeDAL.Instance.GetFlagCheckBWOrder_CNCHob10();
            }
            catch (Exception ex)
            {

                throw ex;
            }
        }
    }
}
