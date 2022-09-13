using BFDataCrawler.DAL;
using BFDataCrawler.Model;
using System;
using System.Collections.Generic;
using System.Data;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace BFDataCrawler.BLL
{
    class OrderBLL
    {
        private static OrderBLL _instance;
        public static OrderBLL Instance
        {
            get {
                if (_instance == null)
                    _instance = new OrderBLL();
                return _instance;
            }
            set
            {
                _instance = value;
            }
        }

        public int AddNewOrders(List<OrderModel> orders)
        {            
            try
            {
                return OrderDAL.Instance.AddNewOrders(orders);
            }
            catch (Exception ex)
            {

                throw ex;
            }
        }

        public List<OrderModel> GetPulleyMasterByProduct(string product)
        {
            try
                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                            {
                return OrderDAL.Instance.GetPulleyMasterByProduct(product);
            }
            catch (Exception ex)
            {

                throw ex;
            }
        }

        public int EditBWCNCHob9(string soPO, string teethShape, int teethQty, double diameter_d, string dmat)
        {
            try
            {
                return OrderDAL.Instance.EditBWCNCHob9(soPO, teethShape, teethQty, diameter_d, dmat);
            }
            catch (Exception ex)
            {

                throw ex;
            }
        }

        public int EditBWCNCHob10(string soPO, string teethShape, int teethQty, double diameter_d, string dmat)
        {
            try
            {
                return OrderDAL.Instance.EditBWCNCHob10(soPO, teethShape, teethQty, diameter_d, dmat);
            }
            catch (Exception ex)
            {

                throw ex;
            }
        }

        public string GetKnifeType(string teethShape, int teethQty)
        {
            try
            {                
                return OrderDAL.Instance.GetKnifeType(teethShape, teethQty);
            }
            catch (Exception ex)
            {

                throw ex;
            }
        }

        public int CheckUnlockSettings(string pass)
        {
            try
            {                
                return OrderDAL.Instance.CheckUnlockSettings(pass);
            }
            catch (Exception ex)
            {
                throw ex;
            }
        }

        public int GetTotalBWOrders()
        {
            try
            {
                
                return OrderDAL.Instance.GetTotalBWOrders();
            }
            catch (Exception ex)
            {

                throw ex;
            }
        }

        public int GetTotalCNCOrders()
        {
            try
            {                
                return OrderDAL.Instance.GetTotalCNCOrders();
            }
            catch (Exception ex)
            {

                throw ex;
            }
        }

        public int UpdateTotalBWOrders(int total)
        {
            try
            {               
                return OrderDAL.Instance.UpdateTotalBWOrders(total);
            }
            catch (Exception ex)
            {
                throw ex;
            }
        }

        public int UpdateTotalCNCOrders(int total)
        {
            try
            {
                return OrderDAL.Instance.UpdateTotalCNCOrders(total);
            }
            catch (Exception ex)
            {
                throw ex;
            }
        }

        public int Update_bwIsSetup(int status)
        {
            try
            {
                
                return OrderDAL.Instance.Update_bwIsSetup(status);
            }
            catch (Exception ex)
            {

                throw ex;
            }
        }

        public int Update_cncIsSetup(int status)
        {
            try
            {
                return OrderDAL.Instance.Update_cncIsSetup(status);
            }
            catch (Exception ex)
            {

                throw ex;
            }
        }
        public int Get_bwIsSetupN1()
        {
            try
            {
                return OrderDAL.Instance.Get_bwIsSetupN1();
            }
            catch (Exception ex)
            {

                throw ex;
            }
        }

        public int Get_cncIsSetupN1()
        {
            try
            {
                return OrderDAL.Instance.Get_cncIsSetupN1();
            }
            catch (Exception ex)
            {

                throw ex;
            }
        }

        //2020-03-06 Start add: thêm chức năng cập nhật thông số Teeth cho BW11 và BW12
        public int Update_BW11_TeethInfo(OrderModel order)
        {
            
            try
            {
                return OrderDAL.Instance.Update_BW11_TeethInfo(order);
            }
            catch (Exception ex)
            {

                throw ex;
            }
        }

        public int Update_BW12_TeethInfo(OrderModel order)
        {
            
            try
            {
                return OrderDAL.Instance.Update_BW12_TeethInfo(order);
            }
            catch (Exception ex)
            {

                throw ex;
            }
        }

        public DataRow GetBW11_TeethInfo()
        {            
            try
            {
                return OrderDAL.Instance.GetBW11_TeethInfo();
            }
            catch (Exception ex)
            {

                throw ex;
            }
        }

        public DataRow GetBW12_TeethInfo()
        {
           
            try
            {
                return OrderDAL.Instance.GetBW12_TeethInfo();
            }
            catch (Exception ex)
            {

                throw ex;
            }
        }
        //2020-03-06 End add

        //2020-03-09 Start add: function lấy các ngày lễ của năm hiện tại
        public List<DateTime> GetHolidaysInYear()
        {            
            try
            {
               
                return OrderDAL.Instance.GetHolidaysInYear();
            }
            catch (Exception ex)
            {
                throw ex;
            }
        }

        //lấy sum giá trị cột PULLEY_LINEB
        public double GetSumPulley_BW(int hour)
        {
            try
            {               
                return OrderDAL.Instance.GetSumPulley_BW(hour);
            }
            catch (Exception ex)
            {

                throw ex;
            }
        }

        //lấy sum giá trị cột PULLEY_LINEB
        public double GetSumPulley_CNC(int hour)
        {
            try
            {
                return OrderDAL.Instance.GetSumPulley_CNC(hour);
            }
            catch (Exception ex)
            {

                throw ex;
            }
        }
        //2020-03-09

        //2020-03-23 Start add: edit IsLockCNC9, 10 để khóa/mở chức năng in khi có hàng F1
        public Tuple<int,int> GetLockStatusCNC9()
        {           
            try
            {               
                return OrderDAL.Instance.GetLockStatusCNC9();
            }
            catch (Exception ex)
            {

                throw ex;
            }
        }

        public Tuple<int,int> GetLockStatusCNC10()
        {
            try
            {
                return OrderDAL.Instance.GetLockStatusCNC10();
            }
            catch (Exception ex)
            {

                throw ex;
            }
        }

        public int SetLockCNC9(int status)
        {            
            try
            {                
                return OrderDAL.Instance.SetLockCNC9(status);
            }
            catch (Exception ex)
            {

                throw ex;
            }
        }

        public int SetLockCNC10(int status)
        {
            try
            {
                return OrderDAL.Instance.SetLockCNC10(status);
            }
            catch (Exception ex)
            {

                throw ex;
            }
        }
        //2020-03-23 End add
    }
}
