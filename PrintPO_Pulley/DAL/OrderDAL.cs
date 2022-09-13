using BDDataCrawler.DataProvider;
using BDDataCrawler.Provider;
using BFDataCrawler.DataProviders;
using BFDataCrawler.Model;
using System;
using System.Collections.Generic;
using System.Data;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace BFDataCrawler.DAL
{
    class OrderDAL
    {
        private static OrderDAL _instance;
        public static OrderDAL Instance
        {
            get
            {
                if (_instance == null)
                    _instance = new OrderDAL();
                return _instance;
            }
            set { _instance = value; }
        }

        public int AddNewOrders(List<OrderModel> orders)
        {
            StringBuilder builder = new StringBuilder();
            builder.AppendLine(string.Format(@"INSERT INTO `BFData` 
            (
             `no`, 
             `sales_order_no`,
             `pier_instruction_no`,
             `manufa_instruction_no`,
             `global_code`,
             `customers`,
             `item_name`,
             `mc`,
             `received_date`,
             `factory_ship_date`,
             `number_of_orders`,
             `number_of_available_instructions`,
             `number_of_repairs`,
             `number_of_instructions`,
             `line`,
             `payward`,
             `major`,
             `special_orders`,
             `method`,
             `destination`,
             `instructions_print_date`,
             `latest_progress`,
             `tack_label_output_date`,
             `completion_instruction_date`,
             `re_print_count`,
             `latest_issue_time`,             
             `material_code1`,
             `material_text1`,
             `amount_used1`,
             `unit1`,
             `material_code2`,
             `material_text2`,
             `amount_used2`,
             `unit2`,
             `material_code3`,
             `material_text3`,
             `amount_used3`,
             `unit3`,
             `material_code4`,
             `material_text4`,
             `amount_used4`,
             `unit4`,
             `material_code5`,
             `material_text5`,
             `amount_used5`,
             `unit5`,
             `material_code6`,
             `material_text6`,
             `amount_used6`,
             `unit6`,
             `material_code7`,
             `material_text7`,
             `amount_used7`,
             `unit7`,
             `material_code8`,
             `material_text8`,
             `amount_used8`,
             `unit8`,
             `inner_code`,
             `classify_code`,
             `createddate`) VALUES")
             );
            //`no`,	`sales_order_no`,	`pier_instruction_no`,	`manufa_instruction_no`,	`global_code`,	
            //`customers`,	`item_name`,	`mc`,	`received_date`,	`factory_ship_date`,	
            //`number_of_orders`,	`number_of_available_instructions`,	`number_of_repairs`,	`number_of_instructions`,	`line`,	
            //`payward`,	`major`,	`special_orders`,	`method`,	`destination`,	
            //`instructions_print_date`,	`latest_progress`,	`tack_label_output_date`,	`completion_instruction_date`,	`re_print_count`,	
            //`latest_issue_time`,	`material_code1`,	`material_text1`,	`amount_used1`,	`unit1`,	
            //`material_code2`,	`material_text2`,	`amount_used2`,	`unit2`,	`material_code3`,	
            //`material_text3`,	`amount_used3`,	`unit3`,	`material_code4`,	`material_text4`,	
            //`amount_used4`,	`unit4`,	`material_code5`,	`material_text5`,	`amount_used5`,	
            //`unit5`,	`material_code6`,	`material_text6`,	`amount_used6`,	`unit6`,	
            //`material_code7`,	`material_text7`,	`amount_used7`,	`unit7`,	`material_code8`,	
            //`material_text8`,	`amount_used8`,	`unit8`,	`inner_code`,	`classify_code`,	`createddate`
            foreach (var od in orders)
            {
                builder.AppendLine(string.Format(@"('{0}','{1}','{2}','{3}','{4}',
                                                    '{5}','{6}','{7}','{8}','{9}',
                                                    '{10}','{11}','{12}','{13}','{14}',
                                                    '{15}','{16}','{17}','{18}','{19}',
                                                    '{20}','{21}','{22}','{23}','{24}',
                                                    '{25}','{26}','{27}','{28}','{29}',
                                                    '{30}','{31}','{32}','{33}','{34}',
                                                    '{35}','{36}','{37}','{38}','{39}',
                                                    '{40}','{41}','{42}','{43}','{44}',
                                                    '{45}','{46}','{47}','{48}','{49}',
                                                    '{50}','{51}','{52}','{53}','{54}',
                                                    '{55}','{56}','{57}','{58}','{59}',NOW()),",
                                                    od.No, od.Sales_Order_No, od.Pier_Instruction_No, od.Manufa_Instruction_No, od.Global_Code,
                                                    od.Customers, od.Item_Name, od.MC, od.Received_Date, od.Factory_Ship_Date,
                                                    od.Number_of_Orders, od.Number_of_Available_Instructions, od.Number_of_Repairs, od.Number_of_Instructions, od.Line,
                                                    od.PayWard, od.Major, od.Special_Orders, od.Method, od.Destination,
                                                    od.Instructions_Print_date, od.Latest_progress, od.Tack_Label_Output_Date, od.Completion_Instruction_Date, od.Re_print_Count,
                                                    od.Latest_issue_time, od.Material_code1, od.Material_text1, od.Amount_used1, od.Unit1,
                                                    od.Material_code2, od.Material_text2, od.Amount_used2, od.Unit2, od.Material_code3,
                                                    od.Material_text3, od.Amount_used3, od.Unit3, od.Material_code4, od.Material_text4,
                                                    od.Amount_used4, od.Unit4, od.Material_code5, od.Material_text5, od.Amount_used5,
                                                    od.Unit5, od.Material_code6, od.Material_text6, od.Amount_used6, od.Unit6,
                                                    od.Material_code7, od.Material_text7, od.Amount_used7, od.Unit7, od.Material_code8,
                                                    od.Material_text8, od.Amount_used8, od.Unit8, od.Inner_Code, od.Classify_Code));
            }

            string query = builder.ToString();
            query = query.Remove(query.LastIndexOf(',')) + ";";
            int result = 0;
            try
            {
                result = MySQLProvider.Instance.ExecuteNonQuery(query);
            }
            catch (Exception ex)
            {

                throw ex;
            }

            return result;
        }

        public List<OrderModel> GetPulleyMasterByProduct(string product)
        {
            
                List<OrderModel> lst = new List<OrderModel>();
                string query = string.Format(@"SELECT * FROM PulleyMaster WHERE Product IN ({0})", product);
                try
                {
                    DataTable dt = MySQLProvider.Instance.ExecuteQuery(query);
                    foreach (DataRow row in dt.Rows)
                    {
                        lst.Add(new OrderModel()
                        {
                            Inner_Code = row["Product"].ToString(),
                            Hobbing = Convert.ToBoolean(row["Hobbing"]),
                            TeethQuantity = Convert.ToInt32(row["TeethQuantity"]),
                            TeethShape = row["TeethShape"].ToString(),
                            Diameter_d = Convert.ToDouble(row["d_HobShaft"])
                        });
                    }
                    return lst;
                }
                catch (Exception ex)
                {

                    throw ex;
                }
            
        }

        public int EditBWHob11(string soPO, string teethShape, int teethQty, double diameter_d)
        {
            try
            {
                string query = string.Format("UPDATE BWHob11 SET SoPO='{0}', TeethShape='{1}', TeethQty= {2}, Diameter_d= {3}",soPO,teethShape,teethQty, diameter_d);
                return MySQLProvider.Instance.ExecuteNonQuery(query);
            }
            catch (Exception ex)
            {

                throw ex;
            }
        }
        public int EditBWHob12(string soPO, string teethShape, int teethQty, double diameter_d)
        {
            try
            {
                string query = string.Format("UPDATE BWHob12 SET SoPO='{0}', TeethShape='{1}', TeethQty= {2}, Diameter_d= {3}", soPO, teethShape, teethQty, diameter_d);
                return MySQLProvider.Instance.ExecuteNonQuery(query);
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
                string query = string.Format("UPDATE CNCHob9 SET SoPO = '{0}', TeethShape = '{1}', TeethQty = {2}, Diameter_d = {3}, DMat = '{4}'", "", teethShape, teethQty, diameter_d,dmat);
                return MySQLProvider.Instance.ExecuteNonQuery(query);
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
                string query = string.Format("UPDATE CNCHob10 SET SoPO = '{0}', TeethShape = '{1}', TeethQty = {2}, Diameter_d = {3}, DMat = '{4}'", "", teethShape, teethQty, diameter_d, dmat);
                return MySQLProvider.Instance.ExecuteNonQuery(query);
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
                string query = string.Format("SELECT * FROM TeethMaster WHERE TeethShape = '{0}' AND (MinTeethQuantity <= {1} AND MaxTeethQuantity >= {1})",teethShape, teethQty);
                DataTable dt = MySQLProvider.Instance.ExecuteQuery(query);                
                return dt.Rows[0]["KnifeType"].ToString();                
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
                string query = string.Format("SELECT COUNT(*) FROM Account WHERE UserName = '{0}' AND Password = '{1}'","admin",pass);
                return Convert.ToInt32(MySQLProvider.Instance.ExecuteScalar(query));
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
                string query = string.Format("SELECT Total FROM OrdersCount WHERE Line ='BW'");
                return Convert.ToInt32(MySQLProvider.Instance.ExecuteScalar(query));
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
                string query = string.Format("SELECT Total FROM OrdersCount WHERE Line ='CNC'");
                return Convert.ToInt32(MySQLProvider.Instance.ExecuteScalar(query));
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
                string query = string.Format(string.Format("UPDATE OrdersCount SET Total = {0} WHERE Line ='BW'",total));
                return MySQLProvider.Instance.ExecuteNonQuery(query);
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
                string query = string.Format(string.Format("UPDATE OrdersCount SET Total = {0} WHERE Line ='CNC'", total));
                return MySQLProvider.Instance.ExecuteNonQuery(query);
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
                string query = string.Format(string.Format("UPDATE OrdersCheckN1 SET bwIsSetupN1 = {0}", status));
                return MySQLProvider.Instance.ExecuteNonQuery(query);
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
                string query = string.Format(string.Format("UPDATE OrdersCheckN1 SET cncIsSetupN1 = {0}", status));
                return MySQLProvider.Instance.ExecuteNonQuery(query);
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
                return Convert.ToInt32(MySQLProvider.Instance.ExecuteScalar("SELECT bwIsSetupN1 FROM OrdersCheckN1"));
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
                return Convert.ToInt32(MySQLProvider.Instance.ExecuteScalar("SELECT cncIsSetupN1 FROM OrdersCheckN1"));
            }
            catch (Exception ex)
            {

                throw ex;
            }
        }


        //2020-03-06 Start add: thêm chức năng cập nhật thông số Teeth cho BW11 và BW12
        public int Update_BW11_TeethInfo(OrderModel order)
        {
            string query = string.Format(@"UPDATE `BWHob11` SET `SoPO`='{0}',`TeethShape`='{1}',`TeethQty`={2},`Diameter_d`={3},`KnifeType`='{4}', `GlobalCode`='{5}'",order.Manufa_Instruction_No, order.TeethShape, order.TeethQuantity, order.Diameter_d, order.KnifeType, order.Global_Code);
            try
            {
                return MySQLProvider.Instance.ExecuteNonQuery(query);
            }
            catch (Exception ex)
            {

                throw ex;
            }
        }

        public int Update_BW12_TeethInfo(OrderModel order)
        {
            string query = string.Format(@"UPDATE `BWHob12` SET `SoPO`='{0}',`TeethShape`='{1}',`TeethQty`={2},`Diameter_d`={3},`KnifeType`='{4}', `GlobalCode`='{5}'", order.Manufa_Instruction_No, order.TeethShape, order.TeethQuantity, order.Diameter_d, order.KnifeType, order.Global_Code);
            try
            {
                return MySQLProvider.Instance.ExecuteNonQuery(query);
            }
            catch (Exception ex)
            {

                throw ex;
            }
        }

        public DataRow GetBW11_TeethInfo()
        {
            string query = string.Format(@"SELECT * FROM BWHob11 LIMIT 1");
            try
            {
                DataRow r = MySQLProvider.Instance.ExecuteQuery(query).Rows[0];
                if(r!= null)
                {
                    return r;
                }
                else return null;
            }
            catch (Exception ex)
            {

                throw ex;
            }
        }

        public DataRow GetBW12_TeethInfo()
        {
            string query = string.Format(@"SELECT * FROM BWHob12 LIMIT 1");
            try
            {
                DataRow r = MySQLProvider.Instance.ExecuteQuery(query).Rows[0];
                if (r != null)
                {
                    return r;
                }
                else return null;
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
            List<DateTime> lst = new List<DateTime>();
            try
            {
                DataTable dt = SQLProvider.Instance.ExecuteQuery("SELECT COnvert(varchar,Holidayday,111) Holiday FROM Holiday WHERE YEAR(Holidayday)= YEAR(GETDATE())");
                foreach (DataRow r in dt.Rows)
                {
                    lst.Add(Convert.ToDateTime(r["Holiday"]));
                }
                return lst;
            }
            catch (Exception ex)
            {
                throw ex;
            }
        }

        //lấy sum giá trị cột PULLEY_LINEB
        public double GetSumPulley_BW(int hour)
        {
            int[] sph1Hours = { 6,7,8, 9,10,11, 12,13,14, 15,16,17, 18,19,20, 21,22,23, 0 };
            int[] sph2Hours = { 2, 3, 4, 5 };
            string query = "";
            if(sph1Hours.Contains(hour))
            {
                query = string.Format(@"SELECT SUM(RemainPT) RemainPT FROM
                                        (SELECT SUM(PULLEY_LINEB) RemainPT FROM sph_hourly WHERE Date_ = TO_CHAR(SYSDATE,'yyyymmdd') AND Time_ >= {0}
                                        UNION
                                        SELECT SUM(PULLEY_LINEB) FROM sph_hourly WHERE Date_ = TO_CHAR(SYSDATE,'yyyymmdd') AND (Time_ >= 2 AND Time_ <6)
                                        ) tempA", hour);
            }
            else if(sph2Hours.Contains(hour))
            {
                query = string.Format(@"SELECT SUM(PULLEY_LINEB) RemainPT FROM sph_hourly WHERE Date_ = TO_CHAR(SYSDATE,'yyyymmdd') AND (Time_ >= {0} AND Time_ <6)", hour);
            }

            try
            {
                var kq = OracleProvider.Instance.ExecuteScalar(query);
                return kq != DBNull.Value ? Convert.ToDouble(kq) : 0;
            }
            catch (Exception ex)
            {

                throw ex;
            }
        }

        //lấy sum giá trị cột PULLEY_LINEB
        public double GetSumPulley_CNC(int hour)
        {
            int[] sph1Hours = { 6, 7, 8, 9, 10, 11, 12, 13, 14, 15, 16, 17, 18, 19, 20, 21, 22, 23, 0 };
            int[] sph2Hours = { 2, 3, 4, 5 };
            string query = "";
            if (sph1Hours.Contains(hour))
            {
                query = string.Format(@"SELECT SUM(RemainPT) RemainPT FROM
                                        (SELECT SUM(PULLEY_LINEA) RemainPT FROM sph_hourly WHERE Date_ = TO_CHAR(SYSDATE,'yyyymmdd') AND Time_ >= {0}
                                        UNION
                                        SELECT SUM(PULLEY_LINEA) FROM sph_hourly WHERE Date_ = TO_CHAR(SYSDATE,'yyyymmdd') AND (Time_ >= 2 AND Time_ <6)
                                        ) tempA", hour);
            }
            else if (sph2Hours.Contains(hour))
            {
                query = string.Format(@"SELECT SUM(PULLEY_LINEA) RemainPT FROM sph_hourly WHERE Date_ = TO_CHAR(SYSDATE,'yyyymmdd') AND (Time_ >= {0} AND Time_ <6)", hour);
            }


            try
            {
                var kq = OracleProvider.Instance.ExecuteScalar(query);
                return kq != DBNull.Value ? Convert.ToDouble(kq) : 0;
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
            string query = "SELECT IsLockedCNC9, BWCNCHob9_HasBWOrders FROM `CNCHob9`";
            int IsLockedCNC9 = 0;
            int BWCNCHob9_HasBWOrders = 0;
            try
            {
                DataTable dt = MySQLProvider.Instance.ExecuteQuery(query);

                IsLockedCNC9 = Convert.ToInt32(dt.Rows[0]["IsLockedCNC9"]);
                BWCNCHob9_HasBWOrders = Convert.ToInt32(dt.Rows[0]["BWCNCHob9_HasBWOrders"]);
                return new Tuple<int, int>(IsLockedCNC9,BWCNCHob9_HasBWOrders);
            }
            catch (Exception ex)
            {

                throw ex;
            }
        }

        public Tuple<int,int> GetLockStatusCNC10()
        {
            string query = "SELECT IsLockedCNC10, BWCNCHob10_HasBWOrders FROM `CNCHob10`";
            int IsLockedCNC10 = 0;
            int BWCNCHob10_HasBWOrders = 0;
            try
            {
                DataTable dt = MySQLProvider.Instance.ExecuteQuery(query);
                IsLockedCNC10 = Convert.ToInt32(dt.Rows[0]["IsLockedCNC10"]);
                BWCNCHob10_HasBWOrders = Convert.ToInt32(dt.Rows[0]["BWCNCHob10_HasBWOrders"]);
                return new Tuple<int,int>(IsLockedCNC10,BWCNCHob10_HasBWOrders);
            }
            catch (Exception ex)
            {

                throw ex;
            }
        }

        public int SetLockCNC9(int status)
        {
            string query = "UPDATE `CNCHob9` SET IsLockedCNC9 = "+status;
            int res = 0;
            try
            {
                res = MySQLProvider.Instance.ExecuteNonQuery(query);
                return res;
            }
            catch (Exception ex)
            {

                throw ex;
            }
        }

        public int SetLockCNC10(int status)
        {
            string query = "UPDATE `CNCHob10` SET IsLockedCNC10 = "+status;
            int res = 0;
            try
            {
                res = MySQLProvider.Instance.ExecuteNonQuery(query);
                return res;
            }
            catch (Exception ex)
            {

                throw ex;
            }
        }
        //2020-03-23 End add

    }
}