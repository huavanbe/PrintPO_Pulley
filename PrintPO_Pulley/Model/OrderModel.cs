using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace BFDataCrawler.Model
{
    class OrderModel
    {
        public OrderModel(DataRow dr)
        {
            //sử dụng cho việc lấy thông tin của PO để đồng bộ BWCNC và CNC
            Manufa_Instruction_No = dr["SoPO"].ToString();
            TeethShape = dr["TeethShape"].ToString();
            TeethQuantity = Convert.ToInt32(dr["TeethQty"]);
            Diameter_d = Convert.ToDouble(dr["Diameter_d"]);
            Material_text1 = dr["DMat"].ToString();
        }
        
        public OrderModel()
        { }
        public string No { get; set; }
        public string Sales_Order_No { get; set; }
        public string Pier_Instruction_No { get; set; }
        public string Manufa_Instruction_No { get; set; }
        public string Global_Code { get; set; }
        
        public string Customers { get; set; }
        public string Item_Name { get; set; }
        public string MC { get; set; }
        public DateTime? Received_Date { get; set; }
        public DateTime? Factory_Ship_Date { get; set; }
        public string Number_of_Orders { get; set; }
        public string Number_of_Available_Instructions { get; set; }
        
        public string Number_of_Repairs { get; set; }
        
        public string Number_of_Instructions { get; set; }
        public string Line { get; set; }
        
        public string PayWard { get; set; }
        
        public string Major { get; set; }
        
        public string Special_Orders { get; set; }
        
        public string Method { get; set; }
        
        public string Destination { get; set; }
        
        public DateTime? Instructions_Print_date { get; set; }
        
        public string Latest_progress { get; set; }
        
        public string Tack_Label_Output_Date { get; set; }
        
        public DateTime? Completion_Instruction_Date { get; set; }
        
        public string Re_print_Count { get; set; }
        
        public string Latest_issue_time { get; set; }
        
        public string History1 { get; set; }
        
        public string History2 { get; set; }
        
        public string History3 { get; set; }
        
        public string History4 { get; set; }
        
        public string Material_code1 { get; set; }
        public string Material_text1 { get; set; }
        
        public string Amount_used1 { get; set; }
        
        public string Unit1 { get; set; }
        
        public string Material_code2 { get; set; }
        
        public string Material_text2 { get; set; }
        
        public string Amount_used2 { get; set; }
        
        public string Unit2 { get; set; }
        
        public string Material_code3 { get; set; }
        
        public string Material_text3 { get; set; }
        
        public string Amount_used3 { get; set; }
        
        public string Unit3 { get; set; }
        
        public string Material_code4 { get; set; }
        
        public string Material_text4 { get; set; }
        
        public string Amount_used4 { get; set; }
        
        public string Unit4 { get; set; }
        
        public string Material_code5 { get; set; }
        
        public string Material_text5 { get; set; }
        
        public string Amount_used5 { get; set; }
        
        public string Unit5 { get; set; }
        
        public string Material_code6 { get; set; }
        
        public string Material_text6 { get; set; }
        
        public string Amount_used6 { get; set; }
        
        public string Unit6 { get; set; }
        
        public string Material_code7 { get; set; }
        
        public string Material_text7 { get; set; }
        
        public string Amount_used7 { get; set; }
        
        public string Unit7 { get; set; }
        
        public string Material_code8 { get; set; }
        
        public string Material_text8 { get; set; }
        
        public string Amount_used8 { get; set; }
        
        public string Unit8 { get; set; }
        public string Inner_Code { get; set; }
        
        public string Classify_Code { get; set; }

        public bool Hobbing { get; set; }
        public int TeethQuantity { get; set; }
        //2019-12-04 Start change: new property for sorting Orders (column that displays TeethShape now change to this)
        public string TeethShape { get; set; }
        //2019-12-04 End change

        //2019-10-05 Start add: new property to identify which PO is BW, which PO is CNC
        public string LineAB { get; set; }
        //2019-10-05 End add

        //2019-10-16 Start add: new property to indentify which PO is line A, line B
        public string LineName { get; set; }
        //2019-10-16 End add

        //2019-11-28 Start add: new property to identify to get d_HobShaft
        public double Diameter_d { get; set; }
        //2019-11-28 End add

       
        public string KnifeType { get; set; }       
    }
}
