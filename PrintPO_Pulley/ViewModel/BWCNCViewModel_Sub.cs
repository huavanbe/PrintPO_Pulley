using BFDataCrawler.Model;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace BFDataCrawler.ViewModel
{
    class BWCNCViewModel_Sub
    {
        private static BWCNCViewModel_Sub instance;
        public static BWCNCViewModel_Sub Instance
        {
            get { return (instance == null) ? (instance = new BWCNCViewModel_Sub()) : instance; }
            set { instance = value; }
        }

        public bool IsF1Order(string orderName)
        {
            string parts = orderName.Remove(2);
            if (!string.IsNullOrEmpty(parts) && parts.ToLower().Contains("f1"))
                return true;
            else return false;
        }

        public List<OrderModel> Sort_F1Orders(List<OrderModel> lst, double d = 0)
        {
            return lst.OrderBy(c => Math.Abs(c.Diameter_d - d)).ThenBy(c=>c.Manufa_Instruction_No).ToList();
        }

        //chia đều số lượng item của list
        public List<List<T>> splitList<T>(List<T> locations, int nSize = 30)
        {
            var list = new List<List<T>>();

            for (int i = 0; i < locations.Count; i += nSize)
            {
                list.Add(locations.GetRange(i, Math.Min(nSize, locations.Count - i)));
            }

            return list;
        }

        public string HTMLTableOrders(List<OrderModel> orders, string machine)
        {
            StringBuilder str = new StringBuilder();
            str.AppendLine("Orders "+machine+": </br>");
            str.AppendLine("<table>");
            str.AppendLine("<head>");
            str.AppendLine("<tr>");
            str.AppendLine("<th>Số PO</th> <th>Tên</th> <th>Ngày xuất</th> <th>Dao</th> <th>Số răng</th> <th>D</th>");
            str.AppendLine("<tr>");
            str.AppendLine("</head>");
            str.AppendLine("<tbody>");
            foreach (var od in orders)
            {
                str.AppendLine("<tr>");
                str.AppendLine(string.Format("<td>{0}</td> <td>{1}</td> <td>{2}</td> <td>{3}</td> <td>{4}</td> <td>{5}</td>",od.Manufa_Instruction_No, od.Item_Name, od.Factory_Ship_Date, od.TeethShape, od.TeethQuantity, od.Diameter_d));
                str.AppendLine("</tr>");
            }
            str.AppendLine("</tbody>");
            str.AppendLine("</table>");
            return str.ToString();
        }

        public string GetF1OrdersInnerCode(string innerCode)
        {
            if (innerCode.StartsWith("F1"))
            {
                var newStr = innerCode.Substring(innerCode.IndexOf('-') + 1, innerCode.Length - (innerCode.IndexOf('-') + 1));
                newStr = newStr.Substring(0, newStr.LastIndexOf('-'));
                return newStr;
            }
            else
                return innerCode;
        }


        //chưa sử dụng..
        public List<OrderModel> SortByWithGlobalCode(List<OrderModel> source, string shape, int qty, double d, string gbcode)
        {
            List<OrderModel> result = new List<OrderModel>();

            //lấy ra PO đầu tiên và các thông số Size, Type, D, V
            OrderModel firstPO = source.FirstOrDefault();
            if (firstPO != null)
            {
                //tìm các PO có cùng Size, D, V, Type như PO đầu
                var new_orders = source.Where(c => c.TeethShape == firstPO.TeethShape && c.TeethQuantity.Equals(firstPO.TeethQuantity) && c.Diameter_d == firstPO.Diameter_d).ToList();
                if (new_orders != null && new_orders.Count > 0)
                {
                    //group các PO có cùng Size, D, V, Type như PO đầu theo GlobalCode
                    var g_new_orders = new_orders.GroupBy(c => c.Global_Code).Select(g => new { gKey = g.Key, gItems = g.ToList().OrderBy(i => i.Manufa_Instruction_No) }).ToList();
                    if (g_new_orders != null && g_new_orders.Count > 0)
                    {
                        //tìm ra group có GlobalCode giống với GlobalCode cho trước
                        var g_withSameGbCode = g_new_orders.Where(g => g.gKey.Equals(gbcode)).ToList();
                        if (g_withSameGbCode != null && g_withSameGbCode.Count > 0)
                        {
                            //thêm group có GlobalCode giống với GlobalCode cho trước vào list 
                            g_withSameGbCode.ForEach(g => result.AddRange(g.gItems));
                            //remove group order khỏi source
                            source = source.Except(result).ToList();
                            //add các order còn lại của source tiếp theo
                            result.AddRange(source);
                        }
                        else
                        {
                            result = source;
                        }
                    }

                }
                else
                {
                    result = source;
                }

            }

            return result;
        }
    }
}
