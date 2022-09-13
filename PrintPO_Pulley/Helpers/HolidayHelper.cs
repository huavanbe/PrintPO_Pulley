using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace BFDataCrawler.Helpers
{
    class HolidayHelper
    {
        private static HolidayHelper instance;

        public static HolidayHelper Instance
        {
            get
            {
                if (instance == null)
                    instance = new HolidayHelper();

                return instance;
            }
            set
            {
                instance = value;
            }
        }

        private IList<string> listHoliday;
        public IList<string> ListHoliday
        {
            get
            {
                if (listHoliday == null)
                {
                    listHoliday = new List<string>();
                    IList<DateTime> holidays = BLL.OrderBLL.Instance.GetHolidaysInYear();
                    foreach (var d in holidays)
                    {
                        listHoliday.Add(d.ToString("yyyyMMdd"));
                    }
                }


                return listHoliday;
            }
            set
            {
                listHoliday = value;
            }
        }

        private string dayN1;
        public string DayN1
        {
            get
            {
                DateTime today = DateTime.Now;
                DateTime flagN1 = today.DayOfWeek == DayOfWeek.Saturday ? today.AddDays(2) : today.AddDays(1); //default add 1, if saturday add 2
                IList<string> holidays = HolidayHelper.Instance.ListHoliday;
                while (holidays.Contains(flagN1.ToString("yyyyMMdd")) || flagN1.DayOfWeek == DayOfWeek.Sunday)
                {
                    flagN1 = flagN1.AddDays(1);
                }

                return dayN1 = flagN1.ToString("yyyyMMdd");
            }
            set { dayN1 = value; }
        }

        private int dayCountN1;
        public int DayCountN1
        {
            get
            {
                DateTime n1 = DateTime.ParseExact(DayN1, "yyyyMMdd", null);
                dayCountN1 = (n1.Date - DateTime.Now.Date).Days;
                return dayCountN1;
            }
        }
    }
}
