using BFDataCrawler.Model;
using System;
using System.Collections.Generic;
using System.Globalization;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;
using System.Threading.Tasks;

namespace BFDataCrawler.Helper
{
    class GeneralUtility
    {
        private static GeneralUtility instance;
        public static GeneralUtility Instance
        {
            get {
                if (instance == null)
                    instance = new GeneralUtility();
                return instance;
            }
            set
            {
                instance = value;
            }
        }

        public int SelectNumPart(string inputtext)
        {
            int a = 0;
            Regex re2 = new Regex(@"([0-9.]+)");
            Match result2 = re2.Match(inputtext);
            string numberPart = result2.Groups[1].Value;
            return Int32.TryParse(numberPart,out a) ? Convert.ToInt32(numberPart) : 0;
        }

        public double SelectNumPartDouble(string inputtext)
        {
            double a = 0;
            Regex re2 = new Regex(@"([0-9.]+)");
            Match result2 = re2.Match(inputtext);
            string numberPart = result2.Groups[1].Value;
            return Double.TryParse(numberPart, out a) ? Convert.ToDouble(numberPart) : 0;
        }

        public string SelectCharPart(string inputtext)
        {
            Regex re1 = new Regex(@"([a-zA-Z]+)(\d+)");
            Match result1 = re1.Match(inputtext);
            //string alphaPart = result1.Groups[1].Value;
            return result1.Value;
        }

       

        public int GetWeekNumberOfMonth(DateTime date)
        {
            date = date.Date;
            DateTime firstMonthDay = new DateTime(date.Year, date.Month, 1);
            DateTime firstMonthMonday = firstMonthDay.AddDays((DayOfWeek.Monday + 7 - firstMonthDay.DayOfWeek) % 7);
            if (firstMonthMonday > date)
            {
                firstMonthDay = firstMonthDay.AddMonths(-1);
                firstMonthMonday = firstMonthDay.AddDays((DayOfWeek.Monday + 7 - firstMonthDay.DayOfWeek) % 7);
            }
            return (date - firstMonthMonday).Days / 7 + 1;
        }
        public int GetIso8601WeekOfYear(DateTime time)
        {
            // Seriously cheat.  If its Monday, Tuesday or Wednesday, then it'll 
            // be the same week# as whatever Thursday, Friday or Saturday are,
            // and we always get those right
            DayOfWeek day = CultureInfo.InvariantCulture.Calendar.GetDayOfWeek(time);
            if (day >= DayOfWeek.Monday && day <= DayOfWeek.Wednesday)
            {
                time = time.AddDays(3);
            }

            // Return the week of our adjusted day
            return CultureInfo.InvariantCulture.Calendar.GetWeekOfYear(time, CalendarWeekRule.FirstFourDayWeek, DayOfWeek.Monday);
        }

    }
}
