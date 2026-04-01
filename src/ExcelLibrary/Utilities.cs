using System;
using System.Collections.Generic;
using System.Globalization;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ExcelLibrary
{
    class Utilities
    {
        public static string ConvertDate(string excelDate, int baseYear)
        {
            DateTime convertedDate = new DateTime();
            const int MYSTERIOUS_CONSTANT = 2;
            convertedDate = DateTime.Parse(string.Format("{0}-01-01", baseYear));
            double daysToAdd = double.Parse(excelDate) - MYSTERIOUS_CONSTANT;
            convertedDate = convertedDate.AddDays(daysToAdd);
            return convertedDate.ToShortDateString();
        }

        public static string ConvertTime(string excelTime)
        {
            double time = double.Parse(excelTime, CultureInfo.GetCultureInfo("en-us"));
            double second = 1 / 86400d;
            double seconds = time / second;
            TimeSpan span = TimeSpan.FromSeconds(seconds);
            return span.ToString();
        }
    }
}
