using System;
using System.Collections.Generic;
using System.Globalization;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;
using System.Threading.Tasks;

namespace testings
{
    public class Program
    {
        static void Main(string[] args)
        {
            string dateString = "20.02.18";
            string day = DateTime.ParseExact(dateString, "dd.MM.yy", CultureInfo.InvariantCulture).DayOfWeek.ToString();

            Console.WriteLine(day);
            Console.ReadKey();
        }
    }
}
