using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Mail.WindowsService.DateClasses
{
    public class Dates
    {

        private static DateTime FirstDayOfTheMonth { get; set; }
        private static DateTime Today { get; set; }

        /// <summary>
        /// Returns first day of the month as DateTime object.
        /// </summary>
        /// <returns></returns>
        public static DateTime ReturnFirstDayOfTheMonth()
        {

            FirstDayOfTheMonth = new DateTime(DateTime.Now.Year, (DateTime.Now.Month + 1), DateTime.Now.Day);

            return FirstDayOfTheMonth;

        }

        /// <summary>
        /// Returns current day as DateTime object.
        /// </summary>
        /// <returns></returns>
        public static DateTime ReturnToday()
        {

            Today = new DateTime(DateTime.Now.Year, (DateTime.Now.Month + 1), DateTime.Now.Day);

            return Today;

        }

    }
}
