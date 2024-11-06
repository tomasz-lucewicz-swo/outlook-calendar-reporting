using OutlookCalendarReporting.ConsoleApp.Calendars;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace OutlookCalendarReporting.ConsoleApp.Class
{
    public class CalendarEntryAggregator
    {
        public DateTime Day { get; set; }

        public List<CalendarEntry> Admin { get; set; }

        public List<CalendarEntry> Training { get; set; }

        public List<CalendarEntry> BDEV { get; set; }

        public List<CalendarEntry> Ignored { get; set; }

        public int NotIgnoredDuration => Admin.Sum(x => x.Duration) + Training.Sum(x => x.Duration) + BDEV.Sum(x => x.Duration);

        public CalendarEntryAggregator(DateTime day)
        {
            Day = day;
            Admin = new List<CalendarEntry>();
            Training = new List<CalendarEntry>();
            BDEV = new List<CalendarEntry>();
            Ignored = new List<CalendarEntry>();
        }            
    }


}
