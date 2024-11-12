using OutlookCalendarReporting.ConsoleApp.Calendars;
using OutlookCalendarReporting.ConsoleApp.Class;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace OutlookCalendarReporting.ConsoleApp.BusinessLogic
{
    public class CalendarBL
    {
        public static List<CalendarEntryAggregator> Aggregate(DateTime fromDate, DateTime toDate, CalendarEntry[] calendarEntries)
        {
            List<CalendarEntryAggregator> result = new List<CalendarEntryAggregator>();

            DateTime day = fromDate;

            while (day <= toDate)
            {
                CalendarEntryAggregator aggregator = new CalendarEntryAggregator(day);

                foreach (var entry in calendarEntries.Where(x => x.Start.Day == day.Day))
                {
                    if (entry.ProjectCode.ToLower() == "mgmt" || entry.ProjectCode.ToLower() == "hr" || entry.ProjectCode.ToLower() == "recruit" || entry.ProjectCode.ToLower() == "travel")
                    {
                        aggregator.Admin.Add(entry);
                    }
                    else if (entry.ProjectCode.ToLower() == "training")
                    {
                        aggregator.Training.Add(entry);
                    }
                    else if (entry.ProjectCode.ToLower() == "bdev")
                    {
                        aggregator.BDEV.Add(entry);
                    }
                    else
                    {
                        aggregator.Ignored.Add(entry);
                    }
                }

                result.Add(aggregator);

                day = day.AddDays(1);
            }

            return result;
        }
    }
}
