using OutlookCalendarReporting.ConsoleApp.BusinessLogic;
using OutlookCalendarReporting.ConsoleApp.Calendars;
using OutlookCalendarReporting.ConsoleApp.Class;
using Spectre.Console;
using Spectre.Console.Cli;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace OutlookCalendarReporting.ConsoleApp.Command
{
    internal class CSVMonthlyTimeEntriesReportCommand : AsyncCommand
    {
        private readonly ICalendar _calendar;

        public CSVMonthlyTimeEntriesReportCommand(ICalendar calendar)
        {
            _calendar = calendar;
        }

        private static (DateTime From, DateTime To) GetDates()
        {
            var fromDate = new DateTime(DateTime.Today.Year, DateTime.Today.Month, 1);

            return (fromDate, fromDate.AddMonths(1).AddDays(-1));
        }

        private CalendarEntry[] GetCalendarEntries(DateTime fromDate, DateTime toDate)
        {
            AnsiConsole.WriteLine("Getting calendar entries...");
            var calendarEntries = _calendar.GetEntries(fromDate, toDate).ToArray();
            AnsiConsole.WriteLine($"Finished getting calendar entries. Total count: {calendarEntries.Length}.");

            return calendarEntries;
        }

        public override async Task<int> ExecuteAsync(CommandContext context)
        {
            if (_calendar.IsAppAvailable == false)
            {
                Console.ForegroundColor = ConsoleColor.Red;
                Console.WriteLine("Outlook (Classic) is not turned on.");

                return -1;
            }

            try
            {
                var (fromDate, toDate) = GetDates();
                var calendarEntries = GetCalendarEntries(fromDate, toDate);

                var aggregatedEntries = CalendarBL.Aggregate(fromDate, toDate, calendarEntries);
                

                AnsiConsole.WriteLine("Time to Report: " + (aggregatedEntries.Sum(x => x.NotIgnoredDuration) / 60.0) + " h\n");

                FormatResponse(aggregatedEntries);

                AnsiConsole.WriteLine("Execution finished");
                return 0;
            }
            catch (Exception exception)
            {
                AnsiConsole.MarkupLineInterpolated($"[red]{exception.Message}[/]");
                return -1;
            }
        }

        private void FormatResponse(List<CalendarEntryAggregator> aggregatedEntries)
        {
            foreach (var entry in aggregatedEntries.OrderBy(x => x.Day))
            {
                if (entry.NotIgnoredDuration > 0)
                {
                    // 8,5%
                    foreach (var calendarEntry in entry.Admin)
                    {
                        AnsiConsole.WriteLine(entry.Day.ToShortDateString() + ";" + calendarEntry.Subject + ";" + "" + ";" + (calendarEntry.Duration / 60.0).ToString().Replace(",", "."));
                    }
                    // 12%
                    foreach (var calendarEntry in entry.Training)
                    {
                        AnsiConsole.WriteLine(entry.Day.ToShortDateString() + ";" + calendarEntry.Subject + ";" + (calendarEntry.Duration / 60.0).ToString().Replace(",",".") + ";" + "");
                    }
                    // 12%
                    foreach (var calendarEntry in entry.BDEV)
                    {
                        AnsiConsole.WriteLine(entry.Day.ToShortDateString() + ";" + calendarEntry.Subject + ";" + (calendarEntry.Duration / 60.0).ToString().Replace(",", ".") + ";" + "");
                    }
                }
            }
        }
    }
}
