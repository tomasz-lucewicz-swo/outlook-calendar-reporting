using Microsoft.Office.Interop.Outlook;
using OutlookCalendarReporting.ConsoleApp.BusinessLogic;
using OutlookCalendarReporting.ConsoleApp.Calendars;
using OutlookCalendarReporting.ConsoleApp.Class;
using Spectre.Console;
using Spectre.Console.Cli;

namespace OutlookCalendarReporting.ConsoleApp.Command
{
    internal class SwoDPWeeklyTimeEntriesReportCommand : AsyncCommand
    {
        private readonly ICalendar _calendar;

        public SwoDPWeeklyTimeEntriesReportCommand(ICalendar calendar)
        {
            _calendar = calendar;
        }

        private static (DateTime From, DateTime To) GetDates()
        {
            var fromDate = DateTime.Today.AddDays(-(int)DateTime.Today.DayOfWeek + (int)DayOfWeek.Monday);

            return (fromDate, fromDate.AddDays(7));
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
            catch (System.Exception exception)
            {
                AnsiConsole.MarkupLineInterpolated($"[red]{exception.Message}[/]");
                return -1;
            }
        }

        

        private CalendarEntry[] GetCalendarEntries(DateTime fromDate, DateTime toDate)
        {
            AnsiConsole.WriteLine("Getting calendar entries...");
            var calendarEntries = _calendar.GetEntries(fromDate, toDate).ToArray();
            AnsiConsole.WriteLine($"Finished getting calendar entries. Total count: {calendarEntries.Length}.");

            return calendarEntries;
        }

        private void FormatResponse(List<CalendarEntryAggregator> aggregatedEntries)
        {
            foreach (var entry in aggregatedEntries.OrderBy(x => x.Day))
            {
                String result = $"Day: {entry.Day.ToShortDateString()}, Not Ingnored Duration: " + (entry.NotIgnoredDuration / 60.0) + " h" + "\n";

                if (entry.Admin.Sum(x => x.Duration) > 0)
                {
                    result += $"Admin:" + (entry.Admin.Sum(x => x.Duration) / 60.0) + " h\n";
                    foreach (var item in entry.Admin)
                    {
                        result += $"\t{item.Subject}\n";
                    }
                }

                if (entry.Training.Sum(x => x.Duration) > 0)
                {
                    result += $"Training:" + (entry.Training.Sum(x => x.Duration) / 60.0) + " h\n";
                    if (entry.Training is not null)
                    {
                        foreach (var item in entry.Training)
                        {
                            result += $"\t{item.Subject}\n";
                        }
                    }
                }

                if (entry.BDEV.Sum(x => x.Duration) > 0)
                {
                    result += $"BDEV:" + (entry.BDEV.Sum(x => x.Duration) / 60.0) + " h\n";
                    foreach (var item in entry.BDEV)
                    {
                        result += $"\t{item.Subject}\n";
                    }
                }

                if (entry.Ignored.Sum(x => x.Duration) > 0)
                {
                    result += $"Ignored:" + (entry.Ignored.Sum(x => x.Duration) / 60.0) + " h\n";
                    foreach (var item in entry.Ignored)
                    {
                        result += $"\t{item.Subject}\n";
                    }
                }

                result += $"\n";

                AnsiConsole.WriteLine(result);
            }
        
        }
    }
}
