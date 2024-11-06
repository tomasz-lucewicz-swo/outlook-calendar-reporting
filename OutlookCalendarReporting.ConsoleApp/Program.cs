using OutlookCalendarReporting.ConsoleApp.Calendars.Outlook;
using OutlookCalendarReporting.ConsoleApp.Calendars;
using Microsoft.Extensions.DependencyInjection;
using Spectre.Console.Cli;
using OutlookCalendarReporting.ConsoleApp.Utils;
using OutlookCalendarReporting.ConsoleApp.Command;
using Spectre.Console;


namespace OutlookCalendarReporting.ConsoleApp
{
    class Program
    {
        internal static ITypeResolver TypeResolver { get; set; } = null!;

        internal static async Task<int> Main(string[] args)
        {
            ServiceCollection services = new();
            services.AddSingleton<ICalendar, OutlookCalendar>();

            TypeRegistrar registrar = new(services);

            Console.WriteLine("Options:");
            Console.WriteLine("1 - Weekly TimeEntries for SwoDP");
            Console.WriteLine("2 - Monthly TimeEntries for CSV / Excel");

            var option = Console.ReadLine();


            switch (option)
            {
                case "1":
                    CommandApp<SwoDPWeeklyTimeEntriesReportCommand> app1 = new(registrar);
                    TypeResolver = registrar.Build();
                    return await app1.RunAsync(args);
                case "2":
                    CommandApp<CSVMonthlyTimeEntriesReportCommand> app2 = new(registrar);
                    TypeResolver = registrar.Build();
                    return await app2.RunAsync(args);
                default:
                    break;
            }

            return 0;
        }

       
    }
}
