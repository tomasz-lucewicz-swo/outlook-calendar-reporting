using System;
using Microsoft.Office.Interop.Outlook;
using Spectre.Console;

namespace OutlookCalendarReporting.ConsoleApp.Calendars.Outlook
{
    internal sealed class OutlookCalendar : ICalendar
    {
        private const string SubjectSeparator = " - ";
        private const char IgnoreChar = '@';

        private readonly Application _app;

        public OutlookCalendar()
        {
            try
            {
                _app = new();
            }
            catch
            {
                _app = null!;
            }
        }

        public bool IsAppAvailable => _app is not null;

        public IEnumerable<CalendarEntry> GetEntries(DateTime from, DateTime to)
        {
            if (_app is null)
            {
                return Array.Empty<CalendarEntry>();
            }

            var mapiNamespace = _app.GetNamespace("MAPI");
            var calendarFolder = mapiNamespace.GetDefaultFolder(OlDefaultFolders.olFolderCalendar);
            var items = calendarFolder.Items;
            items.IncludeRecurrences = false;

            var calendarEntries = items
                .Cast<AppointmentItem>()
                .Where(item => item.Start >= from && item.Start <= to && !item.IsRecurring)
                .Select(TryFrom)
                .Where(calendarEvent => calendarEvent is not null)
                .Cast<CalendarEntry>();

            return calendarEntries;
        }

        private static CalendarEntry? TryFrom(AppointmentItem item)
        {
            if (item.Subject.StartsWith(IgnoreChar))
                return null;

            var projectCode = item.Subject.Split(SubjectSeparator).Select(part => part.Trim()).First();
            return new CalendarEntry(item.Subject, projectCode, item.Start, item.Duration, item.AllDayEvent, item);
        }
    }
}
