namespace OutlookCalendarReporting.ConsoleApp.Calendars
{
    internal interface ICalendar
    {
        IEnumerable<CalendarEntry> GetEntries(DateTime from, DateTime to);

        public bool IsAppAvailable { get; }
    }
}
