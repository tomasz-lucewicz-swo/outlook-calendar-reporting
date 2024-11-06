namespace OutlookCalendarReporting.ConsoleApp.Calendars
{
    public record CalendarEntry
    (
        string Subject,
        string ProjectCode,
        DateTime Start,
        int Duration,
        bool AllDay,
        object SourceObject
    );
}
