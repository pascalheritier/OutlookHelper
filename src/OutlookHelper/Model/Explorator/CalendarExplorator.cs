namespace OutlookHelper
{
    internal abstract class CalendarExplorator // in case we want another explorator, some refactoring will be involved
    {
        public abstract void DisplayCalendarAllYearsSummary(SortedCalendar sortedCalendar);
        public abstract void DisplayCalendarAllWeeksSummary(SortedCalendar sortedCalendar);
        public abstract void DisplayCalendarWeeklySummary(SortedCalendar sortedCalendar, int year, int week);
        public abstract void DisplayCalendarAllDaysSummary(SortedCalendar sortedCalendar);
        public abstract void DisplayCalendarDailySummary(SortedCalendar sortedCalendar, DateTime dayToDisplay);
    }
}