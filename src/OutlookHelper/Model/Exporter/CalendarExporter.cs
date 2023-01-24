namespace OutlookHelper
{
    internal abstract class CalendarExporter
    {
        #region Export calendar data

        public abstract void ExportCalendarData(SortedCalendar sortedCalendar, DateTime dayToExport);
        public abstract void ExportCalendarData(SortedCalendar sortedCalendar, int year, int weekToExport);

        #endregion
    }
}
