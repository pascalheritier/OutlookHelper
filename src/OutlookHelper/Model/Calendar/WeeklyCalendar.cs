using Microsoft.Office.Interop.Outlook;

namespace OutlookHelper
{
    internal class WeeklyCalendar
    {
        public int Week { get; set; }

        public TimeSpan WeeklyTimeSpent { get; set; }

        public List<AppointmentItem> Appointments { get; set; } = new();

        public double ComputeOvertime(double workingPercentage)
        {
            return WeeklyTimeSpent.TotalHours - Utils.DailyHoursBasis * Utils.DaysPerWeek * workingPercentage;
        }
    }
}
