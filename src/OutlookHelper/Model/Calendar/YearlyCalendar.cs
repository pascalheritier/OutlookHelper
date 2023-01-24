using Microsoft.Office.Interop.Outlook;

namespace OutlookHelper
{
    internal class YearlyCalendar
    {
        public int Year { get; set; }

        public TimeSpan YearlyTimeSpent { get; set; }

        public List<WeeklyCalendar> WeeklyCalendars { get; set; } = new();

        public IEnumerable<AppointmentItem> YearlyAppointments => WeeklyCalendars.SelectMany(_W => _W.Appointments);

        public double ComputeOvertime(double workingPercentage)
        {
            return WeeklyCalendars.Sum(_WeeklyCalendar => _WeeklyCalendar.ComputeOvertime(workingPercentage));
        }
    }
}
