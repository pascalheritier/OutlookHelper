using Redmine.Net.Api.Types;
using Redmine.Net.Api;
using System.Collections.Specialized;
using Microsoft.Office.Interop.Outlook;
using System.Text.RegularExpressions;
using Microsoft.Extensions.Logging;

namespace OutlookHelper
{
    internal class RedmineCalendarExporter : CalendarExporter
    {
        #region Members

        private string _serverUrl;
        private string _apiKey;
        private string _targetUserId;
        private List<Activity> _timeEntryActivities;
        List<CategoryToActivityMapping> _categoryToActivityMappings;
        private RedmineManager _manager;
        private ILogger _logger;

        #endregion

        #region Constructor

        public RedmineCalendarExporter(
            string serverUrl,
            string apiKey,
            string targetUserId,
            List<Activity> timeEntryActivities,
            List<CategoryToActivityMapping> categoryToActivityMappings,
            ILogger logger)
        {
            _serverUrl = serverUrl;
            _apiKey = apiKey;
            _targetUserId = targetUserId;
            _timeEntryActivities = timeEntryActivities;
            _categoryToActivityMappings = categoryToActivityMappings;
            _manager = new RedmineManager(_serverUrl, _apiKey);
            _logger = logger;
        }

        #endregion

        #region Export calendar data

        public override void ExportCalendarData(SortedCalendar sortedCalendar, DateTime dayToExport)
        {
            YearlyCalendar? yearlyCalendar = sortedCalendar
                .YearlyCalendars.FirstOrDefault(_Y => _Y.Year == dayToExport.Year);
            if (yearlyCalendar is null)
            {
                _logger.LogError($"Export failed, there is no year {dayToExport.Year} in current calendar.");
                return;
            }

            IEnumerable<AppointmentItem> dailyAppointments = yearlyCalendar.YearlyAppointments.Where(_A => _A.Start.Date == dayToExport);
            if (!dailyAppointments.Any())
            {
                _logger.LogError($"Export failed, there is no appointments for day {dayToExport.ToString(Utils.DisplayDateFormat)} in current calendar.");
                return;
            }

            if (!CheckDailyAppointmentsAreEmpty(dayToExport))
                return;

            ExportCalendarDataAppointments(dailyAppointments);
        }

        public override void ExportCalendarData(SortedCalendar sortedCalendar, int year, int weekToExport)
        {
            WeeklyCalendar? weeklyCalendar = sortedCalendar
                .YearlyCalendars.FirstOrDefault(_Y => _Y.Year == year)?
                .WeeklyCalendars.FirstOrDefault(_W => _W.Week == weekToExport);

            if (weeklyCalendar is null)
            {
                _logger.LogError($"Export failed, there is no week {weekToExport} for year {year} in current calendar.");
                return;
            }

            if (!weeklyCalendar.Appointments.Any())
            {
                _logger.LogError($"Export failed, there is no appointments for week {weeklyCalendar.Week} in current calendar.");
                return;
            }

            if (!CheckWeeklyAppointmentsAreEmpty(year, weekToExport))
                return;

            ExportCalendarDataAppointments(weeklyCalendar.Appointments);
        }

        private void ExportCalendarDataAppointments(IEnumerable<AppointmentItem> appointmentItemsToExport)
        {
            foreach (AppointmentItem appointmentItemToExport in appointmentItemsToExport)
                ExportCalendarDataAppointment(appointmentItemToExport);
        }

        private void ExportCalendarDataAppointment(AppointmentItem appointmentItem)
        {
            try
            {
                // find a corresponding activity that could be mapped from outlook calendar item
                if (!TryGetActivity(appointmentItem, out Activity? foundActivity) || foundActivity is null)
                {
                    if (!_timeEntryActivities.Any())
                        foundActivity = new Activity { Id = 0 };
                    else
                        foundActivity = _timeEntryActivities.First();
                }

                // create a new time entry
                TimeEntry timeEntry = new()
                {
                    SpentOn = appointmentItem.Start,
                    Activity = IdentifiableName.Create<IdentifiableName>(foundActivity.Id),
                    Comments = appointmentItem.Subject,
                    Hours = (decimal)TimeSpan.FromMinutes(appointmentItem.Duration).TotalHours
                };

                // find if an issue is linked to appointment
                Issue? foundIssue;
                if (!TryGetCorrespondingIssue(appointmentItem, out foundIssue) || foundIssue is null)
                    _logger.LogWarning($"Could not find any corresponding issue linked to the current appointment {appointmentItem.Subject}," +
                        $" no issue will be linked to the new time entry.");
                else
                    timeEntry.Issue = IdentifiableName.Create<IdentifiableName>(foundIssue.Id);

                // save the new time entry
                TimeEntry savedTimeEntry = _manager.CreateObject(timeEntry);
                _logger.LogInformation($"Export of appointment '{appointmentItem.Start} - {appointmentItem.Subject}' as a time entry succeeded, entry id is: {savedTimeEntry.Id}.");
            }
            catch (System.Exception e)
            {
                _logger.LogError($"FAILURE: Export of appointment '{appointmentItem.Start} - {appointmentItem.Subject}' failed, with error message: {e.Message} - {e.InnerException?.Message}");
            }
        }

        #endregion

        #region Check empty appointments

        private bool CheckWeeklyAppointmentsAreEmpty(int year, int weekToExport)
        {
            try
            {
                DateTime beginWeekDay = Utils.GetFirstDateOfWeekISO8601(year, weekToExport);
                DateTime endWeekDay = beginWeekDay.AddDays(5);
                var parameters = new NameValueCollection
                {
                    { RedmineKeys.USER_ID, _targetUserId },
                    { RedmineKeys.SPENT_ON, $"><{beginWeekDay.ToString(Utils.RedmineDateFormat)}|{endWeekDay.ToString(Utils.RedmineDateFormat)}" }
                };
                IEnumerable<TimeEntry> weeklyTimeEntries = _manager.GetObjects<TimeEntry>(parameters);
                if (weeklyTimeEntries is null || !weeklyTimeEntries.Any())
                    return true;

                string errorMessage = $"Export failed, there are currently time entries for week {weekToExport}";
                foreach (var timeEntry in _manager.GetObjects<TimeEntry>(parameters))
                    errorMessage += Environment.NewLine + string.Format("Time entry details: {0} / {1} / {2} / {3}h.", timeEntry.SpentOn, timeEntry.Activity.Name, timeEntry.Comments, timeEntry.Hours);
                _logger.LogError(errorMessage);
                return false;
            }
            catch (System.Exception e)
            {
                _logger.LogError($"Export failed for week {weekToExport}, with error message: {e.Message}");
                return false;
            }
        }

        private bool CheckDailyAppointmentsAreEmpty(DateTime dayToExport)
        {
            try
            {
                var parameters = new NameValueCollection
                {
                    { RedmineKeys.USER_ID, _targetUserId },
                    { RedmineKeys.SPENT_ON, dayToExport.ToString(Utils.RedmineDateFormat) }
                };
                IEnumerable<TimeEntry>? dailyTimeEntries = _manager.GetObjects<TimeEntry>(parameters);
                if (dailyTimeEntries is null || !dailyTimeEntries.Any())
                    return true;

                string errorMessage = $"Export failed, there are currently time entries for day {dayToExport.ToString(Utils.DisplayDateFormat)}";
                foreach (var timeEntry in _manager.GetObjects<TimeEntry>(parameters))
                    errorMessage += Environment.NewLine + string.Format("Time entry details: {0} / {1} / {2} / {3}h.", timeEntry.SpentOn, timeEntry.Activity.Name, timeEntry.Comments, timeEntry.Hours);
                _logger.LogError(errorMessage);
                return false;
            }
            catch (System.Exception e)
            {
                _logger.LogError($"Export failed for day {dayToExport.ToString(Utils.DisplayDateFormat)}, with error message: {e.Message}");
                return false;
            }
        }
        #endregion

        #region Get issue for appointment

        private bool TryGetCorrespondingIssue(AppointmentItem appointmentItem, out Issue? foundIssue)
        {
            foundIssue = null;
            Regex regex = new("#([0-9]+)");
            Match match = regex.Match(appointmentItem.Subject);
            if (match.Success)
            {
                if (match.Groups.Count > 1)
                {
                    string? targetIssueId = match.Groups.Values.ElementAt(1)?.Value;
                    if (targetIssueId is not null)
                    {
                        if (TryGetIssueFromOpenIssues(targetIssueId, out foundIssue))
                            return true;
                        if (TryGetIssueFromClosedIssues(targetIssueId, out foundIssue))
                            return true;
                    }
                }
            }
            return false;
        }

        private bool TryGetIssueFromOpenIssues(string targetIssueId, out Issue? foundIssue)
        {
            var parameters = new NameValueCollection
            {
                { RedmineKeys.ISSUE_ID, targetIssueId }
            };
            return TryGetIssue(targetIssueId, parameters, out foundIssue);
        }

        private bool TryGetIssueFromClosedIssues(string targetIssueId, out Issue? foundIssue)
        {
            var parameters = new NameValueCollection
            {
                { RedmineKeys.ISSUE_ID, targetIssueId },
                { RedmineKeys.STATUS_ID, "closed" }
            };
            return TryGetIssue(targetIssueId, parameters, out foundIssue);
        }

        private bool TryGetIssue(string targetIssueId, NameValueCollection parameters, out Issue? foundIssue)
        {
            foundIssue = null;
            try
            {
                foundIssue = _manager.GetObjects<Issue>(parameters)?.FirstOrDefault();
                if (foundIssue is not null)
                    return true;
            }
            catch
            {
                // silent failure, found issue is null
            }
            return false;
        }
        #endregion

        #region Get activity for appointment

        private bool TryGetActivity(AppointmentItem appointmentItem, out Activity? foundActivity)
        {
            foundActivity = _timeEntryActivities.FirstOrDefault(_A => _A.Name == appointmentItem.Categories);
            if (foundActivity is not null)
                return true;

            CategoryToActivityMapping? customMapping = _categoryToActivityMappings.FirstOrDefault(_C => _C.Category == appointmentItem.Categories);
            if (customMapping is not null)
            {
                foundActivity = _timeEntryActivities.FirstOrDefault(_A => _A.Id == customMapping.ActivityId);
                if (foundActivity is not null)
                    return true;
            }
            return false;
        }

        #endregion
    }
}
