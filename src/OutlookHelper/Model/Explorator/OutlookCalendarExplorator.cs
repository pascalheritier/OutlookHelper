using Microsoft.Extensions.Logging;
using Microsoft.Office.Interop.Outlook;
using System.Globalization;

namespace OutlookHelper
{
    // https://stackoverflow.com/questions/53737653/how-to-get-data-from-outlook-calendar
    // https://stackoverflow.com/questions/32399420/could-not-load-file-or-assembly-office-version-15-0-0-0
    internal class OutlookCalendarExplorator : CalendarExplorator
    {
        #region Statics

        private static Func<DateTime, int> WeekProjector = (d) => CultureInfo.CurrentCulture.Calendar.GetWeekOfYear(d, CalendarWeekRule.FirstFourDayWeek, DayOfWeek.Sunday);

        #endregion

        #region Members

        private IEnumerable<YearRange> _weekRangePerYear;
        private IEnumerable<string> _excludedCategories;
        private IEnumerable<string> _excludedSubjects;
        private double _workingPercentage;
        private ILogger _logger;

        #endregion

        #region Constructor

        public OutlookCalendarExplorator(
            IEnumerable<YearRange> weekRangePerYear,
            IEnumerable<string> excludedCategories,
            IEnumerable<string> excludedSubjects,
            double workingPercentage,
            ILogger logger)
        {
            _weekRangePerYear = weekRangePerYear;
            _excludedCategories = excludedCategories;
            _excludedSubjects = excludedSubjects;
            _workingPercentage = workingPercentage;
            _logger = logger;
        }

        #endregion

        #region Explore calendar

        public SortedCalendar? ExploreCalendar()
        {
            // get outlook calendar
            var outlookApplication = new Application();
            NameSpace mapiNamespace = outlookApplication.GetNamespace("MAPI");
            MAPIFolder calendar = mapiNamespace.GetDefaultFolder(OlDefaultFolders.olFolderCalendar);
            if (calendar != null)
            {
                SortedCalendar sortedCalendar = new();
                // get list of all calendar events
                List<AppointmentItem> calendarItems = new();
                for (int i = 1; i <= calendar.Items.Count; i++)
                {
                    if (calendar.Items[i] is not AppointmentItem)
                        continue;
                    calendarItems.Add((AppointmentItem)calendar.Items[i]);
                }

                // group calendar items by year, while excluding unwanted categories (or items with no category) or unwanted subjects
                var yearlySortedCalendarItems = calendarItems
                    .Where(_CI => !string.IsNullOrEmpty(_CI.Categories))
                    .Where(_CI => !_excludedSubjects.Any(_S => _S == _CI.Subject))
                    .Where(_CI => !_excludedCategories.Any(_S => _S == _CI.Categories))
                    .GroupBy(_CI => _CI.Start.Year);
                foreach (var yearlyGroupCalendarItems in yearlySortedCalendarItems)
                {
                    int year = yearlyGroupCalendarItems.Key;
                    YearlyCalendar yearlyCalendar = new YearlyCalendar { Year = year };
                    sortedCalendar.YearlyCalendars.Add(yearlyCalendar);

                    // group calendar items by week
                    var weeklyGroupCalendarItems = yearlyGroupCalendarItems.GroupBy(_C => WeekProjector(_C.Start));

                    // skip items that are not in desired range
                    YearRange? yearRange = _weekRangePerYear.FirstOrDefault(_W => _W.Year == year);
                    if (yearRange is null)
                        continue; // skip year with no range
                    if (yearRange.WeekRange.Count() != 2)
                        throw new NotSupportedException("Year range must define a starting and an ending week!");
                    weeklyGroupCalendarItems = weeklyGroupCalendarItems.Where(_WeeklyGroup => _WeeklyGroup.Key >= yearRange.WeekRange[0] && _WeeklyGroup.Key <= yearRange.WeekRange[1]);

                    // compute yearly data
                    yearlyCalendar.YearlyTimeSpent = TimeSpan.FromMinutes(weeklyGroupCalendarItems.SelectMany(_WeeklyGroup => _WeeklyGroup.Select(_CI => _CI.Duration)).Sum());

                    // loop calendar for each week of current year
                    foreach (var weeklyCalendarItems in weeklyGroupCalendarItems)
                    {
                        int week = weeklyCalendarItems.Key;
                        WeeklyCalendar weeklyCalendar = new WeeklyCalendar { Week = week };
                        yearlyCalendar.WeeklyCalendars.Add(weeklyCalendar);

                        // compute weekly data
                        weeklyCalendar.WeeklyTimeSpent = TimeSpan.FromMinutes(weeklyCalendarItems.Select(_Item => _Item.Duration).Sum());

                        // loop calendar for each calendar item of current week
                        foreach (AppointmentItem sortedCalendarItem in weeklyCalendarItems)
                            weeklyCalendar.Appointments.Add(sortedCalendarItem);
                    }
                }
                return sortedCalendar;
            }
            return null;
        }

        #endregion

        #region Display calendar

        public override void DisplayCalendarAllYearsSummary(SortedCalendar sortedCalendar)
        {
            sortedCalendar.YearlyCalendars.ForEach(_YearlyCalendar => DisplayYearInformation(_YearlyCalendar, _workingPercentage));
        }

        public override void DisplayCalendarAllWeeksSummary(SortedCalendar sortedCalendar)
        {
            foreach (YearlyCalendar yearlyCalendar in sortedCalendar.YearlyCalendars)
            {
                Console.WriteLine($"{Environment.NewLine}--------------------------");
                Console.WriteLine($"YEAR {yearlyCalendar.Year}");
                Console.WriteLine($"--------------------------");

                yearlyCalendar.WeeklyCalendars.ForEach(_WeeklyCalendar => DisplayWeekInformation(_WeeklyCalendar, _workingPercentage));
            }
        }

        public override void DisplayCalendarWeeklySummary(SortedCalendar sortedCalendar, int year, int week)
        {
            Console.WriteLine($"{Environment.NewLine}--------------------------");
            Console.WriteLine($"YEAR {year}");
            Console.WriteLine($"--------------------------");

            YearlyCalendar? yearlyCalendar = sortedCalendar.YearlyCalendars.FirstOrDefault(_YearlyCalendar => _YearlyCalendar.Year == year);
            if (yearlyCalendar == null)
            {
                _logger.LogError($"No year {year} found in calendar.");
                return;
            }

            WeeklyCalendar? weeklyCalendar = yearlyCalendar.WeeklyCalendars.FirstOrDefault(_WeeklyCalendar => _WeeklyCalendar.Week == week);
            if (weeklyCalendar == null)
            {
                _logger.LogError($"No week {week} found in calendar.");
                return;
            }

            this.DisplayWeekInformation(weeklyCalendar, _workingPercentage);
        }

        public override void DisplayCalendarAllDaysSummary(SortedCalendar sortedCalendar)
        {
            foreach (YearlyCalendar yearlyCalendar in sortedCalendar.YearlyCalendars)
            {
                Console.WriteLine($"{Environment.NewLine}--------------------------");
                Console.WriteLine($"YEAR {yearlyCalendar.Year}");
                Console.WriteLine($"--------------------------");

                foreach (WeeklyCalendar weeklyCalendar in yearlyCalendar.WeeklyCalendars)
                {
                    Console.WriteLine($"{Environment.NewLine}--------------------------");
                    Console.WriteLine($"WEEK {weeklyCalendar.Week}");
                    Console.WriteLine($"--------------------------");

                    weeklyCalendar.Appointments.ForEach(_Appointment => DisplayCalendarAppointmentInformation(_Appointment));
                }
            }
        }

        public override void DisplayCalendarDailySummary(SortedCalendar sortedCalendar, DateTime dayToDisplay)
        {
            YearlyCalendar? yearlyCalendar = sortedCalendar.YearlyCalendars.FirstOrDefault(_Y => _Y.Year == dayToDisplay.Year);
            if (yearlyCalendar == null)
            {
                _logger.LogError($"Could not find any appointment for day {dayToDisplay.ToString(Utils.DisplayDateFormat)}");
                return;
            }
            List<AppointmentItem> dailyAppointments = yearlyCalendar.YearlyAppointments.Where(_A => _A.Start.Date == dayToDisplay.Date).ToList();
            TimeSpan dailyTotal = TimeSpan.FromMinutes(dailyAppointments.Sum(_A => _A.Duration));
            double dailyOvertime = dailyTotal.TotalHours - Utils.DailyHoursBasis;

            Console.WriteLine($"{Environment.NewLine}--------------------------");
            Console.WriteLine($"DAY {dayToDisplay.ToString(Utils.DisplayDateFormat)}");
            Console.WriteLine($"Total daily hours: {dailyTotal.Hours}h");
            Console.WriteLine($"Daily overtime balance: {dailyOvertime}h");
            Console.WriteLine($"--------------------------");

            dailyAppointments.ForEach(_Appointment => DisplayCalendarAppointmentInformation(_Appointment));
        }

        #endregion

        #region Display helpers

        private void DisplayYearInformation(YearlyCalendar yearlyCalendar, double workingPercentage)
        {
            Console.WriteLine($"{Environment.NewLine}--------------------------");
            Console.WriteLine($"YEAR {yearlyCalendar.Year}");
            Console.WriteLine($"Total yearly hours: {yearlyCalendar.YearlyTimeSpent.TotalHours}h");
            Console.WriteLine($"Total overtime: {yearlyCalendar.ComputeOvertime(workingPercentage)}h");
            Console.WriteLine($"--------------------------");
        }

        private void DisplayWeekInformation(WeeklyCalendar weeklyCalendar, double workingPercentage)
        {
            Console.WriteLine($"{Environment.NewLine}--------------------------");
            Console.WriteLine($"WEEK {weeklyCalendar.Week}");
            Console.WriteLine($"Total weekly hours: {weeklyCalendar.WeeklyTimeSpent.TotalHours}h");
            Console.WriteLine($"Weekly overtime balance: {weeklyCalendar.ComputeOvertime(workingPercentage)}h");
            Console.WriteLine($"--------------------------");
        }

        private void DisplayCalendarAppointmentInformation(AppointmentItem calendarItem)
        {
            Console.WriteLine($"{Environment.NewLine}--------------------------");
            Console.WriteLine(
                $"Calendar item: {calendarItem.Subject}{Environment.NewLine}" +
                $"Date:{calendarItem.Start.ToString(Utils.DisplayDateFormat)} {calendarItem.Start.ToString(Utils.DisplayHourFormat)} - {calendarItem.End.ToString(Utils.DisplayHourFormat)}{Environment.NewLine}" +
                $"Duration: {TimeSpan.FromMinutes(calendarItem.Duration).TotalHours}h{Environment.NewLine}" +
                $"Category: {calendarItem.Categories}");
            Console.WriteLine($"--------------------------");
        }

        #endregion
    }
}