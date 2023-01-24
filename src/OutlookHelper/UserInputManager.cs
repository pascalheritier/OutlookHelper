using Microsoft.Extensions.Logging;

namespace OutlookHelper
{
    internal class UserInputManager
    {
        #region Members

        private OutlookCalendarExplorator _outlookCalendarExplorator;
        private RedmineCalendarExporter _redmineCalendarExporter;
        private ILogger _logger;

        #endregion

        #region Constructor

        public UserInputManager(AppConfiguration appConfiguration, ILoggerFactory loggerFactory)
        {
            _logger = loggerFactory.CreateLogger<UserInputManager>();

            // create outlook calendar
            _outlookCalendarExplorator = new OutlookCalendarExplorator(
                appConfiguration.ExploratorConfiguration.WeekRangePerYear,
                appConfiguration.ExploratorConfiguration.ExcludedCategories,
                appConfiguration.ExploratorConfiguration.ExcludedSubjects,
                appConfiguration.ExploratorConfiguration.WorkingPercentage,
                loggerFactory.CreateLogger<OutlookCalendarExplorator>());

            // create calendar data exporter
            _redmineCalendarExporter = new RedmineCalendarExporter(
                appConfiguration.ExporterConfiguration.ServerUrl,
                appConfiguration.ExporterConfiguration.ApiKey,
                appConfiguration.ExporterConfiguration.TargetUserId,
                appConfiguration.ExporterConfiguration.TimeEntryActivities,
                appConfiguration.ExporterConfiguration.CustomCategoryToActivityMappings,
                loggerFactory.CreateLogger<RedmineCalendarExporter>());
        }

        #endregion

        #region Run
        public void Run()
        {
            try
            {
                // explore outlook calendar
                SortedCalendar? sortedCalendar = _outlookCalendarExplorator.ExploreCalendar();
                if (sortedCalendar == null)
                {
                    _logger.LogError("Could not find an outlook calendar to explore!");
                    return;
                }

                bool exitRequested = false;
                do
                {
                    exitRequested = GetUserRequestedAction(sortedCalendar);
                }
                while (!exitRequested);
            }
            catch(Exception e)
            {
                _logger.LogCritical($"App critical failure: {e.Message}");
            }
        }

        #endregion

        #region Get user inputs
        private bool GetUserRequestedAction(SortedCalendar sortedCalendar)
        {
            Console.WriteLine();
            Console.WriteLine("-- Please select the desired action: --");
            Console.WriteLine("1. Display outlook calendar summaries for ALL YEARS");
            Console.WriteLine("2. Display outlook calendar summaries for ALL WEEKS");
            Console.WriteLine("3. Display outlook calendar summaries for a SPECIFIC WEEK");
            Console.WriteLine("4. Display outlook calendar summaries for ALL DAYS");
            Console.WriteLine("5. Display outlook calendar summary for a SPECIFIC DAY");
            Console.WriteLine("6. Export a particular day of the outlook calendar");
            Console.WriteLine("7. Export a particular week of the outlook calendar");
            Console.WriteLine("0. Exit");
            Console.WriteLine();
            Console.Write("> ");

            string? input = Console.ReadLine();
            if (uint.TryParse(input, out uint selection))
            {
                switch (selection)
                {
                    case 0:
                        return true;
                    case 1:
                        _outlookCalendarExplorator.DisplayCalendarAllYearsSummary(sortedCalendar);
                        break;
                    case 2:
                        _outlookCalendarExplorator.DisplayCalendarAllWeeksSummary(sortedCalendar);
                        break;
                    case 3:
                        DisplayCalendarDataForWeek(sortedCalendar);
                        break;
                    case 4:
                        _outlookCalendarExplorator.DisplayCalendarAllDaysSummary(sortedCalendar);
                        break;
                    case 5:
                        DisplayCalendarDataForDay(sortedCalendar);
                        break;
                    case 6:
                        ExportCalendarDataForDay(sortedCalendar);
                        break;
                    case 7:
                        ExportCalendarDataForWeek(sortedCalendar);
                        break;
                    default:
                        _logger.LogError("Desired action is not available.");
                        break;
                }
            }
            else
            {
                _logger.LogError("Invalid input: please use a valid integer number.");
            }
            return false;
        }

        #endregion

        #region Display

        private void DisplayCalendarDataForDay(SortedCalendar sortedCalendar)
        {
            Console.WriteLine();
            Console.WriteLine("-- Please enter the day (format 'dd.MM.yyyy') to be displayed: --");
            Console.WriteLine();
            Console.Write("> ");

            string? input = Console.ReadLine();
            if (DateTime.TryParse(input, out DateTime dayToDisplay))
            {
                _outlookCalendarExplorator.DisplayCalendarDailySummary(sortedCalendar, dayToDisplay);
            }
            else
            {
                _logger.LogError($"Invalid input '{input}', please use a valid input format 'dd.MM.yyyy'.");
            }
        }

        private void DisplayCalendarDataForWeek(SortedCalendar sortedCalendar)
        {
            Console.WriteLine();
            Console.WriteLine("-- Please enter the year (integer number) of the week to be displayed: --");
            Console.WriteLine();
            Console.Write("> ");

            string? input = Console.ReadLine();
            if (int.TryParse(input, out int year))
            {
                Console.WriteLine();
                Console.WriteLine("-- Please enter the week (integer number) to be displayed: --");
                Console.WriteLine();
                Console.Write("> ");

                input = Console.ReadLine();
                if (int.TryParse(input, out int weekToDisplay))
                {
                    _outlookCalendarExplorator.DisplayCalendarWeeklySummary(sortedCalendar, year, weekToDisplay);
                }
                else
                {
                    _logger.LogError($"Invalid input '{input}', please use a valid input (integer number).");
                }
            }
            else
            {
                _logger.LogError($"Invalid input '{input}', please use a valid input (integer number).");
            }
        }

        #endregion

        #region Export
        private void ExportCalendarDataForDay(SortedCalendar sortedCalendar)
        {
            Console.WriteLine();
            Console.WriteLine("-- Please enter the day (format 'dd.MM.yyyy') to be exported: --");
            Console.WriteLine();
            Console.Write("> ");

            string? input = Console.ReadLine();
            if (DateTime.TryParse(input, out DateTime dayToExport))
            {
                _redmineCalendarExporter.ExportCalendarData(sortedCalendar, dayToExport);
            }
            else
            {
                _logger.LogError($"Invalid input '{input}', please use a valid input format 'dd.MM.yyyy'.");
            }
        }

        private void ExportCalendarDataForWeek(SortedCalendar sortedCalendar)
        {
            Console.WriteLine();
            Console.WriteLine("-- Please enter the year (integer number) of the week to be exported: --");
            Console.WriteLine();
            Console.Write("> ");

            string? input = Console.ReadLine();
            if (int.TryParse(input, out int year))
            {
                Console.WriteLine();
                Console.WriteLine("-- Please enter the week (integer number) to be exported: --");
                Console.WriteLine();
                Console.Write("> ");

                input = Console.ReadLine();
                if (int.TryParse(input, out int weekToExport))
                {
                    _redmineCalendarExporter.ExportCalendarData(sortedCalendar, year, weekToExport);
                }
                else
                {
                    _logger.LogError($"Invalid input '{input}', please use a valid input (integer number).");
                }
            }
            else
            {
                _logger.LogError($"Invalid input '{input}', please use a valid input (integer number).");
            }
        }

        #endregion
    }
}
