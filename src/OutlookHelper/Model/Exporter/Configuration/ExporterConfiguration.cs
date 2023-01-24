namespace OutlookHelper
{
    internal class ExporterConfiguration
    {
        public string ServerUrl { get; set; }
        public string ApiKey { get; set; }
        public string TargetUserId { get; set; }
        public List<Activity> TimeEntryActivities { get; set; }
        public List<CategoryToActivityMapping> CustomCategoryToActivityMappings { get; set; }
    }
}
