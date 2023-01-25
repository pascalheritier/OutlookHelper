namespace OutlookHelper
{
    internal class ExporterConfiguration
    {
        public string ServerUrl { get; set; } = null!;
        public string ApiKey { get; set; } = null!;
        public string TargetUserId { get; set; } = null!;
        public List<Activity> TimeEntryActivities { get; set; } = null!;
        public List<CategoryToActivityMapping> CustomCategoryToActivityMappings { get; set; } = null!;
    }
}
