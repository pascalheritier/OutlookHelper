namespace OutlookHelper
{
    internal class ExploratorConfiguration
    {
        public double WorkingPercentage { get; set; }
        public List<string> ExcludedSubjects { get; set; }
        public List<string> ExcludedCategories { get; set; }
        public List<YearRange> WeekRangePerYear { get; set; }
    }
}
