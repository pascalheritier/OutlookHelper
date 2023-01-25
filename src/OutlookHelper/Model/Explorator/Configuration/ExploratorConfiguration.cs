namespace OutlookHelper
{
    internal class ExploratorConfiguration
    {
        public double WorkingPercentage { get; set; }
        public List<string> ExcludedSubjects { get; set; } = null!;
        public List<string> ExcludedCategories { get; set; } = null!;
        public List<YearRange> WeekRangePerYear { get; set; } = null!;
    }
}
