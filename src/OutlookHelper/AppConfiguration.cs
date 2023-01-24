namespace OutlookHelper
{
    internal class AppConfiguration
    {
        public ExploratorConfiguration ExploratorConfiguration { get; } = new();

        public ExporterConfiguration ExporterConfiguration { get; } = new();
    }
}
