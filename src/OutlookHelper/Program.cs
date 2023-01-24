using Microsoft.Extensions.Configuration;
using Microsoft.Extensions.DependencyInjection;
using Microsoft.Extensions.Logging;
using NLog;
using NLog.Extensions.Logging;

namespace OutlookHelper;

public class Program
{
    #region App Config

    private static string AppSettingsFileName = "appsettings.json";

    #endregion

    #region Main

    static void Main(string[] args)
    {
        try
        {
            IServiceCollection services = new ServiceCollection();
            ConfigureServices(services);
            IServiceProvider serviceProvider = services.BuildServiceProvider();
            var runner = serviceProvider.GetRequiredService<UserInputManager>();
            runner.Run();
        }
        catch (Exception ex)
        {
            LogManager.GetCurrentClassLogger().Log(NLog.LogLevel.Fatal, $"Critical app failure: {ex.Message}{Environment.NewLine}{ex.StackTrace}");
        }
    }

    #endregion

    #region Services
    private static void ConfigureServices(IServiceCollection services)
    {
        https://github.com/NLog/NLog/wiki/Getting-started-with-.NET-Core-2---Console-application

        IConfiguration config = new ConfigurationBuilder()
        .SetBasePath(System.IO.Directory.GetCurrentDirectory()) //From NuGet Package Microsoft.Extensions.Configuration.Json
        .AddJsonFile(AppSettingsFileName, optional: true, reloadOnChange: true)
        .Build();

        services.AddSingleton<AppConfiguration>(_X => GetAppConfiguration(config));
        services.AddLogging(loggingBuilder =>
        {
            // configure Logging with NLog
            loggingBuilder.ClearProviders();
            loggingBuilder.SetMinimumLevel(Microsoft.Extensions.Logging.LogLevel.Trace);
            loggingBuilder.AddNLog(config);
        });
        services.AddTransient<UserInputManager>();
    } 

    #endregion

    #region Application configuration

    private static AppConfiguration GetAppConfiguration(IConfiguration configuration)
    {
        AppConfiguration appConfiguration = new();
        configuration.Bind(appConfiguration);
        return appConfiguration;
    }

    #endregion
}