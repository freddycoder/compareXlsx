using Microsoft.Extensions.Configuration;
using Microsoft.Extensions.DependencyInjection;
using Serilog;

namespace compareXlsx;

public static class Setup {
    public static (IServiceProvider, ILogger, compareXlsxpOptions) GetAppServices(string[] args)
    {
        var configuration = new ConfigurationBuilder()
            .AddEnvironmentVariables()
            .AddCommandLine(args)
            .Build();

        var option = new compareXlsxpOptions();
        configuration.Bind(option);

        var logger = new LoggerConfiguration()
            .WriteTo.Console(outputTemplate: "{Message:lj}{NewLine}{Exception}")
//            .WriteTo.File("log.txt", outputTemplate: "{Timestamp:yyyy-MM-dd HH:mm:ss.fff zzz} [{Level:u3}] {Message:lj}{NewLine}{Exception}")
            .CreateLogger();

        logger.Debug("Starting comapreXlsx.");

        IServiceCollection serviceCollection = new ServiceCollection()
            .AddMemoryCache()
            .AddLogging(loggingBuilder =>
            {
                loggingBuilder.AddSerilog(logger);
            });

        IServiceProvider provider = serviceCollection.BuildServiceProvider();

        logger.Debug("ServiceProvider is ready.");

        return (provider, logger, option);
    }
}