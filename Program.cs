using CliFx;
using Microsoft.Extensions.DependencyInjection;
using Serilog;

namespace AutoFocus
{
    internal class Program
    {
        private static void SetupLogger()
        {
#if DEBUG
            string logFilePath = "log.txt";
#else
    string appDataPath = Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData);
    string appFolder = Path.Combine(appDataPath, "AutoFocus");
    string logFolder = Path.Combine(appFolder, "Logs");
    if (!Directory.Exists(logFolder))
    {
        Directory.CreateDirectory(logFolder);
    }
    string logFilePath = Path.Combine(logFolder, "log.txt");
#endif

            // Configure Serilog
            Log.Logger = new LoggerConfiguration()
                .MinimumLevel.Information()
                .Enrich.FromLogContext()
                .WriteTo.Async(sink => sink.File(logFilePath, rollingInterval: RollingInterval.Day, rollOnFileSizeLimit: true))
                .WriteTo.Console()
                .CreateLogger();
        }

        public static async Task<int> Main()
        {
            int returnCode = 0;

            SetupLogger();

            try
            {
                Log.Information(new string('-', 50));
                Log.Information("Application is starting up");

                returnCode = await new CliApplicationBuilder()
                    .SetTitle("AutoFocus")
                    .SetDescription("Automate changing Sample Rate and Buffer Size in Focusrite Notifier device settings.")
                    .SetExecutableName("autofocus.exe")
                    .AddCommandsFromThisAssembly()
                    .UseTypeActivator(commandTypes =>
                    {
                        var services = new ServiceCollection();
                        services.AddLogging(builder => builder.AddSerilog(dispose: true));

                        // Register all commands as transient services
                        foreach (var commandType in commandTypes)
                            services.AddTransient(commandType);

                        return services.BuildServiceProvider();
                    })
                    .AllowDebugMode()
                    .Build()
                    .RunAsync();
            }
            catch (Exception ex)
            {
                returnCode = 1;
                Log.Error(ex, "Something went wrong");
            }
            finally
            {
                Log.Information("Application is stopping, goodbye.");
                await Log.CloseAndFlushAsync();
            }

            return returnCode;
        }
    }
}
