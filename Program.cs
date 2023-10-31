using FlaUI.Core.AutomationElements;
using FlaUI.Core.Definitions;
using FlaUI.Core.Input;
using FlaUI.UIA3;
using Serilog;

static void setupLogger()
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
        .WriteTo.Console()
        .WriteTo.File(logFilePath, rollingInterval: RollingInterval.Day, rollOnFileSizeLimit: true)
        .CreateLogger();
}

static void doAutomation(bool checkHiddenTrayIcons, bool checkTrayIcons, string sampleRate, string bufferSize)
{
    using var automation = new UIA3Automation();

    Log.Information("Getting Desktop...");
    var desktop = automation.GetDesktop();
    _ = desktop ?? throw new Exception("Couldn't find Desktop, do you have one?");

    Log.Information("Getting Taskbar...");
    var taskbar = desktop.FindFirstChild(tf => tf.ByClassName("Shell_TrayWnd").And(tf.ByName("Taskbar")));
    _ = taskbar ?? throw new Exception("Couldn't find Taskbar, do you have one?");

    var arrayWithIcons = new List<AutomationElement>();

    if (checkHiddenTrayIcons)
    {
        Log.Information("Getting SystemTrayIcon...");
        var systemTrayIcon = taskbar.FindFirstDescendant(cf => cf.ByAutomationId("SystemTrayIcon"));
        if (systemTrayIcon != null)
        {
            Log.Information("Found SystemTrayIcon, invoking...");
            systemTrayIcon?.AsButton().Invoke();

            Log.Information("Getting tray overflow window...");
            Wait.UntilInputIsProcessed();
            var trayOverflow = desktop.FindFirstChild(tf => tf.ByClassName("TopLevelWindowForOverflowXamlIsland").And(tf.ByName("System tray overflow window.")));
            _ = trayOverflow ?? throw new Exception("Couldn't find tray overflow window, something must have went wrong.");

            Log.Information("Getting hidden icons...");
            var hiddenIcons = trayOverflow.FindAllDescendants(tf => tf.ByAutomationId("NotifyItemIcon"));
            _ = hiddenIcons ?? throw new Exception("Couldn't find hidden icons, something must have went wrong.");

            Log.Information("Found hidden icons, adding to list...");
            arrayWithIcons.AddRange(hiddenIcons);
        }
        else
        {
            Log.Warning("Couldn't find the hidden icons chevron.");
        }
    }

    if (checkTrayIcons)
    {
        Log.Information("Getting tray icons...");
        var trayIcons = taskbar.FindAllDescendants(cf => cf.ByAutomationId("NotifyItemIcon"));

        if (trayIcons != null && trayIcons.Length != 0)
        {
            Log.Information("Found non-hidden tray icons, adding to list...");
            arrayWithIcons.AddRange(trayIcons);
        }
        else
        {
            Log.Warning("Couldn't find any non-hidden tray icons.");
        }
    }

    Log.Information("Getting the Focusrite Notifier icon...");
    var notifierButton = arrayWithIcons.Find(e =>
    {
        try { return e.Name.Contains("Focusrite"); } catch { return false; }
    });
    _ = notifierButton ?? throw new Exception("Couldn't find Focusrite Notifier icon, is it running? Is it not in the tray?");

    Log.Information("Found the Focusrite Notifier icon, invoking...");
    notifierButton.AsButton().Invoke();

    Log.Information("Waiting for context menu...");
    Wait.UntilInputIsProcessed();
    var contextMenu = desktop.FindFirstChild(tf => tf.ByControlType(ControlType.Menu).And(tf.ByName("Context")));
    _ = contextMenu ?? throw new Exception("Couldn't find context menu, something must have went wrong. Maybe adjust wait time?");

    Log.Information("Selecting settings menu item...");
    var settingsItem = contextMenu.FindAllChildren().ToList().Find(e =>
    {
        try { return e.Name.Contains("settings"); } catch { return false; }
    });
    _ = settingsItem ?? throw new Exception("Couldn't find settings item, something must have went wrong.");

    Log.Information("Found settings menu item, invoking...");
    settingsItem.AsButton().Invoke();

    Log.Information("Getting the Device Settings window...");
    Wait.UntilInputIsProcessed(TimeSpan.FromMilliseconds(500));
    var settingsWindow = desktop.FindAllChildren(cf => cf.ByControlType(ControlType.Window)).ToList().Find(e =>
    {
        try { return e.Name.Contains("Device Settings"); } catch { return false; }
    });
    _ = settingsWindow ?? throw new Exception("Couldn't find the Device Settings window, something must have went wrong." +
        " Did the window open? Does it have 'Device Settings' in the title?");

    Log.Information("Getting the Sample Rate combo box...");
    var sampleBox = settingsWindow.FindFirstChild(tf => tf.ByControlType(ControlType.ComboBox).And(tf.ByName("Sample Rate")));
    _ = sampleBox ?? throw new Exception("Couldn't find the Sample Rate combo box, something must have went wrong." +
        " Are you using English?");

    Log.Information($"Trying to set the Sample Rate to '{sampleRate}'...");
    if (sampleBox.AsComboBox().Value == sampleRate)
    {
        Log.Information($"Sample Rate already set to '{sampleRate}', skipping...");
    }
    else
    {
        // https://stackoverflow.com/a/42654772
        sampleBox.AsComboBox().Expand();
        Wait.UntilInputIsProcessed();
        var newSampleRateItem = sampleBox.AsComboBox().Items.FirstOrDefault(it => it.Text.Equals(sampleRate));
        _ = newSampleRateItem ?? throw new Exception($"Couldn't find the '{sampleRate}' Sample Rate option in the combo box," +
            " something must have went wrong. Is that a valid option?");
        newSampleRateItem.Click();
        Wait.UntilInputIsProcessed();
    }

    Log.Information("Getting the Buffer Size combo box...");
    var bufferBox = settingsWindow.FindFirstChild(tf => tf.ByControlType(ControlType.ComboBox).And(tf.ByName("Buffer Size")));
    _ = bufferBox ?? throw new Exception("Couldn't find the Buffer Size combo box, something must have went wrong." +
        " Are you using English?");

    // The UI thread for Focusrite Notifier seems to hang sometimes depending on the sample rate chosen,
    // I wasn't able to find a way to actually determine when it comes back to life :/
    Wait.UntilInputIsProcessed(TimeSpan.FromMilliseconds(4000));

    Log.Information($"Trying to set the Buffer Size to '{bufferSize}'...");
    if (bufferBox.AsComboBox().Value == bufferSize)
    {
        Log.Information($"Buffer Size already set to '{bufferSize}', skipping...");
    } else
    {
        // https://stackoverflow.com/a/42654772
        bufferBox.AsComboBox().Expand();
        Wait.UntilInputIsProcessed();
        var newBufferSizeItem = bufferBox.AsComboBox().Items.FirstOrDefault(it => it.Text.Equals(bufferSize));
        _ = newBufferSizeItem ?? throw new Exception($"Couldn't find the '{bufferSize}' Buffer Size option in the combo box," +
            " something must have went wrong. Is that a valid option?");
        newBufferSizeItem.Click();
        Wait.UntilInputIsProcessed();
    }
}

try
{
    setupLogger();
    Log.Information(new string('-', 50));
    Log.Information("Application is starting up");

    doAutomation(checkHiddenTrayIcons: true, checkTrayIcons: true, sampleRate: "48000", bufferSize: "128");
}
catch (Exception ex)
{
    Log.Error(ex, "Something went wrong");
}
finally
{
    Log.Information("Application is stopping, goodbye.");
    await Log.CloseAndFlushAsync();
}
