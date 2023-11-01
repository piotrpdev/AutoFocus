using CliFx;
using CliFx.Attributes;
using CliFx.Exceptions;
using CliFx.Infrastructure;
using FlaUI.Core.AutomationElements;
using FlaUI.Core.Definitions;
using FlaUI.Core.Input;
using FlaUI.UIA3;
using Microsoft.Extensions.Logging;
using System;
using static System.Windows.Forms.VisualStyles.VisualStyleElement;

namespace AutoFocus;

[Command(Description = "Sets the Sample Rate and Buffer Size.")]
// ReSharper disable once UnusedMember.Global
public class SetSampleAndBufferCommand : ICommand
{
    // ? https://support.focusrite.com/hc/en-gb/articles/360013759200-Fixing-Audio-issues-in-Windows-VoIP-apps-e-g-Discord-Zoom-Teamspeak-
    [CommandOption("sampleRate", 's', Description = "Sample Rate to set.")]
    public string SampleRate { get; init; } = "48000";

    [CommandOption("bufferSize", 'b', Description = "Buffer Size to set.")]
    public string BufferSize { get; init; } = "128";

    [CommandOption("notifierPath", 'n', Description = "Absolute path to 'Focusrite Notifier.exe'")]
    public string NotifierPath { get; init; } = "C:\\Program Files\\Focusrite\\Drivers\\Focusrite Notifier.exe";

    [CommandOption("notifierArgs", 'a', Description = "Arguments to pass when launching 'Focusrite Notifier.exe'")]
    public string NotifierArgs { get; init; } = "40000";

    [CommandOption("fromTray", 'f', Description = "Try launching Focusrite Notifier using the tray icon.")]
    public bool FromTray { get; init; } = false;

    [CommandOption("checkTray", 't', Description = "Check for Focusrite Notifier icon in non-hidden tray icons?")]
    public bool CheckTrayIcons { get; init; } = false;

    [CommandOption("checkHiddenTray", 'x', Description = "Check for Focusrite Notifier icon in hidden tray icons?")]
    public bool CheckHiddenTrayIcons { get; init; } = true;

    [CommandOption("waitAfterSample", 'w', Description = "How long to wait after changing Sample Rate (Notifier often freezes).")]
    public int WaitAfterSample { get; init; } = 500;

    private readonly UIA3Automation _automation = new();

    private readonly ILogger<SetSampleAndBufferCommand> _logger;

    private readonly string _logSeparator = new('-', 50);

    private AutomationElement GetDesktop()
    {
        _logger.LogInformation("Getting Desktop...");
        var desktop = _automation.GetDesktop();
        _ = desktop ?? throw new CommandException("Couldn't find Desktop, do you have one?");
        _logger.LogDebug("{desktop}", desktop);

        return desktop;
    }

    private AutomationElement GetTaskbar(AutomationElement desktop)
    {
        _logger.LogInformation("Getting Taskbar...");
        var taskbar = desktop.FindFirstChild(tf => tf.ByClassName("Shell_TrayWnd").And(tf.ByName("Taskbar")));
        _ = taskbar ?? throw new CommandException("Couldn't find Taskbar, do you have one?");
        _logger.LogDebug("{taskbar}", taskbar);

        return taskbar;
    }

    private void GetHiddenTrayIcons(AutomationElement desktop, AutomationElement taskbar, ref List<AutomationElement> arrayWithIcons)
    {
        _logger.LogInformation("Getting SystemTrayIcon...");
        var systemTrayIcon = taskbar.FindFirstDescendant(cf => cf.ByAutomationId("SystemTrayIcon"));
        if (systemTrayIcon != null)
        {
            _logger.LogDebug("{systemTrayIcon}", systemTrayIcon);

            _logger.LogInformation("Found SystemTrayIcon, invoking...");
            systemTrayIcon.AsButton().Invoke();

            _logger.LogInformation("Getting tray overflow window...");
            Wait.UntilInputIsProcessed();
            var trayOverflow = desktop.FindFirstChild(tf => tf.ByClassName("TopLevelWindowForOverflowXamlIsland").And(tf.ByName("System tray overflow window.")));
            _ = trayOverflow ?? throw new CommandException("Couldn't find tray overflow window, something must have went wrong.");
            _logger.LogDebug("{trayOverflow}", trayOverflow);

            _logger.LogInformation("Getting hidden icons...");
            var hiddenIcons = trayOverflow.FindAllDescendants(tf => tf.ByAutomationId("NotifyItemIcon"));
            _ = hiddenIcons ?? throw new CommandException("Couldn't find hidden icons, something must have went wrong.");
            _logger.LogDebug("{hiddenIcons}", hiddenIcons.Length);

            _logger.LogInformation("Found hidden icons, adding to list...");
            arrayWithIcons.AddRange(hiddenIcons);
        }
        else
        {
            // No error in case the user wants to check hidden and non-hidden icons, and this is normal.
            _logger.LogWarning("Couldn't find the hidden icons chevron.");
        }
    }

    private void GetTrayIcons(AutomationElement taskbar, ref List<AutomationElement> arrayWithIcons)
    {
        _logger.LogInformation("Getting tray icons...");
        var trayIcons = taskbar.FindAllDescendants(cf => cf.ByAutomationId("NotifyItemIcon"));

        if (trayIcons != null && trayIcons.Length != 0)
        {
            _logger.LogDebug("{trayIcons}", trayIcons.Length);

            _logger.LogInformation("Found non-hidden tray icons, adding to list...");
            arrayWithIcons.AddRange(trayIcons);
        }
        else
        {
            _logger.LogWarning("Couldn't find any non-hidden tray icons.");
        }
    }

    private AutomationElement GetNotifierIcon(ref List<AutomationElement> arrayWithIcons)
    {
        _logger.LogInformation("Getting the Focusrite Notifier icon...");
        var notifierIcon = arrayWithIcons.Find(e =>
        {
            try { return e.Name.Contains("Focusrite"); } catch { return false; }
        });
        _ = notifierIcon ?? throw new CommandException("Couldn't find Focusrite Notifier icon, is it running? Is it not in the tray?");
        _logger.LogDebug("{notifierIcon}", notifierIcon);

        return notifierIcon;
    }

    private AutomationElement GetContextMenu(AutomationElement desktop)
    {
        _logger.LogInformation("Getting the context menu...");
        var contextMenu = desktop.FindFirstChild(tf => tf.ByControlType(ControlType.Menu).And(tf.ByName("Context")));
        _ = contextMenu ?? throw new CommandException("Couldn't find context menu, something must have went wrong. Maybe adjust wait time?");
        _logger.LogDebug("{contextMenu}", contextMenu);

        return contextMenu;
    }

    private AutomationElement GetSettingsItem(AutomationElement contextMenu)
    {
        _logger.LogInformation("Selecting settings menu item...");
        var settingsItem = contextMenu.FindAllChildren().ToList().Find(e =>
        {
            try { return e.Name.Contains("settings"); } catch { return false; }
        });
        _ = settingsItem ?? throw new CommandException("Couldn't find settings item, something must have went wrong.");
        _logger.LogDebug("{settingsItem}", settingsItem);

        return settingsItem;
    }

    private void LaunchFromTray(AutomationElement desktop)
    {
        _logger.LogInformation("Trying to launch Focusrite Notifier from tray...");
        var taskbar = GetTaskbar(desktop);

        // ! List is modified by the tray functions.
        var arrayWithIcons = new List<AutomationElement>();
        if (CheckHiddenTrayIcons) GetHiddenTrayIcons(desktop, taskbar, ref arrayWithIcons);
        if (CheckTrayIcons) GetTrayIcons(taskbar, ref arrayWithIcons);
        _logger.LogDebug("{arrayWithIcons}", arrayWithIcons.Count);

        var notifierIcon = GetNotifierIcon(ref arrayWithIcons);

        _logger.LogInformation("Found the Focusrite Notifier icon, invoking...");
        notifierIcon.AsButton().Invoke();
        Wait.UntilInputIsProcessed();

        var contextMenu = GetContextMenu(desktop);
        var settingsItem = GetSettingsItem(contextMenu);

        _logger.LogInformation("Found settings menu item, invoking...");
        settingsItem.AsButton().Invoke();
    }

    private AutomationElement GetSettingsWindow(AutomationElement desktop)
    {
        _logger.LogInformation("Getting the Device Settings window...");
        var settingsWindow = desktop.FindAllChildren(cf => cf.ByControlType(ControlType.Window)).ToList().Find(e =>
        {
            try { return e.Name.Contains("Device Settings"); } catch { return false; }
        });
        _ = settingsWindow ?? throw new CommandException("Couldn't find the Device Settings window, something must have went wrong." +
                                                         " Did the window open? Does it have 'Device Settings' in the title?");
        _logger.LogDebug("{settingsWindow}", settingsWindow);

        return settingsWindow;
    }

    private AutomationElement GetSampleBox(AutomationElement settingsWindow)
    {
        _logger.LogInformation("Getting the Sample Rate combo box...");
        var sampleBox = settingsWindow.FindFirstChild(tf => tf.ByControlType(ControlType.ComboBox).And(tf.ByName("Sample Rate")));
        _ = sampleBox ?? throw new CommandException("Couldn't find the Sample Rate combo box, something must have went wrong." +
                                                    " Are you using English?");
        _logger.LogDebug("{sampleBox}", sampleBox);

        return sampleBox;
    }

    private bool SetSampleRate(AutomationElement sampleBox, string sampleRate)
    {
        _logger.LogInformation("Trying to set the Sample Rate to '{sampleRate}'...", sampleRate);
        if (sampleBox.AsComboBox().Value == sampleRate)
        {
            _logger.LogInformation("Sample Rate already set to '{sampleRate}', skipping...", sampleRate);

            return false;
        }
        // ? https://stackoverflow.com/a/42654772
        sampleBox.AsComboBox().Expand();
        Wait.UntilInputIsProcessed();

        var newSampleRateItem = sampleBox.AsComboBox().Items.FirstOrDefault(it => it.Text.Equals(sampleRate));
        _ = newSampleRateItem ?? throw new CommandException($"Couldn't find the '{sampleRate}' Sample Rate option in the combo box," +
                                                            " something must have went wrong. Is that a valid option?");
        _logger.LogDebug("{newSampleRateItem}", newSampleRateItem);

        newSampleRateItem.Click();

        return true;
    }

    private AutomationElement GetBufferBox(AutomationElement settingsWindow)
    {
        _logger.LogInformation("Getting the Buffer Size combo box...");
        var bufferBox = settingsWindow.FindFirstChild(tf => tf.ByControlType(ControlType.ComboBox).And(tf.ByName("Buffer Size")));
        _ = bufferBox ?? throw new CommandException("Couldn't find the Buffer Size combo box, something must have went wrong." +
                                                    " Are you using English?");
        _logger.LogDebug("{bufferBox}", bufferBox);

        return bufferBox;
    }

    // ReSharper disable once UnusedMethodReturnValue.Local
    private bool SetBufferSize(AutomationElement bufferBox, string bufferSize)
    {
        _logger.LogInformation("Trying to set the Buffer Size to '{bufferSize}'...", bufferSize);
        if (bufferBox.AsComboBox().Value == bufferSize)
        {
            _logger.LogInformation("Buffer Size already set to '{bufferSize}', skipping...", bufferSize);

            return false;
        }

        // ? https://stackoverflow.com/a/42654772
        bufferBox.AsComboBox().Expand();
        Wait.UntilInputIsProcessed();

        var newBufferSizeItem = bufferBox.AsComboBox().Items.FirstOrDefault(it => it.Text.Equals(bufferSize));
        _ = newBufferSizeItem ?? throw new CommandException($"Couldn't find the '{bufferSize}' Buffer Size option in the combo box," + 
                                                            " something must have went wrong. Is that a valid option?");
        _logger.LogDebug("{newBufferSizeItem}", newBufferSizeItem);

        newBufferSizeItem.Click();

        return true;
    }

    private AutomationElement GetCloseButton(AutomationElement settingsWindow)
    {
        _logger.LogInformation("Getting the title bar...");
        var titleBar = settingsWindow.FindFirstChild(tf => tf.ByControlType(ControlType.TitleBar));
        _ = titleBar ??
            throw new CommandException("Couldn't find the title bar, something must have went wrong.");
        _logger.LogDebug("{titleBar}", titleBar);

        _logger.LogInformation("Getting the close button...");
        var closeButton = titleBar.FindFirstChild(tf => tf.ByControlType(ControlType.Button).And(tf.ByName("Close")));
        _ = closeButton ??
            throw new CommandException("Couldn't find the close button, something must have went wrong. Good luck.");
        _logger.LogDebug("{closeButton}", closeButton);

        return closeButton;
    }

    public SetSampleAndBufferCommand(ILogger<SetSampleAndBufferCommand> logger)
    {
        _ = logger ?? throw new ArgumentNullException(nameof(logger));
        _logger = logger;
    }

    public ValueTask ExecuteAsync(IConsole console)
    {
        try
        {
            _logger.LogInformation("{logSeparator}", _logSeparator);
            _logger.LogInformation("Starting 'SetSampleAndBufferCommand'");

            var desktop = GetDesktop();

            if (!FromTray)
            {
                _logger.LogInformation("Trying to launch Focusrite Notifier using exe...");
                // ! ((Application) app).GetMainWindow() always fails here for some reason.
                var app = FlaUI.Core.Application.Launch(NotifierPath, NotifierArgs);
                Wait.UntilInputIsProcessed(TimeSpan.FromMilliseconds(500));
                _logger.LogDebug("{app}", app);
            }
            else
            {
                LaunchFromTray(desktop);
                Wait.UntilInputIsProcessed(TimeSpan.FromMilliseconds(500));
            }

            var settingsWindow = GetSettingsWindow(desktop);

            var sampleBox = GetSampleBox(settingsWindow);

            var isSampleSkipped = !SetSampleRate(sampleBox, SampleRate);
            Wait.UntilInputIsProcessed();

            if (!isSampleSkipped)
            {
                // ? The UI thread for Focusrite Notifier seems to hang sometimes depending on the sample rate chosen,
                // ? I wasn't able to find a way to actually determine when it comes back to life :/
                var waitTime = TimeSpan.FromMilliseconds(WaitAfterSample);
                var formattedTime = waitTime.ToString("ss");
                _logger.LogInformation("Waiting {waitTime} seconds until Notifier is responsive again...", formattedTime);
                Wait.UntilInputIsProcessed(waitTime);
            }

            var bufferBox = GetBufferBox(settingsWindow);

            SetBufferSize(bufferBox, BufferSize);
            Wait.UntilInputIsProcessed();

            var closeButton = GetCloseButton(settingsWindow);

            _logger.LogInformation("Found close button, invoking...");
            closeButton.AsButton().Invoke();
            Wait.UntilInputIsProcessed();

            _logger.LogInformation("Finished");
            _logger.LogInformation("{logSeparator}", _logSeparator);
        }
        catch (Exception ex)
        {
            // ? This is not necessary, I just like using Serilog
            _logger.LogError(ex, "Command threw an error: ");
            throw new CommandException(".");
        }

        return default;
    }
}