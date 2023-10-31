using CliFx;
using CliFx.Attributes;
using CliFx.Exceptions;
using CliFx.Infrastructure;
using FlaUI.Core.AutomationElements;
using FlaUI.Core.Definitions;
using FlaUI.Core.Input;
using FlaUI.UIA3;
using Microsoft.Extensions.Logging;

namespace AutoFocus
{
    [Command(Description = "Sets the Sample Rate and Buffer Size.")]
    public partial class SetSampleAndBufferCommand : ICommand
    {
        // ? https://support.focusrite.com/hc/en-gb/articles/360013759200-Fixing-Audio-issues-in-Windows-VoIP-apps-e-g-Discord-Zoom-Teamspeak-
        [CommandOption("sampleRate", 's', Description = "Sample Rate to set.")]
        public string SampleRate { get; init; } = "48000";

        [CommandOption("bufferSize", 'b', Description = "Buffer Size to set.")]
        public string BufferSize { get; init; } = "128";

        [CommandOption("checkTray", 't', Description = "Check for Focusrite Notifier icon in non-hidden tray icons?")]
        public bool CheckTrayIcons { get; init; } = false;

        [CommandOption("checkHiddenTray", 'x', Description = "Check for Focusrite Notifier icon in hidden tray icons?")]
        public bool CheckHiddenTrayIcons { get; init; } = true;

        readonly UIA3Automation automation = new();

        private readonly ILogger<SetSampleAndBufferCommand> _logger;

        AutomationElement GetDesktop()
        {
            _logger.LogInformation("Getting Desktop...");
            var desktop = automation.GetDesktop();
            _ = desktop ?? throw new CommandException("Couldn't find Desktop, do you have one?");

            return desktop;
        }

        AutomationElement GetTaskbar(AutomationElement desktop)
        {
            _logger.LogInformation("Getting Taskbar...");
            var taskbar = desktop.FindFirstChild(tf => tf.ByClassName("Shell_TrayWnd").And(tf.ByName("Taskbar")));
            _ = taskbar ?? throw new CommandException("Couldn't find Taskbar, do you have one?");

            return taskbar;
        }

        void GetHiddenTrayIcons(AutomationElement desktop, AutomationElement taskbar, ref List<AutomationElement> arrayWithIcons)
        {
            _logger.LogInformation("Getting SystemTrayIcon...");
            var systemTrayIcon = taskbar.FindFirstDescendant(cf => cf.ByAutomationId("SystemTrayIcon"));
            if (systemTrayIcon != null)
            {
                _logger.LogInformation("Found SystemTrayIcon, invoking...");
                systemTrayIcon?.AsButton().Invoke();

                _logger.LogInformation("Getting tray overflow window...");
                Wait.UntilInputIsProcessed();
                var trayOverflow = desktop.FindFirstChild(tf => tf.ByClassName("TopLevelWindowForOverflowXamlIsland").And(tf.ByName("System tray overflow window.")));
                _ = trayOverflow ?? throw new CommandException("Couldn't find tray overflow window, something must have went wrong.");

                _logger.LogInformation("Getting hidden icons...");
                var hiddenIcons = trayOverflow.FindAllDescendants(tf => tf.ByAutomationId("NotifyItemIcon"));
                _ = hiddenIcons ?? throw new CommandException("Couldn't find hidden icons, something must have went wrong.");

                _logger.LogInformation("Found hidden icons, adding to list...");
                arrayWithIcons.AddRange(hiddenIcons);
            }
            else
            {
                // No error in case the user wants to check hidden and non-hidden icons, and this is normal.
                _logger.LogWarning("Couldn't find the hidden icons chevron.");
            }
        }

        void GetTrayIcons(AutomationElement taskbar, ref List<AutomationElement> arrayWithIcons)
        {
            _logger.LogInformation("Getting tray icons...");
            var trayIcons = taskbar.FindAllDescendants(cf => cf.ByAutomationId("NotifyItemIcon"));

            if (trayIcons != null && trayIcons.Length != 0)
            {
                _logger.LogInformation("Found non-hidden tray icons, adding to list...");
                arrayWithIcons.AddRange(trayIcons);
            }
            else
            {
                _logger.LogWarning("Couldn't find any non-hidden tray icons.");
            }
        }

        AutomationElement GetNotifierIcon(ref List<AutomationElement> arrayWithIcons)
        {
            _logger.LogInformation("Getting the Focusrite Notifier icon...");
            var notifierIcon = arrayWithIcons.Find(e =>
            {
                try { return e.Name.Contains("Focusrite"); } catch { return false; }
            });
            _ = notifierIcon ?? throw new CommandException("Couldn't find Focusrite Notifier icon, is it running? Is it not in the tray?");

            return notifierIcon;
        }

        AutomationElement GetContextMenu(AutomationElement desktop)
        {
            _logger.LogInformation("Getting the context menu...");
            var contextMenu = desktop.FindFirstChild(tf => tf.ByControlType(ControlType.Menu).And(tf.ByName("Context")));
            _ = contextMenu ?? throw new CommandException("Couldn't find context menu, something must have went wrong. Maybe adjust wait time?");

            return contextMenu;
        }

        AutomationElement GetSettingsItem(AutomationElement contextMenu)
        {
            _logger.LogInformation("Selecting settings menu item...");
            var settingsItem = contextMenu.FindAllChildren().ToList().Find(e =>
            {
                try { return e.Name.Contains("settings"); } catch { return false; }
            });
            _ = settingsItem ?? throw new CommandException("Couldn't find settings item, something must have went wrong.");

            return settingsItem;
        }

        AutomationElement GetSettingsWindow(AutomationElement desktop)
        {
            _logger.LogInformation("Getting the Device Settings window...");
            var settingsWindow = desktop.FindAllChildren(cf => cf.ByControlType(ControlType.Window)).ToList().Find(e =>
            {
                try { return e.Name.Contains("Device Settings"); } catch { return false; }
            });
            _ = settingsWindow ?? throw new CommandException("Couldn't find the Device Settings window, something must have went wrong." +
                " Did the window open? Does it have 'Device Settings' in the title?");

            return settingsWindow;
        }

        AutomationElement GetSampleBox(AutomationElement settingsWindow)
        {
            _logger.LogInformation("Getting the Sample Rate combo box...");
            var sampleBox = settingsWindow.FindFirstChild(tf => tf.ByControlType(ControlType.ComboBox).And(tf.ByName("Sample Rate")));
            _ = sampleBox ?? throw new CommandException("Couldn't find the Sample Rate combo box, something must have went wrong." +
                " Are you using English?");

            return sampleBox;
        }

        bool SetSampleRate(AutomationElement sampleBox, string sampleRate)
        {
            _logger.LogInformation("Trying to set the Sample Rate to '{sampleRate}'...", sampleRate);
            if (sampleBox.AsComboBox().Value == sampleRate)
            {
                _logger.LogInformation("Sample Rate already set to '{sampleRate}', skipping...", sampleRate);

                return false;
            }
            else
            {
                // ? https://stackoverflow.com/a/42654772
                sampleBox.AsComboBox().Expand();
                Wait.UntilInputIsProcessed();

                var newSampleRateItem = sampleBox.AsComboBox().Items.FirstOrDefault(it => it.Text.Equals(sampleRate));
                _ = newSampleRateItem ?? throw new CommandException($"Couldn't find the '{sampleRate}' Sample Rate option in the combo box," +
                    " something must have went wrong. Is that a valid option?");

                newSampleRateItem.Click();

                return true;
            }
        }

        AutomationElement GetBufferBox(AutomationElement settingsWindow)
        {
            _logger.LogInformation("Getting the Buffer Size combo box...");
            var bufferBox = settingsWindow.FindFirstChild(tf => tf.ByControlType(ControlType.ComboBox).And(tf.ByName("Buffer Size")));
            _ = bufferBox ?? throw new CommandException("Couldn't find the Buffer Size combo box, something must have went wrong." +
                " Are you using English?");

            return bufferBox;
        }

        bool SetBufferSize(AutomationElement bufferBox, string bufferSize)
        {
            _logger.LogInformation("Trying to set the Buffer Size to '{bufferSize}'...", bufferSize);
            if (bufferBox.AsComboBox().Value == bufferSize)
            {
                _logger.LogInformation("Buffer Size already set to '{bufferSize}', skipping...", bufferSize);

                return false;
            }
            else
            {
                // ? https://stackoverflow.com/a/42654772
                bufferBox.AsComboBox().Expand();
                Wait.UntilInputIsProcessed();

                var newBufferSizeItem = bufferBox.AsComboBox().Items.FirstOrDefault(it => it.Text.Equals(bufferSize));
                _ = newBufferSizeItem ?? throw new CommandException($"Couldn't find the '{bufferSize}' Buffer Size option in the combo box," +
                    " something must have went wrong. Is that a valid option?");

                newBufferSizeItem.Click();

                return true;
            }
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
                var desktop = GetDesktop();
                var taskbar = GetTaskbar(desktop);

                // ! List is modified by the tray functions.
                var arrayWithIcons = new List<AutomationElement>();
                if (CheckHiddenTrayIcons) GetHiddenTrayIcons(desktop, taskbar, ref arrayWithIcons);
                if (CheckTrayIcons) GetTrayIcons(taskbar, ref arrayWithIcons);

                var notifierIcon = GetNotifierIcon(ref arrayWithIcons);

                _logger.LogInformation("Found the Focusrite Notifier icon, invoking...");
                notifierIcon.AsButton().Invoke();
                Wait.UntilInputIsProcessed();

                var contextMenu = GetContextMenu(desktop);
                var settingsItem = GetSettingsItem(contextMenu);

                _logger.LogInformation("Found settings menu item, invoking...");
                settingsItem.AsButton().Invoke();
                Wait.UntilInputIsProcessed(TimeSpan.FromMilliseconds(500));

                var settingsWindow = GetSettingsWindow(desktop);

                var sampleBox = GetSampleBox(settingsWindow);

                var isSampleSkipped = !SetSampleRate(sampleBox, SampleRate);
                Wait.UntilInputIsProcessed();

                if (!isSampleSkipped)
                {
                    // ? The UI thread for Focusrite Notifier seems to hang sometimes depending on the sample rate chosen,
                    // ? I wasn't able to find a way to actually determine when it comes back to life :/
                    _logger.LogInformation("Waiting some time until Notifier is responsive again...");
                    Wait.UntilInputIsProcessed(TimeSpan.FromMilliseconds(4000));
                }

                var bufferBox = GetBufferBox(settingsWindow);

                SetBufferSize(bufferBox, BufferSize);
                Wait.UntilInputIsProcessed();

                _logger.LogInformation("Finished");
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
}
