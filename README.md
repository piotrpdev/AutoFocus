# AutoFocus

![example](.github/video/example.mp4)

## Overview
AutoFocus is a tool designed to automate the process of changing Sample Rate and Buffer Size in the Focusrite Notifier device settings. I made it because a lot of different programs change those settings when launched and then you have to manually change them back to prevent crackling in VoIP software (e.g. Discord).

## Requirements
- Windows (only tested on 11)
- Focusrite Notifier ([comes bundled with Focusrite Control](https://support.focusrite.com/hc/en-gb/articles/360013505620-The-Focusrite-Notifier-icon-isn-t-in-the-Windows-taskbar))
    - *make sure you have it set to launch at startup*

## Where do I get it?

Download the latest version from [Releases](https://github.com/piotrpdev/AutoFocus/releases).

## Usage

You can use AutoFocus via the command-line interface. Here is a basic example:

```bash
AutoFocus.exe -s 48000 -b 128
```

This command sets the Sample Rate to 48000 Hz and Buffer Size to 128 (default).

### Options

- `-s, --sample-rate <rate>`: Set the sample rate (in Hz). Default is "48000".
- `-b, --buffer-size <size>`: Set the buffer size. Default is "128".
- `-n, --notifier-path <path>`: Absolute path to 'Focusrite Notifier.exe'. Default is `"C:\\Program Files\\Focusrite\\Drivers\\Focusrite Notifier.exe"`.
- `-a, --notifier-args <args>`: Arguments to pass when launching 'Focusrite Notifier.exe'. Default is "40000".
- `-f, --from-tray`: Try launching Focusrite Notifier using the tray icon. Default is false.
    - `-t, --check-tray`: Check for Focusrite Notifier icon in non-hidden tray icons. Default is false.
    - `-x, --check-hidden-tray`: Check for Focusrite Notifier icon in hidden tray icons. Default is true.
- `-w, --waitAfterSample`, How long to wait after changing Sample Rate (Notifier often freezes). Default is 500.

## Troubleshooting

Debug logs are stored in `%LOCALAPPDATA%/AutoFocus/Logs/*.txt`


## What's under the hood? 👀
- [.NET Core 7](https://learn.microsoft.com/en-us/dotnet/core/introduction) (C#) for good Windows integration
- [FlaUI](https://github.com/FlaUI/FlaUI) (UIA3) for UI automation
- [CliFx](https://github.com/Tyrrrz/CliFx) for CLI functionality
- [Serilog](https://github.com/serilog/serilog) for logging

## License

Licensed under the MIT License. See [LICENSE.md](LICENSE.md) for details.

## Acknowledgements

Thank you:

- Authors of the packages I used
- [Stack Overflow](https://stackoverflow.com/)
- Focusrite for making great audio products, please don't sue me 🙏.
