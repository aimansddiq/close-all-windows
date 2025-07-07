# Close All Windows Script

A simple utility to close opened windows and tidy up your desktop when there are too many windows open. This script closes windows gracefully without terminating the applications themselves.

## Files Included

1. **`close-windows.ps1`** - Advanced PowerShell script with many options
2. **`close-windows-simple.bat`** - Simple batch file for quick use
3. **`README.md`** - This documentation file

## Quick Start

### Option 1: PowerShell Script (Recommended)

```powershell
# Close all windows (excludes current PowerShell by default)
.\close-windows.ps1

# Close all windows including current PowerShell window (script will terminate)
.\close-windows.ps1 -IncludeCurrentWindow

# Close all windows except specific processes
.\close-windows.ps1 -ExcludeProcesses @("chrome", "notepad", "code")

# Enable debug mode to see detailed information
.\close-windows.ps1 -Debug

# Show help
.\close-windows.ps1 -Help
```

### Option 2: Simple Batch File

Double-click `close-windows-simple.bat` or run it from command prompt:

```cmd
close-windows-simple.bat
```

## Features

### PowerShell Script Features
- ✅ Closes windows gracefully (sends WM_CLOSE message)
- ✅ Preserves applications in system tray
- ✅ Excludes system processes automatically
- ✅ **Excludes current PowerShell window by default** (prevents script termination)
- ✅ Option to include current window if needed
- ✅ Option to exclude specific processes
- ✅ **Debug mode with detailed information**
- ✅ **Pause before closing** (shows preview and waits for confirmation)
- ✅ **Auto-restores minimized windows** before closing
- ✅ Detailed logging of what's being closed
- ✅ Summary report of actions taken

### Simple Batch File Features
- ✅ Quick and easy to use
- ✅ No parameters needed
- ✅ Works on any Windows system

## How It Works

The scripts work by:

1. **PowerShell Script**: Uses Windows API to enumerate visible windows and sends WM_CLOSE messages to close them gracefully
2. **Batch Script**: Uses system commands to close common applications

Both methods ensure that:
- Applications aren't forcefully terminated
- System processes are protected
- Applications can save their work before closing

## Safety Features

- **System Protection**: Automatically excludes critical system processes
- **Graceful Closing**: Sends proper close messages, allowing applications to save work
- **No Force Kill**: Doesn't forcefully terminate processes
- **Selective Exclusion**: Can exclude specific applications you want to keep open

## Examples

### Basic Usage
```powershell
# Close all windows (current PowerShell window is preserved by default)
.\close-windows.ps1
```

### Debug Mode
```powershell
# See detailed information about what's happening
.\close-windows.ps1 -Debug
```

### Keep Specific Apps Open
```powershell
# Keep Chrome and VS Code open while closing everything else
.\close-windows.ps1 -ExcludeProcesses @("chrome", "Code")
```

### Safe Mode (Default Behavior)
```powershell
# This is the default - current PowerShell window is automatically preserved
.\close-windows.ps1
```

### Include Current Window (Advanced)
```powershell
# Close everything including current PowerShell (script will terminate)
.\close-windows.ps1 -IncludeCurrentWindow
```

## Troubleshooting

### PowerShell Execution Policy
If you get an execution policy error, run this first:
```powershell
Set-ExecutionPolicy -ExecutionPolicy RemoteSigned -Scope CurrentUser
```

### Some Windows Don't Close
This is normal behavior for:
- System windows
- Windows that have unsaved changes (they may prompt to save)
- Applications that override the close message

### Script Doesn't Run
Make sure you're running from the correct directory:
```powershell
cd "d:\AIMAN\OneDrive\AIMAN\Project\git\close-all-windows"
.\close-windows.ps1
```

## Customization

You can modify the scripts to:
- Add more processes to the exclude list
- Change the delay between window closures
- Add more detailed logging
- Create shortcuts with specific parameters

## Creating a Desktop Shortcut

1. Right-click on desktop → New → Shortcut
2. Enter target: `powershell.exe -File "d:\AIMAN\OneDrive\AIMAN\Project\git\close-all-windows\close-windows.ps1"`
3. Name it "Close All Windows"
4. Change icon if desired

**Note**: The shortcut will preserve the current PowerShell window by default, so it won't terminate the script unexpectedly.

## License

Free to use and modify for personal and commercial purposes.
