# Close All Windows Script

A simple utility to close opened windows and tidy up your desktop when there are too many windows open. This script closes windows gracefully without terminating the applications themselves.

## Files Included

1. **`close-windows.ps1`** - Advanced PowerShell script with many options
2. **`README.md`** - This documentation file

## Quick Start

### PowerShell Script

```powershell
# Close all windows (excludes current PowerShell by default) - runs immediately
.\close-windows.ps1

# Close all windows including current PowerShell window (script will terminate)
.\close-windows.ps1 -IncludeCurrentWindow

# Close all windows except specific processes
.\close-windows.ps1 -ExcludeProcesses @("chrome", "notepad", "code")

# Enable debug mode to see detailed information and pause before closing
.\close-windows.ps1 -Debug

# Force pause before closing (without debug info)
.\close-windows.ps1 -Pause

# Show help
.\close-windows.ps1 -Help
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
- ✅ **Pause only in debug mode** (immediate execution by default)
- ✅ **Auto-restores minimized windows** before closing
- ✅ Detailed logging of what's being closed
- ✅ Summary report of actions taken
- ✅ **Can be compiled to standalone executable** using ps2exe

## How It Works

The script works by using Windows API to enumerate visible windows and sends WM_CLOSE messages to close them gracefully.

This method ensures that:
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
# Close all windows immediately (current PowerShell window is preserved by default)
.\close-windows.ps1
```

### Debug Mode
```powershell
# See detailed information and pause before closing
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
cd "path\to\close-all-windows"
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
2. Enter target: `powershell.exe -File "C:\path\to\close-all-windows\close-windows.ps1"`
3. Name it "Close All Windows"
4. Change icon if desired

**Note**: Replace `C:\path\to\close-all-windows\` with the actual path where you cloned/downloaded the repository. The shortcut will preserve the current PowerShell window by default, so it won't terminate the script unexpectedly.

## Compiling to Standalone Executable

You can compile the PowerShell script into a standalone `.exe` file using **ps2exe** for easier distribution and usage.

### Install ps2exe

```powershell
# Install ps2exe module (run as Administrator if needed)
Install-Module ps2exe -Scope CurrentUser
```

### Compile the Script

```powershell
# Navigate to the script directory
cd "path\to\close-all-windows"

# Compile to executable
ps2exe -inputFile "close-windows.ps1" -outputFile "close-windows.exe" -noOutput
```

### Compilation Options

```powershell
# Basic compilation
ps2exe .\close-windows.ps1 .\close-windows.exe

# Silent execution (no console window)
ps2exe .\close-windows.ps1 .\close-windows.exe -noOutput

# With custom icon and version info
ps2exe .\close-windows.ps1 .\close-windows.exe -noOutput -iconFile "icon.ico" -title "Close All Windows" -description "Desktop Window Cleanup Tool"

# For Windows service/background execution
ps2exe .\close-windows.ps1 .\close-windows.exe -noOutput -noOutput
```

### Using the Compiled Executable

Once compiled, you can:
- **Double-click** `close-windows.exe` to run with default settings
- **Command line**: `close-windows.exe -Debug` or `close-windows.exe -ExcludeProcesses chrome,notepad`
- **Create shortcuts** directly to the `.exe` file
- **Distribute** the single executable file without requiring PowerShell script execution policies

### Benefits of Compilation

- ✅ **No PowerShell execution policy issues**
- ✅ **Single file distribution**
- ✅ **Faster startup time**
- ✅ **Professional appearance**
- ✅ **Can run silently** (with `-noOutput`)
- ✅ **Custom icon support**

## License

Free to use and modify for personal and commercial purposes.
