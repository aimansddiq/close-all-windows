# Close All Windows Script
# This script closes opened windows without terminating the applications
# Useful for tidying up desktop when there are too many windows open

param(
    [switch]$IncludeCurrentWindow,
    [string[]]$ExcludeProcesses = @(),
    [switch]$Help,
    [switch]$Debug,
    [switch]$Pause
)

if ($Help) {
    Write-Host @"
Close All Windows Script
========================

Usage:
  .\close-windows.ps1 [OPTIONS]

Options:
  -IncludeCurrentWindow    Also close the current PowerShell window (script will terminate)
  -ExcludeProcesses        Array of process names to exclude (e.g., "notepad", "chrome")
  -Debug                   Enable debug mode with extra information and pause before closing
  -Pause                   Force pause before closing (even without debug mode)
  -Help                    Show this help message

Examples:
  .\close-windows.ps1                                    # Close all windows immediately (no pause)
  .\close-windows.ps1 -Debug                             # Debug mode with pause for confirmation
  .\close-windows.ps1 -Pause                             # Force pause before closing
  .\close-windows.ps1 -ExcludeProcesses @("chrome", "notepad")  # Exclude specific processes

"@
    return
}

# Add Windows API types
Add-Type @"
    using System;
    using System.Runtime.InteropServices;
    using System.Text;
    using System.Collections.Generic;
    
    public class WindowHelper {
        [DllImport("user32.dll")]
        public static extern bool EnumWindows(EnumWindowsProc enumProc, IntPtr lParam);
        
        [DllImport("user32.dll")]
        public static extern bool IsWindowVisible(IntPtr hWnd);
        
        [DllImport("user32.dll")]
        public static extern int GetWindowText(IntPtr hWnd, StringBuilder lpString, int nMaxCount);
        
        [DllImport("user32.dll")]
        public static extern int GetWindowTextLength(IntPtr hWnd);
        
        [DllImport("user32.dll")]
        public static extern uint GetWindowThreadProcessId(IntPtr hWnd, out uint lpdwProcessId);
        
        [DllImport("user32.dll")]
        public static extern bool PostMessage(IntPtr hWnd, uint Msg, IntPtr wParam, IntPtr lParam);
        
        [DllImport("user32.dll")]
        public static extern bool SendMessage(IntPtr hWnd, uint Msg, IntPtr wParam, IntPtr lParam);
        
        [DllImport("user32.dll")]
        public static extern IntPtr GetForegroundWindow();
        
        [DllImport("user32.dll")]
        public static extern bool IsWindow(IntPtr hWnd);
        
        [DllImport("user32.dll")]
        public static extern bool ShowWindow(IntPtr hWnd, int nCmdShow);
        
        [DllImport("kernel32.dll")]
        public static extern uint GetCurrentProcessId();
        
        public const uint WM_CLOSE = 0x0010;
        public const uint WM_SYSCOMMAND = 0x0112;
        public const uint SC_CLOSE = 0xF060;
        public const int SW_RESTORE = 9;
        
        public delegate bool EnumWindowsProc(IntPtr hWnd, IntPtr lParam);
        
        public static List<WindowInfo> GetVisibleWindows() {
            List<WindowInfo> windows = new List<WindowInfo>();
            
            EnumWindows(delegate(IntPtr hWnd, IntPtr param) {
                if (IsWindowVisible(hWnd)) {
                    int length = GetWindowTextLength(hWnd);
                    if (length > 0) {
                        StringBuilder builder = new StringBuilder(length + 1);
                        GetWindowText(hWnd, builder, builder.Capacity);
                        
                        uint processId;
                        GetWindowThreadProcessId(hWnd, out processId);
                        
                        windows.Add(new WindowInfo {
                            Handle = hWnd,
                            Title = builder.ToString(),
                            ProcessId = processId
                        });
                    }
                }
                return true;
            }, IntPtr.Zero);
            
            return windows;
        }
        
        public static void CloseWindow(IntPtr hWnd) {
            // Try multiple methods to close the window
            // Method 1: Send WM_SYSCOMMAND with SC_CLOSE (most reliable)
            SendMessage(hWnd, WM_SYSCOMMAND, (IntPtr)SC_CLOSE, IntPtr.Zero);
            
            // Method 2: Post WM_CLOSE message (fallback)
            PostMessage(hWnd, WM_CLOSE, IntPtr.Zero, IntPtr.Zero);
        }
        
        public static bool IsValidWindow(IntPtr hWnd) {
            return IsWindow(hWnd);
        }
        
        public static void RestoreWindow(IntPtr hWnd) {
            ShowWindow(hWnd, SW_RESTORE);
        }
    }
    
    public class WindowInfo {
        public IntPtr Handle { get; set; }
        public string Title { get; set; }
        public uint ProcessId { get; set; }
    }
"@

function Get-ProcessName($processId) {
    try {
        $process = Get-Process -Id $processId -ErrorAction SilentlyContinue
        return $process.ProcessName
    }
    catch {
        return $null
    }
}

function Close-AllWindows {
    param(
        [bool]$IncludeCurrentWindow,
        [string[]]$ExcludeProcesses,
        [bool]$Debug,
        [bool]$Pause
    )
    
    $currentProcessId = [WindowHelper]::GetCurrentProcessId()
    $currentWindow = [WindowHelper]::GetForegroundWindow()
    $windows = [WindowHelper]::GetVisibleWindows()
    
    $closedCount = 0
    $skippedCount = 0
    
    Write-Host "Scanning for visible windows..." -ForegroundColor Yellow
    Write-Host "Found $($windows.Count) visible windows" -ForegroundColor Green
    
    if ($Debug) {
        Write-Host "Current Process ID: $currentProcessId" -ForegroundColor Magenta
        Write-Host "Current Window Handle: $currentWindow" -ForegroundColor Magenta
    }
    
    Write-Host ""
    
    # Display all windows first for debugging
    if ($Debug) {
        Write-Host "Windows found:" -ForegroundColor Yellow
        foreach ($window in $windows) {
            $processName = Get-ProcessName $window.ProcessId
            Write-Host "  - Handle: $($window.Handle) | Title: '$($window.Title)' | Process: $processName (PID: $($window.ProcessId))" -ForegroundColor Gray
        }
        Write-Host ""
    }
    
    # Pause logic: pause if Debug mode is on OR if Pause is explicitly requested
    if ($Debug -or $Pause) {
        Write-Host "Press any key to continue with closing windows, or Ctrl+C to abort..." -ForegroundColor Red
        $null = $Host.UI.RawUI.ReadKey("NoEcho,IncludeKeyDown")
        Write-Host ""
    } else {
        Write-Host "Proceeding with window closure..." -ForegroundColor Yellow
        Write-Host ""
    }
    
    foreach ($window in $windows) {
        $processName = Get-ProcessName $window.ProcessId
        $shouldSkip = $false
        $skipReason = ""
        
        # Validate window handle first
        if (-not [WindowHelper]::IsValidWindow($window.Handle)) {
            $shouldSkip = $true
            $skipReason = "Invalid window handle"
        }
        
        # Skip current window by default (unless explicitly included)
        if (-not $shouldSkip -and -not $IncludeCurrentWindow -and $window.Handle -eq $currentWindow) {
            $shouldSkip = $true
            $skipReason = "Current PowerShell window (excluded by default)"
        }
        
        # Skip if process is in exclude list
        if (-not $shouldSkip -and $processName -and $ExcludeProcesses -contains $processName) {
            $shouldSkip = $true
            $skipReason = "Excluded process: $processName"
        }
        
        # Skip Windows system processes
        $systemProcesses = @("dwm", "winlogon", "csrss", "explorer", "taskmgr", "lsass", "services", "svchost")
        if (-not $shouldSkip -and $processName -and $systemProcesses -contains $processName.ToLower()) {
            $shouldSkip = $true
            $skipReason = "System process: $processName"
        }
        
        # Skip empty titles or system dialogs
        if (-not $shouldSkip -and ([string]::IsNullOrWhiteSpace($window.Title) -or $window.Title.Length -lt 2)) {
            $shouldSkip = $true
            $skipReason = "Empty or system window title"
        }
        
        if ($shouldSkip) {
            Write-Host "  [SKIP] $($window.Title) ($skipReason)" -ForegroundColor Gray
            $skippedCount++
        }
        else {
            try {
                Write-Host "  [CLOSE] Attempting to close: '$($window.Title)' (Process: $processName)" -ForegroundColor Cyan
                
                if ($Debug) {
                    Write-Host "    Debug: Window Handle = $($window.Handle), Valid = $([WindowHelper]::IsValidWindow($window.Handle))" -ForegroundColor Magenta
                }
                
                # Try to restore window first (in case it's minimized)
                [WindowHelper]::RestoreWindow($window.Handle)
                Start-Sleep -Milliseconds 50
                
                # Attempt to close the window
                [WindowHelper]::CloseWindow($window.Handle)
                
                # Wait a bit and check if window still exists
                Start-Sleep -Milliseconds 200
                
                if ([WindowHelper]::IsValidWindow($window.Handle)) {
                    Write-Host "    Warning: Window may still be open (application might have prevented closure)" -ForegroundColor Yellow
                } else {
                    Write-Host "    Success: Window closed successfully" -ForegroundColor Green
                }
                
                $closedCount++
                Start-Sleep -Milliseconds 100  # Small delay between closures
            }
            catch {
                Write-Host "  [ERROR] Failed to close: $($window.Title) - Error: $($_.Exception.Message)" -ForegroundColor Red
            }
        }
    }
    
    Write-Host ""
    Write-Host "Summary:" -ForegroundColor Yellow
    Write-Host "  Windows closed: $closedCount" -ForegroundColor Green
    Write-Host "  Windows skipped: $skippedCount" -ForegroundColor Gray
}

# Main execution
try {
    Write-Host "Close All Windows Script" -ForegroundColor Magenta
    Write-Host "========================" -ForegroundColor Magenta
    Write-Host ""
    
    if ($ExcludeProcesses.Count -gt 0) {
        Write-Host "Excluding processes: $($ExcludeProcesses -join ', ')" -ForegroundColor Yellow
    }
    
    if ($IncludeCurrentWindow) {
        Write-Host "WARNING: Current PowerShell window will be closed (script will terminate after execution)" -ForegroundColor Red
    } else {
        Write-Host "Current PowerShell window will be preserved" -ForegroundColor Green
    }
    
    Write-Host ""
    
    Close-AllWindows -IncludeCurrentWindow $IncludeCurrentWindow -ExcludeProcesses $ExcludeProcesses -Debug $Debug -Pause $Pause
    
    Write-Host ""
    Write-Host "Desktop cleanup completed!" -ForegroundColor Green
}
catch {
    Write-Host "Error: $($_.Exception.Message)" -ForegroundColor Red
    exit 1
}
