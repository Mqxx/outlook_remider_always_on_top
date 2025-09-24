# Configuration
$partialTitle = "Erinnerung(en)"
$checkIntervalMs = 500
# -------------



Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing

$notify = New-Object System.Windows.Forms.NotifyIcon
$notify.Icon = [System.Drawing.SystemIcons]::Information
$notify.Visible = $true
$notify.Text = "OutlookReminder Always-on-Top"

$contextMenu = New-Object System.Windows.Forms.ContextMenuStrip

$exitItem = New-Object System.Windows.Forms.ToolStripMenuItem "Exit"
$exitItem.Add_Click({
    $notify.Visible = $false
    $notify.Dispose()
    [System.Windows.Forms.Application]::Exit()
})
$contextMenu.Items.Add($exitItem) | Out-Null

$notify.ContextMenuStrip = $contextMenu

$signature = @"
[DllImport("kernel32.dll")]
public static extern bool SetConsoleTitle(string lpConsoleTitle);
"@

Add-Type -MemberDefinition $signature -Name "Kernel32" -Namespace WinAPI
[WinAPI.Kernel32]::SetConsoleTitle("OutlookReminderAlwaysOnTop") | Out-Null

Add-Type @"
using System;
using System.Text;
using System.Runtime.InteropServices;

public class WinAPI {
    public delegate bool EnumWindowsProc(IntPtr hWnd, IntPtr lParam);

    [DllImport("user32.dll")]
    [return: MarshalAs(UnmanagedType.Bool)]
    public static extern bool EnumWindows(EnumWindowsProc lpEnumFunc, IntPtr lParam);

    [DllImport("user32.dll", CharSet=CharSet.Unicode)]
    public static extern int GetWindowText(IntPtr hWnd, StringBuilder lpString, int nMaxCount);

    [DllImport("user32.dll")]
    [return: MarshalAs(UnmanagedType.Bool)]
    public static extern bool IsWindowVisible(IntPtr hWnd);

    [DllImport("user32.dll")]
    [return: MarshalAs(UnmanagedType.Bool)]
    public static extern bool SetWindowPos(IntPtr hWnd, IntPtr hWndInsertAfter, int X, int Y, int cx, int cy, uint uFlags);

    public static readonly IntPtr HWND_TOPMOST = new IntPtr(-1);
    public const uint SWP_NOSIZE = 0x0001;
    public const uint SWP_NOMOVE = 0x0002;
    public const uint SWP_SHOWWINDOW = 0x0040;
}
"@

Write-Host "Checking for partial window title '$partialTitle' every $($checkIntervalMs)ms"
Write-Host ""

$timer = New-Object System.Windows.Forms.Timer
$timer.Interval = $checkIntervalMs
$timer.add_Tick({
    $callback = [WinAPI+EnumWindowsProc]{
        param($hWnd, $lParam)
        $sb = New-Object System.Text.StringBuilder 256
        [WinAPI]::GetWindowText($hWnd, $sb, $sb.Capacity) | Out-Null
        $text = $sb.ToString()

        if ($text -like "*$partialTitle*") {
            if ([WinAPI]::IsWindowVisible($hWnd)) {
                [WinAPI]::SetWindowPos($hWnd, [WinAPI]::HWND_TOPMOST, 0,0,0,0, `
                    [WinAPI]::SWP_NOMOVE -bor [WinAPI]::SWP_NOSIZE -bor [WinAPI]::SWP_SHOWWINDOW)
                Write-Host "[$(Get-Date -Format "HH:mm:ss")] Window '$text' set to Always on Top."
            }
        }
        return $true
    }
    [WinAPI]::EnumWindows($callback, [IntPtr]::Zero) | Out-Null
})
$timer.Start()

[System.Windows.Forms.Application]::Run()
