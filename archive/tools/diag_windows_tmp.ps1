# diag_windows_tmp.ps1 - list visible windows of a given process (diagnostic)
param([int]$TargetPid = 19304)

Add-Type @"
using System;
using System.Runtime.InteropServices;
using System.Text;
public class WinEnum {
  [DllImport("user32.dll")] public static extern bool EnumWindows(EnumWindowsProc cb, IntPtr l);
  public delegate bool EnumWindowsProc(IntPtr h, IntPtr l);
  [DllImport("user32.dll")] public static extern int GetWindowText(IntPtr h, StringBuilder s, int n);
  [DllImport("user32.dll")] public static extern bool IsWindowVisible(IntPtr h);
  [DllImport("user32.dll")] public static extern uint GetWindowThreadProcessId(IntPtr h, out uint pid);
}
"@

$target = [uint32]$TargetPid
$found = New-Object System.Collections.ArrayList

$cb = [WinEnum+EnumWindowsProc]{
    param($h, $l)
    $pid2 = [uint32]0
    [WinEnum]::GetWindowThreadProcessId($h, [ref]$pid2) | Out-Null
    if ($pid2 -eq $target -and [WinEnum]::IsWindowVisible($h)) {
        $sb = New-Object System.Text.StringBuilder 256
        [WinEnum]::GetWindowText($h, $sb, 256) | Out-Null
        $t = $sb.ToString()
        if ($t -ne "") { [void]$found.Add($t) }
    }
    return $true
}

[WinEnum]::EnumWindows($cb, [IntPtr]::Zero) | Out-Null
if ($found.Count -eq 0) { Write-Output "NO_VISIBLE_WINDOWS" }
foreach ($w in $found) { Write-Output ("WINDOW: " + $w) }
