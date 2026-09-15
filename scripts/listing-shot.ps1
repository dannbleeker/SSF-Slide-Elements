# The AppSource listing screenshot, taken end to end.
#
#   powershell -ExecutionPolicy Bypass -File scripts\listing-shot.ps1 -Out shot.png
#
# Windows and a real PowerPoint only, so nothing here runs in CI. What CI does
# hold is `scripts/ribbon-cache.mjs`, the one piece of this with logic in it.
# `docs/LISTING.md` is the recipe and says which decisions are the owner's.
#
# ASCII ONLY, deliberately. Windows PowerShell 5.1 reads a .ps1 with no byte
# order mark as ANSI, so a UTF-8 em dash arrives as three characters and the
# file stops parsing - the first draft of this script died on an em dash inside
# an error message, with the parser blaming a line thirty below it.
#
# WHY A SCRIPT RATHER THAN A SET OF CLICKS
#
# The store wants 1366x768 exactly, and three things in an ordinary PowerPoint
# window do not belong on a public page: the signed-in account's avatar, an
# "Upgrade your plan" button, and any other add-in's ribbon group. Getting all
# of that right by hand, repeatably, is the part nobody does twice the same way.
#
# HOW EACH IS HANDLED, none of it by retouching the picture
#
# * The avatar and the upsell live in the TITLE BAR. So the window is sized to
#   1366 x (768 + CropTop) and the title bar is cropped off. Every pixel kept is
#   a real pixel of a real window; nothing is painted, moved or invented, which
#   is what `docs/LISTING.md` requires of this image.
# * DPI awareness FIRST, or Windows lies about both the size asked for and the
#   size measured, and the capture comes back at the wrong pixel count - which
#   for a store listing is the entire requirement. Measured 2026-09-14.
# * A maximised window ignores MoveWindow's height (it came back 1366x2180), so
#   the window is restored first and the size is CHECKED rather than assumed.
# * PrintWindow with PW_RENDERFULLCONTENT, not a desktop BitBlt: it renders
#   through DWM, which is what works on a console session with no viewer
#   attached, and it captures the window rather than whatever is in front of it.
# * Other add-ins' ribbon groups are NOT handled here. That is
#   `scripts/ribbon-cache.mjs`, run separately and with a backup, because it
#   edits a file belonging to PowerPoint.

[CmdletBinding()]
param(
  [string]$Out = "listing-1366x768.png",
  # The ribbon button that opens the pane, as the ribbon itself names it.
  [string]$Button = "Slide elements",
  # The title bar's height in pixels. 48 on this machine at 100% scaling;
  # the script prints what it cropped so a different chrome can be corrected.
  [int]$CropTop = 48,
  # Skip creating a deck and opening the pane - capture what is already there.
  [switch]$CaptureOnly
)

$ErrorActionPreference = "Stop"
$WIDE = 1366
$TALL = 768

Add-Type -AssemblyName System.Drawing
Add-Type -AssemblyName UIAutomationClient
Add-Type -AssemblyName UIAutomationTypes

$native = @'
using System;
using System.Runtime.InteropServices;
using System.Drawing;

public class Listing {
  [DllImport("user32.dll")] public static extern bool SetProcessDPIAware();
  [DllImport("user32.dll")] public static extern bool MoveWindow(IntPtr h, int x, int y, int w, int t, bool repaint);
  [DllImport("user32.dll")] public static extern bool GetWindowRect(IntPtr h, out RECT r);
  [DllImport("user32.dll")] public static extern bool PrintWindow(IntPtr h, IntPtr dc, uint flags);
  [DllImport("user32.dll")] public static extern bool SetForegroundWindow(IntPtr h);
  [DllImport("user32.dll")] public static extern bool ShowWindow(IntPtr h, int cmd);
  [StructLayout(LayoutKind.Sequential)] public struct RECT { public int Left, Top, Right, Bottom; }

  public static void Shot(IntPtr h, string path, int cropTop, int w, int t) {
    RECT r; GetWindowRect(h, out r);
    using (Bitmap full = new Bitmap(r.Right - r.Left, r.Bottom - r.Top)) {
      using (Graphics g = Graphics.FromImage(full)) {
        IntPtr dc = g.GetHdc();
        PrintWindow(h, dc, 2);   // PW_RENDERFULLCONTENT
        g.ReleaseHdc(dc);
      }
      using (Bitmap cut = new Bitmap(w, t)) {
        using (Graphics g2 = Graphics.FromImage(cut)) {
          g2.DrawImage(full, new Rectangle(0, 0, w, t), new Rectangle(0, cropTop, w, t), GraphicsUnit.Pixel);
        }
        cut.Save(path, System.Drawing.Imaging.ImageFormat.Png);
      }
    }
  }
}
'@
Add-Type -TypeDefinition $native -ReferencedAssemblies System.Drawing, System.Drawing.Primitives

[void][Listing]::SetProcessDPIAware()

if (-not (Get-Process POWERPNT -ErrorAction SilentlyContinue)) { throw "start PowerPoint first" }

if (-not $CaptureOnly) {
  # An EMPTY deck: the validators' deck reads as a test fixture on a store page.
  $ppt = New-Object -ComObject PowerPoint.Application
  $pres = $ppt.Presentations.Add(-1)
  $pres.PageSetup.SlideWidth = 960
  $pres.PageSetup.SlideHeight = 540
  $null = $pres.Slides.Add(1, 12)   # ppLayoutBlank
  while ($pres.Slides.Count -gt 1) { $pres.Slides.Item($pres.Slides.Count).Delete() }
  Start-Sleep -Seconds 3

  # The pane, opened from the ribbon BY NAME rather than by remembered pixels.
  $proc = Get-Process POWERPNT
  $root = [System.Windows.Automation.AutomationElement]::RootElement
  $cond = New-Object System.Windows.Automation.PropertyCondition(
    [System.Windows.Automation.AutomationElement]::ProcessIdProperty, $proc.Id)
  $window = $root.FindFirst([System.Windows.Automation.TreeScope]::Children, $cond)
  if (-not $window) { throw "no PowerPoint window" }

  $byName = New-Object System.Windows.Automation.PropertyCondition(
    [System.Windows.Automation.AutomationElement]::NameProperty, "Home")
  $byTab = New-Object System.Windows.Automation.PropertyCondition(
    [System.Windows.Automation.AutomationElement]::ControlTypeProperty,
    [System.Windows.Automation.ControlType]::TabItem)
  $homeTab = $window.FindFirst([System.Windows.Automation.TreeScope]::Descendants,
    (New-Object System.Windows.Automation.AndCondition($byName, $byTab)))
  if ($homeTab) {
    $homeTab.GetCurrentPattern([System.Windows.Automation.SelectionItemPattern]::Pattern).Select()
    Start-Sleep -Milliseconds 1500
  }

  $btn = $window.FindFirst([System.Windows.Automation.TreeScope]::Descendants,
    (New-Object System.Windows.Automation.AndCondition(
      (New-Object System.Windows.Automation.PropertyCondition(
        [System.Windows.Automation.AutomationElement]::NameProperty, $Button)),
      (New-Object System.Windows.Automation.PropertyCondition(
        [System.Windows.Automation.AutomationElement]::ControlTypeProperty,
        [System.Windows.Automation.ControlType]::Button)))))
  if (-not $btn) { throw "no ribbon button called '$Button' - is the add-in installed?" }
  $btn.GetCurrentPattern([System.Windows.Automation.InvokePattern]::Pattern).Invoke()
  Write-Output "opened the pane; giving it 30s to fetch the catalogue"
  Start-Sleep -Seconds 30
}

$hwnd = (Get-Process POWERPNT).MainWindowHandle
if ($hwnd -eq [IntPtr]::Zero) { throw "PowerPoint has no main window" }
$want = $TALL + $CropTop

$r = New-Object Listing+RECT
for ($try = 1; $try -le 5; $try++) {
  [void][Listing]::ShowWindow($hwnd, 9)   # SW_RESTORE; a maximised window ignores the size
  Start-Sleep -Milliseconds 900
  [void][Listing]::MoveWindow($hwnd, 40, 40, $WIDE, $want, $true)
  Start-Sleep -Milliseconds 900
  [void][Listing]::GetWindowRect($hwnd, [ref]$r)
  if (($r.Right - $r.Left) -eq $WIDE -and ($r.Bottom - $r.Top) -eq $want) { break }
  Write-Output ("  attempt " + $try + ": " + ($r.Right - $r.Left) + "x" + ($r.Bottom - $r.Top) + ", wanted " + $WIDE + "x" + $want)
}
if (($r.Right - $r.Left) -ne $WIDE -or ($r.Bottom - $r.Top) -ne $want) {
  throw ("could not size the window to " + $WIDE + "x" + $want + "; the capture would be the wrong size and look right")
}
[void][Listing]::SetForegroundWindow($hwnd)
Start-Sleep -Seconds 4

[Listing]::Shot($hwnd, $Out, $CropTop, $WIDE, $TALL)
$img = [System.Drawing.Image]::FromFile($Out)
Write-Output ("wrote " + $img.Width + "x" + $img.Height + " to " + $Out + " (cropped " + $CropTop + "px of title bar)")
$img.Dispose()
