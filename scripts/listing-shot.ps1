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
  [switch]$CaptureOnly,
  # Print the ribbon chrome this PowerPoint is showing, and take no picture.
  # This is how $HOST_TABS and $HOST_GROUPS below were set, and how to re-set
  # them after an Office update renames something.
  [switch]$ListChrome
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

# WHAT MAY APPEAR IN THE PICTURE, AND THE CERTIFICATION THAT SET IT
#
# Microsoft failed the 2026-09-24 submission on 2026-09-30 under policy
# 100.3.2.2: "Images contain references to another add-in, kindly update the
# images." The reference was a Script Lab TAB, present in all five pictures.
# Script Lab is a STORE add-in (wa104380862), and this repository's own
# `docs/LISTING.md` had reasoned, in writing, that leaving it in was acceptable.
#
# The gate in place could not have caught it. `ribbon-cache.mjs` takes a list of
# ids to DROP, so it removes only what somebody remembered to name, and the two
# names it was given were the sideloaded siblings. A store add-in puts itself
# back at every start, which is the one case a drop list cannot hold.
#
# So the question asked here is the other way round. Not "has everything named
# been removed", which is a question about somebody's memory, but "is anything
# PRESENT that we did not expect", which is a question about the picture. It is
# asked at the shutter, against the window about to be photographed, because
# that is the only moment at which the answer is about the picture.
#
# PROVENANCE: read off this machine with -ListChrome on 2026-10-02, PowerPoint
# on Windows, en-GB, with all three SSF add-ins and Script Lab installed. The
# reading corrected two things a person would get wrong from looking at a
# screenshot of the ribbon, which is why it was taken rather than assumed:
#
#  * "File" is NOT a TabItem. It is drawn like a tab and is not one, so it
#    never appears in this sweep. It is left in the list below because an
#    add-in naming a tab "File" is not a thing worth being clever about.
#  * The SELECTED TAB's own name appears again as a GROUP ("Home" was in both
#    lists). So a name that is an allowed TAB is allowed as a group too, which
#    is what `Assert-NoOtherAddin` does, rather than listing the tab names
#    twice and having the two copies drift.
#
# An Office update that renames a tab or a group makes this REFUSE rather than
# ship, which is the way round it should fail. Re-read with -ListChrome and
# widen it deliberately, never to get a capture through.
$HOST_TABS = @(
  "File", "Home", "Insert", "Draw", "Design", "Transitions", "Animations",
  "Slide Show", "Record", "Review", "View", "Help",
  # Contextual tabs. These are PowerPoint's own and appear when something on
  # the slide is selected. NOT in the 2026-10-02 reading, because nothing was
  # selected when it was taken: they are allowed by reasoning, so that a
  # legitimate capture is not refused for one, and not by measurement.
  "Shape Format", "Picture Format", "Graphics Format", "Table Design",
  "Layout", "Chart Design", "Format"
)

# Ribbon groups, as the 2026-10-02 reading returned them. The siblings draw
# GROUPS rather than tabs - that reading had "SSF Merge" and "SSF Charts" in it
# - so this is the half of the gate that catches them, and it catches them by
# what is on the ribbon rather than by an id somebody remembered to list.
$HOST_GROUPS = @(
  "Clipboard", "Slides", "Font", "Paragraph", "Drawing", "Editing",
  "Voice", "Add-ins", "Copilot"
)

# This add-in's own ribbon group, which is the one thing that SHOULD be there.
$OUR_GROUP = "SSF Slide Elements"

function Get-PptWindow {
  $proc = Get-Process POWERPNT -ErrorAction Stop
  $root = [System.Windows.Automation.AutomationElement]::RootElement
  $cond = New-Object System.Windows.Automation.PropertyCondition(
    [System.Windows.Automation.AutomationElement]::ProcessIdProperty, $proc.Id)
  $w = $root.FindFirst([System.Windows.Automation.TreeScope]::Children, $cond)
  if (-not $w) { throw "no PowerPoint window" }
  return $w
}

# Every ribbon TAB name, and every GROUP name under the ribbon.
#
# Tabs are read window-wide: TabItem is specific enough that nothing else in
# PowerPoint answers to it.
#
# Groups are read from inside the ribbon ONLY. Window-wide they came back clean
# on 2026-10-02, but that reading was taken with the TASK PANE SHUT, and the
# pane is a WebView2 whose HTML can expose groups of its own. Scoping to the
# ribbon means this add-in's own pane can never widen or trip its own gate.
#
# The container is the pane NAMED "Lower Ribbon" - by Name, not AutomationId,
# which was the first guess and does not exist: every ribbon element on this
# machine carries an empty AutomationId. Not finding it is a REFUSAL rather
# than a skipped check, because a gate that silently checks nothing is the
# thing this file is being changed to stop.
function Get-RibbonChrome($window) {
  $byType = {
    param($type)
    New-Object System.Windows.Automation.PropertyCondition(
      [System.Windows.Automation.AutomationElement]::ControlTypeProperty, $type)
  }

  $tabs = @()
  foreach ($e in $window.FindAll([System.Windows.Automation.TreeScope]::Descendants,
      (& $byType ([System.Windows.Automation.ControlType]::TabItem)))) {
    if ($e.Current.Name) { $tabs += $e.Current.Name }
  }

  $ribbon = $null
  foreach ($name in @("Lower Ribbon", "Ribbon")) {
    $ribbon = $window.FindFirst([System.Windows.Automation.TreeScope]::Descendants,
      (New-Object System.Windows.Automation.AndCondition(
        (New-Object System.Windows.Automation.PropertyCondition(
          [System.Windows.Automation.AutomationElement]::NameProperty, $name)),
        (New-Object System.Windows.Automation.PropertyCondition(
          [System.Windows.Automation.AutomationElement]::ControlTypeProperty,
          [System.Windows.Automation.ControlType]::Pane)))))
    if ($ribbon) { break }
  }

  $groups = @()
  if ($ribbon) {
    foreach ($e in $ribbon.FindAll([System.Windows.Automation.TreeScope]::Descendants,
        (& $byType ([System.Windows.Automation.ControlType]::Group)))) {
      if ($e.Current.Name) { $groups += $e.Current.Name }
    }
  }

  return @{
    Tabs        = @($tabs | Select-Object -Unique)
    Groups      = @($groups | Select-Object -Unique)
    FoundRibbon = [bool]$ribbon
  }
}

# THE GATE. Refuses the shutter when anything is on the ribbon that is not
# PowerPoint's own and is not this add-in.
function Assert-NoOtherAddin($window) {
  $chrome = Get-RibbonChrome $window
  if (-not $chrome.FoundRibbon) {
    throw ("could not find the ribbon container, so the check for other add-ins' " +
      "ribbon GROUPS did not run. Refusing rather than taking a picture that " +
      "nothing checked. Run -ListChrome to see what the tree looks like.")
  }

  $strangeTabs = @($chrome.Tabs | Where-Object { $HOST_TABS -notcontains $_ })
  # A group whose name is an allowed TAB is the selected tab's own wrapper, not
  # an add-in. Measured 2026-10-02: "Home" came back in both lists.
  $strangeGroups = @($chrome.Groups | Where-Object {
      $HOST_GROUPS -notcontains $_ -and $HOST_TABS -notcontains $_ -and $_ -ne $OUR_GROUP
    })

  if (-not ($chrome.Groups -contains $OUR_GROUP)) {
    throw ("this add-in's own ribbon group '" + $OUR_GROUP + "' is not on the ribbon, " +
      "so this is not a picture of the product. Groups seen: " +
      ($chrome.Groups -join ", "))
  }

  if ($strangeTabs.Count -gt 0 -or $strangeGroups.Count -gt 0) {
    $what = @()
    if ($strangeTabs.Count -gt 0) { $what += "tab(s): " + ($strangeTabs -join ", ") }
    if ($strangeGroups.Count -gt 0) { $what += "group(s): " + ($strangeGroups -join ", ") }
    throw ("the ribbon carries something this script does not recognise - " +
      ($what -join "; ") + ". AppSource failed the 2026-09-24 submission for " +
      "exactly this (policy 100.3.2.2, another add-in in the pictures), so no " +
      "picture is taken. Take the add-in off the ribbon, or - if that name is " +
      "PowerPoint's own and new - re-read with -ListChrome and widen the list " +
      "in this script deliberately.")
  }

  Write-Output ("ribbon checked: " + $chrome.Tabs.Count + " tabs, " +
    $chrome.Groups.Count + " groups, nothing but PowerPoint and " + $OUR_GROUP)
}

if (-not (Get-Process POWERPNT -ErrorAction SilentlyContinue)) { throw "start PowerPoint first" }

if ($ListChrome) {
  $chrome = Get-RibbonChrome (Get-PptWindow)
  Write-Output "TABS:"
  foreach ($t in $chrome.Tabs) { Write-Output ("  " + $t) }
  if ($chrome.FoundRibbon) {
    Write-Output "GROUPS under the ribbon:"
    foreach ($g in $chrome.Groups) { Write-Output ("  " + $g) }
  } else {
    Write-Output "GROUPS: the ribbon container was not found, so none were read."
  }
  return
}

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

# ASKED HERE, AND NOT EARLIER, DELIBERATELY.
#
# The window is sized, foregrounded and settled; the next statement is the
# shutter. A check that ran before the pane was opened would be a check on a
# different window, and a check that ran after the file was written would be
# the mistake this whole gate exists to correct - the 2026-09-24 pictures were
# looked at, by a person, and the Script Lab tab was in all five.
Assert-NoOtherAddin (Get-PptWindow)

# THE CAPTURE LANDS BESIDE ITS DESTINATION, NOT ON IT.
#
# The check below can REJECT this capture, and a rejected capture must not have
# already replaced a good picture. Writing straight to $Out threw the error
# after the damage: `-Out docs\listing-screenshot.png` is an ordinary way to run
# this, and a blank frame overwrote the committed shot and then complained. The
# staging file is left behind on a failure, deliberately, because looking at
# what PrintWindow actually returned is the first thing anyone does next.
$stage = "$Out.staging.png"
[Listing]::Shot($hwnd, $stage, $CropTop, $WIDE, $TALL)

# A BLANK FRAME IS THE FAILURE THIS TOOL CANNOT OTHERWISE SEE.
#
# Measured 2026-09-16: PrintWindow returned an all-black bitmap - PowerPoint
# had been left in an odd state by a teaching callout - and the script wrote it
# out and reported "wrote 1366x768" exactly as if nothing were wrong. The file
# is the right size, the right format, and a picture of nothing. Worse, the
# first check written for it sampled the PNG's bytes and called it fine,
# because an opaque image carries an alpha channel of 255 beside colour bytes
# of 0.
#
# So the pixels are read back. A real pane shot is mostly light and has
# hundreds of colours in it; a dead capture has one.
$img = [System.Drawing.Bitmap]::FromFile($stage)
try {
  $seen = New-Object 'System.Collections.Generic.HashSet[int]'
  $lit = 0
  $n = 0
  for ($y = 4; $y -lt $img.Height; $y += 16) {
    for ($x = 4; $x -lt $img.Width; $x += 16) {
      $c = $img.GetPixel($x, $y)
      [void]$seen.Add(($c.R -shl 16) -bor ($c.G -shl 8) -bor $c.B)
      if ([int]$c.R + [int]$c.G + [int]$c.B -gt 60) { $lit++ }
      $n++
    }
  }
  $share = if ($n -gt 0) { $lit / $n } else { 0 }
  Write-Output ("checked " + $n + " pixels: " + $seen.Count + " colours, " + [int]($share * 100) + "% lit")
  if ($seen.Count -lt 20 -or $share -lt 0.3) {
    throw ("that capture is blank or nearly - " + $seen.Count + " colours, " + [int]($share * 100) +
      "% lit. PrintWindow can return an empty bitmap when PowerPoint is mid-dialog or has a callout open. Clear it and run again. " +
      $Out + " was NOT touched; the capture is at " + $stage + " to look at.")
  }
  $wide = $img.Width
  $tall = $img.Height
} finally {
  # Before the move, or the file is still open and the move fails.
  $img.Dispose()
}
Move-Item -LiteralPath $stage -Destination $Out -Force
Write-Output ("wrote " + $wide + "x" + $tall + " to " + $Out + " (cropped " + $CropTop + "px of title bar)")
