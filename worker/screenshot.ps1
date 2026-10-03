# --- screenshot of the STATISTICA window ----------------------------------
# Brings the query window on screen, activates the relevant document (data
# table or graph) and captures real screen pixels with CopyFromScreen.
# PrintWindow is not used because it does not render the MDI document children.
function Ensure-ShotTypes {
  Add-Type -AssemblyName System.Drawing
  Add-Type -AssemblyName System.Windows.Forms
  if (-not ([System.Management.Automation.PSTypeName]'StaShot.Win32').Type) {
    Add-Type @"
using System;
using System.Text;
using System.Runtime.InteropServices;
namespace StaShot {
  public class Win32 {
    public delegate bool EnumProc(IntPtr hWnd, IntPtr lParam);
    [DllImport("user32.dll")] public static extern bool GetWindowRect(IntPtr hWnd, out RECT lpRect);
    [DllImport("user32.dll")] public static extern bool SetForegroundWindow(IntPtr hWnd);
    [DllImport("user32.dll")] public static extern IntPtr GetForegroundWindow();
    [DllImport("user32.dll")] public static extern bool ShowWindow(IntPtr hWnd, int nCmdShow);
    [DllImport("user32.dll")] public static extern bool BringWindowToTop(IntPtr hWnd);
    [DllImport("user32.dll")] public static extern void SwitchToThisWindow(IntPtr hWnd, bool fAltTab);
    [DllImport("user32.dll")] public static extern IntPtr FindWindowEx(IntPtr parent, IntPtr childAfter, string cls, string title);
    [DllImport("user32.dll")] public static extern bool EnumChildWindows(IntPtr hWnd, EnumProc cb, IntPtr lParam);
    [DllImport("user32.dll")] public static extern int GetClassName(IntPtr hWnd, StringBuilder s, int n);
    [DllImport("user32.dll")] public static extern int GetWindowText(IntPtr hWnd, StringBuilder s, int n);
    [DllImport("user32.dll")] public static extern bool IsWindowVisible(IntPtr hWnd);
    [StructLayout(LayoutKind.Sequential)] public struct RECT { public int Left; public int Top; public int Right; public int Bottom; }
    public static IntPtr ActiveDocument(IntPtr main) {
      IntPtr mdi = FindWindowEx(main, IntPtr.Zero, "MDIClient", null);
      if (mdi == IntPtr.Zero) { return IntPtr.Zero; }
      IntPtr found = IntPtr.Zero;
      EnumChildWindows(mdi, (h, l) => {
        if (found != IntPtr.Zero) { return false; }
        var cls = new StringBuilder(256); GetClassName(h, cls, 256);
        var txt = new StringBuilder(512); GetWindowText(h, txt, 512);
        RECT r; GetWindowRect(h, out r);
        if (IsWindowVisible(h) && txt.Length > 0 && cls.ToString().StartsWith("Afx:") && (r.Right - r.Left) > 200) { found = h; return false; }
        return true;
      }, IntPtr.Zero);
      return found;
    }
  }
}
"@
  }
}

function Save-ScreenShot($app, $out, $mode) {
  Ensure-ShotTypes
  $procId = [int]$app.ProcessID
  $proc = $null
  for ($i = 0; $i -lt 24; $i++) {
    $proc = Get-Process -Id $procId -ErrorAction SilentlyContinue
    if ($proc -and $proc.MainWindowHandle -ne 0) { break }
    Start-Sleep -Milliseconds 250
  }
  if (-not $proc -or $proc.MainWindowHandle -eq 0) { throw 'STATISTICA main window not found' }
  $main = $proc.MainWindowHandle
  # Restore if minimized, then maximize and force it to the foreground
  # (SetForegroundWindow alone is often ignored by the foreground lock).
  $null = [StaShot.Win32]::ShowWindow($main, 9)
  Start-Sleep -Milliseconds 200
  $null = [StaShot.Win32]::ShowWindow($main, 3)
  try { [Microsoft.VisualBasic.Interaction]::AppActivate($procId) } catch { }
  [StaShot.Win32]::SwitchToThisWindow($main, $true)
  $null = [StaShot.Win32]::BringWindowToTop($main)
  $null = [StaShot.Win32]::SetForegroundWindow($main)
  $foreground = $false
  for ($i = 0; $i -lt 25; $i++) {
    if ([StaShot.Win32]::GetForegroundWindow() -eq $main) { $foreground = $true; break }
    Start-Sleep -Milliseconds 150
    $null = [StaShot.Win32]::SetForegroundWindow($main)
  }
  Start-Sleep -Milliseconds 900

  if ($mode -eq 'screen') {
    # Whole virtual desktop (all monitors): nothing is clipped and the app is
    # captured wherever it was maximized.
    $b = [System.Windows.Forms.SystemInformation]::VirtualScreen
    $x = $b.X
    $y = $b.Y
    $w = $b.Width
    $h = $b.Height
  }
  else {
    $target = $main
    if ($mode -eq 'document') {
      $doc = [StaShot.Win32]::ActiveDocument($main)
      if ($doc -ne [IntPtr]::Zero) { $target = $doc }
    }
    $r = New-Object StaShot.Win32+RECT
    $null = [StaShot.Win32]::GetWindowRect($target, [ref]$r)
    $x = $r.Left
    $y = $r.Top
    $w = $r.Right - $r.Left
    $h = $r.Bottom - $r.Top
  }
  if ($w -le 0 -or $h -le 0) { throw "bad capture size ${w}x${h}" }
  $bmp = New-Object System.Drawing.Bitmap $w, $h
  $gfx = [System.Drawing.Graphics]::FromImage($bmp)
  $gfx.CopyFromScreen($x, $y, 0, 0, (New-Object System.Drawing.Size $w, $h))
  $gfx.Dispose()
  $dir = [System.IO.Path]::GetDirectoryName($out)
  if ($dir -and -not (Test-Path -LiteralPath $dir)) { $null = New-Item -ItemType Directory -Path $dir -Force }
  $ext = [System.IO.Path]::GetExtension($out).ToLower()
  if ($ext -eq '.jpg' -or $ext -eq '.jpeg') { $bmp.Save($out, [System.Drawing.Imaging.ImageFormat]::Jpeg) }
  else { $bmp.Save($out, [System.Drawing.Imaging.ImageFormat]::Png) }
  $bmp.Dispose()
  return @{ bytes = (Wait-File $out); foreground = $foreground }
}

function Invoke-screenshot($app, $ss, $req) {
  $out = [string]$req.out
  if ($out -eq '') { throw 'out is required' }
  $mode = 'screen'
  if ($null -ne $req.mode -and "$($req.mode)" -ne '') { $mode = [string]$req.mode }
  $target = $ss
  # Build the requested content while still hidden.
  if ($null -ne $req.module -and $null -ne $ss) {
    try {
      $ana = comCall $app 'Analysis' @([int]$req.module, $ss)
      $d = $ana.Dialog
      if ($req.variables) { $null = comSetOne $d 'Variables' ([string]$req.variables) }
      if ($req.properties) {
        foreach ($p in $req.properties.PSObject.Properties) {
          try { $null = comSetOne $d $p.Name (Convert-JsonValue $p.Value) } catch { }
        }
      }
      if ($req.run) {
        try { $null = comCall $ana 'Run' @() } catch { }
        $d = $ana.Dialog
      }
      $key = 'Graphs'
      if ($req.result) { $key = [string]$req.result }
      try { $v = comGet $d $key @(); if ($null -ne $v) { $target = $v } } catch { }
    }
    catch { throw "screenshot: could not build content: $($_.Exception.Message)" }
  }
  # Bring the relevant document to the front so its pixels are on screen.
  if ($null -ne $target) { try { $target.Activate() } catch { } }
  $wait = 1500
  if ($null -ne $req.waitMs) { $wait = [int]$req.waitMs }
  try { $app.Visible = $true } catch { throw "screenshot: could not show window: $($_.Exception.Message)" }
  Start-Sleep -Milliseconds $wait
  try { $shot = Save-ScreenShot $app $out $mode } catch { throw "screenshot: capture failed: $($_.Exception.Message)" }
  return @{ out = $out; bytes = $shot.bytes; mode = $mode; foreground = $shot.foreground }
}
