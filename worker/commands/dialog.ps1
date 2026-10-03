# --- capture an analysis dialog window ------------------------------------
# STATISTICA creates the analysis dialog as a hidden top-level "#32770" window.
# This command brings it to the front and captures either just the dialog or the
# full screen, so the report can include the actual setup panels. It does NOT
# close STATISTICA (the dialog is cancelled, the application stays open).
function Ensure-DialogTypes {
  Add-Type -AssemblyName System.Drawing
  Add-Type -AssemblyName System.Windows.Forms
  if (-not ([System.Management.Automation.PSTypeName]'StaDlg.Win32').Type) {
    Add-Type @"
using System; using System.Text; using System.Runtime.InteropServices;
namespace StaDlg {
  public class Win32 {
    public delegate bool EnumProc(IntPtr hWnd, IntPtr lParam);
    [DllImport("user32.dll")] public static extern bool EnumWindows(EnumProc cb, IntPtr lParam);
    [DllImport("user32.dll")] public static extern uint GetWindowThreadProcessId(IntPtr hWnd, out uint pid);
    [DllImport("user32.dll")] public static extern int GetClassName(IntPtr hWnd, StringBuilder s, int n);
    [DllImport("user32.dll")] public static extern int GetWindowText(IntPtr hWnd, StringBuilder s, int n);
    [DllImport("user32.dll")] public static extern bool GetWindowRect(IntPtr hWnd, out RECT r);
    [DllImport("user32.dll")] public static extern bool ShowWindow(IntPtr hWnd, int n);
    [DllImport("user32.dll")] public static extern bool SetForegroundWindow(IntPtr hWnd);
    [DllImport("user32.dll")] public static extern bool BringWindowToTop(IntPtr hWnd);
    [StructLayout(LayoutKind.Sequential)] public struct RECT { public int Left, Top, Right, Bottom; }
    public static IntPtr MainFrame(uint pid) {
      IntPtr found = IntPtr.Zero;
      EnumWindows((h, l) => {
        uint p; GetWindowThreadProcessId(h, out p);
        if (p != pid) { return true; }
        var c = new StringBuilder(64); GetClassName(h, c, 64);
        if (c.ToString() == "XTPMainFrame") { found = h; return false; }
        return true;
      }, IntPtr.Zero);
      return found;
    }
    public static IntPtr FindDialog(uint pid, string titleContains) {
      IntPtr found = IntPtr.Zero;
      EnumWindows((h, l) => {
        uint p; GetWindowThreadProcessId(h, out p);
        if (p != pid) { return true; }
        var c = new StringBuilder(64); GetClassName(h, c, 64);
        if (c.ToString() != "#32770") { return true; }
        if (!string.IsNullOrEmpty(titleContains)) {
          var t = new StringBuilder(512); GetWindowText(h, t, 512);
          if (t.ToString().IndexOf(titleContains) < 0) { return true; }
        }
        found = h; return false;
      }, IntPtr.Zero);
      return found;
    }
    public static string Title(IntPtr h) { var t = new StringBuilder(512); GetWindowText(h, t, 512); return t.ToString(); }
  }
}
"@
  }
}

function Invoke-dialog($app, $ss, $req) {
  Ensure-DialogTypes
  $out = [string]$req.out
  if ($out -eq '') { throw 'out is required' }
  $mode = 'screen'
  if ($null -ne $req.mode -and "$($req.mode)" -ne '') { $mode = [string]$req.mode }
  $module = [int]$req.module

  try { $app.Visible = $true } catch { }
  $ana = comCall $app 'Analysis' @($module, $ss)
  $d = $ana.Dialog
  if ($req.variables) { $null = comSetOne $d 'Variables' ([string]$req.variables) }
  if ($req.properties) {
    foreach ($p in $req.properties.PSObject.Properties) {
      try { $null = comSetOne $d $p.Name (Convert-JsonValue $p.Value) } catch { }
    }
  }
  if ($req.run) { try { $null = comCall $ana 'Run' @() } catch { } }
  Start-Sleep -Milliseconds 1200

  $procId = [uint32]([int]$app.ProcessID)
  $key = ''
  if ($req.title) { $key = [string]$req.title }
  $hwnd = $null
  for ($i = 0; $i -lt 10 -and ($null -eq $hwnd -or $hwnd -eq [IntPtr]::Zero); $i++) {
    $hwnd = [StaDlg.Win32]::FindDialog($procId, $key)
    if ($null -eq $hwnd -or $hwnd -eq [IntPtr]::Zero) { Start-Sleep -Milliseconds 300 }
  }
  if ($null -eq $hwnd -or $hwnd -eq [IntPtr]::Zero) { throw 'analysis dialog window not found' }
  $titleText = [StaDlg.Win32]::Title($hwnd)

  # Bring the application frame and then the dialog to the foreground.
  $main = [StaDlg.Win32]::MainFrame($procId)
  if ($null -ne $main -and $main -ne [IntPtr]::Zero) {
    $null = [StaDlg.Win32]::ShowWindow($main, 3)
    $null = [StaDlg.Win32]::SetForegroundWindow($main)
  }
  $null = [StaDlg.Win32]::ShowWindow($hwnd, 5)
  $null = [StaDlg.Win32]::BringWindowToTop($hwnd)
  $null = [StaDlg.Win32]::SetForegroundWindow($hwnd)
  Start-Sleep -Milliseconds 1000

  if ($mode -eq 'screen') {
    $b = [System.Windows.Forms.SystemInformation]::VirtualScreen
    $x = $b.X; $y = $b.Y; $w = $b.Width; $h = $b.Height
  }
  else {
    $r = New-Object StaDlg.Win32+RECT
    $null = [StaDlg.Win32]::GetWindowRect($hwnd, [ref]$r)
    $x = $r.Left; $y = $r.Top; $w = $r.Right - $r.Left; $h = $r.Bottom - $r.Top
  }
  if ($w -le 0 -or $h -le 0) { throw "bad capture size ${w}x${h}" }
  $bmp = New-Object System.Drawing.Bitmap $w, $h
  $g = [System.Drawing.Graphics]::FromImage($bmp)
  $g.CopyFromScreen($x, $y, 0, 0, (New-Object System.Drawing.Size $w, $h))
  $g.Dispose()
  $dir = [System.IO.Path]::GetDirectoryName($out)
  if ($dir -and -not (Test-Path -LiteralPath $dir)) { $null = New-Item -ItemType Directory -Path $dir -Force }
  $bmp.Save($out, [System.Drawing.Imaging.ImageFormat]::Png)
  $bmp.Dispose()

  # Close only the dialog; leave STATISTICA running.
  try { $null = comCall $d 'Cancel' @() } catch { }

  return @{ out = $out; bytes = (Wait-File $out); title = $titleText; width = $w; height = $h; mode = $mode }
}
