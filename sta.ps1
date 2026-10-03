param(
  [Parameter(Mandatory=$true)][string]$RequestFile,
  [Parameter(Mandatory=$true)][string]$ResponseFile
)

$ErrorActionPreference = 'Stop'
$MISSING = -999999998.0
$inv = [System.Globalization.CultureInfo]::InvariantCulture

# Make the process DPI-aware before any window is used: otherwise Windows
# virtualizes the screen and the screenshot is smaller than the real display.
try {
  Add-Type -Namespace StaDpi -Name Win -MemberDefinition '[System.Runtime.InteropServices.DllImport("user32.dll")] public static extern bool SetProcessDPIAware();'
  $null = [StaDpi.Win]::SetProcessDPIAware()
}
catch { }

Add-Type -AssemblyName Microsoft.VisualBasic
$CB = [Microsoft.VisualBasic.CallType]

# Worker code is split into small files and dot-sourced into this script scope.
. (Join-Path $PSScriptRoot 'worker\com.ps1')
. (Join-Path $PSScriptRoot 'worker\sheet.ps1')
. (Join-Path $PSScriptRoot 'worker\result.ps1')
. (Join-Path $PSScriptRoot 'worker\commands\structure.ps1')
. (Join-Path $PSScriptRoot 'worker\commands\io.ps1')
. (Join-Path $PSScriptRoot 'worker\commands\analysis.ps1')
. (Join-Path $PSScriptRoot 'worker\commands\macro.ps1')
. (Join-Path $PSScriptRoot 'worker\commands\open.ps1')
. (Join-Path $PSScriptRoot 'worker\screenshot.ps1')

# --- request handling -----------------------------------------------------
$app = $null
$result = $null
$attached = $false
$keepAlive = $false
try {
  $req = [System.IO.File]::ReadAllText($RequestFile, [System.Text.Encoding]::UTF8) | ConvertFrom-Json
  $cmd = [string]$req.cmd
  if ($req.attach) { $attached = $true }
  if ($cmd -eq 'open') { $keepAlive = $true }

  if ($cmd -eq 'info') {
    $result = @{ ok = $true; cmd = $cmd }
    $app = New-ComApp $attached
    if (-not $attached) {
      $app.Visible = $false
      Disable-Alerts $app
    }
    $result.attached = $attached
    $result.version = [string]$app.Version
    $result.versionEx = [string]$app.VersionEx
    $result.exe = [string]$app.Path
    $result.pid = $app.ProcessID
    if (-not $attached) { Close-App $app; $app = $null }
  }
  else {
    if ($keepAlive -and -not $attached) {
      # Reuse the running instance when there is one, so that a later
      # GetActiveObject-based attach always resolves to the same window.
      try { $app = [System.Runtime.InteropServices.Marshal]::GetActiveObject('STATISTICA.Application') }
      catch { $app = New-Object -ComObject 'STATISTICA.Application' }
      $app.Visible = $true
    }
    else {
      $app = New-ComApp $attached
      if (-not $attached) {
        $app.Visible = $false
        Disable-Alerts $app
      }
    }
    $ss = $null
    if ($req.path) { $ss = Open-Sheet $app ([string]$req.path) $req.sheet }
    elseif ($attached -and $cmd -ne 'import') {
      $ss = $app.ActiveSpreadsheet
      if ($null -eq $ss) { throw 'no active spreadsheet in the running STATISTICA instance' }
      $script:sheetName = [string]$ss.Name
      $script:sheetIndex = 1
      try { $script:sheetIndex = [int]$ss.Index } catch { $script:sheetIndex = 1 }
    }
    elseif ($cmd -notin @('import', 'open')) { throw "path is required for command '$cmd'" }

    $fn = "Invoke-$cmd"
    if (Get-Command -Name $fn -CommandType Function -ErrorAction SilentlyContinue) {
      $result = & $fn $app $ss $req
    }
    else { throw "unknown command: $cmd" }

    if ($req.save -and $cmd -notin @('export_csv', 'save_as', 'import')) {
      $sv = [string]$req.save
      if (Test-Path -LiteralPath $sv) { Remove-Item -LiteralPath $sv -Force }
      $null = comCall $ss 'SaveCopyAs' @($sv, $true)
      $result.saved = $sv
    }

    if (-not $attached -and -not $keepAlive) { Close-App $app; $app = $null }
  }

  $out = @{ ok = $true; result = $result }
}
catch {
  $msg = $_.Exception.Message
  $hr = $null
  try { $hr = '0x{0:X8}' -f $_.Exception.HResult } catch { }
  $out = @{ ok = $false; error = $msg; hresult = $hr }
}
finally {
  if ($null -ne $app -and -not $attached -and -not $keepAlive) {
    try { Close-AllDocuments $app } catch { }
    try { $app.Quit() } catch { }
  }
}

[System.IO.File]::WriteAllText($ResponseFile, ($out | ConvertTo-Json -Depth 24 -Compress), (New-Object System.Text.UTF8Encoding($false)))
