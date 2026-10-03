# --- STATISTICA BASIC (SVB) macro execution -------------------------------
# Opens a macro (from inline `code` or a `.svb` file), makes the opened
# spreadsheet active so `ActiveSpreadsheet` inside the macro resolves to it,
# then executes it. Used for custom recurrent models (DWLS/Lowess/EWPR) that
# the methodology supplies as SVB source.
function Invoke-macro($app, $ss, $req) {
  $code = [string]$req.code
  $src = [string]$req.source
  $svbPath = $null
  if ($src -ne '') {
    if (-not (Test-Path -LiteralPath $src)) { throw "macro file not found: $src" }
    $svbPath = $src
  }
  else {
    if ($code -eq '') { throw 'code or source is required' }
    $svbPath = Join-Path $env:TEMP ("sta-macro-" + [guid]::NewGuid().ToString('N') + ".svb")
    [System.IO.File]::WriteAllText($svbPath, $code, [System.Text.Encoding]::Default)
  }
  $null = $app.Open($svbPath)
  Start-Sleep -Milliseconds 500
  $m = comGet $app 'Macros' @()
  if ($null -eq $m) { throw 'macro was not opened' }
  if ($null -ne $ss) {
    $ss.Activate()
    Start-Sleep -Milliseconds 200
  }
  $null = comCall $m 'Execute' @()
  return @{ macro = [string](comGet $m 'Name' @()); executed = $true }
}
