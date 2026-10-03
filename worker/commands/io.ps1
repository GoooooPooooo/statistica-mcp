# --- import / export / save ----------------------------------------------
function Invoke-export_csv($app, $ss, $req) {
  $out = [string]$req.out
  if ($out -eq '') { throw 'out is required' }
  if ([System.IO.Path]::GetExtension($out) -eq '') { $out = $out + '.csv' }
  $dir = [System.IO.Path]::GetDirectoryName($out)
  if ($dir -and -not (Test-Path -LiteralPath $dir)) { $null = New-Item -ItemType Directory -Path $dir -Force }
  if (Test-Path -LiteralPath $out) { Remove-Item -LiteralPath $out -Force }
  $null = comCall $ss 'SaveAs' @($out, $true)
  $len = Wait-File $out
  return @{ out = $out; bytes = $len }
}

function Invoke-save_as($app, $ss, $req) {
  $out = [string]$req.out
  if ($out -eq '') { throw 'out is required' }
  $dir = [System.IO.Path]::GetDirectoryName($out)
  if ($dir -and -not (Test-Path -LiteralPath $dir)) { $null = New-Item -ItemType Directory -Path $dir -Force }
  $copy = $true
  if ($null -ne $req.copy -and $req.copy -eq $false) { $copy = $false }
  $exists = Test-Path -LiteralPath $out
  if ($exists -and -not $req.overwrite) { throw "file already exists (pass overwrite=true): $out" }
  if ($copy) { $null = comCall $ss 'SaveCopyAs' @($out, $true) }
  else { $null = comCall $ss 'SaveAs' @($out, $true) }
  $len = Wait-File $out
  return @{ out = $out; bytes = $len; copy = $copy }
}

function Invoke-import($app, $ss, $req) {
  $src = [string]$req.source
  if ($src -eq '') { throw 'source is required' }
  if (-not (Test-Path -LiteralPath $src)) { throw "file not found: $src" }
  $fmt = [string]$req.format
  $newss = $null
  if ($fmt -eq 'xls' -or $fmt -eq 'excel') {
    $sheetNo = 1
    if ($null -ne $req.sheetNumber) { $sheetNo = [int]$req.sheetNumber }
    $varsFromRow = $true
    if ($null -ne $req.variableNamesFromFirstRow) { $varsFromRow = [bool]$req.variableNamesFromFirstRow }
    $casesFromCol = $false
    if ($null -ne $req.caseNamesFromFirstColumn) { $casesFromCol = [bool]$req.caseNamesFromFirstColumn }
    $newss = comCall $app 'ImportXLSAsSpreadsheet' @($src, $sheetNo, $casesFromCol, $varsFromRow, $false, 0, 0, 0, 0)
  }
  else {
    $varsFromRow = $true
    if ($null -ne $req.variableNamesFromFirstRow) { $varsFromRow = [bool]$req.variableNamesFromFirstRow }
    $casesFromCol = $false
    if ($null -ne $req.caseNamesFromFirstColumn) { $casesFromCol = [bool]$req.caseNamesFromFirstColumn }
    # scTextImportQualifierDoubleQuote = 1073741825
    if ($null -ne $req.separator -and "$($req.separator)" -ne '') {
      $newss = comCall $app 'ImportTextAutoEx' @($src, [string]$req.separator, 1, 1073741825, $true, $casesFromCol, $varsFromRow, $true)
    }
    else {
      $newss = comCall $app 'ImportTextAuto' @($src, $null, 1, 1073741825, $true, $casesFromCol, $varsFromRow)
    }
  }
  if ($null -eq $newss) { throw "import produced no spreadsheet" }
  $result = @{
    cases = [int](comGet $newss 'NumberOfCases' @())
    variables = [int](comGet $newss 'NumberOfVariables' @())
    name = [string](comGet $newss 'Name' @())
  }
  if ($req.save) {
    $sv = [string]$req.save
    if (Test-Path -LiteralPath $sv) { Remove-Item -LiteralPath $sv -Force }
    $null = comCall $newss 'SaveCopyAs' @($sv, $true)
    $result.saved = $sv
    $result.bytes = Wait-File $sv
  }
  return $result
}
