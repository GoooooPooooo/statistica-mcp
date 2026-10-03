# --- spreadsheet access ---------------------------------------------------
function Get-VarIndex($ss, $key) {
  if ($null -eq $key) { throw 'variable key is required (index or name)' }
  if ($key -is [string]) {
    $name = [string]$key
    if ($name.Trim() -eq '') { throw 'empty variable name' }
    if ($name -match '^\d+$') { return Get-VarIndex $ss ([int]$name) }
    $i = $null
    try { $i = comCall $ss 'VariableNumber' @($name) } catch { $i = $null }
    if ($null -ne $i -and [int]$i -gt 0) { return [int]$i }
    # fall back to a linear scan (new variables added in this session are not
    # always visible to VariableNumber until the file is saved)
    $nv = [int]$ss.NumberOfVariables
    $cleanHit = 0
    for ($k = 1; $k -le $nv; $k++) {
      $raw = [string](comGet $ss 'VariableName' @($k))
      if ($raw -eq $name) { return [int]$k }
      if ((Get-CleanName $raw) -eq $name -and $cleanHit -eq 0) { $cleanHit = $k }
    }
    if ($cleanHit -gt 0) { return [int]$cleanHit }
    throw "variable not found: '$name'"
  }
  $n = [int]$key
  if ($n -lt 1 -or $n -gt [int]$ss.NumberOfVariables) {
    throw "variable index out of range: $n (sheet has $($ss.NumberOfVariables) variables)"
  }
  return $n
}

function Open-Sheet($app, $path, $sheet) {
  if (-not (Test-Path -LiteralPath $path)) { throw "file not found: $path" }
  $before = 0
  try { $before = [int]$app.Spreadsheets.Count } catch { $before = 0 }
  $null = comCall $app 'Open' @($path)
  $count = [int]$app.Spreadsheets.Count
  if ($count -eq 0) { throw "no spreadsheet in file: $path" }

  # Workbooks open each sheet as a separate item; the newest ones start at $before+1.
  $first = $before + 1
  if ($first -gt $count) { $first = 1 }
  $idx = $first
  if ($null -ne $sheet -and "$sheet" -ne '') {
    if ($sheet -is [string] -and $sheet -notmatch '^\d+$') {
      $found = 0
      for ($i = 1; $i -le $count; $i++) {
        $ss = $app.Spreadsheets.Item($i)
        if ([string]$ss.Name -eq $sheet) { $found = $i; break }
      }
      if ($found -eq 0) {
        $names = @()
        for ($i = 1; $i -le $count; $i++) { $names += [string]$app.Spreadsheets.Item($i).Name }
        throw "sheet named '$sheet' not found. available: $($names -join ', ')"
      }
      $idx = $found
    }
    else {
      $base = 0
      if ($before -gt 0) { $base = $before }
      $idx = $base + [int]$sheet
    }
  }
  if ($idx -lt 1 -or $idx -gt $count) { throw "sheet index out of range: $idx (file has $count sheets)" }
  $script:sheetIndex = $idx
  $ss = $app.Spreadsheets.Item($idx)
  $script:sheetName = [string]$ss.Name
  return $ss
}

function Get-VarType($ss, $idx) { return [int](comGet $ss 'VariableType' @($idx)) }
function Get-VarMissing($ss, $idx) {
  $m = $null
  try { $m = [double](comGet $ss 'VariableMissingData' @($idx)) } catch { $m = $null }
  return $m
}

function Read-Variable($ss, $idx, $offset, $limit) {
  $nCases = [int]$ss.NumberOfCases
  if ($null -eq $offset -or [int]$offset -lt 1) { $offset = 1 }
  else { $offset = [int]$offset }
  $n = $nCases - $offset + 1
  if ($n -lt 0) { $n = 0 }
  if ($null -ne $limit -and [int]$limit -ge 0 -and [int]$limit -lt $n) { $n = [int]$limit }

  $type = Get-VarType $ss $idx
  $mdCode = Get-VarMissing $ss $idx
  $out = [object[]]::new($n)

  if ($type -eq 1) {
    if ($n -gt 0) {
      $raw = comCall $ss 'GetData' @($offset, $n, $idx, 1, $true)
      for ($i = 0; $i -lt $n; $i++) { $out[$i] = [string]$raw[$i] }
    }
  }
  else {
    if ($n -gt 0) {
      $col = comGet $ss 'VData' @($idx)
      for ($i = 0; $i -lt $n; $i++) {
        $v = [double]$col[$offset - 1 + $i]
        if ($v -eq $MISSING -or ($null -ne $mdCode -and -not [double]::IsNaN($mdCode) -and $v -eq $mdCode)) {
          $out[$i] = $null
        }
        else { $out[$i] = ToStr $v }
      }
    }
  }
  return @{
    index = $idx
    name = [string](comGet $ss 'VariableName' @($idx))
    cleanName = Get-CleanName ([string](comGet $ss 'VariableName' @($idx)))
    longName = [string](comGet $ss 'VariableLongName' @($idx))
    type = $type
    typeLength = [int](comGet $ss 'VariableTypeLength' @($idx))
    offset = $offset
    count = $n
    values = $out
  }
}

function Write-Variable($ss, $key, $values) {
  $idx = Get-VarIndex $ss $key
  $nCases = [int]$ss.NumberOfCases
  $type = Get-VarType $ss $idx
  $values = @($values)

  if ($type -eq 1) {
    if ($values.Count -gt $nCases) { throw "too many values for text column $idx ($($values.Count) > $nCases cases)" }
    for ($i = 0; $i -lt $values.Count; $i++) {
      $v = $values[$i]
      if ($null -eq $v -or ([string]$v) -eq '') { continue }
      $null = comCall $ss 'SetData' @(($i + 1), $idx, [string]$v)
    }
    return @{ index = $idx; name = [string](comGet $ss 'VariableName' @($idx)); written = $values.Count; type = 'text' }
  }

  if ($values.Count -ne $nCases) {
    throw "numeric column $idx requires exactly $nCases values (got $($values.Count)); resize with set_size first"
  }
  $d = [double[]]::new($nCases)
  for ($i = 0; $i -lt $nCases; $i++) {
    $v = Num $values[$i]
    if ($null -eq $v) { $d[$i] = $MISSING } else { $d[$i] = $v }
  }
  $null = comSet $ss 'VData' @($idx, $d)
  return @{ index = $idx; name = [string](comGet $ss 'VariableName' @($idx)); written = $nCases; type = 'numeric' }
}

# --- named variable creation / whole-column rewrites ---------------------
function Add-OneVariable($ss, $name, $type) {
  $nv = [int]$ss.NumberOfVariables
  $null = comCall $ss 'AddVariables' @([string]$name, $nv, 1)
  $idx = $nv + 1
  $t = 0
  if ($null -ne $type) { $t = [int]$type }
  if ($t -eq 1) { $null = comCall $ss 'VariableSetTextType' @($idx, 10) }
  elseif ($t -ne 0) { $null = comSet $ss 'VariableType' @($idx, $t) }
  return $idx
}

function Read-CaseNames($ss, $limit) {
  $nc = [int]$ss.NumberOfCases
  $n = $nc
  if ($null -ne $limit -and [int]$limit -lt $nc) { $n = [int]$limit }
  $out = [object[]]::new($n)
  for ($i = 1; $i -le $n; $i++) {
    $m = $null
    try { $m = [string](comGet $ss 'CaseName' @($i)) } catch { $m = $null }
    $out[$i - 1] = $m
  }
  return $out
}
