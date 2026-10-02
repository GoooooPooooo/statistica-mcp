param(
  [Parameter(Mandatory=$true)][string]$RequestFile,
  [Parameter(Mandatory=$true)][string]$ResponseFile
)

$ErrorActionPreference = 'Stop'
$MISSING = -999999998.0
$inv = [System.Globalization.CultureInfo]::InvariantCulture

Add-Type -AssemblyName Microsoft.VisualBasic
$CB = [Microsoft.VisualBasic.CallType]

# --- low level COM access -------------------------------------------------
function comGet($o, $n, $par) {
  if ($null -eq $par) { $par = @() }
  $p = @($par)
  $a = [object[]]::new($p.Count)
  for ($i = 0; $i -lt $p.Count; $i++) { $a[$i] = $p[$i] }
  return [Microsoft.VisualBasic.Interaction]::CallByName($o, $n, $CB::Get, $a)
}
function comSet($o, $n, $par) {
  if ($null -eq $par) { $par = @() }
  $p = @($par)
  $a = [object[]]::new($p.Count)
  for ($i = 0; $i -lt $p.Count; $i++) { $a[$i] = $p[$i] }
  return [Microsoft.VisualBasic.Interaction]::CallByName($o, $n, $CB::Let, $a)
}
function comCall($o, $n, $par) {
  if ($null -eq $par) { $par = @() }
  $p = @($par)
  $a = [object[]]::new($p.Count)
  for ($i = 0; $i -lt $p.Count; $i++) { $a[$i] = $p[$i] }
  return [Microsoft.VisualBasic.Interaction]::CallByName($o, $n, $CB::Method, $a)
}
# set a property to exactly one value, even when that value is itself an array
function comSetOne($o, $n, $v) {
  $a = [object[]]::new(1)
  $a[0] = $v
  return [Microsoft.VisualBasic.Interaction]::CallByName($o, $n, $CB::Let, $a)
}

function Wait-File($path) {
  for ($i = 0; $i -lt 20; $i++) {
    if (Test-Path -LiteralPath $path) {
      try { return (Get-Item -LiteralPath $path).Length } catch { return 0 }
    }
    Start-Sleep -Milliseconds 250
  }
  return 0
}

function Num($v) {
  if ($null -eq $v) { return $null }
  if ($v -is [string]) {
    if ($v -eq '') { return $null }
    return [double]::Parse($v, $inv)
  }
  return [double]$v
}
function ToStr($v) {
  if ($null -eq $v) { return $null }
  return ([double]$v).ToString('R', $inv)
}
function Get-CleanName($raw) {
  if ($null -eq $raw) { return $raw }
  $s = [string]$raw
  if ($s.StartsWith('!~\')) {
    $i = $s.LastIndexOf(',')
    if ($i -ge 0 -and $i -lt $s.Length - 1) { return $s.Substring($i + 1) }
  }
  return $s
}
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

# --- analysis result marshalling -----------------------------------------
function Read-WholeSpreadsheet($doc) {
  $nc = [int](comGet $doc 'NumberOfCases' @())
  $nv = [int](comGet $doc 'NumberOfVariables' @())
  $maxVars = 400
  if ($nv -gt $maxVars) { $nv = $maxVars }
  $names = [object[]]::new($nv)
  $columns = [object[]]::new($nv)
  for ($i = 1; $i -le $nv; $i++) {
    $r = Read-Variable $doc $i $null $null
    $names[$i - 1] = @{
      index = $i
      name = $r.name
      cleanName = $r.cleanName
      longName = $r.longName
      type = $r.type
    }
    $columns[$i - 1] = $r.values
  }
  return @{ kind = 'table'; cases = $nc; variables = $nv; names = $names; columns = $columns }
}

function Convert-Result($v) {
  if ($null -eq $v) { return $null }
  if ($v -is [string] -or $v -is [bool] -or $v -is [int] -or $v -is [long] -or $v -is [double] -or $v -is [decimal]) {
    return $v
  }
  if ($v -is [System.Array]) {
    $items = [object[]]::new($v.Length)
    for ($i = 0; $i -lt $v.Length; $i++) { $items[$i] = Convert-Result $v[$i] }
    return @{ kind = 'array'; items = $items }
  }
  # COM object: table?
  $nc = $null
  try { $nc = [int](comGet $v 'NumberOfCases' @()) } catch { $nc = $null }
  if ($null -ne $nc) { return Read-WholeSpreadsheet $v }
  # collection?
  $cnt = $null
  try { $cnt = [int](comGet $v 'Count' @()) } catch { $cnt = $null }
  if ($null -ne $cnt) {
    $limit = $cnt
    if ($limit -gt 200) { $limit = 200 }
    $items = [object[]]::new($limit)
    for ($i = 1; $i -le $limit; $i++) {
      try { $items[$i - 1] = Convert-Result (comGet $v 'Item' @($i)) }
      catch { $items[$i - 1] = $null }
    }
    return @{ kind = 'collection'; count = $cnt; items = $items }
  }
  $nm = $null
  try { $nm = [string](comGet $v 'Name' @()) } catch { $nm = $null }
  return @{ kind = 'document'; name = $nm; type = $v.GetType().FullName }
}

function Convert-JsonValue($v) {
  if ($v -is [int64]) { return [int]$v }
  if ($v -is [System.Array]) {
    $allNum = $true
    foreach ($x in $v) {
      if (-not ($x -is [int] -or $x -is [long] -or $x -is [double] -or $x -is [int64] -or $x -is [int32])) { $allNum = $false; break }
    }
    if ($allNum) {
      $ints = [int[]]::new($v.Length)
      for ($i = 0; $i -lt $v.Length; $i++) { $ints[$i] = [int]$v[$i] }
      return $ints
    }
    return [object[]]@($v)
  }
  return $v
}

# --- request handling -----------------------------------------------------
$app = $null
$result = $null
try {
  $req = [System.IO.File]::ReadAllText($RequestFile, [System.Text.Encoding]::UTF8) | ConvertFrom-Json
  $cmd = [string]$req.cmd

  if ($cmd -eq 'info') {
    $result = @{ ok = $true; cmd = $cmd }
    $app = New-Object -ComObject 'STATISTICA.Application'
    $app.Visible = $false
    $result.version = [string]$app.Version
    $result.versionEx = [string]$app.VersionEx
    $result.exe = [string]$app.Path
    $result.pid = $app.ProcessID
    $app.Quit()
    $app = $null
  }
  else {
    $app = New-Object -ComObject 'STATISTICA.Application'
    $app.Visible = $false
    $ss = $null
    if ($req.path) { $ss = Open-Sheet $app ([string]$req.path) $req.sheet }
    elseif ($cmd -ne 'import') { throw "path is required for command '$cmd'" }

    switch ($cmd) {

      'describe' {
        $nv = [int]$ss.NumberOfVariables
        $nc = [int]$ss.NumberOfCases
        $vars = [object[]]::new($nv)
        for ($i = 1; $i -le $nv; $i++) {
          $vars[$i - 1] = @{
            index = $i
            name = [string](comGet $ss 'VariableName' @($i))
            cleanName = Get-CleanName ([string](comGet $ss 'VariableName' @($i)))
            longName = [string](comGet $ss 'VariableLongName' @($i))
            type = (Get-VarType $ss $i)
            typeLength = [int](comGet $ss 'VariableTypeLength' @($i))
            measurementType = [int](comGet $ss 'VariableMeasurementType' @($i, $true))
            missingValue = Get-VarMissing $ss $i
          }
        }
        $sheetNames = @()
        $sheetCount = [int]$app.Spreadsheets.Count
        for ($i = 1; $i -le $sheetCount; $i++) { $sheetNames += [string]$app.Spreadsheets.Item($i).Name }
        $result = @{
          cases = $nc
          variables = $nv
          sheetName = $script:sheetName
          sheetIndex = $script:sheetIndex
          sheets = $sheetNames
          varInfo = $vars
        }
        break
      }

      'read' {
        $nc = [int]$ss.NumberOfCases
        $keys = @()
        if ($req.variables) {
          foreach ($k in @($req.variables)) { $keys += , $k }
        }
        else {
          for ($i = 1; $i -le [int]$ss.NumberOfVariables; $i++) { $keys += , $i }
        }
        $limit = $null
        if ($null -ne $req.limit) { $limit = [int]$req.limit }
        $offset = $null
        if ($null -ne $req.offset) { $offset = [int]$req.offset }
        $data = [object[]]::new($keys.Count)
        for ($i = 0; $i -lt $keys.Count; $i++) {
          $data[$i] = Read-Variable $ss (Get-VarIndex $ss $keys[$i]) $offset $limit
        }
        $result = @{ cases = $nc; count = $keys.Count; data = $data }
        break
      }

      'write' {
        $written = [object[]]::new(@($req.columns).Count)
        $i = 0
        foreach ($col in @($req.columns)) {
          $key = $col.index
          if ($null -eq $key) { $key = $col.name }
          $written[$i] = Write-Variable $ss $key $col.values
          $i++
        }
        $result = @{ cases = [int]$ss.NumberOfCases; written = $written }
        break
      }

      'add_variables' {
        $name = [string]$req.name
        if ($name -eq '') { throw 'name is required' }
        $count = 1
        if ($null -ne $req.count) { $count = [int]$req.count }
        if ($count -lt 1) { throw 'count must be >= 1' }
        $after = 0
        if ($null -ne $req.after) { $after = [int]$req.after }
        if ($after -lt 0) { $after = 0 }
        if ($after -gt [int]$ss.NumberOfVariables) { $after = [int]$ss.NumberOfVariables }
        $null = comCall $ss 'AddVariables' @($name, $after, $count)
        $added = [object[]]::new($count)
        $j = 0
        for ($k = $after + 1; $k -le $after + $count; $k++) {
          $added[$j] = @{ index = $k; name = [string](comGet $ss 'VariableName' @($k)) }
          $j++
        }
        if ($null -ne $req.type) {
          $t = [int]$req.type
          $len = 10
          if ($null -ne $req.typeLength) { $len = [int]$req.typeLength }
          for ($k = $after + 1; $k -le $after + $count; $k++) {
            if ($t -eq 1) { $null = comCall $ss 'VariableSetTextType' @($k, $len) }
            else { $null = comSet $ss 'VariableType' @($k, $t) }
          }
        }
        if ($req.longName) {
          for ($k = $after + 1; $k -le $after + $count; $k++) { $null = comSet $ss 'VariableLongName' @($k, [string]$req.longName) }
        }
        $result = @{ variables = [int]$ss.NumberOfVariables; added = $added }
        break
      }

      'rename' {
        if (-not $req.renames) { throw 'renames is required' }
        $done = [object[]]::new(@($req.renames).Count)
        $i = 0
        foreach ($prop in $req.renames.PSObject.Properties) {
          $idx = Get-VarIndex $ss $prop.Name
          $null = comSet $ss 'VariableName' @($idx, [string]$prop.Value)
          $done[$i] = @{ index = $idx; from = $prop.Name; to = [string]$prop.Value }
          $i++
        }
        if ($req.longNames) {
          foreach ($prop in $req.longNames.PSObject.Properties) {
            $idx = Get-VarIndex $ss $prop.Name
            $null = comSet $ss 'VariableLongName' @($idx, [string]$prop.Value)
          }
        }
        $result = @{ renamed = $done; variables = [int]$ss.NumberOfVariables }
        break
      }

      'delete_variables' {
        $from = Get-VarIndex $ss $req.from
        $to = Get-VarIndex $ss $req.to
        if ($to -lt $from) { $t = $from; $from = $to; $to = $t }
        $null = comCall $ss 'DeleteVariables' @($from, $to)
        $result = @{ variables = [int]$ss.NumberOfVariables; deletedFrom = $from; deletedTo = $to }
        break
      }

      'set_size' {
        $nc = [int]$ss.NumberOfCases
        if ($null -ne $req.cases) { $nc = [int]$req.cases }
        $nv = [int]$ss.NumberOfVariables
        if ($null -ne $req.variables) { $nv = [int]$req.variables }
        $null = comCall $ss 'SetSize' @($nc, $nv)
        $result = @{ cases = [int]$ss.NumberOfCases; variables = [int]$ss.NumberOfVariables }
        break
      }

      'export_csv' {
        $out = [string]$req.out
        if ($out -eq '') { throw 'out is required' }
        if ([System.IO.Path]::GetExtension($out) -eq '') { $out = $out + '.csv' }
        $dir = [System.IO.Path]::GetDirectoryName($out)
        if ($dir -and -not (Test-Path -LiteralPath $dir)) { $null = New-Item -ItemType Directory -Path $dir -Force }
        if (Test-Path -LiteralPath $out) { Remove-Item -LiteralPath $out -Force }
        $null = comCall $ss 'SaveAs' @($out, $true)
        $len = Wait-File $out
        $result = @{ out = $out; bytes = $len }
        break
      }

      'save_as' {
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
        $result = @{ out = $out; bytes = $len; copy = $copy }
        break
      }

      'import' {
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
        break
      }

      'describe_analysis' {
        $module = [int]$req.module
        $ana = comCall $app 'Analysis' @($module, $ss)
        $members = @()
        foreach ($m in ($ana.Dialog | Get-Member)) {
          $members += @{ kind = $m.MemberType.ToString(); name = $m.Name }
        }
        $result = @{
          module = $module
          name = [string]$ana.Name
          state = [int]$ana.State
          members = $members
        }
        break
      }

      'analysis' {
        $module = [int]$req.module
        $ana = comCall $app 'Analysis' @($module, $ss)
        $out = [ordered]@{}
        $runCount = 0
        $results = @{}
        $graphOut = [ordered]@{}
        $warnings = [System.Collections.ArrayList]::new()
        foreach ($step in @($req.steps)) {
          if ($step.PSObject.Properties.Name -contains 'set') {
            $d = $ana.Dialog
            foreach ($p in $step.set.PSObject.Properties) {
              try { $null = comSetOne $d $p.Name (Convert-JsonValue $p.Value) }
              catch { $null = $warnings.Add("could not set property '$($p.Name)': $($_.Exception.Message)") }
            }
          }
          elseif ($step.PSObject.Properties.Name -contains 'call') {
            $d = $ana.Dialog
            $args = @()
            if ($null -ne $step.args) { $args = @($step.args) }
            try { $null = comCall $d ([string]$step.call) $args }
            catch { $null = $warnings.Add("could not call '$($step.call)': $($_.Exception.Message)") }
          }
          elseif ($step.PSObject.Properties.Name -contains 'run') {
            try {
              $rc = [int](comCall $ana 'Run' @())
              $out["run$runCount"] = $rc
            }
            catch { $null = $warnings.Add("run failed: $($_.Exception.Message)") }
            $runCount++
          }
          elseif ($step.PSObject.Properties.Name -contains 'saveGraph') {
            $key = 'Graphs'
            if ($null -ne $step.result) { $key = [string]$step.result }
            $outPath = [string]$step.saveGraph
            $d = $ana.Dialog
            $v = $null
            try { $v = comGet $d $key @() } catch { $v = $null }
            if ($null -eq $v) {
              for ($k = 1; $k -le 6; $k++) {
                try { $v = comGet $d $key @($k) } catch { $v = $null }
                if ($null -ne $v) { break }
              }
            }
            if ($null -eq $v) { $null = $warnings.Add("saveGraph: no graph in '$key'") }
            else {
              $list = @()
              $cnt = $null
              try { $cnt = [int](comGet $v 'Count' @()) } catch { $cnt = $null }
              if ($null -ne $cnt) {
                for ($gi = 1; $gi -le $cnt; $gi++) { try { $list += , (comGet $v 'Item' @($gi)) } catch { } }
              }
              else { $list += , $v }
              $idx = 1
              foreach ($g in $list) {
                $path = $outPath
                if ($list.Count -gt 1) {
                  $ext = [System.IO.Path]::GetExtension($outPath)
                  $base = $outPath
                  if ($ext) { $base = $outPath.Substring(0, $outPath.Length - $ext.Length) }
                  $path = "$base`_$idx$ext"
                }
                $dir = [System.IO.Path]::GetDirectoryName($path)
                if ($dir -and -not (Test-Path -LiteralPath $dir)) { $null = New-Item -ItemType Directory -Path $dir -Force }
                if (Test-Path -LiteralPath $path) { Remove-Item -LiteralPath $path -Force }
                try {
                  $null = comCall $g 'SaveAs' @($path)
                  $graphOut["graph$($idx - 1)"] = @{ out = $path; bytes = (Wait-File $path) }
                }
                catch { $null = $warnings.Add("saveGraph failed: $($_.Exception.Message)") }
                $idx++
              }
            }
          }
          elseif ($step.PSObject.Properties.Name -contains 'result') {
            $key = [string]$step.result
            $d = $ana.Dialog
            $v = $null
            try { $v = comGet $d $key @() } catch { $v = $null }
            if ($null -eq $v) {
              for ($k = 1; $k -le 4; $k++) {
                try { $v = comGet $d $key @($k) } catch { $v = $null }
                if ($null -ne $v) { break }
              }
            }
            $name = $key
            $suffix = 1
            while ($results.ContainsKey($name)) { $suffix++; $name = "$key#$suffix" }
            $results[$name] = Convert-Result $v
          }
        }
        $result = @{ module = $module; name = [string]$ana.Name; steps = $out; results = $results; graphs = $graphOut; warnings = $warnings }
        break
      }

      default { throw "unknown command: $cmd" }
    }

    if ($req.save -and $cmd -notin @('export_csv', 'save_as', 'import')) {
      $sv = [string]$req.save
      if (Test-Path -LiteralPath $sv) { Remove-Item -LiteralPath $sv -Force }
      $null = comCall $ss 'SaveCopyAs' @($sv, $true)
      $result.saved = $sv
    }

    $app.Quit()
    $app = $null
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
  if ($null -ne $app) {
    try { $app.Quit() } catch { }
  }
}

[System.IO.File]::WriteAllText($ResponseFile, ($out | ConvertTo-Json -Depth 24 -Compress), (New-Object System.Text.UTF8Encoding($false)))
