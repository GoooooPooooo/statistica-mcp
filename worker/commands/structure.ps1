# --- structure / data commands -------------------------------------------
function Invoke-describe($app, $ss, $req) {
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
  return @{
    cases = $nc
    variables = $nv
    sheetName = $script:sheetName
    sheetIndex = $script:sheetIndex
    sheets = $sheetNames
    varInfo = $vars
  }
}

function Invoke-read($app, $ss, $req) {
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
  return @{ cases = $nc; count = $keys.Count; data = $data }
}

function Invoke-write($app, $ss, $req) {
  $written = [object[]]::new(@($req.columns).Count)
  $i = 0
  foreach ($col in @($req.columns)) {
    $key = $col.index
    if ($null -eq $key) { $key = $col.name }
    $written[$i] = Write-Variable $ss $key $col.values
    $i++
  }
  return @{ cases = [int]$ss.NumberOfCases; written = $written }
}

function Invoke-addwrite($app, $ss, $req) {
  # Add one or more named variables and write their values in a single COM session.
  $added = [object[]]::new(@($req.columns).Count)
  $i = 0
  foreach ($col in @($req.columns)) {
    $name = [string]$col.name
    if ($name -eq '') { throw 'addwrite column name is required' }
    $type = 0
    if ($null -ne $col.type) { $type = [int]$col.type }
    $idx = Add-OneVariable $ss $name $type
    $vals = @($col.values)
    $w = Write-Variable $ss $idx $vals
    $added[$i] = @{ index = $idx; name = [string](comGet $ss 'VariableName' @($idx)); written = $w.written }
    $i++
  }
  return @{ variables = [int]$ss.NumberOfVariables; cases = [int]$ss.NumberOfCases; added = $added }
}

function Invoke-formula($app, $ss, $req) {
  # Assign a formula to a variable (stored in its long name) and recompute it.
  $idx = Get-VarIndex $ss $req.variable
  $formula = [string]$req.formula
  if ($formula -ne '' -and -not $formula.StartsWith('=')) { $formula = '=' + $formula }
  $recalc = $true
  if ($null -ne $req.recalculate -and $req.recalculate -eq $false) { $recalc = $false }
  if ($formula -ne '') { Set-LongName $ss $idx $formula }
  if ($recalc) { $null = comCall $ss 'Recalculate' @($idx) }
  $preview = Read-Variable $ss $idx 1 5
  return @{
    index = $idx
    name = [string](comGet $ss 'VariableName' @($idx))
    longName = [string](comGet $ss 'VariableLongName' @($idx))
    values = $preview.values
  }
}

function Invoke-add_variables($app, $ss, $req) {
  $name = [string]$req.name
  if ($name -eq '') { throw 'name is required' }
  $count = 1
  if ($null -ne $req.count) { $count = [int]$req.count }
  if ($count -lt 1) { throw 'count must be >= 1' }
  $nv = [int]$ss.NumberOfVariables
  if ($null -eq $req.after -or "$($req.after)" -eq '') { $after = $nv } else { $after = [int]$req.after }
  if ($after -lt 0) { $after = 0 }
  if ($after -gt $nv) { $after = $nv }
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
  return @{ variables = [int]$ss.NumberOfVariables; added = $added }
}

function Invoke-rename($app, $ss, $req) {
  if (-not $req.renames) { throw 'renames is required' }
  $props = @($req.renames.PSObject.Properties)
  $done = [object[]]::new($props.Count)
  $i = 0
  foreach ($prop in $props) {
    $idx = Get-VarIndex $ss $prop.Name
    $null = comSet $ss 'VariableName' @($idx, [string]$prop.Value)
    $done[$i] = @{ index = $idx; from = $prop.Name; to = [string]$prop.Value }
    $i++
  }
  if ($req.longNames) {
    foreach ($prop in $req.longNames.PSObject.Properties) {
      $idx = Get-VarIndex $ss $prop.Name
      Set-LongName $ss $idx ([string]$prop.Value)
    }
  }
  return @{ renamed = $done; variables = [int]$ss.NumberOfVariables }
}

function Invoke-delete_variables($app, $ss, $req) {
  $from = Get-VarIndex $ss $req.from
  $to = Get-VarIndex $ss $req.to
  if ($to -lt $from) { $t = $from; $from = $to; $to = $t }
  $null = comCall $ss 'DeleteVariables' @($from, $to)
  return @{ variables = [int]$ss.NumberOfVariables; deletedFrom = $from; deletedTo = $to }
}

function Invoke-set_size($app, $ss, $req) {
  $nc = [int]$ss.NumberOfCases
  if ($null -ne $req.cases) { $nc = [int]$req.cases }
  $nv = [int]$ss.NumberOfVariables
  if ($null -ne $req.variables) { $nv = [int]$req.variables }
  $null = comCall $ss 'SetSize' @($nc, $nv)
  return @{ cases = [int]$ss.NumberOfCases; variables = [int]$ss.NumberOfVariables }
}

function Invoke-case_names($app, $ss, $req) {
  $nc = [int]$ss.NumberOfCases
  if ($null -ne $req.names) {
    $arr = @($req.names)
    $limit = $arr.Count
    if ($limit -gt $nc) { $limit = $nc }
    for ($i = 0; $i -lt $limit; $i++) {
      $v = $arr[$i]
      if ($null -eq $v) { continue }
      $null = comSet $ss 'CaseName' @(($i + 1), [string]$v)
    }
  }
  $showLimit = 200
  if ($null -ne $req.limit) { $showLimit = [int]$req.limit }
  $names = Read-CaseNames $ss $showLimit
  $nonEmpty = 0
  foreach ($n in $names) { if ($null -ne $n -and "$n" -ne '') { $nonEmpty++ } }
  return @{ cases = $nc; count = $names.Count; named = $nonEmpty; names = $names }
}

function Invoke-sort($app, $ss, $req) {
  $keys = @($req.variables)
  if ($keys.Count -eq 0) { throw 'sort requires at least one variable' }
  $orders = @()
  if ($null -ne $req.order) { $orders = @($req.order) }
  # apply from the least significant key to the most significant one
  for ($k = $keys.Count - 1; $k -ge 0; $k--) {
    $idx = Get-VarIndex $ss $keys[$k]
    $ord = 0
    if ($k -lt $orders.Count -and $null -ne $orders[$k]) {
      $o = "$($orders[$k])"
      if ($o -eq '1' -or $o -eq 'desc' -or $o -eq 'descending' -or $o -eq 'true') { $ord = 1 }
    }
    $type = Get-VarType $ss $idx
    $by = 0
    if ($type -eq 1) { $by = 1 }
    $null = comCall $ss 'SortDataEx' @([object[]]@($idx), [object[]]@($ord), [object[]]@($by), $false, $true)
  }
  $preview = @()
  $pc = [Math]::Min(5, [int]$ss.NumberOfCases)
  for ($i = 1; $i -le $pc; $i++) {
    $m = $null
    try { $m = [string](comGet $ss 'CaseName' @($i)) } catch { $m = $null }
    $preview += $m
  }
  return @{ cases = [int]$ss.NumberOfCases; variables = [int]$ss.NumberOfVariables; sortedBy = @($keys); firstCaseNames = $preview }
}

function Invoke-select($app, $ss, $req) {
  # Keep only the 1-based case numbers listed in req.cases, rewriting every column.
  if ($null -eq $req.cases) { throw 'select requires a case list' }
  $keep = @($req.cases | ForEach-Object { [int]$_ })
  $nv = [int]$ss.NumberOfVariables
  $nc = [int]$ss.NumberOfCases
  foreach ($c in $keep) {
    if ($c -lt 1 -or $c -gt $nc) { throw "case index out of range: $c (sheet has $nc cases)" }
  }
  $keepNames = $true
  if ($null -ne $req.caseNames -and $req.caseNames -eq $false) { $keepNames = $false }

  $cols = [object[]]::new($nv)
  for ($v = 1; $v -le $nv; $v++) {
    $r = Read-Variable $ss $v $null $null
    $filtered = [object[]]::new($keep.Count)
    for ($j = 0; $j -lt $keep.Count; $j++) { $filtered[$j] = $r.values[$keep[$j] - 1] }
    $cols[$v - 1] = $filtered
  }
  $names = $null
  if ($keepNames) { $names = Read-CaseNames $ss $null }

  $null = comCall $ss 'SetSize' @($keep.Count, $nv)
  for ($v = 1; $v -le $nv; $v++) {
    $null = Write-Variable $ss $v $cols[$v - 1]
  }
  if ($keepNames) {
    for ($j = 0; $j -lt $keep.Count; $j++) {
      $n = $names[$keep[$j] - 1]
      if ($null -ne $n -and "$n" -ne '') { $null = comSet $ss 'CaseName' @(($j + 1), [string]$n) }
    }
  }
  return @{ cases = [int]$ss.NumberOfCases; variables = [int]$ss.NumberOfVariables; kept = $keep.Count; removed = ($nc - $keep.Count) }
}

function Invoke-set_measurement($app, $ss, $req) {
  $idx = Get-VarIndex $ss $req.variable
  $map = @{ unspecified = 0; auto = 1; continuous = 2; categorical = 3; ordinal = 4 }
  $t = $req.type
  if ($null -eq $t) { throw 'type is required' }
  if ($t -is [string]) {
    $key = $t.ToLower()
    if ($map.ContainsKey($key)) { $t = $map[$key] }
    elseif ($t -match '^\d+$') { $t = [int]$t }
    else { throw "unknown measurement type: $($req.type) (use unspecified/auto/continuous/categorical/ordinal)" }
  }
  $t = [int]$t
  if ($t -lt 0 -or $t -gt 4) { throw "measurement type out of range: $t" }
  $null = comSet $ss 'VariableMeasurementType' @($idx, $t)
  $now = [int](comGet $ss 'VariableMeasurementType' @($idx, $true))
  return @{ index = $idx; name = [string](comGet $ss 'VariableName' @($idx)); measurementType = $now }
}

function Invoke-labels($app, $ss, $req) {
  $idx = Get-VarIndex $ss $req.variable
  $applied = @()
  if ($req.clear) {
    try { $null = comCall $ss 'RemoveAllLabels' @($idx) } catch { }
  }
  if ($null -ne $req.labels) {
    foreach ($p in $req.labels.PSObject.Properties) {
      $num = [double]::Parse($p.Name, $inv)
      $label = ''
      $desc = ''
      $val = $p.Value
      if ($val -is [string]) { $label = [string]$val }
      elseif ($null -ne $val) {
        if ($null -ne $val.label) { $label = [string]$val.label }
        if ($null -ne $val.description) { $desc = [string]$val.description }
      }
      if ($label -eq '' -and $desc -eq '') { continue }
      $null = comCall $ss 'SetTextLabel' @($idx, $num, $label, $desc)
      $applied += @{ value = $num; label = $label }
    }
  }
  $count = $null
  try { $count = [int](comGet $ss 'NumberOfTextLabels' @($idx)) } catch { $count = $null }
  return @{ index = $idx; name = [string](comGet $ss 'VariableName' @($idx)); count = $count; applied = $applied }
}

function Invoke-sheets($app, $ss, $req) {
  $sheets = @()
  $n = [int]$app.Spreadsheets.Count
  for ($i = 1; $i -le $n; $i++) {
    $s = $app.Spreadsheets.Item($i)
    $sheets += @{ index = $i; name = [string]$s.Name; cases = [int]$s.NumberOfCases; variables = [int]$s.NumberOfVariables }
  }
  return @{ count = $n; active = $script:sheetName; sheets = $sheets }
}
