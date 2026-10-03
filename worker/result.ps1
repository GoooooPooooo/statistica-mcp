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
