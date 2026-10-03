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
# VariableLongName is an indexed property with a "Set" accessor: CallByName Let fails,
# so it is assigned natively (this also enables formula variables, see the 'formula' command).
function Set-LongName($ss, $idx, $value) {
  $sv = [string]$value
  try {
    $ss.VariableLongName($idx) = $sv
  }
  catch {
    $a = [object[]]::new(2)
    $a[0] = $idx
    $a[1] = $sv
    $null = $ss.GetType().InvokeMember('VariableLongName', [System.Reflection.BindingFlags]::SetProperty, $null, $ss, $a)
  }
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
