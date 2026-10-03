# --- persistent visible window --------------------------------------------
# Keeps the application running after the worker exits (the entry point skips
# Quit for this command), so the user can type values into the table and later
# calls can use attach=true to edit/graph the same window without reopening it.
function Invoke-open($app, $ss, $req) {
  $app.Visible = $true
  Start-Sleep -Milliseconds 400
  $result = @{
    pid = [int]$app.ProcessID
    version = [string]$app.Version
    exe = [string]$app.Path
    opened = $false
  }
  if ($null -ne $ss) {
    $result.opened = $true
    $result.sheetName = $script:sheetName
    $result.sheetIndex = $script:sheetIndex
    try { $ss.Activate() } catch { }
  }
  return $result
}
