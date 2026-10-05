# --- analysis engine ------------------------------------------------------
function Invoke-describe_analysis($app, $ss, $req) {
  $module = [int]$req.module
  $ana = comCall $app 'Analysis' @($module, $ss)
  $members = @()
  foreach ($m in ($ana.Dialog | Get-Member)) {
    $members += @{ kind = $m.MemberType.ToString(); name = $m.Name }
  }
  return @{
    module = $module
    name = [string]$ana.Name
    state = [int]$ana.State
    members = $members
  }
}

function Invoke-analysis($app, $ss, $req) {
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
        if ($v -is [System.Array]) {
          foreach ($item in $v) { $list += , $item }
        }
        else {
          $cnt = $null
          try { $cnt = [int](comGet $v 'Count' @()) } catch { $cnt = $null }
          if ($null -ne $cnt) {
            for ($gi = 1; $gi -le $cnt; $gi++) { try { $list += , (comGet $v 'Item' @($gi)) } catch { } }
          }
          else { $list += , $v }
        }
        # Keep only actual graphs: a spreadsheet result exposes NumberOfCases,
        # a graph document does not. Without this filter a table/graph pair
        # (e.g. Autocorrelations) would also export the table and force a
        # numeric suffix on every output file.
        $graphs = @()
        foreach ($item in $list) {
          $nc = $null
          try { $nc = comGet $item 'NumberOfCases' @() } catch { $nc = $null }
          if ($null -eq $nc) { $graphs += , $item }
        }
        if ($graphs.Count -eq 0) { $null = $warnings.Add("saveGraph: no graph in '$key'") }
        $idx = 1
        foreach ($g in $graphs) {
          $path = $outPath
          if ($graphs.Count -gt 1) {
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
  return @{ module = $module; name = [string]$ana.Name; steps = $out; results = $results; graphs = $graphOut; warnings = $warnings }
}
