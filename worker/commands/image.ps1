# --- image utilities (no STATISTICA needed) -------------------------------
# Combines several PNG images into one, stacking them vertically (default) or
# laying them horizontally, scaling each to a common width/height. Used to build
# two-panel report figures (e.g. series+model above, residuals below).
function Invoke-combine_images($req) {
  Add-Type -AssemblyName System.Drawing
  $images = @($req.images)
  if ($images.Count -lt 1) { throw 'images is required (non-empty array)' }
  $out = [string]$req.out
  if ($out -eq '') { throw 'out is required' }
  $direction = 'vertical'
  if ($null -ne $req.direction -and "$($req.direction)" -ne '') { $direction = [string]$req.direction }
  $gap = 16
  if ($null -ne $req.gap) { $gap = [int]$req.gap }

  $loaded = @()
  foreach ($p in $images) {
    if (-not (Test-Path -LiteralPath ([string]$p))) { throw "image not found: $p" }
    $loaded += , ([System.Drawing.Image]::FromFile([string]$p))
  }

  try {
    if ($direction -eq 'horizontal') {
      $h = 0
      foreach ($im in $loaded) { if ($im.Height -gt $h) { $h = $im.Height } }
      $w = 0
      $scaled = @()
      foreach ($im in $loaded) { $sw = [int][Math]::Round($im.Width * $h / $im.Height); $scaled += $sw; $w += $sw }
      $w += $gap * ($loaded.Count - 1)
      $bmp = New-Object System.Drawing.Bitmap $w, $h
      $g = [System.Drawing.Graphics]::FromImage($bmp)
      $g.Clear([System.Drawing.Color]::White)
      $g.InterpolationMode = [System.Drawing.Drawing2D.InterpolationMode]::HighQualityBicubic
      $x = 0
      for ($i = 0; $i -lt $loaded.Count; $i++) {
        $g.DrawImage($loaded[$i], $x, 0, $scaled[$i], $h)
        $x += $scaled[$i] + $gap
      }
    }
    else {
      $w = 0
      foreach ($im in $loaded) { if ($im.Width -gt $w) { $w = $im.Width } }
      $h = 0
      $scaled = @()
      foreach ($im in $loaded) { $sh = [int][Math]::Round($im.Height * $w / $im.Width); $scaled += $sh; $h += $sh }
      $h += $gap * ($loaded.Count - 1)
      $bmp = New-Object System.Drawing.Bitmap $w, $h
      $g = [System.Drawing.Graphics]::FromImage($bmp)
      $g.Clear([System.Drawing.Color]::White)
      $g.InterpolationMode = [System.Drawing.Drawing2D.InterpolationMode]::HighQualityBicubic
      $y = 0
      for ($i = 0; $i -lt $loaded.Count; $i++) {
        $g.DrawImage($loaded[$i], 0, $y, $w, $scaled[$i])
        $y += $scaled[$i] + $gap
      }
    }
    $g.Dispose()
    $dir = [System.IO.Path]::GetDirectoryName($out)
    if ($dir -and -not (Test-Path -LiteralPath $dir)) { $null = New-Item -ItemType Directory -Path $dir -Force }
    $bmp.Save($out, [System.Drawing.Imaging.ImageFormat]::Png)
    $bmp.Dispose()
  }
  finally {
    foreach ($im in $loaded) { $im.Dispose() }
  }
  return @{ out = $out; bytes = (Wait-File $out); count = $loaded.Count; direction = $direction }
}
