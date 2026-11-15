# Opens Print Pricer with the print that Bambu Studio sliced last.
# Run it from the "Send to Pricer" shortcut that setup.ps1 creates.
#
# record-slice.js leaves one note per sliced plate in %LOCALAPPDATA%\PrintPricer\slices.
# This script reads those notes and opens the pricer with them packed into the address.
param(
  # Address of the pricer. Saved rates and quotes belong to an address, so use the one you normally open.
  [string]$PricerUrl = 'https://andreaspapachristodoulou087-collab.github.io/PrintPricer/print-pricer.html',
  # Prints the address instead of opening it.
  [switch]$PrintOnly
)

$ErrorActionPreference = 'Stop'
$invariant = [Globalization.CultureInfo]::InvariantCulture
$store = Join-Path $env:LOCALAPPDATA 'PrintPricer\slices'

# "1d 2h 3m 4s" -> seconds
function ConvertTo-Seconds([string]$text) {
  $total = 0
  foreach ($part in [regex]::Matches($text, '(\d+)\s*([dhms])')) {
    $n = [int]$part.Groups[1].Value
    switch ($part.Groups[2].Value) {
      'd' { $total += $n * 86400 }
      'h' { $total += $n * 3600 }
      'm' { $total += $n * 60 }
      's' { $total += $n }
    }
  }
  $total
}

# Callers wrap the result in @(): PowerShell hands back a one-item list as the bare item.
function Split-List([string]$text, [char]$separator) {
  @($text.Split($separator) | ForEach-Object { $_.Trim().Trim('"') })
}

function ConvertTo-JsonText([string]$text) {
  '"' + ((($text -replace '[\x00-\x1f]', ' ') -replace '\\', '\\') -replace '"', '\"') + '"'
}

# Reads one note left by record-slice.js. Returns nothing when it holds no usable print.
function Read-Slice([string]$file) {
  $path = ''; $time = [long]0; $seconds = 0
  $weights = @(); $ids = @(); $types = @(); $colours = @()
  foreach ($line in Get-Content -LiteralPath $file -Encoding Unicode) {
    if ($line -match '^path=(.+)$') { $path = $Matches[1] }
    elseif ($line -match '^time=(\d+)$') { $time = [long]$Matches[1] }
    elseif ($line -match 'total estimated time:\s*(.+)$') { $seconds = ConvertTo-Seconds $Matches[1] }
    elseif ($line -match 'estimated printing time \(normal mode\)\s*=\s*(.+)$') { if (-not $seconds) { $seconds = ConvertTo-Seconds $Matches[1] } }
    elseif ($line -match '^; total filament weight \[g\]\s*:\s*(.+)$') { $weights = @(Split-List $Matches[1] ',') }
    elseif ($line -match '^; filament:\s*(.+)$') { $ids = @(Split-List $Matches[1] ',') }
    elseif ($line -match '^; filament_type\s*=\s*(.*)$') { $types = @(Split-List $Matches[1] ';') }
    elseif ($line -match '^; filament_colour\s*=\s*(.*)$') { $colours = @(Split-List $Matches[1] ';') }
  }

  # The weights cover only the filaments that were used. "; filament:" says which slots those are;
  # older Bambu Studio versions leave it out, and then the slots are taken in order.
  $filaments = @()
  $oneType = @($types | Select-Object -Unique).Count -eq 1
  for ($i = 0; $i -lt $weights.Count; $i++) {
    $grams = 0.0
    if (-not [double]::TryParse($weights[$i], [Globalization.NumberStyles]::Float, $invariant, [ref]$grams)) { continue }
    if ($grams -le 0) { continue }
    $slot = -1
    if ($ids.Count -eq $weights.Count) { $id = 0; if ([int]::TryParse($ids[$i], [ref]$id)) { $slot = $id - 1 } }
    elseif ($types.Count -eq $weights.Count) { $slot = $i }
    $type = 'Filament'; $colour = ''
    if ($slot -ge 0 -and $slot -lt $types.Count -and $types[$slot]) { $type = $types[$slot] }
    elseif ($oneType -and $types[0]) { $type = $types[0] }
    if ($slot -ge 0 -and $slot -lt $colours.Count) { $colour = $colours[$slot] }
    $filaments += @{ Type = $type; Colour = $colour; Grams = $grams }
  }

  if (-not $path -or (-not $seconds -and -not $filaments.Count)) { return }
  @{ Path = $path; Time = $time; Seconds = $seconds; Filaments = $filaments }
}

# Bambu Studio names its temporary G-code ".<process id>.<counter>.gcode", which leads back to the
# window the slice came from. That window's title is the project name.
function Get-ProjectName([string]$gcodePath) {
  if ([IO.Path]::GetFileName($gcodePath) -notmatch '^\.(\d+)\.') { return '' }
  $studio = Get-Process -Id ([int]$Matches[1]) -ErrorAction SilentlyContinue
  if (-not $studio -or $studio.ProcessName -notmatch 'bambu') { return '' }
  $title = ('' + $studio.MainWindowTitle).Trim().TrimStart('*').Trim()
  $title = ($title -replace '\s*-\s*Bambu\s?Studio.*$', '') -replace '^Bambu\s?Studio\s*-\s*', ''
  if ($title -match '^(Bambu\s?Studio|Untitled)?$') { return '' }
  $title
}

function Format-Slice($slice) {
  $filaments = @($slice.Filaments | ForEach-Object {
    '[' + (ConvertTo-JsonText $_.Type) + ',' + (ConvertTo-JsonText $_.Colour) + ',' + $_.Grams.ToString('0.##', $invariant) + ']'
  })
  '[' + $slice.Time + ',' + $slice.Seconds + ',[' + ($filaments -join ',') + ']]'
}

try {
  $slices = @()
  if (Test-Path -LiteralPath $store) {
    foreach ($note in Get-ChildItem -LiteralPath $store -Filter '*.txt' -File) {
      if ($note.LastWriteTime -lt (Get-Date).AddDays(-30)) { Remove-Item -LiteralPath $note.FullName -Force -ErrorAction SilentlyContinue; continue }
      try { $slice = Read-Slice $note.FullName } catch { $slice = $null }
      if ($slice) { $slices += $slice }
    }
  }

  if ($slices.Count) {
    # The newest slice, plus the other plates of the same project that are still open in Bambu Studio.
    $latest = $slices | Sort-Object { $_.Time } -Descending | Select-Object -First 1
    $folder = [IO.Path]::GetDirectoryName($latest.Path)
    $plates = @($slices | Where-Object {
      $_.Path -eq $latest.Path -or ([IO.Path]::GetDirectoryName($_.Path) -eq $folder -and (Test-Path -LiteralPath $_.Path))
    } | Sort-Object { $_.Time } | Select-Object -Last 12)

    $name = Get-ProjectName $latest.Path
    # Windows will not open very long addresses, so the oldest plates are left out if there are too many.
    do {
      $json = '{"v":1,"n":' + (ConvertTo-JsonText $name) + ',"p":[' + (@($plates | ForEach-Object { Format-Slice $_ }) -join ',') + ']}'
      $packed = [Convert]::ToBase64String([Text.Encoding]::UTF8.GetBytes($json)).TrimEnd('=').Replace('+', '-').Replace('/', '_')
      $tooLong = $packed.Length -gt 1800 -and $plates.Count -gt 1
      if ($tooLong) { $plates = @($plates | Select-Object -Skip 1) }
    } while ($tooLong)
  } else {
    $packed = [Convert]::ToBase64String([Text.Encoding]::UTF8.GetBytes('{"v":1,"e":"none"}')).TrimEnd('=')
  }

  $url = $PricerUrl + '#slice=' + $packed
  if ($PrintOnly) { $url } else { Start-Process $url }
} catch {
  if ($PrintOnly) { throw }
  $popup = New-Object -ComObject WScript.Shell
  [void]$popup.Popup("Could not send the print to Print Pricer.`n`n" + $_.Exception.Message, 0, 'Send to Pricer', 48)
}
