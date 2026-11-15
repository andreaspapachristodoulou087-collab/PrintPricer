# Sets up "Send to Pricer" on this computer. Safe to run again after an update.
#   1. Copies the two helper scripts to a fixed place, so they keep working if this folder moves.
#   2. Puts a "Send to Pricer" shortcut on the Desktop.
#   3. Shows the line to paste into Bambu Studio, and copies it to the clipboard.
param(
  [string]$ShortcutFolder = [Environment]::GetFolderPath('Desktop')
)

$ErrorActionPreference = 'Stop'
$here = Split-Path -Parent $MyInvocation.MyCommand.Path
$target = Join-Path $env:LOCALAPPDATA 'PrintPricer'

New-Item -ItemType Directory -Force -Path $target | Out-Null
foreach ($file in 'record-slice.js', 'send-to-pricer.ps1') {
  Copy-Item -LiteralPath (Join-Path $here $file) -Destination $target -Force
}

# The shortcut's icon is the pricer's own icon, scaled to the size Windows expects inside an .ico file.
$icon = Join-Path $target 'print-pricer.ico'
$png = Join-Path (Split-Path -Parent $here) 'apple-touch-icon.png'
if (Test-Path -LiteralPath $png) {
  try {
    Add-Type -AssemblyName System.Drawing
    $source = [Drawing.Image]::FromFile($png)
    $scaled = New-Object Drawing.Bitmap 256, 256
    $canvas = [Drawing.Graphics]::FromImage($scaled)
    $canvas.InterpolationMode = [Drawing.Drawing2D.InterpolationMode]::HighQualityBicubic
    $canvas.DrawImage($source, 0, 0, 256, 256)
    $canvas.Dispose(); $source.Dispose()
    $stream = New-Object IO.MemoryStream
    $scaled.Save($stream, [Drawing.Imaging.ImageFormat]::Png)
    $scaled.Dispose()
    $image = $stream.ToArray()
    # .ico header: one image, 256x256 (written as 0), 32-bit, starting right after these 22 bytes.
    $header = [byte[]](0, 0, 1, 0, 1, 0, 0, 0, 0, 0, 1, 0, 32, 0) + [BitConverter]::GetBytes([uint32]$image.Length) + [BitConverter]::GetBytes([uint32]22)
    [IO.File]::WriteAllBytes($icon, [byte[]]($header + $image))
  } catch {
    Write-Warning "The shortcut icon could not be made, so the shortcut gets a plain one. ($($_.Exception.Message))"
  }
}

$shell = New-Object -ComObject WScript.Shell
$shortcut = $shell.CreateShortcut((Join-Path $ShortcutFolder 'Send to Pricer.lnk'))
$shortcut.TargetPath = Join-Path $env:WINDIR 'System32\WindowsPowerShell\v1.0\powershell.exe'
$shortcut.Arguments = '-NoProfile -ExecutionPolicy Bypass -WindowStyle Hidden -File "' + (Join-Path $target 'send-to-pricer.ps1') + '"'
$shortcut.WorkingDirectory = $target
$shortcut.WindowStyle = 7
$shortcut.Description = 'Open Print Pricer with the print Bambu Studio sliced last'
if (Test-Path -LiteralPath $icon) { $shortcut.IconLocation = "$icon,0" }
$shortcut.Save()

$line = '"' + (Join-Path $env:WINDIR 'System32\wscript.exe') + '" //B //Nologo "' + (Join-Path $target 'record-slice.js') + '"'
$copied = $true
try { Set-Clipboard -Value $line } catch { $copied = $false }

Write-Host ''
Write-Host 'Send to Pricer is set up.' -ForegroundColor Green
Write-Host "  Shortcut: $(Join-Path $ShortcutFolder 'Send to Pricer.lnk')"
Write-Host "  Scripts:  $target"
Write-Host ''
Write-Host 'One step left, in Bambu Studio:'
Write-Host '  1. In the Process settings, turn on "Advanced", open the "Others" tab'
Write-Host '     and find "Post-processing Scripts".'
Write-Host '  2. Paste this line into the box:'
Write-Host ''
Write-Host "       $line" -ForegroundColor Cyan
Write-Host ''
if ($copied) { Write-Host '     (It is already on the clipboard, so Ctrl+V pastes it.)' }
Write-Host '  3. Save the process preset. The line belongs to the preset, so add it to each one you slice with.'
Write-Host ''
Write-Host 'After that: slice a plate, then double-click "Send to Pricer".'
