# Create shortcuts.ps1
# Makes the app icon (if missing) and the "Find and Replace" shortcuts on the
# Desktop and in this folder. Run it again if you move this folder.
$ErrorActionPreference = 'Stop'
$AppDir = Split-Path -Parent $MyInvocation.MyCommand.Path
$IconPath = Join-Path $AppDir 'Find and Replace.ico'
Add-Type -AssemblyName System.Drawing

# Draws the icon: a page with an old line crossed out in red, a new line in
# green, and a magnifying glass
function New-IconImage([int]$Size) {
    $k = $Size / 256.0
    $bmp = New-Object System.Drawing.Bitmap($Size, $Size, [System.Drawing.Imaging.PixelFormat]::Format32bppArgb)
    $g = [System.Drawing.Graphics]::FromImage($bmp)
    $g.SmoothingMode = 'AntiAlias'
    $g.Clear([System.Drawing.Color]::Transparent)
    $x1 = 14 * $k; $y1 = 8 * $k; $x2 = 190 * $k; $y2 = 238 * $k; $fold = 50 * $k
    $page = New-Object System.Drawing.Drawing2D.GraphicsPath
    $page.AddLines([System.Drawing.PointF[]]@(
        (New-Object System.Drawing.PointF($x1, $y1)), (New-Object System.Drawing.PointF(($x2 - $fold), $y1)),
        (New-Object System.Drawing.PointF($x2, ($y1 + $fold))), (New-Object System.Drawing.PointF($x2, $y2)),
        (New-Object System.Drawing.PointF($x1, $y2))))
    $page.CloseFigure()
    $g.FillPath([System.Drawing.Brushes]::White, $page)
    $edge = New-Object System.Drawing.Pen([System.Drawing.Color]::FromArgb(255, 70, 84, 110), [single][Math]::Max(1, 10 * $k))
    $g.DrawPath($edge, $page)
    $corner = New-Object System.Drawing.Drawing2D.GraphicsPath
    $corner.AddLines([System.Drawing.PointF[]]@(
        (New-Object System.Drawing.PointF(($x2 - $fold), $y1)), (New-Object System.Drawing.PointF(($x2 - $fold), ($y1 + $fold))),
        (New-Object System.Drawing.PointF($x2, ($y1 + $fold)))))
    $corner.CloseFigure()
    $g.FillPath((New-Object System.Drawing.SolidBrush([System.Drawing.Color]::FromArgb(255, 205, 214, 230))), $corner)
    $g.DrawPath($edge, $corner)
    # The old words, crossed out
    $grey = New-Object System.Drawing.SolidBrush([System.Drawing.Color]::FromArgb(255, 150, 160, 178))
    $g.FillRectangle($grey, [single](38 * $k), [single](70 * $k), [single](112 * $k), [single][Math]::Max(2, 24 * $k))
    $strike = New-Object System.Drawing.Pen([System.Drawing.Color]::FromArgb(255, 215, 40, 40), [single][Math]::Max(1, 11 * $k))
    $g.DrawLine($strike, [single](30 * $k), [single](82 * $k), [single](158 * $k), [single](82 * $k))
    # The new words
    $green = New-Object System.Drawing.SolidBrush([System.Drawing.Color]::FromArgb(255, 30, 150, 60))
    $g.FillRectangle($green, [single](38 * $k), [single](112 * $k), [single](104 * $k), [single][Math]::Max(2, 26 * $k))
    foreach ($y in 156, 184) { $g.FillRectangle($grey, [single](38 * $k), [single]($y * $k), [single](70 * $k), [single][Math]::Max(1, 11 * $k)) }
    # The magnifying glass
    $glass = New-Object System.Drawing.SolidBrush([System.Drawing.Color]::FromArgb(210, 205, 228, 255))
    $r = 44 * $k; $cx = 168 * $k; $cy = 176 * $k
    $g.FillEllipse($glass, [single]($cx - $r), [single]($cy - $r), [single](2 * $r), [single](2 * $r))
    $ring = New-Object System.Drawing.Pen([System.Drawing.Color]::FromArgb(255, 45, 55, 75), [single][Math]::Max(1.5, 15 * $k))
    $g.DrawEllipse($ring, [single]($cx - $r), [single]($cy - $r), [single](2 * $r), [single](2 * $r))
    $handle = New-Object System.Drawing.Pen([System.Drawing.Color]::FromArgb(255, 45, 55, 75), [single][Math]::Max(2, 26 * $k))
    $handle.StartCap = 'Round'; $handle.EndCap = 'Round'
    $g.DrawLine($handle, [single](202 * $k), [single](210 * $k), [single](238 * $k), [single](246 * $k))
    $g.Dispose()
    return $bmp
}

# Saves 16, 32 and 48 pixel pictures (classic format) and a 256 pixel one (PNG) in one .ico file
function Save-Icon([string]$Path) {
    $images = @()
    foreach ($size in 16, 32, 48, 256) {
        $bmp = New-IconImage $size
        $ms = New-Object System.IO.MemoryStream
        if ($size -eq 256) { $bmp.Save($ms, [System.Drawing.Imaging.ImageFormat]::Png) }
        else {
            $w = New-Object System.IO.BinaryWriter($ms)
            $maskRow = [int]([Math]::Ceiling($size / 32.0) * 4)
            foreach ($v in 40, $size, ($size * 2)) { $w.Write([int]$v) }
            $w.Write([int16]1); $w.Write([int16]32)
            foreach ($v in 0, ($size * $size * 4 + $maskRow * $size), 0, 0, 0, 0) { $w.Write([int]$v) }
            for ($y = $size - 1; $y -ge 0; $y--) {
                for ($x = 0; $x -lt $size; $x++) { $c = $bmp.GetPixel($x, $y); $w.Write([byte]$c.B); $w.Write([byte]$c.G); $w.Write([byte]$c.R); $w.Write([byte]$c.A) }
            }
            $w.Write((New-Object byte[] ($maskRow * $size)))
            $w.Flush()
        }
        $images += , @($size, $ms.ToArray())
        $bmp.Dispose()
    }
    $out = New-Object System.IO.MemoryStream
    $bw = New-Object System.IO.BinaryWriter($out)
    $bw.Write([int16]0); $bw.Write([int16]1); $bw.Write([int16]$images.Count)
    $offset = 6 + 16 * $images.Count
    foreach ($img in $images) {
        $dim = $img[0]; if ($dim -ge 256) { $dim = 0 }
        $bw.Write([byte]$dim); $bw.Write([byte]$dim); $bw.Write([byte]0); $bw.Write([byte]0)
        $bw.Write([int16]1); $bw.Write([int16]32); $bw.Write([int]$img[1].Length); $bw.Write([int]$offset)
        $offset += $img[1].Length
    }
    foreach ($img in $images) { $bw.Write([byte[]]$img[1]) }
    $bw.Flush()
    [System.IO.File]::WriteAllBytes($Path, $out.ToArray())
}

function New-Shortcut([string]$Path) {
    $shell = New-Object -ComObject WScript.Shell
    $lnk = $shell.CreateShortcut($Path)
    $lnk.TargetPath = Join-Path $env:WINDIR 'System32\WindowsPowerShell\v1.0\powershell.exe'
    $lnk.Arguments = '-NoProfile -STA -ExecutionPolicy Bypass -WindowStyle Hidden -File "' + (Join-Path $AppDir 'Find and Replace.ps1') + '"'
    $lnk.WorkingDirectory = $AppDir
    $lnk.IconLocation = "$IconPath,0"
    $lnk.WindowStyle = 7   # start minimized, so no black PowerShell box flashes up
    $lnk.Description = 'Find and replace words, or whole beginnings and endings, in .rtf, .doc, .docx and .txt documents'
    $lnk.Save()
}

if (-not (Test-Path -LiteralPath $IconPath)) { Save-Icon $IconPath }
$desktop = [Environment]::GetFolderPath('Desktop')
New-Shortcut (Join-Path $desktop 'Find and Replace.lnk')
New-Shortcut (Join-Path $AppDir 'Find and Replace.lnk')
"Shortcuts made on the Desktop ($desktop) and in $AppDir"