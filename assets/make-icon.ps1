# Generates assets\app.ico — multi-size icon with a "convert" motif.
# Small sizes are stored as 32-bpp BMP entries (max compatibility), 128/256 as PNG.
$ErrorActionPreference = 'Stop'
Add-Type -AssemblyName System.Drawing

$outDir = Split-Path -Parent $MyInvocation.MyCommand.Path
$icoPath = Join-Path $outDir 'app.ico'
$sizes = @(16, 24, 32, 48, 64, 128, 256)

function New-IconPng([int]$px) {
    $s = $px / 256.0
    $bmp = New-Object System.Drawing.Bitmap($px, $px)
    $g = [System.Drawing.Graphics]::FromImage($bmp)
    $g.SmoothingMode = 'AntiAlias'
    $g.PixelOffsetMode = 'HighQuality'
    $g.InterpolationMode = 'HighQualityBicubic'

    # rounded-square gradient background
    $rect = New-Object System.Drawing.Rectangle(0, 0, $px, $px)
    $radius = [int](58 * $s)
    $path = New-Object System.Drawing.Drawing2D.GraphicsPath
    $d = $radius * 2
    $path.AddArc(0, 0, $d, $d, 180, 90)
    $path.AddArc($px - $d, 0, $d, $d, 270, 90)
    $path.AddArc($px - $d, $px - $d, $d, $d, 0, 90)
    $path.AddArc(0, $px - $d, $d, $d, 90, 90)
    $path.CloseFigure()
    $brush = New-Object System.Drawing.Drawing2D.LinearGradientBrush($rect,
        [System.Drawing.Color]::FromArgb(91, 120, 240),    # #5B78F0
        [System.Drawing.Color]::FromArgb(139, 92, 246),    # #8B5CF6
        35)
    $g.FillPath($brush, $path)

    # two opposing arrows (convert motif)
    $white = [System.Drawing.Brushes]::White
    $penW = [Math]::Max(2, [int](30 * $s))
    $pen = New-Object System.Drawing.Pen($white, $penW)
    $pen.StartCap = 'Round'; $pen.EndCap = 'Round'
    $pen.Alignment = 'Center'

    # top arrow → : bar + head
    $y1 = 98 * $s
    $g.DrawLine($pen, (70 * $s), $y1, (146 * $s), $y1)
    $head1 = @(
        (New-Object System.Drawing.PointF((140 * $s), (72 * $s))),
        (New-Object System.Drawing.PointF((140 * $s), (124 * $s))),
        (New-Object System.Drawing.PointF((194 * $s), $y1)))
    $g.FillPolygon($white, $head1)

    # bottom arrow ← : bar + head
    $y2 = 158 * $s
    $g.DrawLine($pen, (186 * $s), $y2, (110 * $s), $y2)
    $head2 = @(
        (New-Object System.Drawing.PointF((116 * $s), (132 * $s))),
        (New-Object System.Drawing.PointF((116 * $s), (184 * $s))),
        (New-Object System.Drawing.PointF((62 * $s), $y2)))
    $g.FillPolygon($white, $head2)

    $g.Dispose()
    return $bmp
}

function Get-PngBytes([System.Drawing.Bitmap]$bmp) {
    $ms = New-Object System.IO.MemoryStream
    $bmp.Save($ms, [System.Drawing.Imaging.ImageFormat]::Png)
    return ,$ms.ToArray()
}

function Get-IcoBmpBytes([System.Drawing.Bitmap]$bmp) {
    # ICO " BMP" entry: BITMAPINFOHEADER + BGRA XOR mask (bottom-up) + AND mask (all zero)
    $w = $bmp.Width; $h = $bmp.Height
    $rect = New-Object System.Drawing.Rectangle(0, 0, $w, $h)
    $data = $bmp.LockBits($rect, 'ReadOnly', [System.Drawing.Imaging.PixelFormat]::Format32bppArgb)
    $xor = New-Object byte[] ($data.Stride * $h)
    [System.Runtime.InteropServices.Marshal]::Copy($data.Scan0, $xor, 0, $xor.Length)
    $bmp.UnlockBits($data)

    # stride == w*4 for 32bpp; rows are stored top-down in memory — flip for bottom-up BMP
    $rowLen = $w * 4
    $flipped = New-Object byte[] ($rowLen * $h)
    for ($r = 0; $r -lt $h; $r++) {
        [Array]::Copy($xor, $r * $data.Stride, $flipped, ($h - 1 - $r) * $rowLen, $rowLen)
    }
    # convert ARGB memory layout to BGRA with straight alpha (GDI+ 32bppArgb is already BGRA in memory)
    $maskStride = [int][Math]::Ceiling($w / 32.0) * 4
    $andMask = New-Object byte[] ($maskStride * $h)   # zeros — alpha channel carries transparency

    $header = New-Object byte[] 40
    $bw = [System.BitConverter]::GetBytes([uint32]40);  [Array]::Copy($bw, 0, $header, 0, 4)
    $bw = [System.BitConverter]::GetBytes([int32]$w);   [Array]::Copy($bw, 0, $header, 4, 4)
    $bw = [System.BitConverter]::GetBytes([int32]($h * 2)); [Array]::Copy($bw, 0, $header, 8, 4)
    $bw = [System.BitConverter]::GetBytes([uint16]1);   [Array]::Copy($bw, 0, $header, 12, 2)
    $bw = [System.BitConverter]::GetBytes([uint16]32);  [Array]::Copy($bw, 0, $header, 14, 2)
    $bw = [System.BitConverter]::GetBytes([uint32]($flipped.Length + $andMask.Length)); [Array]::Copy($bw, 0, $header, 20, 4)

    $ms = New-Object System.IO.MemoryStream
    $ms.Write($header, 0, 40)
    $ms.Write($flipped, 0, $flipped.Length)
    $ms.Write($andMask, 0, $andMask.Length)
    return ,$ms.ToArray()
}

# ---- assemble ICO ----
$entries = @()
foreach ($px in $sizes) {
    $bmp = New-IconPng $px
    if ($px -ge 128) { $blob = Get-PngBytes $bmp; $isPng = $true }
    else             { $blob = Get-IcoBmpBytes $bmp; $isPng = $false }
    $entries += ,@($px, $blob, $isPng)
    $bmp.Dispose()
}

$fs = [System.IO.File]::Create($icoPath)
$bw2 = New-Object System.IO.BinaryWriter($fs)
# ICONDIR: reserved=0, type=1, count
$bw2.Write([uint16]0); $bw2.Write([uint16]1); $bw2.Write([uint16]$entries.Count)
$offset = 6 + 16 * $entries.Count
foreach ($e in $entries) {
    $px = $e[0]; $blob = $e[1]
    $dim = if ($px -ge 256) { 0 } else { $px }
    $bw2.Write([byte]$dim)          # width
    $bw2.Write([byte]$dim)          # height
    $bw2.Write([byte]0)             # palette
    $bw2.Write([byte]0)             # reserved
    $bw2.Write([uint16]1)           # planes
    $bw2.Write([uint16]32)          # bpp
    $bw2.Write([uint32]$blob.Length)
    $bw2.Write([uint32]$offset)
    $offset += $blob.Length
}
foreach ($e in $entries) { $bw2.Write([byte[]]$e[1]) }
$bw2.Flush(); $bw2.Close()

# validate
$check = New-Object System.Drawing.Icon($icoPath)
Write-Host ("app.ico written ({0} bytes), sizes: {1}, loads OK, primary {2}x{3}" -f (Get-Item $icoPath).Length, ($sizes -join ','), $check.Size.Width, $check.Size.Height)
