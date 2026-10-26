# 노란 포스트잇 아이콘(끄적.ico)을 그린다. 크기별 PNG를 ICO 하나에 담는다.
param([string]$Out)
Add-Type -AssemblyName System.Drawing

function Draw([int]$s) {
    $bmp = New-Object System.Drawing.Bitmap $s, $s
    $g = [System.Drawing.Graphics]::FromImage($bmp)
    $g.SmoothingMode = 'AntiAlias'
    $m = [math]::Max(1, [int]($s * 0.08)); $w = $s - 2 * $m; $fold = [int]($w * 0.28)
    $pts = @(
        (New-Object System.Drawing.Point $m, $m),
        (New-Object System.Drawing.Point ($m + $w), $m),
        (New-Object System.Drawing.Point ($m + $w), ($m + $w - $fold)),
        (New-Object System.Drawing.Point ($m + $w - $fold), ($m + $w)),
        (New-Object System.Drawing.Point $m, ($m + $w)))
    $g.FillPolygon((New-Object System.Drawing.SolidBrush ([System.Drawing.Color]::FromArgb(255, 214, 79))), $pts)
    $corner = @(
        (New-Object System.Drawing.Point ($m + $w), ($m + $w - $fold)),
        (New-Object System.Drawing.Point ($m + $w - $fold), ($m + $w - $fold)),
        (New-Object System.Drawing.Point ($m + $w - $fold), ($m + $w)))
    $g.FillPolygon((New-Object System.Drawing.SolidBrush ([System.Drawing.Color]::FromArgb(224, 170, 40))), $corner)
    if ($s -ge 24) {
        $pen = New-Object System.Drawing.Pen ([System.Drawing.Color]::FromArgb(150, 110, 20)), ([math]::Max(1, $s / 20))
        foreach ($k in 1..3) {
            $y = $m + [int]($w * (0.18 + 0.2 * $k))
            $x2 = if ($k -eq 3) { $m + $w - $fold - [int]($w * 0.08) } else { $m + $w - [int]($w * 0.16) }
            $g.DrawLine($pen, $m + [int]($w * 0.16), $y, $x2, $y)
        }
    }
    $g.Dispose()
    $ms = New-Object System.IO.MemoryStream
    $bmp.Save($ms, [System.Drawing.Imaging.ImageFormat]::Png)
    $bmp.Dispose()
    return ,$ms.ToArray()
}

$sizes = 16, 24, 32, 48, 256
$pngs = foreach ($s in $sizes) { ,(Draw $s) }
$fs = [System.IO.File]::Create($Out)
$bw = New-Object System.IO.BinaryWriter $fs
$bw.Write([uint16]0); $bw.Write([uint16]1); $bw.Write([uint16]$sizes.Count)
$offset = 6 + 16 * $sizes.Count
for ($i = 0; $i -lt $sizes.Count; $i++) {
    $s = $sizes[$i]; $d = if ($s -ge 256) { 0 } else { $s }
    $bw.Write([byte]$d); $bw.Write([byte]$d); $bw.Write([byte]0); $bw.Write([byte]0)
    $bw.Write([uint16]1); $bw.Write([uint16]32)
    $bw.Write([uint32]$pngs[$i].Length); $bw.Write([uint32]$offset)
    $offset += $pngs[$i].Length
}
foreach ($p in $pngs) { $bw.Write($p) }
$bw.Close()
