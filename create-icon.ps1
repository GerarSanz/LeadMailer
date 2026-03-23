#Requires -Version 5.1
param(
    [string]$SourcePng = "app-icon.png"
)

Set-Location $PSScriptRoot
Add-Type -AssemblyName System.Drawing

$sourcePath = Join-Path $PSScriptRoot $SourcePng
$destDir    = Join-Path $PSScriptRoot "LeadMailer\Resources"
$destIco    = Join-Path $destDir "app.ico"

New-Item -ItemType Directory -Path $destDir -Force | Out-Null

$sizes  = @(256, 128, 64, 48, 32, 16)
$images = @()

# ?? Fuente: PNG personalizado o icono generado automaticamente ????????????????
if (Test-Path $sourcePath) {
    Write-Host "  Fuente : $sourcePath" -ForegroundColor DarkGray
    Write-Host "  Destino: $destIco" -ForegroundColor DarkGray
    Write-Host ""

    $srcBitmap = [System.Drawing.Bitmap]::new($sourcePath)

    foreach ($size in $sizes) {
        $bmp = [System.Drawing.Bitmap]::new($size, $size,
               [System.Drawing.Imaging.PixelFormat]::Format32bppArgb)
        $g   = [System.Drawing.Graphics]::FromImage($bmp)
        $g.InterpolationMode  = [System.Drawing.Drawing2D.InterpolationMode]::HighQualityBicubic
        $g.SmoothingMode      = [System.Drawing.Drawing2D.SmoothingMode]::HighQuality
        $g.PixelOffsetMode    = [System.Drawing.Drawing2D.PixelOffsetMode]::HighQuality
        $g.CompositingQuality = [System.Drawing.Drawing2D.CompositingQuality]::HighQuality
        $g.DrawImage($srcBitmap, 0, 0, $size, $size)
        $g.Dispose()

        $ms = [System.IO.MemoryStream]::new()
        $bmp.Save($ms, [System.Drawing.Imaging.ImageFormat]::Png)
        $images += , $ms.ToArray()
        $ms.Dispose()
        $bmp.Dispose()

        Write-Host "  [OK] ${size}x${size} px  (desde PNG)" -ForegroundColor DarkGray
    }
    $srcBitmap.Dispose()

} else {
    Write-Host "  [!] app-icon.png no encontrado. Generando icono por defecto..." -ForegroundColor Yellow
    Write-Host "      (coloca app-icon.png en la raiz para usar tu propio icono)" -ForegroundColor DarkGray
    Write-Host ""

    # Colores del icono por defecto: fondo azul oscuro + letras LM en blanco
    $bgColor   = [System.Drawing.Color]::FromArgb(255, 30, 80, 180)
    $textColor = [System.Drawing.Color]::White

    foreach ($size in $sizes) {
        $bmp = [System.Drawing.Bitmap]::new($size, $size,
               [System.Drawing.Imaging.PixelFormat]::Format32bppArgb)
        $g   = [System.Drawing.Graphics]::FromImage($bmp)
        $g.SmoothingMode     = [System.Drawing.Drawing2D.SmoothingMode]::HighQuality
        $g.TextRenderingHint = [System.Drawing.Text.TextRenderingHint]::AntiAliasGridFit

        # Fondo redondeado
        $brush = [System.Drawing.SolidBrush]::new($bgColor)
        $r     = [int]($size * 0.15)
        $rect  = [System.Drawing.Rectangle]::new(0, 0, $size - 1, $size - 1)
        $path  = [System.Drawing.Drawing2D.GraphicsPath]::new()
        $path.AddArc($rect.X, $rect.Y, $r * 2, $r * 2, 180, 90)
        $path.AddArc($rect.Right - $r * 2, $rect.Y, $r * 2, $r * 2, 270, 90)
        $path.AddArc($rect.Right - $r * 2, $rect.Bottom - $r * 2, $r * 2, $r * 2, 0, 90)
        $path.AddArc($rect.X, $rect.Bottom - $r * 2, $r * 2, $r * 2, 90, 90)
        $path.CloseFigure()
        $g.FillPath($brush, $path)
        $brush.Dispose()
        $path.Dispose()

        # Texto "LM"
        if ($size -ge 32) {
            $fontSize   = [float]($size * 0.38)
            $font       = [System.Drawing.Font]::new("Segoe UI", $fontSize, [System.Drawing.FontStyle]::Bold)
            $textBrush  = [System.Drawing.SolidBrush]::new($textColor)
            $text       = "LM"
            $textSize   = $g.MeasureString($text, $font)
            $x          = ($size - $textSize.Width)  / 2
            $y          = ($size - $textSize.Height) / 2
            $g.DrawString($text, $font, $textBrush, $x, $y)
            $font.Dispose()
            $textBrush.Dispose()
        }

        $g.Dispose()

        $ms = [System.IO.MemoryStream]::new()
        $bmp.Save($ms, [System.Drawing.Imaging.ImageFormat]::Png)
        $images += , $ms.ToArray()
        $ms.Dispose()
        $bmp.Dispose()

        Write-Host "  [OK] ${size}x${size} px  (generado)" -ForegroundColor DarkGray
    }
}

# ?? Construir archivo .ico ????????????????????????????????????????????????????
$mem    = [System.IO.MemoryStream]::new()
$writer = [System.IO.BinaryWriter]::new($mem)

$writer.Write([uint16]0)
$writer.Write([uint16]1)
$writer.Write([uint16]$sizes.Count)

$dataOffset = 6 + 16 * $sizes.Count
for ($i = 0; $i -lt $sizes.Count; $i++) {
    $sz = $sizes[$i]
    $writer.Write([byte]$(if ($sz -eq 256) { 0 } else { $sz }))
    $writer.Write([byte]$(if ($sz -eq 256) { 0 } else { $sz }))
    $writer.Write([byte]0)
    $writer.Write([byte]0)
    $writer.Write([uint16]1)
    $writer.Write([uint16]32)
    $writer.Write([uint32]$images[$i].Length)
    $writer.Write([uint32]$dataOffset)
    $dataOffset += $images[$i].Length
}
foreach ($img in $images) { $writer.Write($img) }

$writer.Flush()
[System.IO.File]::WriteAllBytes($destIco, $mem.ToArray())
$writer.Dispose()
$mem.Dispose()

$kb = [math]::Round((Get-Item $destIco).Length / 1KB, 1)
Write-Host ""
Write-Host "  [OK] app.ico generado ($($sizes.Count) tamanos - $kb KB)" -ForegroundColor Green
Write-Host "       $destIco" -ForegroundColor DarkGray
Write-Host ""