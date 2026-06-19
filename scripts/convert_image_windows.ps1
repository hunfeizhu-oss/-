param(
    [Parameter(Mandatory = $true)]
    [string]$SourcePath,

    [Parameter(Mandatory = $true)]
    [string]$TargetPath,

    [Parameter(Mandatory = $true)]
    [int]$TargetWidth,

    [Parameter(Mandatory = $true)]
    [int]$TargetHeight
)

$ErrorActionPreference = "Stop"
Add-Type -AssemblyName System.Drawing
$image = $null
$bitmap = $null
$graphics = $null

try {
    $image = [System.Drawing.Image]::FromFile($SourcePath)
    $sourceRatio = $image.Width / $image.Height
    $targetRatio = $TargetWidth / $TargetHeight

    if ($sourceRatio -gt $targetRatio) {
        $drawHeight = $TargetHeight
        $drawWidth = [int][Math]::Ceiling($TargetHeight * $sourceRatio)
    }
    else {
        $drawWidth = $TargetWidth
        $drawHeight = [int][Math]::Ceiling($TargetWidth / $sourceRatio)
    }

    $offsetX = [int][Math]::Floor(($TargetWidth - $drawWidth) / 2)
    $offsetY = [int][Math]::Floor(($TargetHeight - $drawHeight) / 2)
    $bitmap = New-Object System.Drawing.Bitmap($TargetWidth, $TargetHeight, [System.Drawing.Imaging.PixelFormat]::Format32bppArgb)
    $graphics = [System.Drawing.Graphics]::FromImage($bitmap)
    $graphics.Clear([System.Drawing.Color]::Transparent)
    $graphics.InterpolationMode = [System.Drawing.Drawing2D.InterpolationMode]::HighQualityBicubic
    $graphics.PixelOffsetMode = [System.Drawing.Drawing2D.PixelOffsetMode]::HighQuality
    $graphics.SmoothingMode = [System.Drawing.Drawing2D.SmoothingMode]::HighQuality
    $graphics.DrawImage($image, $offsetX, $offsetY, $drawWidth, $drawHeight)
    $bitmap.Save($TargetPath, [System.Drawing.Imaging.ImageFormat]::Png)
}
finally {
    if ($null -ne $graphics) {
        $graphics.Dispose()
    }
    if ($null -ne $bitmap) {
        $bitmap.Dispose()
    }
    if ($null -ne $image) {
        $image.Dispose()
    }
}
