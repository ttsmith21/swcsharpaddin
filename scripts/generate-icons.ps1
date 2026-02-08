<#
.SYNOPSIS
    Extends toolbar strip PNGs from 4 icons to 6 icons.
.DESCRIPTION
    Copies existing 4 icons (play, warning, checkmark, arrow) to positions 0-3,
    then draws:
      - Icon 4: Teal document page with folded corner (for Analyze Drawing)
      - Icon 5: Gray gear/cog (for Settings)
    Overwrites strip PNGs at all 6 resolutions (20, 32, 40, 64, 96, 128).
#>

param(
    [string]$IconDir = (Join-Path $PSScriptRoot "..\Icons")
)

Add-Type -AssemblyName System.Drawing

$sizes = @(20, 32, 40, 64, 96, 128)

foreach ($size in $sizes) {
    $filePath = Join-Path $IconDir "ToolbarStrip_${size}.png"

    if (-not (Test-Path $filePath)) {
        Write-Warning "Missing: $filePath - skipping"
        continue
    }

    $oldImg = [System.Drawing.Image]::FromFile($filePath)
    $oldWidth = $oldImg.Width
    $iconCount = [int]($oldWidth / $size)
    $newWidth = $size * 6

    # Create new wider strip
    $newImg = New-Object System.Drawing.Bitmap($newWidth, $size, [System.Drawing.Imaging.PixelFormat]::Format32bppArgb)
    $g = [System.Drawing.Graphics]::FromImage($newImg)
    $g.SmoothingMode = [System.Drawing.Drawing2D.SmoothingMode]::AntiAlias
    $g.InterpolationMode = [System.Drawing.Drawing2D.InterpolationMode]::HighQualityBicubic
    $g.PixelOffsetMode = [System.Drawing.Drawing2D.PixelOffsetMode]::HighQuality

    # Clear to transparent
    $g.Clear([System.Drawing.Color]::Transparent)

    # Copy existing icons (positions 0..iconCount-1)
    $srcRect = New-Object System.Drawing.Rectangle(0, 0, $oldWidth, $size)
    $dstRect = New-Object System.Drawing.Rectangle(0, 0, $oldWidth, $size)
    $g.DrawImage($oldImg, $dstRect, $srcRect, [System.Drawing.GraphicsUnit]::Pixel)
    $oldImg.Dispose()

    # --- Icon 4: Teal document with folded corner ---
    $x4 = $size * 4
    $margin = [Math]::Max(2, [int]($size * 0.15))
    $foldSize = [Math]::Max(3, [int]($size * 0.22))
    $docColor = [System.Drawing.Color]::FromArgb(255, 0, 150, 136)  # Teal 500
    $docColorLight = [System.Drawing.Color]::FromArgb(255, 128, 203, 196)  # Lighter teal for fold
    $lineColor = [System.Drawing.Color]::FromArgb(255, 0, 105, 92)  # Dark teal for lines

    # Document body (rectangle with top-right corner cut)
    $docLeft = $x4 + $margin
    $docTop = $margin
    $docRight = $x4 + $size - $margin
    $docBottom = $size - $margin
    $docW = $docRight - $docLeft
    $docH = $docBottom - $docTop

    # Build the page shape path (with folded corner)
    $path = New-Object System.Drawing.Drawing2D.GraphicsPath
    $path.AddLine($docLeft, $docTop, ($docRight - $foldSize), $docTop)
    $path.AddLine(($docRight - $foldSize), $docTop, $docRight, ($docTop + $foldSize))
    $path.AddLine($docRight, ($docTop + $foldSize), $docRight, $docBottom)
    $path.AddLine($docRight, $docBottom, $docLeft, $docBottom)
    $path.CloseFigure()

    $brushDoc = New-Object System.Drawing.SolidBrush($docColor)
    $g.FillPath($brushDoc, $path)
    $brushDoc.Dispose()

    # Draw the fold triangle
    $foldPath = New-Object System.Drawing.Drawing2D.GraphicsPath
    $foldPath.AddLine(($docRight - $foldSize), $docTop, ($docRight - $foldSize), ($docTop + $foldSize))
    $foldPath.AddLine(($docRight - $foldSize), ($docTop + $foldSize), $docRight, ($docTop + $foldSize))
    $foldPath.CloseFigure()
    $brushFold = New-Object System.Drawing.SolidBrush($docColorLight)
    $g.FillPath($brushFold, $foldPath)
    $brushFold.Dispose()

    # Draw text lines on the page
    $linePen = New-Object System.Drawing.Pen($lineColor, [Math]::Max(1, [int]($size * 0.06)))
    $lineLeft = $docLeft + [int]($docW * 0.15)
    $lineRight = $docRight - [int]($docW * 0.15)
    $lineSpacing = [int]($docH * 0.22)
    $lineStart = $docTop + [int]($docH * 0.35)
    for ($li = 0; $li -lt 3; $li++) {
        $ly = $lineStart + ($li * $lineSpacing)
        $lr = if ($li -eq 2) { $lineLeft + [int](($lineRight - $lineLeft) * 0.6) } else { $lineRight }
        if ($ly -lt ($docBottom - $margin/2)) {
            $g.DrawLine($linePen, $lineLeft, $ly, $lr, $ly)
        }
    }
    $linePen.Dispose()
    $path.Dispose()
    $foldPath.Dispose()

    # --- Icon 5: Gray gear/cog ---
    $x5 = $size * 5
    $cx = $x5 + ($size / 2.0)
    $cy = $size / 2.0
    $outerR = ($size / 2.0) - $margin
    $innerR = $outerR * 0.62
    $holeR = $outerR * 0.3
    $teeth = 8
    $toothDepth = $outerR - $innerR
    $toothHalfAngle = [Math]::PI / $teeth * 0.45

    $gearColor = [System.Drawing.Color]::FromArgb(255, 117, 117, 117)  # Gray 600

    # Build gear path
    $gearPath = New-Object System.Drawing.Drawing2D.GraphicsPath
    $points = New-Object System.Collections.ArrayList
    for ($t = 0; $t -lt $teeth; $t++) {
        $angle = (2 * [Math]::PI * $t / $teeth) - ([Math]::PI / 2)

        # Inner start
        $a1 = $angle - $toothHalfAngle * 1.5
        [void]$points.Add([System.Drawing.PointF]::new(
            [float]($cx + $innerR * [Math]::Cos($a1)),
            [float]($cy + $innerR * [Math]::Sin($a1))
        ))
        # Outer start
        $a2 = $angle - $toothHalfAngle
        [void]$points.Add([System.Drawing.PointF]::new(
            [float]($cx + $outerR * [Math]::Cos($a2)),
            [float]($cy + $outerR * [Math]::Sin($a2))
        ))
        # Outer end
        $a3 = $angle + $toothHalfAngle
        [void]$points.Add([System.Drawing.PointF]::new(
            [float]($cx + $outerR * [Math]::Cos($a3)),
            [float]($cy + $outerR * [Math]::Sin($a3))
        ))
        # Inner end
        $a4 = $angle + $toothHalfAngle * 1.5
        [void]$points.Add([System.Drawing.PointF]::new(
            [float]($cx + $innerR * [Math]::Cos($a4)),
            [float]($cy + $innerR * [Math]::Sin($a4))
        ))
    }
    $gearPath.AddPolygon($points.ToArray([System.Drawing.PointF]))

    # Cut out center hole
    $holePath = New-Object System.Drawing.Drawing2D.GraphicsPath
    $holePath.AddEllipse([float]($cx - $holeR), [float]($cy - $holeR), [float]($holeR * 2), [float]($holeR * 2))

    # Combine: gear minus hole
    $combinedPath = New-Object System.Drawing.Drawing2D.GraphicsPath
    $combinedPath.AddPath($gearPath, $false)
    $combinedPath.AddPath($holePath, $false)
    $combinedPath.FillMode = [System.Drawing.Drawing2D.FillMode]::Alternate

    $brushGear = New-Object System.Drawing.SolidBrush($gearColor)
    $g.FillPath($brushGear, $combinedPath)
    $brushGear.Dispose()
    $gearPath.Dispose()
    $holePath.Dispose()
    $combinedPath.Dispose()

    # Save
    $g.Dispose()

    # Delete old file and save new one
    [System.IO.File]::Delete($filePath)
    $newImg.Save($filePath, [System.Drawing.Imaging.ImageFormat]::Png)
    $newImg.Dispose()

    Write-Host "Updated: $filePath (${size}x${size}, 6 icons)"
}

Write-Host "`nIcon generation complete. Strip PNGs now contain 6 icons."
