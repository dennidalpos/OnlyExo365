#Requires -Version 7.0

[CmdletBinding()]
param(
    [string]$OutputIcoPath = "src/ExchangeAdmin.Presentation/Assets/AppIcon.ico",
    [string]$OutputPngPath = "src/ExchangeAdmin.Presentation/Assets/AppIcon.png"
)

Set-StrictMode -Version Latest
$ErrorActionPreference = "Stop"

Add-Type -AssemblyName System.Drawing
Add-Type @"
using System;
using System.Runtime.InteropServices;

public static class NativeIconMethods
{
    [DllImport("user32.dll", CharSet = CharSet.Auto)]
    public static extern bool DestroyIcon(IntPtr handle);
}
"@

$scriptDirectory = Split-Path -Parent $MyInvocation.MyCommand.Path
$repositoryRoot = Split-Path -Parent $scriptDirectory

function Resolve-RepoPath {
    param([string]$PathValue)

    if ([System.IO.Path]::IsPathRooted($PathValue)) {
        return [System.IO.Path]::GetFullPath($PathValue)
    }

    return [System.IO.Path]::GetFullPath((Join-Path $repositoryRoot $PathValue))
}

function New-RoundedRectanglePath {
    param(
        [System.Drawing.RectangleF]$Bounds,
        [float]$Radius
    )

    $diameter = $Radius * 2
    $path = [System.Drawing.Drawing2D.GraphicsPath]::new()
    $path.AddArc($Bounds.X, $Bounds.Y, $diameter, $diameter, 180, 90)
    $path.AddArc($Bounds.Right - $diameter, $Bounds.Y, $diameter, $diameter, 270, 90)
    $path.AddArc($Bounds.Right - $diameter, $Bounds.Bottom - $diameter, $diameter, $diameter, 0, 90)
    $path.AddArc($Bounds.X, $Bounds.Bottom - $diameter, $diameter, $diameter, 90, 90)
    $path.CloseFigure()
    return $path
}

function New-IconBitmap {
    param([int]$Size)

    $bitmap = [System.Drawing.Bitmap]::new($Size, $Size, [System.Drawing.Imaging.PixelFormat]::Format32bppArgb)
    $graphics = [System.Drawing.Graphics]::FromImage($bitmap)

    try {
        $graphics.SmoothingMode = [System.Drawing.Drawing2D.SmoothingMode]::AntiAlias
        $graphics.InterpolationMode = [System.Drawing.Drawing2D.InterpolationMode]::HighQualityBicubic
        $graphics.PixelOffsetMode = [System.Drawing.Drawing2D.PixelOffsetMode]::HighQuality
        $graphics.Clear([System.Drawing.Color]::Transparent)

        $padding = $Size * 0.08
        $cardBounds = [System.Drawing.RectangleF]::new($padding, $padding, $Size - (2 * $padding), $Size - (2 * $padding))
        $radius = $Size * 0.18

        $backgroundPath = New-RoundedRectanglePath -Bounds $cardBounds -Radius $radius
        try {
            $backgroundBrush = [System.Drawing.Drawing2D.LinearGradientBrush]::new(
                [System.Drawing.PointF]::new(0, 0),
                [System.Drawing.PointF]::new($Size, $Size),
                [System.Drawing.Color]::FromArgb(255, 20, 62, 117),
                [System.Drawing.Color]::FromArgb(255, 7, 126, 167))
            try {
                $graphics.FillPath($backgroundBrush, $backgroundPath)
            }
            finally {
                $backgroundBrush.Dispose()
            }

            $accentBrush = [System.Drawing.SolidBrush]::new([System.Drawing.Color]::FromArgb(72, 255, 255, 255))
            try {
                $graphics.FillEllipse(
                    $accentBrush,
                    [System.Drawing.RectangleF]::new($Size * 0.58, $Size * 0.12, $Size * 0.26, $Size * 0.18))
            }
            finally {
                $accentBrush.Dispose()
            }
        }
        finally {
            $backgroundPath.Dispose()
        }

        $envelopeBounds = [System.Drawing.RectangleF]::new($Size * 0.18, $Size * 0.28, $Size * 0.64, $Size * 0.43)
        $envelopePath = New-RoundedRectanglePath -Bounds $envelopeBounds -Radius ($Size * 0.05)
        try {
            $envelopeBrush = [System.Drawing.SolidBrush]::new([System.Drawing.Color]::FromArgb(245, 250, 252, 255))
            $envelopePen = [System.Drawing.Pen]::new([System.Drawing.Color]::FromArgb(255, 13, 84, 128), [Math]::Max(2, $Size * 0.028))
            try {
                $graphics.FillPath($envelopeBrush, $envelopePath)
                $graphics.DrawPath($envelopePen, $envelopePath)

                $flapPen = [System.Drawing.Pen]::new([System.Drawing.Color]::FromArgb(255, 13, 84, 128), [Math]::Max(2, $Size * 0.028))
                try {
                    $left = $envelopeBounds.Left
                    $right = $envelopeBounds.Right
                    $top = $envelopeBounds.Top
                    $bottom = $envelopeBounds.Bottom
                    $centerX = $envelopeBounds.Left + ($envelopeBounds.Width / 2)
                    $centerY = $envelopeBounds.Top + ($envelopeBounds.Height * 0.53)

                    $graphics.DrawLine($flapPen, $left, $top, $centerX, $centerY)
                    $graphics.DrawLine($flapPen, $right, $top, $centerX, $centerY)
                    $graphics.DrawLine($flapPen, $left, $bottom, $centerX, $centerY)
                    $graphics.DrawLine($flapPen, $right, $bottom, $centerX, $centerY)
                }
                finally {
                    $flapPen.Dispose()
                }
            }
            finally {
                $envelopePen.Dispose()
                $envelopeBrush.Dispose()
            }
        }
        finally {
            $envelopePath.Dispose()
        }

        $shieldPath = [System.Drawing.Drawing2D.GraphicsPath]::new()
        try {
            $shieldTop = $Size * 0.54
            $shieldLeft = $Size * 0.57
            $shieldWidth = $Size * 0.22
            $shieldHeight = $Size * 0.24
            $shieldPath.AddPolygon([System.Drawing.PointF[]]@(
                    [System.Drawing.PointF]::new($shieldLeft, $shieldTop),
                    [System.Drawing.PointF]::new($shieldLeft + $shieldWidth, $shieldTop),
                    [System.Drawing.PointF]::new($shieldLeft + ($shieldWidth * 0.92), $shieldTop + ($shieldHeight * 0.55)),
                    [System.Drawing.PointF]::new($shieldLeft + ($shieldWidth * 0.5), $shieldTop + $shieldHeight),
                    [System.Drawing.PointF]::new($shieldLeft + ($shieldWidth * 0.08), $shieldTop + ($shieldHeight * 0.55))
                ))

            $shieldBrush = [System.Drawing.Drawing2D.LinearGradientBrush]::new(
                [System.Drawing.PointF]::new($shieldLeft, $shieldTop),
                [System.Drawing.PointF]::new($shieldLeft, $shieldTop + $shieldHeight),
                [System.Drawing.Color]::FromArgb(255, 33, 210, 183),
                [System.Drawing.Color]::FromArgb(255, 17, 154, 142))
            $shieldPen = [System.Drawing.Pen]::new([System.Drawing.Color]::FromArgb(255, 235, 255, 250), [Math]::Max(1.5, $Size * 0.02))
            try {
                $graphics.FillPath($shieldBrush, $shieldPath)
                $graphics.DrawPath($shieldPen, $shieldPath)

                $checkPen = [System.Drawing.Pen]::new([System.Drawing.Color]::White, [Math]::Max(2, $Size * 0.03))
                $checkPen.StartCap = [System.Drawing.Drawing2D.LineCap]::Round
                $checkPen.EndCap = [System.Drawing.Drawing2D.LineCap]::Round
                try {
                    $graphics.DrawLines($checkPen, [System.Drawing.PointF[]]@(
                            [System.Drawing.PointF]::new($shieldLeft + ($shieldWidth * 0.26), $shieldTop + ($shieldHeight * 0.58)),
                            [System.Drawing.PointF]::new($shieldLeft + ($shieldWidth * 0.44), $shieldTop + ($shieldHeight * 0.76)),
                            [System.Drawing.PointF]::new($shieldLeft + ($shieldWidth * 0.76), $shieldTop + ($shieldHeight * 0.38))
                        ))
                }
                finally {
                    $checkPen.Dispose()
                }
            }
            finally {
                $shieldPen.Dispose()
                $shieldBrush.Dispose()
            }
        }
        finally {
            $shieldPath.Dispose()
        }

        return $bitmap
    }
    catch {
        $graphics.Dispose()
        $bitmap.Dispose()
        throw
    }
    finally {
        if ($null -ne $graphics) {
            $graphics.Dispose()
        }
    }
}

$resolvedIcoPath = Resolve-RepoPath -PathValue $OutputIcoPath
$resolvedPngPath = Resolve-RepoPath -PathValue $OutputPngPath

New-Item -Path (Split-Path -Parent $resolvedIcoPath) -ItemType Directory -Force | Out-Null
New-Item -Path (Split-Path -Parent $resolvedPngPath) -ItemType Directory -Force | Out-Null

$bitmap = New-IconBitmap -Size 256
try {
    $bitmap.Save($resolvedPngPath, [System.Drawing.Imaging.ImageFormat]::Png)

    $iconHandle = $bitmap.GetHicon()
    try {
        $icon = [System.Drawing.Icon]::FromHandle($iconHandle)
        try {
            $iconStream = [System.IO.File]::Open($resolvedIcoPath, [System.IO.FileMode]::Create, [System.IO.FileAccess]::Write)
            try {
                $icon.Save($iconStream)
            }
            finally {
                $iconStream.Dispose()
            }
        }
        finally {
            $icon.Dispose()
        }
    }
    finally {
        [NativeIconMethods]::DestroyIcon($iconHandle) | Out-Null
    }
}
finally {
    $bitmap.Dispose()
}

Write-Host "Generated icon assets:" -ForegroundColor Cyan
Write-Host "  ICO: $resolvedIcoPath" -ForegroundColor Green
Write-Host "  PNG: $resolvedPngPath" -ForegroundColor Green
