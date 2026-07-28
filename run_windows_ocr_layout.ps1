param(
    [Parameter(Mandatory = $true)]
    [string]$InputDirectory,
    [Parameter(Mandatory = $true)]
    [string]$OutputJson,
    [string]$LanguageTag = "ko",
    [switch]$Resume
)

$ErrorActionPreference = "Stop"
Add-Type -AssemblyName System.Runtime.WindowsRuntime
Add-Type -AssemblyName System.Drawing
[void][Windows.Storage.StorageFile, Windows.Storage, ContentType = WindowsRuntime]
[void][Windows.Storage.FileAccessMode, Windows.Storage, ContentType = WindowsRuntime]
[void][Windows.Graphics.Imaging.BitmapDecoder, Windows.Graphics.Imaging, ContentType = WindowsRuntime]
[void][Windows.Graphics.Imaging.BitmapTransform, Windows.Graphics.Imaging, ContentType = WindowsRuntime]
[void][Windows.Graphics.Imaging.BitmapBounds, Windows.Graphics.Imaging, ContentType = WindowsRuntime]
[void][Windows.Graphics.Imaging.BitmapPixelFormat, Windows.Graphics.Imaging, ContentType = WindowsRuntime]
[void][Windows.Graphics.Imaging.BitmapAlphaMode, Windows.Graphics.Imaging, ContentType = WindowsRuntime]
[void][Windows.Graphics.Imaging.ExifOrientationMode, Windows.Graphics.Imaging, ContentType = WindowsRuntime]
[void][Windows.Graphics.Imaging.ColorManagementMode, Windows.Graphics.Imaging, ContentType = WindowsRuntime]
[void][Windows.Media.Ocr.OcrEngine, Windows.Foundation, ContentType = WindowsRuntime]
[void][Windows.Globalization.Language, Windows.Globalization, ContentType = WindowsRuntime]

$asTaskMethod = [System.WindowsRuntimeSystemExtensions].GetMethods() |
    Where-Object {
        $_.Name -eq "AsTask" -and
        $_.IsGenericMethod -and
        $_.GetParameters().Count -eq 1 -and
        $_.GetGenericArguments().Count -eq 1 -and
        $_.ReturnType.Name -eq 'Task`1'
    } |
    Select-Object -First 1

function Wait-WinRtOperation {
    param($Operation, [Type]$ResultType, [int]$TimeoutMs = 60000)
    $task = $asTaskMethod.MakeGenericMethod($ResultType).Invoke($null, @($Operation))
    if (-not $task.Wait($TimeoutMs)) {
        throw [TimeoutException]::new("WinRT layout OCR operation exceeded ${TimeoutMs}ms")
    }
    return $task.Result
}

function Invoke-OcrSourceWithLayoutDirect {
    param([string]$Path, $Engine)

    $storageFile = Wait-WinRtOperation ([Windows.Storage.StorageFile]::GetFileFromPathAsync($Path)) ([Windows.Storage.StorageFile])
    $stream = Wait-WinRtOperation ($storageFile.OpenAsync([Windows.Storage.FileAccessMode]::Read)) ([Windows.Storage.Streams.IRandomAccessStream])
    try {
        $decoder = Wait-WinRtOperation ([Windows.Graphics.Imaging.BitmapDecoder]::CreateAsync($stream)) ([Windows.Graphics.Imaging.BitmapDecoder])
        $width = [uint32]$decoder.PixelWidth
        $height = [uint32]$decoder.PixelHeight
        $tileHeight = [uint32]2200
        $texts = @()
        $words = @()
        $tileCount = 0
        $fileStopwatch = [Diagnostics.Stopwatch]::StartNew()

        for ($top = [uint32]0; $top -lt $height; $top = [uint32]($top + $tileHeight)) {
            if ($fileStopwatch.ElapsedMilliseconds -gt 120000) {
                throw [TimeoutException]::new("Image layout OCR exceeded 120000ms")
            }
            $currentHeight = [uint32][Math]::Min([double]$tileHeight, [double]($height - $top))
            $bounds = New-Object Windows.Graphics.Imaging.BitmapBounds
            $bounds.X = [uint32]0
            $bounds.Y = $top
            $bounds.Width = $width
            $bounds.Height = $currentHeight

            $transform = New-Object Windows.Graphics.Imaging.BitmapTransform
            $transform.Bounds = $bounds
            $bitmap = Wait-WinRtOperation (
                $decoder.GetSoftwareBitmapAsync(
                    [Windows.Graphics.Imaging.BitmapPixelFormat]::Bgra8,
                    [Windows.Graphics.Imaging.BitmapAlphaMode]::Premultiplied,
                    $transform,
                    [Windows.Graphics.Imaging.ExifOrientationMode]::IgnoreExifOrientation,
                    [Windows.Graphics.Imaging.ColorManagementMode]::DoNotColorManage
                )
            ) ([Windows.Graphics.Imaging.SoftwareBitmap])
            try {
                $result = Wait-WinRtOperation ($Engine.RecognizeAsync($bitmap)) ([Windows.Media.Ocr.OcrResult])
                if ($result.Text) { $texts += $result.Text }
                $lineIndex = 0
                foreach ($line in $result.Lines) {
                    $wordIndex = 0
                    foreach ($word in $line.Words) {
                        $rect = $word.BoundingRect
                        $words += [pscustomobject]@{
                            text = $word.Text
                            x = [double]$rect.X
                            y = [double]$rect.Y + [double]$top
                            width = [double]$rect.Width
                            height = [double]$rect.Height
                            line_index = $lineIndex
                            word_index = $wordIndex
                        }
                        $wordIndex++
                    }
                    $lineIndex++
                }
                $tileCount++
            }
            finally {
                if ($null -ne $bitmap) { $bitmap.Dispose() }
            }
        }
        return [pscustomobject]@{
            text = ($texts -join "`n").Trim()
            words = $words
            tile_count = $tileCount
            original_width = [int]$width
            original_height = [int]$height
            scaled_width = [int]$width
            scaled_height = [int]$height
            scale = 1.0
        }
    }
    finally {
        if ($null -ne $stream) { $stream.Dispose() }
    }
}

function Invoke-OcrFileWithLayout {
    param([string]$Path, $Engine, [int]$OffsetY)
    $storageFile = Wait-WinRtOperation ([Windows.Storage.StorageFile]::GetFileFromPathAsync($Path)) ([Windows.Storage.StorageFile])
    $stream = Wait-WinRtOperation ($storageFile.OpenAsync([Windows.Storage.FileAccessMode]::Read)) ([Windows.Storage.Streams.IRandomAccessStream])
    try {
        $decoder = Wait-WinRtOperation ([Windows.Graphics.Imaging.BitmapDecoder]::CreateAsync($stream)) ([Windows.Graphics.Imaging.BitmapDecoder])
        $bitmap = Wait-WinRtOperation ($decoder.GetSoftwareBitmapAsync()) ([Windows.Graphics.Imaging.SoftwareBitmap])
        try {
            $result = Wait-WinRtOperation ($Engine.RecognizeAsync($bitmap)) ([Windows.Media.Ocr.OcrResult])
            $words = @()
            $lineIndex = 0
            foreach ($line in $result.Lines) {
                $wordIndex = 0
                foreach ($word in $line.Words) {
                    $rect = $word.BoundingRect
                    $words += [pscustomobject]@{
                        text = $word.Text
                        x = [double]$rect.X
                        y = [double]$rect.Y + $OffsetY
                        width = [double]$rect.Width
                        height = [double]$rect.Height
                        line_index = $lineIndex
                        word_index = $wordIndex
                    }
                    $wordIndex++
                }
                $lineIndex++
            }
            return [pscustomobject]@{
                text = $result.Text
                words = $words
            }
        }
        finally {
            if ($null -ne $bitmap) { $bitmap.Dispose() }
        }
    }
    finally {
        if ($null -ne $stream) { $stream.Dispose() }
    }
}

function Get-OcrTilesWithMetadata {
    param([string]$Path, [string]$TempDirectory)
    $image = [System.Drawing.Image]::FromFile($Path)
    try {
        $maxWidth = 2200
        $tileHeight = 2200
        $scale = [Math]::Min(1.0, $maxWidth / [double]$image.Width)
        $scaledWidth = [Math]::Max(1, [int][Math]::Round($image.Width * $scale))
        $scaledHeight = [Math]::Max(1, [int][Math]::Round($image.Height * $scale))

        $scaled = New-Object System.Drawing.Bitmap($scaledWidth, $scaledHeight)
        $graphics = [System.Drawing.Graphics]::FromImage($scaled)
        try {
            $graphics.InterpolationMode = [System.Drawing.Drawing2D.InterpolationMode]::HighQualityBicubic
            $graphics.DrawImage($image, 0, 0, $scaledWidth, $scaledHeight)
        }
        finally {
            $graphics.Dispose()
        }

        $tiles = @()
        $tileIndex = 0
        for ($top = 0; $top -lt $scaledHeight; $top += $tileHeight) {
            $height = [Math]::Min($tileHeight, $scaledHeight - $top)
            $tile = New-Object System.Drawing.Bitmap($scaledWidth, $height)
            $tileGraphics = [System.Drawing.Graphics]::FromImage($tile)
            try {
                $sourceRectangle = New-Object System.Drawing.Rectangle(0, $top, $scaledWidth, $height)
                $destinationRectangle = New-Object System.Drawing.Rectangle(0, 0, $scaledWidth, $height)
                $tileGraphics.DrawImage($scaled, $destinationRectangle, $sourceRectangle, [System.Drawing.GraphicsUnit]::Pixel)
            }
            finally {
                $tileGraphics.Dispose()
            }
            $tilePath = Join-Path $TempDirectory (([IO.Path]::GetFileNameWithoutExtension($Path)) + "_tile_" + $tileIndex + ".png")
            $tile.Save($tilePath, [System.Drawing.Imaging.ImageFormat]::Png)
            $tile.Dispose()
            $tiles += [pscustomobject]@{
                path = $tilePath
                offset_y = $top
            }
            $tileIndex++
        }
        $scaled.Dispose()
        return [pscustomobject]@{
            tiles = $tiles
            scaled_width = $scaledWidth
            scaled_height = $scaledHeight
            original_width = $image.Width
            original_height = $image.Height
            scale = $scale
        }
    }
    finally {
        $image.Dispose()
    }
}

$inputPath = (Resolve-Path -LiteralPath $InputDirectory).Path
$outputPath = [IO.Path]::GetFullPath($OutputJson)
$language = New-Object Windows.Globalization.Language($LanguageTag)
$engine = [Windows.Media.Ocr.OcrEngine]::TryCreateFromLanguage($language)
if ($null -eq $engine) {
    throw "OCR engine unavailable for language: $LanguageTag"
}

$tempDirectory = Join-Path ([IO.Path]::GetTempPath()) ("homestyle_ocr_layout_" + [Guid]::NewGuid().ToString("N"))
[void](New-Item -ItemType Directory -Path $tempDirectory)

try {
    $files = Get-ChildItem -LiteralPath $inputPath -File |
        Where-Object { $_.Extension.ToLowerInvariant() -in @(".jpg", ".jpeg", ".png", ".gif", ".bmp") } |
        Sort-Object Name

    $existingByFile = @{}
    if ($Resume -and (Test-Path -LiteralPath $outputPath)) {
        try {
            $oldResults = Get-Content -LiteralPath $outputPath -Raw -Encoding UTF8 | ConvertFrom-Json
            foreach ($oldRow in @($oldResults)) {
                if ($oldRow.status -eq "SUCCESS" -and $oldRow.file) {
                    $existingByFile[[string]$oldRow.file] = $oldRow
                }
            }
        }
        catch {
            Write-Warning "Previous layout OCR output could not be reused: $($_.Exception.Message)"
        }
    }

    $results = @()
    $reusedCount = 0
    $progressPath = $outputPath + ".progress"

    foreach ($file in $files) {
        if ($existingByFile.ContainsKey($file.Name)) {
            $results += $existingByFile[$file.Name]
            $reusedCount++
            continue
        }
        $started = Get-Date
        try {
            $metadata = Invoke-OcrSourceWithLayoutDirect -Path $file.FullName -Engine $engine
            $text = $metadata.text
            $results += [pscustomobject]@{
                file = $file.Name
                status = "SUCCESS"
                language = $LanguageTag
                tile_count = $metadata.tile_count
                character_count = $text.Length
                text = $text
                original_width = $metadata.original_width
                original_height = $metadata.original_height
                scaled_width = $metadata.scaled_width
                scaled_height = $metadata.scaled_height
                scale = $metadata.scale
                words = $metadata.words
                error = ""
                elapsed_ms = [int]((Get-Date) - $started).TotalMilliseconds
            }
        }
        catch {
            $results += [pscustomobject]@{
                file = $file.Name
                status = "ERROR"
                language = $LanguageTag
                tile_count = 0
                character_count = 0
                text = ""
                words = @()
                error = $_.Exception.Message + " @ " + $_.ScriptStackTrace
                elapsed_ms = [int]((Get-Date) - $started).TotalMilliseconds
            }
        }
        finally {
            Get-ChildItem -LiteralPath $tempDirectory -File -ErrorAction SilentlyContinue | Remove-Item -Force
        }
        if ($results.Count % 20 -eq 0) {
            "processed=$($results.Count)/$($files.Count) reused=$reusedCount" |
                Set-Content -LiteralPath $progressPath -Encoding UTF8
            $results | ConvertTo-Json -Depth 8 |
                Set-Content -LiteralPath $outputPath -Encoding UTF8
        }
    }
    $results | ConvertTo-Json -Depth 8 | Set-Content -LiteralPath $outputPath -Encoding UTF8
    $successCount = @($results | Where-Object status -eq 'SUCCESS').Count
    "processed=$($results.Count)/$($files.Count) reused=$reusedCount complete=1" |
        Set-Content -LiteralPath $progressPath -Encoding UTF8
    Write-Output "files=$($results.Count) success=$successCount output=$outputPath"
}
finally {
    if (Test-Path -LiteralPath $tempDirectory) {
        Remove-Item -LiteralPath $tempDirectory -Recurse -Force
    }
}
