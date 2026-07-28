param(
    [Parameter(Mandatory = $true)]
    [string]$InputDirectory,
    [Parameter(Mandatory = $true)]
    [string]$OutputJson,
    [string]$LanguageTag = "ko"
)

$ErrorActionPreference = "Stop"
Add-Type -AssemblyName System.Runtime.WindowsRuntime
Add-Type -AssemblyName System.Drawing
[void][Windows.Storage.StorageFile, Windows.Storage, ContentType = WindowsRuntime]
[void][Windows.Storage.FileAccessMode, Windows.Storage, ContentType = WindowsRuntime]
[void][Windows.Graphics.Imaging.BitmapDecoder, Windows.Graphics.Imaging, ContentType = WindowsRuntime]
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
    param($Operation, [Type]$ResultType)
    $task = $asTaskMethod.MakeGenericMethod($ResultType).Invoke($null, @($Operation))
    $task.Wait()
    return $task.Result
}

function Invoke-OcrFile {
    param([string]$Path, $Engine)
    $storageFile = Wait-WinRtOperation ([Windows.Storage.StorageFile]::GetFileFromPathAsync($Path)) ([Windows.Storage.StorageFile])
    $stream = Wait-WinRtOperation ($storageFile.OpenAsync([Windows.Storage.FileAccessMode]::Read)) ([Windows.Storage.Streams.IRandomAccessStream])
    try {
        $decoder = Wait-WinRtOperation ([Windows.Graphics.Imaging.BitmapDecoder]::CreateAsync($stream)) ([Windows.Graphics.Imaging.BitmapDecoder])
        $bitmap = Wait-WinRtOperation ($decoder.GetSoftwareBitmapAsync()) ([Windows.Graphics.Imaging.SoftwareBitmap])
        try {
            $result = Wait-WinRtOperation ($Engine.RecognizeAsync($bitmap)) ([Windows.Media.Ocr.OcrResult])
            return $result.Text
        }
        finally {
            if ($null -ne $bitmap) { $bitmap.Dispose() }
        }
    }
    finally {
        if ($null -ne $stream) { $stream.Dispose() }
    }
}

function Get-OcrTiles {
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
            $tiles += $tilePath
            $tileIndex++
        }
        $scaled.Dispose()
        return $tiles
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

$tempDirectory = Join-Path ([IO.Path]::GetTempPath()) ("homestyle_ocr_" + [Guid]::NewGuid().ToString("N"))
[void](New-Item -ItemType Directory -Path $tempDirectory)

try {
    $results = @()
    $files = Get-ChildItem -LiteralPath $inputPath -File |
        Where-Object { $_.Extension.ToLowerInvariant() -in @(".jpg", ".jpeg", ".png", ".gif", ".bmp") } |
        Sort-Object Name

    foreach ($file in $files) {
        $started = Get-Date
        try {
            $tiles = Get-OcrTiles -Path $file.FullName -TempDirectory $tempDirectory
            $texts = foreach ($tile in $tiles) {
                Invoke-OcrFile -Path $tile -Engine $engine
            }
            $text = ($texts -join "`n").Trim()
            $results += [pscustomobject]@{
                file = $file.Name
                status = "SUCCESS"
                language = $LanguageTag
                tile_count = $tiles.Count
                character_count = $text.Length
                text = $text
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
                error = $_.Exception.Message + " @ " + $_.ScriptStackTrace
                elapsed_ms = [int]((Get-Date) - $started).TotalMilliseconds
            }
        }
        finally {
            Get-ChildItem -LiteralPath $tempDirectory -File -ErrorAction SilentlyContinue | Remove-Item -Force
        }
    }
    $results | ConvertTo-Json -Depth 5 | Set-Content -LiteralPath $outputPath -Encoding UTF8
    $successCount = @($results | Where-Object status -eq 'SUCCESS').Count
    Write-Output "files=$($results.Count) success=$successCount output=$outputPath"
}
finally {
    if (Test-Path -LiteralPath $tempDirectory) {
        Remove-Item -LiteralPath $tempDirectory -Recurse -Force
    }
}
