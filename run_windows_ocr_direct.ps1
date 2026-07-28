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
        throw [TimeoutException]::new("WinRT OCR operation exceeded ${TimeoutMs}ms")
    }
    return $task.Result
}

function Invoke-OcrFileDirect {
    param([string]$Path, $Engine)

    $storageFile = Wait-WinRtOperation ([Windows.Storage.StorageFile]::GetFileFromPathAsync($Path)) ([Windows.Storage.StorageFile])
    $stream = Wait-WinRtOperation ($storageFile.OpenAsync([Windows.Storage.FileAccessMode]::Read)) ([Windows.Storage.Streams.IRandomAccessStream])
    try {
        $decoder = Wait-WinRtOperation ([Windows.Graphics.Imaging.BitmapDecoder]::CreateAsync($stream)) ([Windows.Graphics.Imaging.BitmapDecoder])
        $width = [uint32]$decoder.PixelWidth
        $height = [uint32]$decoder.PixelHeight
        $tileHeight = [uint32]2200
        $texts = @()
        $tileCount = 0
        $fileStopwatch = [Diagnostics.Stopwatch]::StartNew()

        for ($top = [uint32]0; $top -lt $height; $top = [uint32]($top + $tileHeight)) {
            if ($fileStopwatch.ElapsedMilliseconds -gt 120000) {
                throw [TimeoutException]::new("Image OCR exceeded 120000ms")
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
                $texts += $result.Text
                $tileCount++
            }
            finally {
                if ($null -ne $bitmap) { $bitmap.Dispose() }
            }
        }
        return [pscustomobject]@{
            text = ($texts -join "`n").Trim()
            tile_count = $tileCount
        }
    }
    finally {
        if ($null -ne $stream) { $stream.Dispose() }
    }
}

$inputPath = (Resolve-Path -LiteralPath $InputDirectory).Path
$outputPath = [IO.Path]::GetFullPath($OutputJson)
$language = New-Object Windows.Globalization.Language($LanguageTag)
$engine = [Windows.Media.Ocr.OcrEngine]::TryCreateFromLanguage($language)
if ($null -eq $engine) {
    throw "OCR engine unavailable for language: $LanguageTag"
}

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
        Write-Warning "Previous OCR output could not be reused: $($_.Exception.Message)"
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
        $ocrResult = Invoke-OcrFileDirect -Path $file.FullName -Engine $engine
        $text = $ocrResult.text
        $results += [pscustomobject]@{
            file = $file.Name
            status = "SUCCESS"
            language = $LanguageTag
            tile_count = $ocrResult.tile_count
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
    if ($results.Count % 50 -eq 0) {
        "processed=$($results.Count)/$($files.Count) reused=$reusedCount" |
            Set-Content -LiteralPath $progressPath -Encoding UTF8
    }
    if ($results.Count % 50 -eq 0) {
        $results | ConvertTo-Json -Depth 5 |
            Set-Content -LiteralPath $outputPath -Encoding UTF8
    }
}

$results | ConvertTo-Json -Depth 5 | Set-Content -LiteralPath $outputPath -Encoding UTF8
$successCount = @($results | Where-Object status -eq 'SUCCESS').Count
"processed=$($results.Count)/$($files.Count) reused=$reusedCount complete=1" |
    Set-Content -LiteralPath $progressPath -Encoding UTF8
Write-Output "files=$($results.Count) success=$successCount output=$outputPath"
