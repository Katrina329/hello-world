$ErrorActionPreference = 'Stop'
Add-Type -AssemblyName System.Runtime.WindowsRuntime
$null = [Windows.Media.Ocr.OcrEngine, Windows.Foundation, ContentType=WindowsRuntime]
$null = [Windows.Graphics.Imaging.BitmapDecoder, Windows.Foundation, ContentType=WindowsRuntime]
$null = [Windows.Storage.StorageFile, Windows.Foundation, ContentType=WindowsRuntime]
$null = [Windows.Globalization.Language, Windows.Foundation, ContentType=WindowsRuntime]
$asTaskGeneric = ([System.WindowsRuntimeSystemExtensions].GetMethods() | Where-Object { $_.Name -eq 'AsTask' -and $_.GetParameters().Count -eq 1 -and $_.GetParameters()[0].ParameterType.Name -eq 'IAsyncOperation`1' })[0]
function Await($WinRtOp, $ResultType) { $asTask = $asTaskGeneric.MakeGenericMethod($ResultType); $netTask = $asTask.Invoke($null, @($WinRtOp)); $netTask.Wait(-1) | Out-Null; $netTask.Result }
function Ocr-Lines([string]$path, [string]$langtag) {
  $file = Await ([Windows.Storage.StorageFile]::GetFileFromPathAsync($path)) ([Windows.Storage.StorageFile])
  $stream = Await ($file.OpenReadAsync()) ([Windows.Storage.Streams.IRandomAccessStreamWithContentType])
  $decoder = Await ([Windows.Graphics.Imaging.BitmapDecoder]::CreateAsync($stream)) ([Windows.Graphics.Imaging.BitmapDecoder])
  $bitmap = Await ($decoder.GetSoftwareBitmapAsync()) ([Windows.Graphics.Imaging.SoftwareBitmap])
  $lang = New-Object Windows.Globalization.Language $langtag
  $engine = [Windows.Media.Ocr.OcrEngine]::TryCreateFromLanguage($lang)
  $result = Await ($engine.RecognizeAsync($bitmap)) ([Windows.Media.Ocr.OcrResult])
  $result.Lines
}
$dir = Join-Path (Get-Location).Path 'ocr_big'
$files = @('p0010_code.png','p0030_code.png','p0050_code.png','p0050_code2.png')
$out = New-Object System.Collections.Generic.List[string]
foreach ($name in $files) {
  $p = Join-Path $dir $name
  $lines = Ocr-Lines $p 'en-GB'
  $out.Add('===== ' + $name + ' =====')
  foreach ($line in $lines) {
    $r = $line.Words[0].BoundingRect
    $s = 'Y=' + $r.Y + ' X=' + $r.X + '  ' + $line.Text
    $out.Add($s)
  }
  $out.Add('')
}
$outPath = Join-Path (Get-Location).Path 'ocr_lines_box.txt'
[System.IO.File]::WriteAllLines($outPath, $out, (New-Object System.Text.UTF8Encoding $false))
Write-Output ('saved ' + $outPath)
