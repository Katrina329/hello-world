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
  foreach ($line in $result.Lines) { $line.Text }
}
$dir = Join-Path (Get-Location).Path 'ocr_big'
$files = Get-ChildItem (Join-Path $dir '*.png') | Sort-Object Name
$out = New-Object System.Collections.Generic.List[string]
foreach ($f in $files) {
  $out.Add('===== ' + $f.Name + ' =====')
  $out.AddRange([string[]](Ocr-Lines $f.FullName 'en-GB'))
  $out.Add('')
}
$outPath = Join-Path (Get-Location).Path 'ocr_big_en.txt'
[System.IO.File]::WriteAllLines($outPath, $out, (New-Object System.Text.UTF8Encoding $false))
Write-Output ('saved ' + $outPath)
