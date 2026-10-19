$ErrorActionPreference = 'Stop'
$srcFile = Get-ChildItem (Join-Path (Get-Location).Path '*.docx') | Sort-Object LastWriteTime -Descending | Select-Object -First 1
$src = $srcFile.FullName
$pdf = Join-Path (Get-Location).Path 'render_out\report.pdf'
New-Item -ItemType Directory -Force -Path (Split-Path $pdf) | Out-Null
$word = New-Object -ComObject Word.Application
$word.Visible = $false
$word.DisplayAlerts = 0
try { $word.Options.DoNotPromptForConvert = $true } catch {}
$doc = $word.Documents.Open($src, $false, $true)
$doc.ExportAsFixedFormat($pdf, 17)
$doc.Close($false)
$word.Quit()
[System.Runtime.InteropServices.Marshal]::ReleaseComObject($word) | Out-Null
Write-Output ('pdf saved ' + $pdf)
