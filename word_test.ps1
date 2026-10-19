$ErrorActionPreference = 'Stop'
$srcFile = Get-ChildItem (Join-Path (Get-Location).Path '*.docx') | Sort-Object LastWriteTime -Descending | Select-Object -First 1
$src = $srcFile.FullName
Write-Output ('opening ' + $src)
$word = New-Object -ComObject Word.Application
$word.Visible = $false
$word.DisplayAlerts = 0
$doc = $word.Documents.Open($src, $false, $true)
Write-Output ('paragraphs ' + $doc.Paragraphs.Count)
$doc.Close($false)
$word.Quit()
[System.Runtime.InteropServices.Marshal]::ReleaseComObject($word) | Out-Null
Write-Output 'done'
