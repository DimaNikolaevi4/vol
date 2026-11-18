param(
    [Parameter(Mandatory = $true)][string]$src,
    [string]$pdf = ""
)
# Обновляет поля и оглавление в docx через Word и сохраняет файл.
# Использование: powershell -File update_fields_word.ps1 -src "00.Программа....docx" [-pdf preview.pdf]
$ErrorActionPreference = "Stop"
if (-not (Test-Path $src)) { throw "not found: $src" }

$word = New-Object -ComObject Word.Application
$word.Visible = $false
$word.DisplayAlerts = 0
try {
    $doc = $word.Documents.Open($src, $false, $false)
    $doc.Fields.Update() | Out-Null
    if ($doc.TablesOfContents.Count -gt 0) { $doc.TablesOfContents.Item(1).Update() }
    $doc.Repaginate()
    Write-Output ("pages=" + $doc.ComputeStatistics(2))
    Write-Output ("toc_entries=" + $doc.TablesOfContents.Count)
    if ($pdf -ne "") { $doc.SaveAs([ref]$pdf, [ref]17) }
    $doc.Save()
    $doc.Close([ref]$false)
}
finally {
    $word.Quit()
}
Write-Output "done"
