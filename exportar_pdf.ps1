# Exporta el informe a PDF con Word y devuelve el número de páginas.
param([string]$Docx, [string]$Pdf)
$word = New-Object -ComObject Word.Application
$word.Visible = $false
$doc = $word.Documents.Open($Docx, $false, $true)
$doc.Repaginate()
$paginas = $doc.ComputeStatistics(2)
$doc.ExportAsFixedFormat($Pdf, 17, $false, 0)
$doc.Close($false)
$word.Quit()
[System.Runtime.Interopservices.Marshal]::ReleaseComObject($word) | Out-Null
Write-Output ("PAGINAS=" + $paginas)
