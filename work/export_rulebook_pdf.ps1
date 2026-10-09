$ErrorActionPreference = 'Stop'
$inputDocx = 'C:\Users\richc\Documents\Codex\2026-10-02\4\outputs\島嶼共和_原型規則_v0.1.docx'
$outputPdf = 'C:\Users\richc\Documents\Codex\2026-10-02\4\outputs\島嶼共和_原型規則_v0.1.pdf'

$word = $null
$document = $null
try {
    $word = New-Object -ComObject Word.Application
    $word.Visible = $false
    $word.DisplayAlerts = 0
    $document = $word.Documents.Open($inputDocx, $false, $true)
    $document.Repaginate()
    $pages = $document.ComputeStatistics(2)
    $document.ExportAsFixedFormat($outputPdf, 17)
    Write-Output "pages=$pages"
}
finally {
    if ($null -ne $document) {
        $document.Close(0)
        [void][System.Runtime.InteropServices.Marshal]::ReleaseComObject($document)
    }
    if ($null -ne $word) {
        $word.Quit()
        [void][System.Runtime.InteropServices.Marshal]::ReleaseComObject($word)
    }
    [GC]::Collect()
    [GC]::WaitForPendingFinalizers()
}
