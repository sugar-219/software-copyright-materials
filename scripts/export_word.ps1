param([Parameter(Mandatory=$true)][string]$Job)
$ErrorActionPreference = 'Stop'
$exportJob = Get-Content -LiteralPath $Job -Raw -Encoding UTF8 | ConvertFrom-Json
$word = $null
$currentDocument = $null
$results = @()
try {
    $word = New-Object -ComObject Word.Application
    $word.Visible = $false
    $word.DisplayAlerts = 0
    $word.AutomationSecurity = 3
    foreach ($item in $exportJob.documents) {
        $inputPath = [System.IO.Path]::GetFullPath($item.docx)
        $pdfPath = [System.IO.Path]::GetFullPath($item.pdf)
        if ([System.IO.Path]::GetExtension($inputPath) -ne '.docx') { throw 'Only DOCX input is allowed' }
        if (-not (Test-Path -LiteralPath $inputPath)) { throw "Missing document: $inputPath" }
        if (Test-Path -LiteralPath $pdfPath) { throw "PDF already exists; use a new work/output directory: $pdfPath" }
        $currentDocument = $word.Documents.Open($inputPath, $false, $false, $false)
        $currentDocument.Fields.Update() | Out-Null
        $currentDocument.Repaginate()
        $pageCount = $currentDocument.ComputeStatistics(2)
        $pageStats = @()
        for ($page = 1; $page -le $pageCount; $page++) {
            $startRange = $currentDocument.GoTo(1, 1, $page)
            $start = $startRange.Start
            if ($page -lt $pageCount) {
                $nextRange = $currentDocument.GoTo(1, 1, ($page + 1))
                $end = $nextRange.Start
                [void][System.Runtime.InteropServices.Marshal]::ReleaseComObject($nextRange)
            } else { $end = $currentDocument.Content.End }
            $pageRange = $currentDocument.Range($start, $end)
            $pageStats += [pscustomobject]@{ page = $page; word_lines = $pageRange.ComputeStatistics(1) }
            [void][System.Runtime.InteropServices.Marshal]::ReleaseComObject($pageRange)
            [void][System.Runtime.InteropServices.Marshal]::ReleaseComObject($startRange)
        }
        $currentDocument.Save()
        # Export the same saved editable document; never re-create its layout independently.
        $currentDocument.ExportAsFixedFormat($pdfPath, 17)
        $results += [pscustomobject]@{ kind = $item.kind; docx = $inputPath; pdf = $pdfPath; pages = $pageCount; page_statistics = $pageStats }
        $currentDocument.Close(0)
        [void][System.Runtime.InteropServices.Marshal]::ReleaseComObject($currentDocument)
        $currentDocument = $null
    }
    $reportPath = Join-Path $exportJob.work_dir 'word-export-report.json'
    @{ documents = $results; renderer = 'Microsoft Word'; status = 'exported_pending_visual_review' } |
        ConvertTo-Json -Depth 8 | Set-Content -LiteralPath $reportPath -Encoding UTF8
    Write-Output "Exported $($results.Count) documents; report: $reportPath"
} finally {
    if ($null -ne $currentDocument) {
        $currentDocument.Close(0)
        [void][System.Runtime.InteropServices.Marshal]::ReleaseComObject($currentDocument)
    }
    if ($null -ne $word) {
        $word.Quit()
        [void][System.Runtime.InteropServices.Marshal]::ReleaseComObject($word)
    }
    [GC]::Collect()
    [GC]::WaitForPendingFinalizers()
}
