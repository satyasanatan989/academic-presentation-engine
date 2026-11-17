$pptx = "C:\Users\prabh\Documents\antigravity\PPT\samples\phytochemicals_25slides_master_deck.pptx"
$pdf = "C:\Users\prabh\Documents\antigravity\PPT\samples\phytochemicals_25slides_master_deck.pdf"
$galDir = "C:\Users\prabh\Documents\antigravity\PPT\samples\gallery_phytochemicals_25"

if (-not (Test-Path $galDir)) {
    New-Item -ItemType Directory -Force -Path $galDir | Out-Null
}

$ppt = New-Object -ComObject PowerPoint.Application
try {
    $pres = $ppt.Presentations.Open($pptx, [Microsoft.Office.Core.MsoTriState]::msoFalse, [Microsoft.Office.Core.MsoTriState]::msoFalse, [Microsoft.Office.Core.MsoTriState]::msoFalse)
    $pres.SaveAs($pdf, 32)
    Write-Host "Saved PDF: $pdf"

    $idx = 1
    foreach ($slide in $pres.Slides) {
        $outFile = Join-Path $galDir ("Slide_{0:D2}.PNG" -f $idx)
        $slide.Export($outFile, "PNG", 1920, 1080)
        $idx++
    }
    Write-Host "Exported $($idx-1) slides to $galDir"
    $pres.Close()
}
finally {
    $ppt.Quit()
    [System.Runtime.Interopservices.Marshal]::ReleaseComObject($ppt) | Out-Null
    [System.GC]::Collect()
    [System.GC]::WaitForPendingFinalizers()
}
