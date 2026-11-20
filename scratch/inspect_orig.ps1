$ppt = New-Object -ComObject PowerPoint.Application
$pres = $ppt.Presentations.Open('C:\Users\gargt\OneDrive\Desktop\ppt midsem sem 7\MinorProject_Backup_Original.pptx', [Microsoft.Office.Core.MsoTriState]::msoFalse, [Microsoft.Office.Core.MsoTriState]::msoFalse, [Microsoft.Office.Core.MsoTriState]::msoFalse)
Write-Host "Original slide count: $($pres.Slides.Count)"
for ($i = 1; $i -le $pres.Slides.Count; $i++) {
    $s = $pres.Slides.Item($i)
    $title = ""
    foreach ($sh in $s.Shapes) {
        if ($sh.HasTextFrame -and $sh.Top -lt 60) {
            $title = $sh.TextFrame.TextRange.Text -replace "`r", " " -replace "`n", " "
            break
        }
    }
    Write-Host "Slide $i : $title"
}
$pres.Close()
$ppt.Quit()
