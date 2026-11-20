$ppt = New-Object -ComObject PowerPoint.Application
$pres = $ppt.Presentations.Open('C:\Users\gargt\OneDrive\Desktop\ppt midsem sem 7\MinorProject_Backup_Original.pptx', [Microsoft.Office.Core.MsoTriState]::msoFalse, [Microsoft.Office.Core.MsoTriState]::msoFalse, [Microsoft.Office.Core.MsoTriState]::msoFalse)
Write-Host "Total slides in backup: $($pres.Slides.Count)"
for ($i = 1; $i -le $pres.Slides.Count; $i++) {
    $sl = $pres.Slides.Item($i)
    $t = ""
    foreach ($sh in $sl.Shapes) {
        if ($sh.HasTextFrame -and $sh.TextFrame.HasText -and $sh.Top -lt 70) {
            $t = $sh.TextFrame.TextRange.Text.Trim()
            break
        }
    }
    Write-Host "Slide $i`: $t"
}
$pres.Close()
$ppt.Quit()
