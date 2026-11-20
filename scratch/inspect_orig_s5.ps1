$ppt = New-Object -ComObject PowerPoint.Application
$pres = $ppt.Presentations.Open('C:\Users\gargt\OneDrive\Desktop\ppt midsem sem 7\MinorProject_Backup_Original.pptx', [Microsoft.Office.Core.MsoTriState]::msoFalse, [Microsoft.Office.Core.MsoTriState]::msoFalse, [Microsoft.Office.Core.MsoTriState]::msoFalse)
$s5 = $pres.Slides.Item(5)
Write-Host "Slide 5 in backup shape count: $($s5.Shapes.Count)"
for ($i = 1; $i -le $s5.Shapes.Count; $i++) {
    $sh = $s5.Shapes.Item($i)
    Write-Host "Shape $i : Name=$($sh.Name) Type=$($sh.Type) Top=$($sh.Top) Left=$($sh.Left) Width=$($sh.Width) Height=$($sh.Height)"
    if ($sh.HasTable) {
        $tbl = $sh.Table
        Write-Host "  Table: $($tbl.Rows.Count) rows, $($tbl.Columns.Count) cols"
    } elseif ($sh.HasTextFrame) {
        $txt = $sh.TextFrame.TextRange.Text -replace "`r", " " -replace "`n", " "
        Write-Host "  Text: $($txt)"
    }
}
$pres.Close()
$ppt.Quit()
