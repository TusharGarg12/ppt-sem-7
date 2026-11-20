$ppt = New-Object -ComObject PowerPoint.Application
$pres = $ppt.Presentations.Open('C:\Users\gargt\OneDrive\Desktop\ppt midsem sem 7\MinorProject.pptx', [Microsoft.Office.Core.MsoTriState]::msoFalse, [Microsoft.Office.Core.MsoTriState]::msoFalse, [Microsoft.Office.Core.MsoTriState]::msoFalse)
$s = $pres.Slides.Item(17)
Write-Host "Slide 17 shapes:"
foreach ($sh in $s.Shapes) {
    Write-Host "  Name: $($sh.Name), Type: $($sh.Type)"
    if ($sh.HasTable) {
        $tbl = $sh.Table
        Write-Host "  Table: $($tbl.Rows.Count) rows, $($tbl.Columns.Count) cols"
        $c1 = $tbl.Cell(1,1)
        Write-Host "  Cell 1,1 Text: '$($c1.Shape.TextFrame.TextRange.Text)'"
    }
}
$pres.Close()
$ppt.Quit()
