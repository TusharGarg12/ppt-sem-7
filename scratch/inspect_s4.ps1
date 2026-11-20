$ppt = New-Object -ComObject PowerPoint.Application
$pres = $ppt.Presentations.Open('C:\Users\gargt\OneDrive\Desktop\ppt midsem sem 7\MinorProject.pptx', [Microsoft.Office.Core.MsoTriState]::msoFalse, [Microsoft.Office.Core.MsoTriState]::msoFalse, [Microsoft.Office.Core.MsoTriState]::msoFalse)
$s4 = $pres.Slides.Item(4)
Write-Host "Slide 4 shapes count: $($s4.Shapes.Count)"
for ($i = 1; $i -le $s4.Shapes.Count; $i++) {
    $sh = $s4.Shapes.Item($i)
    Write-Host "Shape $i : Name=$($sh.Name) Type=$($sh.Type) Top=$($sh.Top) Left=$($sh.Left) Width=$($sh.Width) Height=$($sh.Height)"
    if ($sh.HasTable) {
        $tbl = $sh.Table
        Write-Host "  Table: $($tbl.Rows.Count) rows, $($tbl.Columns.Count) cols"
        for ($r = 1; $r -le $tbl.Rows.Count; $r++) {
            $rowTxt = @()
            for ($c = 1; $c -le $tbl.Columns.Count; $c++) {
                $cellTxt = $tbl.Cell($r, $c).Shape.TextFrame.TextRange.Text -replace "`r", " " -replace "`n", " "
                $rowTxt += $cellTxt
            }
            Write-Host "    Row $r : $($rowTxt -join ' | ')"
        }
    } elseif ($sh.HasTextFrame) {
        $txt = $sh.TextFrame.TextRange.Text -replace "`r", " " -replace "`n", " "
        Write-Host "  Text: $($txt.Substring(0, [System.Math]::Min(100, $txt.Length)))"
    }
}
$pres.Close()
$ppt.Quit()
