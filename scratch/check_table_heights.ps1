$ppt = New-Object -ComObject PowerPoint.Application
$ppt.WindowState = 2
$pres = $ppt.Presentations.Open('C:\Users\gargt\OneDrive\Desktop\ppt midsem sem 7\MinorProject.pptx', [Microsoft.Office.Core.MsoTriState]::msoFalse, [Microsoft.Office.Core.MsoTriState]::msoFalse, [Microsoft.Office.Core.MsoTriState]::msoTrue)

$s4 = $pres.Slides.Item(4)
foreach ($sh in $s4.Shapes) {
    if ($sh.HasTable) {
        $t = $sh.Table
        Write-Host "Slide 4 Table Top=$($sh.Top), Height=$($sh.Height)"
        for ($r = 1; $r -le $t.Rows.Count; $r++) {
            Write-Host "Row $r height: $($t.Rows.Item($r).Height)"
            $cell = $t.Cell($r, 1)
            Write-Host "  Cell($r,1) text: '$($cell.Shape.TextFrame.TextRange.Text)'"
            Write-Host "  Cell($r,1) font size: $($cell.Shape.TextFrame.TextRange.Font.Size), font color: $($cell.Shape.TextFrame.TextRange.Font.Color.RGB), fill: $($cell.Shape.Fill.ForeColor.RGB)"
        }
    }
}

$s17 = $pres.Slides.Item(17)
foreach ($sh in $s17.Shapes) {
    if ($sh.HasTable) {
        $t = $sh.Table
        Write-Host "Slide 17 Table Top=$($sh.Top), Height=$($sh.Height)"
        for ($r = 1; $r -le $t.Rows.Count; $r++) {
            Write-Host "Row $r height: $($t.Rows.Item($r).Height)"
            $cell = $t.Cell($r, 1)
            Write-Host "  Cell($r,1) text: '$($cell.Shape.TextFrame.TextRange.Text)'"
        }
    }
}

$pres.Close()
$ppt.Quit()
