$ppt = New-Object -ComObject PowerPoint.Application
$ppt.WindowState = 2
$pres = $ppt.Presentations.Open('C:\Users\gargt\OneDrive\Desktop\ppt midsem sem 7\MinorProject.pptx', [Microsoft.Office.Core.MsoTriState]::msoFalse, [Microsoft.Office.Core.MsoTriState]::msoFalse, [Microsoft.Office.Core.MsoTriState]::msoTrue)

$s4 = $pres.Slides.Item(4)
$table = $null
foreach ($sh in $s4.Shapes) {
    if ($sh.HasTable) {
        $table = $sh.Table
        break
    }
}

if ($table -ne $null) {
    Write-Host "Table rows: $($table.Rows.Count), cols: $($table.Columns.Count)"
    for ($c = 1; $c -le 4; $c++) {
        $cell = $table.Cell(1, $c)
        Write-Host "Cell (1,$c) text: '$($cell.Shape.TextFrame.TextRange.Text)'"
        Write-Host "Cell (1,$c) fill type: $($cell.Shape.Fill.Type), ForeColor RGB: $($cell.Shape.Fill.ForeColor.RGB)"
    }
    
    # Check borders of Cell(1,1)
    for ($b = 1; $b -le $cell.Borders.Count; $b++) {
        $border = $cell.Borders.Item($b)
        Write-Host "Border $b Visible: $($border.Visible), Weight: $($border.Weight)"
    }
}

$pres.Close()
$ppt.Quit()
