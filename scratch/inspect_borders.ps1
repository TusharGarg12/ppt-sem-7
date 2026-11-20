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
    $cell = $table.Cell(1, 1)
    Write-Host "Cell (1,1) borders count: $($cell.Borders.Count)"
    for ($i = 1; $i -le $cell.Borders.Count; $i++) {
        $b = $cell.Borders.Item($i)
        Write-Host "Border $($i) : Visible=$($b.Visible), Weight=$($b.Weight), ForeColor=$($b.ForeColor.RGB)"
    }
}

$pres.Close()
$ppt.Quit()
