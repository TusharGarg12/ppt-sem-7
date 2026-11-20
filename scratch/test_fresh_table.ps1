$ppt = New-Object -ComObject PowerPoint.Application
$ppt.WindowState = 2
$pres = $ppt.Presentations.Open('C:\Users\gargt\OneDrive\Desktop\ppt midsem sem 7\MinorProject_Backup_Original.pptx', [Microsoft.Office.Core.MsoTriState]::msoFalse, [Microsoft.Office.Core.MsoTriState]::msoFalse, [Microsoft.Office.Core.MsoTriState]::msoTrue)

$s4 = $pres.Slides.Item(4)

# Delete existing body shapes
$toDel = @()
foreach ($sh in $s4.Shapes) {
    $isRedLine = ($sh.Top -ge 40 -and $sh.Top -le 65 -and $sh.Height -le 10)
    $isLogo    = ($sh.Top -lt 55 -and $sh.Left -gt 870)
    $isNum     = ($sh.HasTextFrame -and $sh.Top -gt 500 -and $sh.Left -gt 700)
    if (-not ($isRedLine -or $isLogo -or $isNum)) {
        $toDel += $sh
    }
}
foreach ($sh in $toDel) { $sh.Delete() }

# Add fresh table
$tblShape = $s4.Shapes.AddTable(3, 3, 25, 68, 600, 150)
$tbl = $tblShape.Table

for ($r = 1; $r -le 3; $r++) {
    for ($c = 1; $c -le 3; $c++) {
        $cell = $tbl.Cell($r, $c)
        $cell.Shape.TextFrame.TextRange.Text = "R$r C$c"
        
        # ONLY borders 1 to 4!
        for ($b = 1; $b -le 4; $b++) {
            $border = $cell.Borders.Item($b)
            $border.ForeColor.RGB = 0xCBD5E1
            $border.Weight = 0.8
            $border.Visible = -1
        }
        
        $cell.Shape.Fill.Transparency = 0.0
        $cell.Shape.Fill.Solid()
        if ($r -eq 1) {
            $cell.Shape.Fill.ForeColor.RGB = 0x1C1CB7
            $cell.Shape.TextFrame.TextRange.Font.Color.RGB = 0xFFFFFF
        } else {
            $cell.Shape.Fill.ForeColor.RGB = 0xF8FAFC
        }
    }
}

$s4.Export("C:\Users\gargt\OneDrive\Desktop\ppt midsem sem 7\scratch\verify_fresh_table_no_diagonals.png", "PNG", 1920, 1080)
$pres.Close()
$ppt.Quit()
Write-Host "Exported verify_fresh_table_no_diagonals.png"
