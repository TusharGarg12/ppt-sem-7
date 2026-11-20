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
    Write-Host "Table Style: $($table.Style.Name)"
    $table.ApplyStyleHeadingRows = $false
    $table.ApplyStyleBandedRows = $false
    $table.ApplyStyleFirstRow = $false
    $table.ApplyStyleLastRow = $false

    for ($r = 1; $r -le $table.Rows.Count; $r++) {
        for ($c = 1; $c -le $table.Columns.Count; $c++) {
            $cell = $table.Cell($r, $c)
            # Turn OFF diagonal borders
            $cell.Borders.Item(5).Visible = 0
            $cell.Borders.Item(6).Visible = 0
            
            # Set borders 1 to 4
            for ($b = 1; $b -le 4; $b++) {
                $cell.Borders.Item($b).ForeColor.RGB = 0xCBD5E1
                $cell.Borders.Item($b).Weight = 0.8
                $cell.Borders.Item($b).Visible = -1
            }

            if ($r -eq 1) {
                $cell.Shape.Fill.Solid()
                $cell.Shape.Fill.ForeColor.RGB = 0x881414
                $cell.Shape.TextFrame.TextRange.Font.Color.RGB = 0xFFFFFF
                $cell.Shape.TextFrame.TextRange.Font.Bold = -1
                $cell.Shape.TextFrame.TextRange.Font.Size = [single]11.0
            } else {
                if ($r % 2 -eq 0) {
                    $cell.Shape.Fill.Solid()
                    $cell.Shape.Fill.ForeColor.RGB = 0xFFFFFF
                } else {
                    $cell.Shape.Fill.Solid()
                    $cell.Shape.Fill.ForeColor.RGB = 0xF8FAFC
                }
            }
        }
    }
}

$s4.Export("C:\Users\gargt\OneDrive\Desktop\ppt midsem sem 7\scratch\preview_s4_test2.png", "PNG", 1920, 1080)
$pres.Close()
$ppt.Quit()
Write-Host "Exported preview_s4_test2.png"
