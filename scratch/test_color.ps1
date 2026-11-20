$ppt = New-Object -ComObject PowerPoint.Application
$ppt.WindowState = 2
$pres = $ppt.Presentations.Add([Microsoft.Office.Core.MsoTriState]::msoTrue)
$slide = $pres.Slides.Add(1, 12)

$tblShape = $slide.Shapes.AddTable(3, 3, 50, 50, 600, 300)
$tbl = $tblShape.Table

# Cell 1: 0x881414
$c1 = $tbl.Cell(1, 1)
$c1.Shape.Fill.Solid()
$c1.Shape.Fill.ForeColor.RGB = 0x881414
$c1.Shape.TextFrame.TextRange.Text = "0x881414"
$c1.Shape.TextFrame.TextRange.Font.Color.RGB = 0xFFFFFF

# Cell 2: 0x000088
$c2 = $tbl.Cell(1, 2)
$c2.Shape.Fill.Solid()
$c2.Shape.Fill.ForeColor.RGB = 0x000088
$c2.Shape.TextFrame.TextRange.Text = "0x000088"
$c2.Shape.TextFrame.TextRange.Font.Color.RGB = 0xFFFFFF

# Cell 3: 0x1C1CB7 (Crimson BGR)
$c3 = $tbl.Cell(1, 3)
$c3.Shape.Fill.Solid()
$c3.Shape.Fill.ForeColor.RGB = 0x1C1CB7
$c3.Shape.TextFrame.TextRange.Text = "0x1C1CB7"
$c3.Shape.TextFrame.TextRange.Font.Color.RGB = 0xFFFFFF

# Row 2 Cell 1: set borders 1..4 only
for ($b = 1; $b -le 4; $b++) {
    $c1.Borders.Item($b).ForeColor.RGB = 0xCBD5E1
    $c1.Borders.Item($b).Weight = 1.0
}

$slide.Export("C:\Users\gargt\OneDrive\Desktop\ppt midsem sem 7\scratch\test_color_export.png", "PNG", 800, 600)
$pres.Close()
$ppt.Quit()
Write-Host "Exported test_color_export.png"
