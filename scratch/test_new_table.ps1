$ppt = New-Object -ComObject PowerPoint.Application
$ppt.WindowState = 2
$pres = $ppt.Presentations.Add([Microsoft.Office.Core.MsoTriState]::msoTrue)
$slide = $pres.Slides.Add(1, 12) # 12 = ppLayoutBlank

$tblShape = $slide.Shapes.AddTable(3, 3, 50, 50, 400, 200)
$tbl = $tblShape.Table

Write-Host "Default Table Style: $($tbl.Style.Name)"
Write-Host "Default Table Style Id: $($tbl.Style.Id)"

# Try setting cell fill
$c = $tbl.Cell(1, 1)
$c.Shape.Fill.Solid()
$c.Shape.Fill.ForeColor.RGB = 0x000088 # Crimson
$c.Shape.TextFrame.TextRange.Text = "HELLO WORLD"
$c.Shape.TextFrame.TextRange.Font.Color.RGB = 0xFFFFFF

$slide.Export("C:\Users\gargt\OneDrive\Desktop\ppt midsem sem 7\scratch\test_new_table.png", "PNG", 800, 600)
$pres.Close()
$ppt.Quit()
Write-Host "Exported test_new_table.png"
