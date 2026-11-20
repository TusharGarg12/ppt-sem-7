$ppt = New-Object -ComObject PowerPoint.Application
$ppt.WindowState = 2
$pres = $ppt.Presentations.Open('C:\Users\gargt\OneDrive\Desktop\ppt midsem sem 7\MinorProject.pptx', [Microsoft.Office.Core.MsoTriState]::msoFalse, [Microsoft.Office.Core.MsoTriState]::msoFalse, [Microsoft.Office.Core.MsoTriState]::msoTrue)

$s4 = $pres.Slides.Item(4)
$tbl = $null
foreach ($sh in $s4.Shapes) {
    if ($sh.HasTable) {
        $tbl = $sh.Table
        break
    }
}

if ($tbl -ne $null) {
    $c1 = $tbl.Cell(1, 1)
    Write-Host "Cell (1,1) fill visible: $($c1.Shape.Fill.Visible)"
    Write-Host "Cell (1,1) fill ForeColor: $($c1.Shape.Fill.ForeColor.RGB)"
    Write-Host "Cell (1,1) fill BackColor: $($c1.Shape.Fill.BackColor.RGB)"
    Write-Host "Cell (1,1) fill Transparency: $($c1.Shape.Fill.Transparency)"
    
    # Try setting transparency to 0 and solid fill
    $c1.Shape.Fill.Transparency = 0.0
    $c1.Shape.Fill.Solid()
    $c1.Shape.Fill.ForeColor.RGB = 0x1C1CB7 # Crimson
}

$s4.Export("C:\Users\gargt\OneDrive\Desktop\ppt midsem sem 7\scratch\test_s4_c1_fill.png", "PNG", 1920, 1080)
$pres.Close()
$ppt.Quit()
Write-Host "Exported test_s4_c1_fill.png"
