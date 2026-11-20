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
    Write-Host "Found table!"
    $cell = $table.Cell(1, 1)
    try {
        $cell.Shape.TextFrame.TextRange.Text = "TEST HEADER"
        Write-Host "Text set successfully!"
    } catch {
        Write-Host "Text set failed: $_"
    }

    try {
        $cell.Shape.Fill.Solid()
        $cell.Shape.Fill.ForeColor.RGB = 0x881414
        Write-Host "Fill set successfully!"
    } catch {
        Write-Host "Fill set failed: $_"
    }

    try {
        $cell.Shape.TextFrame.TextRange.Font.Color.RGB = 0xFFFFFF
        $cell.Shape.TextFrame.TextRange.Font.Bold = -1
        Write-Host "Font color set successfully!"
    } catch {
        Write-Host "Font color set failed: $_"
    }
}

$s4.Export("C:\Users\gargt\OneDrive\Desktop\ppt midsem sem 7\scratch\preview_tbl_test.png", "PNG", 1920, 1080)
$pres.Close()
$ppt.Quit()
