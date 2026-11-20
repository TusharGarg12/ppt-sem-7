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
    Write-Host "Borders count: $($cell.Borders.Count)"
    try {
        for ($b = 1; $b -le $cell.Borders.Count; $b++) {
            $border = $cell.Borders.Item($b)
            $border.ForeColor.RGB = 0xCBD5E1
            $border.Weight = 0.8
            $border.Visible = -1
        }
        Write-Host "Borders set successfully!"
    } catch {
        Write-Host "Border set failed: $_"
    }
}

$s4.Export("C:\Users\gargt\OneDrive\Desktop\ppt midsem sem 7\scratch\preview_border_test.png", "PNG", 1920, 1080)
$pres.Close()
$ppt.Quit()
