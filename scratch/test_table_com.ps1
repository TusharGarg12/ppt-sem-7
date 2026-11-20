$ppt = New-Object -ComObject PowerPoint.Application
$ppt.WindowState = 2
$pres = $ppt.Presentations.Open('C:\Users\gargt\OneDrive\Desktop\ppt midsem sem 7\MinorProject.pptx', [Microsoft.Office.Core.MsoTriState]::msoFalse, [Microsoft.Office.Core.MsoTriState]::msoFalse, [Microsoft.Office.Core.MsoTriState]::msoFalse)

$s4 = $pres.Slides.Item(4)

# Test creating a test table
Write-Host "Slide 4 current shapes:"
foreach ($sh in $s4.Shapes) {
    Write-Host "  Name=$($sh.Name), Type=$($sh.Type)"
}

$pres.Close()
$ppt.Quit()
