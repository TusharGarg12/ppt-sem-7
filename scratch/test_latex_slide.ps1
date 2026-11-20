$ppt = New-Object -ComObject PowerPoint.Application
$ppt.Visible = [Microsoft.Office.Core.MsoTriState]::msoFalse
$pres = $ppt.Presentations.Open('C:\Users\gargt\OneDrive\Desktop\ppt midsem sem 7\MinorProject.pptx', [Microsoft.Office.Core.MsoTriState]::msoFalse, [Microsoft.Office.Core.MsoTriState]::msoFalse, [Microsoft.Office.Core.MsoTriState]::msoFalse)

$slide5 = $pres.Slides.Item(5)

# Inspect shapes on Slide 5
Write-Host "Shapes on Slide 5 before test:"
foreach ($sh in $slide5.Shapes) {
    Write-Host ("  " + $sh.Name + " | Type=" + $sh.Type + " | Top=" + $sh.Top + " | Left=" + $sh.Left + " | W=" + $sh.Width + " | H=" + $sh.Height)
}

$pres.Close()
$ppt.Quit()
