$ppt = New-Object -ComObject PowerPoint.Application
$pres = $ppt.Presentations.Open('C:\Users\gargt\OneDrive\Desktop\ppt midsem sem 7\MinorProject.pptx', [Microsoft.Office.Core.MsoTriState]::msoFalse, [Microsoft.Office.Core.MsoTriState]::msoFalse, [Microsoft.Office.Core.MsoTriState]::msoFalse)
for ($i = 1; $i -le $pres.Slides.Count; $i++) {
    $sl = $pres.Slides.Item($i)
    $titles = @()
    foreach ($sh in $sl.Shapes) {
        if ($sh.HasTextFrame -and $sh.Top -lt 60 -and $sh.TextFrame.HasText) {
            $t = $sh.TextFrame.TextRange.Text.Trim()
            if ($t.Length -gt 0) {
                $titles += "[Id $($sh.Id) L=$([Math]::Round($sh.Left,1)) W=$([Math]::Round($sh.Width,1))]: $t"
            }
        }
    }
    Write-Host "Slide $i`: $($titles -join ' | ')"
}
$pres.Close()
$ppt.Quit()
