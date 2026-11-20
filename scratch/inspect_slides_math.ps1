$ppt = New-Object -ComObject PowerPoint.Application
$pres = $ppt.Presentations.Open('C:\Users\gargt\OneDrive\Desktop\ppt midsem sem 7\MinorProject.pptx', [Microsoft.Office.Core.MsoTriState]::msoTrue, [Microsoft.Office.Core.MsoTriState]::msoFalse, [Microsoft.Office.Core.MsoTriState]::msoFalse)
for ($i = 1; $i -le $pres.Slides.Count; $i++) {
    $slide = $pres.Slides.Item($i)
    $title = ""
    try { $title = $slide.Shapes.Title.TextFrame.TextRange.Text } catch {}
    Write-Host ("=== Slide $i : $title ===")
    foreach ($sh in $slide.Shapes) {
        if ($sh.HasTextFrame -and $sh.TextFrame.HasText) {
            $t = $sh.TextFrame.TextRange.Text.Trim()
            if ($t.Length -gt 0) {
                Write-Host ("  [Shape $($sh.Name) Left=$($sh.Left) Top=$($sh.Top) W=$($sh.Width) H=$($sh.Height)]")
                $lines = $t -split "`r`n|`n|`r"
                foreach ($l in $lines) {
                    if ($l.Trim().Length -gt 0) {
                        Write-Host ("    " + $l.Trim())
                    }
                }
            }
        }
    }
}
$pres.Close()
$ppt.Quit()
