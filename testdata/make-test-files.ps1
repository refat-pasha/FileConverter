# Generates test files for the File Converter test suite.
# Images via System.Drawing; PPTX/XLSX via the installed MS Office (COM).
# Note: test.docx is not regenerated here — it lives in ./minidocx (a hand-built
# minimal OOXML package, zipped to test.docx once) so the suite never depends on Word.
$ErrorActionPreference = 'Stop'
$dir = Split-Path -Parent $MyInvocation.MyCommand.Path
Write-Host "Creating test files in $dir"

# --- images ---
Add-Type -AssemblyName System.Drawing
$bmp = New-Object System.Drawing.Bitmap(240, 140)
$g = [System.Drawing.Graphics]::FromImage($bmp)
$g.Clear([System.Drawing.Color]::FromArgb(40, 60, 110))
$font = New-Object System.Drawing.Font('Arial', 26, [System.Drawing.FontStyle]::Bold)
$g.DrawString('Test image', $font, [System.Drawing.Brushes]::White, 20, 45)
$g.Dispose()
$bmp.Save((Join-Path $dir 'test.png'), [System.Drawing.Imaging.ImageFormat]::Png)
$bmp.Save((Join-Path $dir 'test.jpg'), [System.Drawing.Imaging.ImageFormat]::Jpeg)
$bmp.Save((Join-Path $dir 'test.bmp'), [System.Drawing.Imaging.ImageFormat]::Bmp)
$bmp.Dispose()

# --- docx ---
$word = New-Object -ComObject Word.Application
$word.Visible = $false
$doc = $word.Documents.Add()
$range = $doc.Range()
$range.Text = "File Converter test document.`r`nSecond line with some text to convert."
$doc.SaveAs2((Join-Path $dir 'test.docx'), 16) | Out-Null
$doc.Close($false)
$word.Quit()
Write-Host 'test.docx OK'

# --- pptx ---
$ppt = New-Object -ComObject PowerPoint.Application
$pres = $ppt.Presentations.Add(-1) # msoTrue (with window is unavoidable for Add)
$slide = $pres.Slides.Add(1, 2)    # ppLayoutText
$slide.Shapes.Title.TextFrame.TextRange.Text = 'Converter test deck'
$slide.Shapes.Placeholders(2).TextFrame.TextRange.Text = 'Slide one body text'
$pres.SaveAs((Join-Path $dir 'test.pptx'), 24) # ppSaveAsOpenXMLPresentation
$pres.Close()
$ppt.Quit()
Write-Host 'test.pptx OK'

# --- xlsx ---
$xl = New-Object -ComObject Excel.Application
$xl.Visible = $false
$xl.DisplayAlerts = $false
$wb = $xl.Workbooks.Add()
$ws = $wb.Worksheets.Item(1)
$ws.Cells.Item(1, 1) = 'Name'
$ws.Cells.Item(1, 2) = 'Value'
$ws.Cells.Item(2, 1) = 'alpha'
$ws.Cells.Item(2, 2) = 42
$ws.Cells.Item(3, 1) = 'beta'
$ws.Cells.Item(3, 2) = 7.5
$wb.SaveAs((Join-Path $dir 'test.xlsx'), 51) # xlOpenXMLWorkbook
$wb.Close($false)
$xl.Quit()
Write-Host 'test.xlsx OK'

Get-ChildItem $dir -File | Select-Object Name, Length | Format-Table -AutoSize
