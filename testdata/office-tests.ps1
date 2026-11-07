# Office-path end-to-end tests, run OUTSIDE the DSH sandbox (medium integrity) via WMI.
$ErrorActionPreference = 'Continue'
$log  = 'D:\Codes\AI Agents\Deepseek Harness\fc-src\testdata\office-tests.log'
$exe  = 'D:\Codes\AI Agents\Deepseek Harness\fc-src\bin\Debug\net10.0-windows\FileConverter.exe'
$td   = 'D:\Codes\AI Agents\Deepseek Harness\fc-src\testdata'
$root = Join-Path $td 'out'
$env:FILECONVERTER_OFFICE_TIMEOUT_MS = '180000'

function Log($m) { Add-Content -Path $log -Value $m }
Set-Content -Path $log -Value "START $(Get-Date -Format o)"

$pass = 0; $fail = 0
function RunTest($name, $input, $extraArgs, $checkFile, $magic) {
    $outDir = Join-Path $root ([System.IO.Path]::GetFileName(($extraArgs | Select-Object -Last 1)))
    $cmdline = "`"$exe`" `"$input`" $extraArgs"
    $o = cmd /c "$cmdline 2>&1" | Out-String
    $code = $LASTEXITCODE
    $ok = ($code -eq 0)
    if ($ok) {
        $ok = (Test-Path $checkFile) -and ((Get-Item $checkFile).Length -gt 0)
        if ($ok -and $magic) {
            $b = [System.IO.File]::ReadAllBytes($checkFile)
            $head = [System.Text.Encoding]::ASCII.GetString($b[0..4])
            $ok = $head -like "$magic*"
        }
    }
    $size = if (Test-Path $checkFile) { (Get-Item $checkFile).Length } else { 0 }
    if ($ok) { $script:pass++; Log "PASS $name ($size bytes)" }
    else {
        $script:fail++
        $detail = ($o.Trim() -replace "`r`n", ' | ')
        if ($detail.Length -gt 300) { $detail = $detail.Substring(0, 300) }
        Log "FAIL $name (exit=$code, size=$size): $detail"
    }
}

RunTest 'xlsx->pdf' "$td\test.xlsx" "--out `"$root\xls`"" "$root\xls\test.pdf" '%PDF'
RunTest 'docx->pdf' "$td\test.docx" "--out `"$root\word`"" "$root\word\test.pdf" '%PDF'
if (Test-Path "$root\word\test.pdf") {
    RunTest 'pdf->docx' "$root\word\test.pdf" "--to docx --out `"$root\word2`"" "$root\word2\test.docx" 'PK'
} else {
    Log 'SKIP pdf->docx (no input pdf)'
}

Log "DONE pass=$pass fail=$fail $(Get-Date -Format o)"
