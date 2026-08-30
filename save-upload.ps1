# save-upload.ps1  (helper: persist an uploaded PDF into its paper folder)
#  Purpose: take a base64-encoded temp text file and decode it into a real PDF,
#           creating the target folder if needed. Then remove the temp base64 file.
#  Usage (via terminal):
#      powershell -NoProfile -ExecutionPolicy Bypass -File save-upload.ps1 -B64File "...\_upload_tmp.txt" -OutPdf "...\<folder>\<name>.pdf"

param([Parameter(Mandatory=$true)][string]$B64File, [Parameter(Mandatory=$true)][string]$OutPdf)
$ErrorActionPreference = "Stop"
$dir = Split-Path $OutPdf -Parent
New-Item -ItemType Directory -Path $dir -Force | Out-Null
$bytes = [Convert]::FromBase64String((Get-Content -Raw $B64File))
[IO.File]::WriteAllBytes($OutPdf, $bytes)
if (Test-Path $OutPdf) {
  Write-Host ("SAVED " + (Get-Item $OutPdf).Length)
} else {
  Write-Host "SAVE_FAILED"
}
Remove-Item -Force -ErrorAction SilentlyContinue $B64File
