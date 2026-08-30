# make-ppt-view.ps1
#  Purpose: convert a .pptx into a printable view (<PPT名>.pdf) + its built-in speaker notes (<PPT名>.md),
#           placed in the same paper folder. Then load the PPT in window1 -> page 3 (rehearsal).
#  Usage (in project root):
#      powershell -ExecutionPolicy Bypass -File make-ppt-view.ps1 -Pptx "AgentPoison\AgentPoison_GroupTalk.pptx"
#  Deps: Microsoft PowerPoint installed (for PDF), python + python-pptx (for notes).
#  Note: notes are extracted FIRST (python reads the pptx), then PowerPoint converts to PDF -- to avoid file locks.

param([Parameter(Mandatory=$true)][string]$Pptx)
$ErrorActionPreference = "Stop"
if (-not (Test-Path $Pptx)) { throw "File not found: $Pptx" }
$Pptx = (Resolve-Path $Pptx).Path
$dir  = Split-Path $Pptx -Parent
$base = [System.IO.Path]::GetFileNameWithoutExtension($Pptx)
$pdf   = Join-Path $dir ($base + '.pdf')   # view, named after the ppt
$notes = Join-Path $dir ($base + '.md')    # notes, named after the ppt

Write-Host "== 1/2 Extract speaker notes -> <PPT>.md via python-pptx =="
# PPTX/notes paths are non-ASCII; pass them via a temp JSON so PS5.1 -> python arg encoding is not an issue.
$cfg = Join-Path $env:TEMP 'ppt_cfg.json'
[System.IO.File]::WriteAllText($cfg, (ConvertTo-Json @{ pptx = $Pptx; out = $notes } -Compress), (New-Object System.Text.UTF8Encoding($false)))
$py = @'
import sys, io, json
cfg = json.load(open(sys.argv[1], encoding='utf-8'))
from pptx import Presentation
prs = Presentation(cfg['pptx'])
out = []
for i, s in enumerate(prs.slides, 1):
    txt = ""
    try:
        if s.has_notes_slide and s.notes_slide.notes_text_frame is not None:
            txt = s.notes_slide.notes_text_frame.text.strip()
    except Exception:
        txt = ""
    out.append("## Slide %d\n" % i)
    out.append((txt if txt else "(no notes on this slide)") + "\n")
    out.append("")
io.open(cfg['out'], 'w', encoding='utf-8').write("\n".join(out))
print("notes ok")
'@
$tmpPy = Join-Path $env:TEMP 'extract_notes_tmp.py'
[System.IO.File]::WriteAllText($tmpPy, $py, (New-Object System.Text.UTF8Encoding($false)))
python $tmpPy $cfg
Write-Host ("Notes done: {0} ({1} bytes)" -f $notes, (Get-Item $notes).Length)

Write-Host "== 2/2 Convert PPT -> PDF via PowerPoint =="
$ppt = New-Object -ComObject PowerPoint.Application
try {
  $pres = $ppt.Presentations.Open($Pptx, $true, $false, $false)   # ReadOnly, Untitled, WithWindow=false
  $pres.SaveAs($pdf, 32)                                          # 32 = ppSaveAsPDF
  $pres.Close()
} finally {
  $ppt.Quit()
}
Start-Sleep -Milliseconds 400
Write-Host ("PDF done: {0} ({1} bytes)" -f $pdf, (Get-Item $pdf).Length)
Write-Host "Done. Back to window1 -> page 3, load that PPT name to see its slides + notes."
