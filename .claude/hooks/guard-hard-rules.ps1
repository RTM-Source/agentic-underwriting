# PreToolUse hook - enforces the hard rules CLAUDE.md can only request.
# Input: hook JSON on stdin ({tool_name, tool_input:{command}}). Exit 0 = allow, exit 2 = block
# (stderr is fed back to Claude as the reason). ASCII only: PS 5.1 reads BOM-less UTF-8 as ANSI.

$raw = [Console]::In.ReadToEnd()
try { $j = $raw | ConvertFrom-Json } catch { exit 0 }

$cmd = ""
if ($j -and $j.tool_input -and $j.tool_input.command) { $cmd = [string]$j.tool_input.command }
if ($cmd -eq "") { exit 0 }

# HARD RULE 1 - kill ONLY the Excel this run spawned; never a windowless SWEEP (memory never-blanket-kill-excel).
# The old "MainWindowHandle -eq 0 is exempt" carve-out is GONE: a COM Excel is windowless its ENTIRE life,
# so `? { $_.MainWindowHandle -eq 0 } | Stop-Process` also force-kills a CONCURRENT run's live Excel in
# another terminal (the 6-terminal 3-pack corrupts mid-.Save() with no crash). Sanctioned forms only:
# Clear-OrphanExcel (concurrency-safe startup sweep) / Stop-TrackedExcel / New-ExcelTracked (lib/HF-Build-Lib.ps1),
# a lockfile-scoped kill (mentions uw_excel_pids), or a targeted Stop-Process -Id <pid>.
if ($cmd -match '(?i)\b(Stop-Process|taskkill|pkill|killall|kill)\b' -and
    $cmd -match '(?i)excel' -and
    $cmd -notmatch '(?i)(Stop-TrackedExcel|Clear-OrphanExcel|uw_excel_pids|-Id\b)') {
  [Console]::Error.WriteLine('BLOCKED by guard-hard-rules hook: unsafe Excel kill. A COM Excel is windowless its whole life, so a MainWindowHandle-eq-0 sweep also kills a CONCURRENT run''s Excel in another terminal. Dot-source lib/HF-Build-Lib.ps1 and use Clear-OrphanExcel (safe startup sweep) + New-ExcelTracked/Stop-TrackedExcel, or target a specific Stop-Process -Id <pid>. (memory never-blanket-kill-excel)')
  exit 2
}

# HARD RULE 2 - COM SaveAs to a OneDrive path hangs silently forever (memory word-com-saveas-hang).
if ($cmd -match '(?i)\.SaveAs\(' -and $cmd -match '(?i)OneDrive') {
  [Console]::Error.WriteLine('BLOCKED by guard-hard-rules hook: COM SaveAs to a OneDrive path hangs silently with no error. Copy-Item a seed file to the target path first, then Open and $wb.Save() in place - or build in %TEMP% and Copy-Item out. (memory word-com-saveas-hang)')
  exit 2
}

exit 0
