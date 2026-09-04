# Shared HF / PnL-map helper library (dot-sourced by per-community build scripts). KEEP - reusable.
# COM-BATCHED build: per-row/column writes go through one Range.Value2 assignment (a 2-D object[,])
# instead of cell-by-cell, cutting a build from ~thousands of COM round-trips to a few hundred.
# Output workbook is byte-identical to the cell-by-cell version - only the number of COM hops changes.
#
# CRITICAL COM-binding rule (learned the hard way):
#   When this file is dot-sourced, PowerShell shares ONE cached COM binding for the `.Value2=` setter
#   across every site in the file. A scalar `$cell.Value2 = "text"` binds it to String and then
#   POISONS the array assignment in SetBlock (object[,] -> String InvalidCastException).
#   => Every SCALAR write here goes through SetV (reflection InvokeMember, which binds per-call and
#      never touches the PS adapter cache). SetBlock's direct `$rng.Value2 = $arr` is the ONLY
#      PS-adapter `.Value2=` site, so it is only ever exercised with an array. Do NOT add a bare
#      `$cell.Value2 = <scalar>` anywhere in THIS file - use SetV. (The cache is per-FILE: tested,
#      a top-level `$cell.Value2 = "text"` in a per-deal script does NOT poison this lib's SetBlock,
#      so per-deal scripts may keep using direct `.Value2=` for their own scalars. The rule only
#      bites if a single file mixes scalar `.Value2=` AND its own block `.Value2=array`.)
#   The cache is per-PROPERTY, not just Value2 - it bites ANY COM property a file writes with two
#   different .NET types. Seen 2026-07-27: `ColumnWidth=7.86` then `ColumnWidth=11` (Double then
#   Int32) threw InvalidCastException, same for RowHeight. Fix: give each property one type
#   ([double[]] width arrays, `30.0` not `30`), or route it through SetV-style reflection.
#
# Caller contract (see Build-DealA-HF.ps1):
#   1. $xl = New-Excel              # EnableEvents/ScreenUpdating already off
#   2. open/add the workbook, THEN  $xl.Calculation = -4135   # xlManual - workbook must exist first
#   3. build with the HF-* / Map-* writers below
#   4. $xl.CalculateFull(); $xl.Calculation = -4105   # recalc once, restore automatic
#   5. $wb.Save()
[System.Threading.Thread]::CurrentThread.CurrentCulture=[System.Globalization.CultureInfo]'en-US'
$ACCT = '_(* #,##0_);_(* \(#,##0\);_(* "-"_);_(@_)'
$SLATE=11507343; $WHITE=16777215; $GREY=5263440
$BLUE=16711680; $RED=255; $CREAM=13369343

# Scalar write: reflection InvokeMember binds per-call from the actual arg type (avoids the
# cached-call-site String-then-Double InvalidCast). Use for SINGLE cells only.
# Also unwraps PSObject: a scalar that came out of Sort-Object/Select-Object is PSObject-wrapped and
# reaches Excel as an IDispatch it cannot store -> "Exception from HRESULT: 0x800A03EC" on a write
# that looks perfectly valid (the cell is writable, the value prints as a plain Int32). Verified
# 2026-07-27 on the O&O .RR build: `[int[]]@(... | Sort-Object -Unique)` fixed it at source,
# this unwrap is the backstop.
function SetV($cell,$val){
  if($val -is [psobject] -and $null -ne $val.psobject.BaseObject){ $val = $val.psobject.BaseObject }
  [void]$cell.GetType().InvokeMember("Value2",[System.Reflection.BindingFlags]::SetProperty,$null,$cell,@([object]$val))
}

# Block write: assign a 2-D object[,] to a Range in ONE COM hop via direct property assignment.
# Must be direct (NOT InvokeMember - that enumerates the array -> DISP_E_NOTACOLLECTION / 0x800A03EC).
# This is the only `.Value2=` site in the file (see CRITICAL rule above). A "=..." string element is
# entered as a live formula. Light retry for transient COM rejections.
function SetBlock($rng,$arr){ for($k=0;$k -lt 5;$k++){ try{ $rng.Value2=$arr; return }catch{ if($k -eq 4){throw}; Start-Sleep -Milliseconds 50 } } }

function New-Excel {
  $x=$null; for($i=1;$i -le 6 -and $x -eq $null;$i++){ try{ $x=New-Object -ComObject Excel.Application }catch{ Start-Sleep -Seconds 2 } }
  $x.Visible=$false; $x.DisplayAlerts=$false; $x.ScreenUpdating=$false; $x.EnableEvents=$false
  return $x   # NOTE: set $x.Calculation=-4135 AFTER opening a workbook (it errors with none open)
}

# ---- Concurrency-safe Excel cleanup (opt-in; see memory never-blanket-kill-excel) -------------------
# The blanket windowless kill (`? {$_.MainWindowHandle -eq 0} | Stop-Process`) also terminates ANY other
# live COM automation — a COM Excel is windowless its whole life — so it corrupts a peer run's in-flight
# workbook. These helpers kill ONLY the Excel this PowerShell process spawned, tracked in a per-PID
# lockfile (per-$PID so a peer process never reads our list, nor we theirs). A windowed user Excel is
# never listed. Prefer New-ExcelTracked + Stop-TrackedExcel over the blanket kill; better still, never
# run two COM automations at once.
$script:ExcelPidFile = Join-Path $env:TEMP ("uw_excel_pids_{0}.txt" -f $PID)
function New-ExcelTracked {
  $before = @((Get-Process EXCEL -ErrorAction SilentlyContinue).Id)
  $x = New-Excel
  Start-Sleep -Milliseconds 250
  $mine = @((Get-Process EXCEL -ErrorAction SilentlyContinue).Id) | Where-Object { $_ -notin $before }
  if($mine){ Add-Content -LiteralPath $script:ExcelPidFile -Value $mine }
  return $x
}
function Stop-TrackedExcel {
  if(Test-Path -LiteralPath $script:ExcelPidFile){
    foreach($procId in (Get-Content -LiteralPath $script:ExcelPidFile | Sort-Object -Unique)){
      Get-Process -Id $procId -ErrorAction SilentlyContinue | Stop-Process -Force -ErrorAction SilentlyContinue
    }
    Remove-Item -LiteralPath $script:ExcelPidFile -ErrorAction SilentlyContinue
  }
}
# Clear-OrphanExcel: concurrency-SAFE replacement for the blanket windowless kill. Sweeps ONLY leftover
# Excel that NO live tracked run owns, so it is safe at the top of every run even with sibling automations
# active in other terminals (the 6-terminal 3-pack). How it stays safe:
#   * protected = Excel PIDs listed in per-PS-PID lockfiles (uw_excel_pids_<pspid>.txt) whose owning
#     PowerShell process is STILL ALIVE -> those are a peer run's live Excel; never touched.
#   * stale lockfiles (owner PowerShell dead) are deleted; the Excel they named are fair game.
#   * a windowless Excel is killed only if NOT protected AND older than $GraceSeconds -- the grace window
#     covers the ~250ms in New-ExcelTracked before a just-spawned sibling has written its lockfile.
# A windowed (hand-opened) Excel has MainWindowHandle != 0 and is never a candidate. Returns killed PIDs.
function Clear-OrphanExcel([int]$GraceSeconds = 60){
  $protected = New-Object 'System.Collections.Generic.HashSet[int]'
  Get-ChildItem -LiteralPath $env:TEMP -Filter 'uw_excel_pids_*.txt' -ErrorAction SilentlyContinue | ForEach-Object {
    $ownerAlive = $false
    if($_.Name -match 'uw_excel_pids_(\d+)\.txt'){
      $ownerAlive = [bool](Get-Process -Id ([int]$Matches[1]) -ErrorAction SilentlyContinue)
    }
    if($ownerAlive){
      foreach($line in (Get-Content -LiteralPath $_.FullName -ErrorAction SilentlyContinue)){
        $n = 0; if([int]::TryParse($line.Trim(), [ref]$n)){ [void]$protected.Add($n) }
      }
    } else {
      Remove-Item -LiteralPath $_.FullName -ErrorAction SilentlyContinue   # stale: owning PowerShell gone
    }
  }
  $cut = (Get-Date).AddSeconds(-$GraceSeconds)
  $killed = @()
  Get-Process EXCEL -ErrorAction SilentlyContinue | Where-Object { $_.MainWindowHandle -eq 0 } | ForEach-Object {
    if($protected.Contains($_.Id)){ return }                    # owned by a live tracked run
    $st = $null; try { $st = $_.StartTime } catch { return }    # cannot read age -> do NOT touch
    if($st -lt $cut){ try { $_ | Stop-Process -Force -ErrorAction Stop; $killed += $_.Id } catch {} }
  }
  return $killed
}
function HF-SetMonthly($r,$vals){
  # 12 monthly values + the Q row-total formula, written as one E:Q block
  $a=New-Object 'object[,]' 1,13
  for($i=0;$i -lt 12;$i++){ $a[0,$i]=[double]$vals[$i] }
  $a[0,12]="=SUM(E${r}:P${r})"
  $rng=$script:ws.Range($script:ws.Cells($r,5),$script:ws.Cells($r,17))
  SetBlock $rng $a
  $rng.NumberFormat=$ACCT
}
function HF-Band($r,[bool]$banner){
  $rng=$script:ws.Range($script:ws.Cells($r,3),$script:ws.Cells($r,17))
  if($banner){ $rng.Interior.Color=$SLATE; $rng.Font.Bold=$true; $rng.Font.Color=$WHITE }
  else { $rng.Font.Bold=$true; $rng.Borders(8).LineStyle=1; $rng.Borders(8).Weight=2; $rng.Borders(9).LineStyle=1; $rng.Borders(9).Weight=2 }
}
function HF-Line($code,$label,$vals){
  $script:r++; $row=$script:r
  # S (code) + T (label) in one write
  $st=New-Object 'object[,]' 1,2
  $st[0,0]=[string]$code; $st[0,1]=[string]$label
  SetBlock ($script:ws.Range($script:ws.Cells($row,19),$script:ws.Cells($row,20))) $st
  $script:ws.Cells($row,3).Formula="=IF(S$row="""",T$row,S$row&""  -  ""&T$row)"
  $cC=$script:ws.Cells($row,3); $cC.Font.Name="Arial"; $cC.Font.Size=10
  HF-SetMonthly $row $vals
  return $row
}
function HF-Total($label,$memberRows,$marker){
  $script:r++; $row=$script:r
  if($marker){ SetV $script:ws.Cells($row,1) $marker }
  SetV $script:ws.Cells($row,3) ([string]$label)
  $a=New-Object 'object[,]' 1,13
  foreach($cc in 5..17){ $L=[char](64+$cc); $parts=($memberRows | ForEach-Object { "$L$_" }) -join "+"; if($parts -eq ""){$parts="0"}; $a[0,($cc-5)]="=$parts" }
  $rng=$script:ws.Range($script:ws.Cells($row,5),$script:ws.Cells($row,17))
  SetBlock $rng $a
  $rng.NumberFormat=$ACCT
  HF-Band $row $false
  return $row
}
function HF-SumRows($label,$rows,$marker){
  $script:r++; $row=$script:r
  if($marker){ SetV $script:ws.Cells($row,1) $marker }
  SetV $script:ws.Cells($row,3) ([string]$label)
  $a=New-Object 'object[,]' 1,13
  foreach($cc in 5..17){ $L=[char](64+$cc); $parts=($rows | ForEach-Object { "$L$_" }) -join "+"; if($parts -eq ""){$parts="0"}; $a[0,($cc-5)]="=$parts" }
  $rng=$script:ws.Range($script:ws.Cells($row,5),$script:ws.Cells($row,17))
  SetBlock $rng $a
  $rng.NumberFormat=$ACCT
  HF-Band $row $false
  return $row
}
function HF-Header($name, $dates){
  $c=$script:ws.Cells(2,3); SetV $c $name; $c.Font.Name="Tahoma"; $c.Font.Size=8; $c.Font.Color=$GREY; $c.HorizontalAlignment=-4108
  $c=$script:ws.Cells(3,3); SetV $c "Statement (12 months)"; $c.Font.Name="Tahoma"; $c.Font.Size=12; $c.Font.Bold=$true; $c.HorizontalAlignment=-4108
  $c=$script:ws.Cells(4,3); SetV $c ("Period = "+$dates[0].ToString("MMM yyyy")+" - "+$dates[11].ToString("MMM yyyy")); $c.HorizontalAlignment=-4108
  SetV $script:ws.Cells(5,3) "Book = Accrual"
}
function HF-RevBanner($dates){
  SetV $script:ws.Cells(6,3) "REVENUE"; SetV $script:ws.Cells(6,1) "x"
  $a=New-Object 'object[,]' 1,12
  for($i=0;$i -lt 12;$i++){ $a[0,$i]=[double]$dates[$i].ToOADate() }
  $rng=$script:ws.Range($script:ws.Cells(6,5),$script:ws.Cells(6,16))
  SetBlock $rng $a
  $rng.NumberFormat="mmm-yy"
  SetV $script:ws.Cells(6,17) "Total"; HF-Band 6 $true
}
function HF-Finish(){
  $script:ws.Columns.Item(3).ColumnWidth=58; $script:ws.Columns.Item(4).ColumnWidth=2
  $script:ws.Range($script:ws.Cells(1,5),$script:ws.Cells(1,17)).EntireColumn.ColumnWidth=11   # E:Q in one set
  $script:ws.Columns.Item(19).Hidden=$true; $script:ws.Columns.Item(20).Hidden=$true
  $script:wb.Windows.Item(1).DisplayGridlines=$false
}
# $TAGS: parsed at dot-source time from the canonical vocabulary file, so the reference stays the
# single source of truth (CLAUDE.md: "use those tags, never an invented or renamed one"). Tags are
# the `- ` bullets AFTER the first `## ` section heading — the file's usage-rules bullets precede it.
# Order is preserved (it is the List-of-Tags display order on the Mapping sheet).
function Get-MasterTags([string]$path){
  if(-not (Test-Path -LiteralPath $path)){ throw "Master tag list not found: $path" }
  $tags=@(); $inTags=$false
  foreach($ln in (Get-Content -LiteralPath $path)){
    if($ln -match '^## '){ $inTags=$true; continue }
    if($inTags -and $ln -match '^- (.+)$'){ $tags += $matches[1].Trim() }
  }
  if($tags.Count -lt 60 -or $tags[-1] -ne "Ignore"){ throw "master-tags.md parse sanity failed ($($tags.Count) tags, last='$($tags[-1])') - has the file changed shape?" }
  return $tags
}
$TAGS = Get-MasterTags (Join-Path $PSScriptRoot '..\..\.claude\skills\pnl-mapping\References\master-tags.md')
function Map-Header(){
  SetV $script:mw.Cells(2,1) "Check"; SetV $script:mw.Cells(2,3) "Tag"; SetV $script:mw.Cells(2,4) "Questions & Comments"; SetV $script:mw.Cells(2,5) "List of Tags"
  foreach($c in 1,3,4,5){ $h=$script:mw.Cells(2,$c); $h.Font.Name="Aptos Narrow"; $h.Font.Size=9; $h.Font.Bold=$true; $h.HorizontalAlignment=-4108 }
  $script:mw.Cells(2,2).Font.Name="Aptos Narrow"; $script:mw.Cells(2,2).Font.Size=11
  foreach($c in 1,3,4,5){ $bd=$script:mw.Cells(2,$c).Borders(9); $bd.LineStyle=1; $bd.Weight=2; $bd.ColorIndex=-4105 }
  $script:mw.Cells(2,2).Borders(9).LineStyle=-4142
  $script:mr=3
}
function Map-Struct($text){ $c=$script:mw.Cells($script:mr,3); SetV $c ([string]$text); $c.Font.Name="Tahoma"; $c.Font.Size=8; $c.Font.Bold=$true; $script:mr++ }
function Map-Line($raw,$tag,$flag){
  $r=$script:mr
  # A (Check formula) + B (tag) + C (raw) + D (flag) in one write; "=..." enters as a formula
  $v=New-Object 'object[,]' 1,4
  $v[0,0]='=IF(COUNTIF($E:$E,B'+$r+'),"Yes","No")'
  $v[0,1]=[string]$tag
  $v[0,2]=[string]$raw
  $v[0,3]=[string]$flag
  SetBlock ($script:mw.Range($script:mw.Cells($r,1),$script:mw.Cells($r,4))) $v
  $script:mw.Cells($r,1).Font.Name="Aptos Narrow"; $script:mw.Cells($r,1).Font.Size=9
  $b=$script:mw.Cells($r,2); $b.Font.Name="Aptos Narrow"; $b.Font.Size=9; $b.Interior.Color=$CREAM
  if($tag -eq "Ignore"){ $b.Font.Color=$RED } else { $b.Font.Color=$BLUE }
  $script:mw.Cells($r,3).Font.Name="Tahoma"; $script:mw.Cells($r,3).Font.Size=8
  if($flag -ne ""){ $script:mw.Cells($r,4).Font.Name="Tahoma"; $script:mw.Cells($r,4).Font.Size=8 }
  $script:mr++
}
function Map-Finish(){
  # 63-tag List of Tags in one vertical write, then range-level font/fill, then redden 'Ignore'
  $n=$script:TAGS.Count
  $a=New-Object 'object[,]' $n,1
  for($i=0;$i -lt $n;$i++){ $a[$i,0]=[string]$script:TAGS[$i] }
  $rng=$script:mw.Range($script:mw.Cells(3,5),$script:mw.Cells(2+$n,5))
  SetBlock $rng $a
  $rng.Font.Name="Aptos Narrow"; $rng.Font.Size=9; $rng.Font.Color=$BLUE; $rng.Interior.Color=$CREAM
  for($i=0;$i -lt $n;$i++){ if($script:TAGS[$i] -eq "Ignore"){ $script:mw.Cells(3+$i,5).Font.Color=$RED } }
  $script:mw.Columns.Item(1).ColumnWidth=8.43; $script:mw.Columns.Item(2).ColumnWidth=19.71; $script:mw.Columns.Item(3).ColumnWidth=44.86; $script:mw.Columns.Item(4).ColumnWidth=17.43; $script:mw.Columns.Item(5).ColumnWidth=21.14
  $script:wb.Windows.Item(1).DisplayGridlines=$false
}
