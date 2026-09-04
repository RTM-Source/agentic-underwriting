# .RR build engine (v1, 2026-07-14) - see .claude/skills/rr-formatting/scripts/ENGINE-SPEC.md
# Dot-source, then: Invoke-RRBuild -ConfigPath <config.json>
#
# Split of responsibility (ENGINE-SPEC.md):
#   AGENT  - parses the raw export, resolves judgment calls, emits config.json + units.csv.
#   ENGINE - everything right of blockStart: unit block, tiering helper, main summary, Second
#            Residents, Rent Adjustments, Check, move-in boxes, Notes, reconcile block, the full
#            formatting contract, gates, rule-21 build-in-place flow. The engine IS the golden,
#            made executable. It never re-derives judgment - it trusts units.csv + config.json.
#
# House style follows lib/HF-Build-Lib.ps1 (dot-sourced below for SetV/SetBlock/New-Excel/tracking).
# Locked-rule citations appear inline as "rule N".

. (Join-Path $PSScriptRoot "HF-Build-Lib.ps1")

# ---- column-letter helpers (rule: case-collision - distinct, alias-safe names) ----------------
function RRColL([int]$n){ $rs=""; while($n -gt 0){ $rm=($n-1)%26; $rs=[char](65+$rm)+$rs; $n=[int](($n-$rm-1)/26) }; return $rs }
function RRColIdx([string]$L){ $ri=0; foreach($ch in $L.ToCharArray()){ $ri = $ri*26 + ([int][char]$ch - 64) }; return $ri }
function RRAbsR($col,$r1,$r2){ $rl=RRColL $col; return ('${0}${1}:${0}${2}' -f $rl,$r1,$r2) }
function RRNumOf($x){ if($null -eq $x){return 0.0}; $rt="$x" -replace '[,\$\s]',''; if($rt -eq ''){return 0.0}; try{ return [double]$rt }catch{ return 0.0 } }

# ---- style constants (formatting.md - BGR longs, not RGB) --------------------------------------
$script:rrYEL=13434879; $script:rrNAV=4990985; $script:rrWHT=16777215; $script:rrBLU=16711680
$script:rrORG=14745600; $script:rrRED=255; $script:rrBLK=0
$script:rrACC='_(* #,##0_);_(* (#,##0);_(* "-"_);_(@_)'
$script:rrPCT='_(#,##0.0%_);(#,##0.0%);_("' + [char]0x2013 + '"_)_%;_(@_)_%'
$script:rrDOLLAR='$#,##0_);($#,##0)'
$script:rrErrCodes=@(-2146826281,-2146826273,-2146826265,-2146826259,-2146826252,-2146826246,-2146826288)

function RRStyleYellow($rng){ $rng.Interior.Color=$script:rrYEL; $rng.Font.Name='Calibri'; $rng.Font.Size=11; $rng.Font.Bold=$true; $rng.Font.Color=$script:rrBLK; $rng.WrapText=$true; $rng.HorizontalAlignment=1 }
function RRStyleNavy($rng){ $rng.Interior.Color=$script:rrNAV; $rng.Font.Name='Arial'; $rng.Font.Size=10; $rng.Font.Bold=$true; $rng.Font.Color=$script:rrWHT; $rng.WrapText=$true; $rng.HorizontalAlignment=-4108 }
# Sub-block title rows (Second Residents / Rent Adjustments / Check / Notes) are PLAIN BOLD, NO fill -
# never navy-bannered (formatting.md §5, golden O&O Community 1-RR_v23 verified). Black-bold, except Check = red-bold.
function RRStylePlain($rng,$color){ $rng.Interior.Color=$script:rrWHT; $rng.Font.Name='Arial'; $rng.Font.Size=10; $rng.Font.Bold=$true; $rng.Font.Color=$color; $rng.WrapText=$false }
function RRTotalBorder($rng){ $rng.Font.Bold=$true; $rng.Borders.Item(8).LineStyle=1; $rng.Borders.Item(8).Weight=2; $rng.Borders.Item(9).LineStyle=-4142 }

# =================================================================================================
function Invoke-RRBuild {
  param([Parameter(Mandatory=$true)][string]$ConfigPath)

  [System.Threading.Thread]::CurrentThread.CurrentCulture=[System.Globalization.CultureInfo]'en-US'
  $ErrorActionPreference = "Stop"

  if(-not (Test-Path -LiteralPath $ConfigPath)){ throw "config not found: $ConfigPath" }
  $cfg = Get-Content -Raw -LiteralPath $ConfigPath | ConvertFrom-Json
  $cfgDir = Split-Path -Parent (Resolve-Path -LiteralPath $ConfigPath)

  $unitsCsvPath = if($cfg.PSObject.Properties.Name -contains 'unitsCsvPath' -and $cfg.unitsCsvPath){ $cfg.unitsCsvPath } else { Join-Path $cfgDir "units.csv" }
  if(-not (Test-Path -LiteralPath $unitsCsvPath)){ throw "units.csv not found: $unitsCsvPath" }
  $units = Import-Csv -LiteralPath $unitsCsvPath | Sort-Object { [int]$_.anchorRow }
  if(-not $units -or $units.Count -eq 0){ throw "units.csv is empty" }

  # ---- resolve source (OneDrive may move files mid-session - re-resolve by leaf name) ----------
  $src = $cfg.sourcePath
  if(-not (Test-Path -LiteralPath $src)){
    $leaf = Split-Path $src -Leaf
    $found = Get-ChildItem -Path (Split-Path $src -Parent) -Filter $leaf -ErrorAction SilentlyContinue
    if($found){ $src = $found[0].FullName } else { throw "source not found: $($cfg.sourcePath)" }
  }
  if(-not (Test-Path -LiteralPath $cfg.dealFolder)){ New-Item -ItemType Directory -Path $cfg.dealFolder -Force | Out-Null }

  $emd = [char]0x2014
  $finalName    = "{0}{1}RR_v{2}.xlsx" -f $cfg.community,$emd,$cfg.version
  $buildingName = "~building {0} {1}" -f $emd,$finalName
  $finalPath    = Join-Path $cfg.dealFolder $finalName
  $buildingPath = Join-Path $cfg.dealFolder $buildingName

  # ---- rule 21 step 1: filesystem copy under the building marker -------------------------------
  Copy-Item -LiteralPath $src -Destination $buildingPath -Force

  $xl=$null; $wb=$null; $success=$false
  $result = [ordered]@{
    community=$cfg.community; finalPath=$finalPath; blockStart=$null; headerRow=$null
    unitCount=$units.Count; reconcile=$null; errorCount=$null; errorCells=@()
    checkValues=$null; headersPresent=$null; saveOk=$false
  }

  try {
    $xl = New-ExcelTracked
    $wb = $null
    for($k=0;$k -lt 6 -and -not $wb;$k++){ try{ $wb=$xl.Workbooks.Open($buildingPath,0,$false) }catch{ Start-Sleep -Milliseconds 500 } }
    if(-not $wb){ throw "could not open $buildingPath" }
    $xl.Calculation=-4135   # manual while writing
    $ws = $wb.Worksheets.Item(1)

    # ---- geometry is RELATIVE (rule 15) ----
    $ur = $ws.UsedRange
    $lastRawCol = $ur.Column + $ur.Columns.Count - 1
    $blockStart = $lastRawCol + 1
    $hdrRow = [int]$cfg.headerRow
    $result.blockStart = $blockStart; $result.headerRow = $hdrRow

    $unitFirst = [int]$cfg.unitRowFirst; $unitLast = [int]$cfg.unitRowLast
    $daysInPeriod = if($cfg.PSObject.Properties.Name -contains 'daysInPeriod'){ [double]$cfg.daysInPeriod } else { 30.0 }
    $deprorate = [bool]$cfg.deprorate

    # =============================================================================================
    # PER-UNIT BLOCK (17 cols, offsets +0..+16 from blockStart)
    # =============================================================================================
    $wOrder=$blockStart+0;  $wUnit=$blockStart+1;  $wUType=$blockStart+2; $wPSP=$blockStart+3
    $wCareT=$blockStart+4;  $wSqFt=$blockStart+5;  $wRType=$blockStart+6; $wCap=$blockStart+7
    $wOccC=$blockStart+8;   $wMoveIn=$blockStart+9;$wMkt=$blockStart+10;  $wRent=$blockStart+11
    $wCareF=$blockStart+12; $wSec=$blockStart+13;  $wUCode=$blockStart+14;$wBCode=$blockStart+15
    $wGCode=$blockStart+16

    $LUnit=RRColL $wUnit; $LUType=RRColL $wUType; $LCareT=RRColL $wCareT; $LSqFt=RRColL $wSqFt
    $LCap=RRColL $wCap; $LOccC=RRColL $wOccC; $LMoveIn=RRColL $wMoveIn

    $cMkt = $cfg.colMap.market; $cAct = $cfg.colMap.actual; $cUnit = $cfg.colMap.unit
    $cCare = if($cfg.colMap.PSObject.Properties.Name -contains 'care'){ $cfg.colMap.care } else { $null }
    $cSec  = if($cfg.colMap.PSObject.Properties.Name -contains 'secOcc'){ $cfg.colMap.secOcc } else { $null }

    # Days-in-period absolute ref (written later, at the summary's value column, rTOT+22) - the
    # de-proration factor needs a forward reference, so compute the address now.
    $sCol = $blockStart+23   # main summary first column
    # row plan must be computed before we know rTOT (needed for $DAYSREF) - build it now.
    # ---- dynamic row plan (populated care levels get tier rows; absent levels get a 3-row
    #      placeholder then their Total row - matches the verified Deal C acceptance ref) ----
    $tiersByCare = [ordered]@{ IL=@(); AL=@(); MC=@() }
    foreach($t in $cfg.tiers){ $care = ($t.group -split '\|')[0].Trim(); $tiersByCare[$care] += ,$t }

    $cursor = $hdrRow+1
    $levelInfo = @{}
    $careOrder = @('IL','AL','MC')
    for($cix=0; $cix -lt $careOrder.Count; $cix++){
      $care = $careOrder[$cix]
      $tiers = $tiersByCare[$care]
      # NOTE (fix 2026-07-29, Deal E): $rows MUST be reset to @() on the
      # absent-level branch. PowerShell loop locals persist across iterations, so the old code left
      # an absent level holding either $null (when the FIRST care level is absent -> "Cannot index
      # into a null array" downstream, which is every AL/MC-only deal) or the PREVIOUS level's row
      # numbers (silently writing phantom move-in-box formulas onto a populated level's rows).
      $rows=@()
      if($tiers.Count -gt 0){
        foreach($t in $tiers){ $rows+=$cursor; $cursor++ }
        $totalRow=$cursor; $cursor++
      } else {
        $cursor += 3
        $totalRow=$cursor; $cursor++
      }
      $levelInfo[$care] = @{ tiers=$tiers; rows=$rows; totalRow=$totalRow }
      # blank-row separator at EVERY care-group transition (formatting.md §5 row plan: Total IL/16=blank/
      # AL rows 17.., Total AL/25=blank/MC rows.., Total MC/29=blank/Total) - not just before the grand Total.
      if($cix -lt $careOrder.Count-1){ $cursor += 1 }
    }
    $cursor += 1
    $rTOT = $cursor

    $daysRow = $rTOT+22
    $LDaysVal = RRColL ($sCol+1)
    $DAYSREF = '$'+$LDaysVal+'$'+$daysRow

    # ---- build the dense unit-block write array in ONE bulk range (rule 6: no per-cell loops) ----
    $minRow = ($units | ForEach-Object {[int]$_.anchorRow} | Measure-Object -Minimum).Minimum
    $maxRow = ($units | ForEach-Object {[int]$_.anchorRow} | Measure-Object -Maximum).Maximum
    $nRows = $maxRow-$minRow+1
    $blk = New-Object 'object[,]' $nRows,17
    for($i=0;$i -lt $nRows;$i++){ for($j=0;$j -lt 17;$j++){ $blk[$i,$j]="" } }

    $ordN=0
    foreach($u in $units){
      $a=[int]$u.anchorRow; $off=$a-$minRow; $ordN++
      $cap=[int]$u.capacity; $occ=[int]$u.occupied
      $leaves = @($u.leafRows -split ';' | Where-Object { $_ -ne '' })
      $moveInDay = 0
      $moveOA = ""
      if($u.moveIn -and $u.moveIn.Trim() -ne ""){
        $dt=[datetime]::Parse($u.moveIn,[System.Globalization.CultureInfo]::InvariantCulture)
        $moveOA=[double]$dt.ToOADate(); $moveInDay=$dt.Day
      }

      $blk[$off,0] = [double]$ordN                                             # Order
      $blk[$off,1] = "={0}{1}" -f $cUnit,$a                                    # Unit
      $blk[$off,2] = [string]$u.unitType                                       # Unit Type (hardcode - source-agnostic)
      $blk[$off,3] = [string]$u.psp                                            # P/SP
      $blk[$off,4] = [string]$u.careType                                       # Care Type (hardcode)
      $blk[$off,5] = [double]$u.sqft                                           # SqFt
      $blk[$off,6] = ""                                                        # Resident Type (reserved)
      $blk[$off,7] = [double]$cap                                              # Capacity (STRUCTURAL, rule 2 - from units.csv, not SqFt>0)
      $blk[$off,8] = [double]$occ                                              # Occupancy (from units.csv)
      $blk[$off,9] = $(if($moveOA -ne ""){ $moveOA } else { "" })              # Move-in Date

      $sumTerm = { param($col) '=' + (($leaves | ForEach-Object { 'IFERROR(VALUE({0}{1}),0)' -f $col,$_ }) -join '+') }

      if($cap -eq 0){
        # structural non-bed (S-fold row, rule 17): entire row zeroed
        $blk[$off,10]=0.0; $blk[$off,11]=0.0; $blk[$off,12]=0.0; $blk[$off,13]=0.0
      } else {
        $blk[$off,10] = & $sumTerm $cMkt                                        # Market Rent - always populates (bed exists)
        if($occ -eq 0){
          $blk[$off,11]=0.0                                                     # vacant -> literal 0 (rule 17)
        } else {
          $rentSum = & $sumTerm $cAct
          if($deprorate -and $moveInDay -gt 1){
            $blk[$off,11] = "=({0})*{1}/({1}-DAY({2}{3})+1)" -f $rentSum.TrimStart('='),$DAYSREF,$LMoveIn,$a
          } else { $blk[$off,11] = $rentSum }
        }
        if(-not $cCare){ $blk[$off,12]=0.0 }
        elseif($occ -eq 0){ $blk[$off,12]=0.0 }
        else {
          $careSum = & $sumTerm $cCare
          if($deprorate -and $moveInDay -gt 1){
            $blk[$off,12] = "=({0})*{1}/({1}-DAY({2}{3})+1)" -f $careSum.TrimStart('='),$DAYSREF,$LMoveIn,$a
          } else { $blk[$off,12] = $careSum }
        }
        if(-not $cSec){ $blk[$off,13]=0.0 } elseif($occ -eq 0){ $blk[$off,13]=0.0 } else { $blk[$off,13] = & $sumTerm $cSec }
      }
      $blk[$off,14] = '=IF({0}{1}="","",{2}{1}&" | "&{3}{1}&" | "&{4}{1})' -f $LUnit,$a,$LCareT,$LUType,$LSqFt   # Unit Code
      $blk[$off,15] = [string]$u.basicUnitCode                                 # Basic Unit Code (hardcode)
      $blk[$off,16] = [string]$u.groupUnitCode                                 # Group Unit Code (hardcode - the join key)
    }
    $blkRng = $ws.Range($ws.Cells($minRow,$blockStart),$ws.Cells($maxRow,$blockStart+16))
    SetBlock $blkRng $blk

    # ---- unit-block header (yellow) ----
    $hdrLabels=@('Order','Unit','Unit Type','P/SP','Care Type','SqFt','Resident Type','Capacity','Occupancy','Move in Date','Market Rent','In-Place Rent','In-Place Care','2nd Resident','Unit Code','Basic Unit Code','Group Unit Code')
    $hdrArr = New-Object 'object[,]' 1,17
    for($c=0;$c -lt 17;$c++){ $hdrArr[0,$c]=$hdrLabels[$c] }
    SetBlock ($ws.Range($ws.Cells($hdrRow,$blockStart),$ws.Cells($hdrRow,$blockStart+16))) $hdrArr
    RRStyleYellow ($ws.Range($ws.Cells($hdrRow,$blockStart),$ws.Cells($hdrRow,$blockStart+16)))
    $ws.Rows.Item($hdrRow).RowHeight=30

    # ---- unit-block number formats + all-black font (rule 13 - per-unit block not colour-coded) ----
    $ws.Range($ws.Cells($minRow,$wSqFt),$ws.Cells($maxRow,$wSqFt)).NumberFormat='#,##0'
    $ws.Range($ws.Cells($minRow,$wCap),$ws.Cells($maxRow,$wOccC)).NumberFormat='0'
    $ws.Range($ws.Cells($minRow,$wMoveIn),$ws.Cells($maxRow,$wMoveIn)).NumberFormat='m/d/yyyy'
    $ws.Range($ws.Cells($minRow,$wMkt),$ws.Cells($maxRow,$wSec)).NumberFormat=$script:rrDOLLAR
    $ws.Range($ws.Cells($minRow,$blockStart),$ws.Cells($maxRow,$blockStart+16)).Font.Color=$script:rrBLK
    $ws.Range($ws.Cells($minRow,1),$ws.Cells($maxRow,1)).EntireRow.RowHeight=15

    # ---- absolute ranges the analysis grouping uses ----
    $R_SQ =RRAbsR $wSqFt  $unitFirst $unitLast
    $R_CAP=RRAbsR $wCap   $unitFirst $unitLast
    $R_OCC=RRAbsR $wOccC  $unitFirst $unitLast
    $R_MI =RRAbsR $wMoveIn $unitFirst $unitLast
    $R_MK =RRAbsR $wMkt   $unitFirst $unitLast
    $R_RT =RRAbsR $wRent  $unitFirst $unitLast
    $R_CR =RRAbsR $wCareF $unitFirst $unitLast
    $R_SE =RRAbsR $wSec   $unitFirst $unitLast
    $R_GRP=RRAbsR $wGCode $unitFirst $unitLast
    $R_CT =RRAbsR $wCareT $unitFirst $unitLast
    $R_UN =RRAbsR $wUnit  $unitFirst $unitLast
    $LOrder=RRColL $wOrder; $LMkt=RRColL $wMkt; $LRent=RRColL $wRent; $LCareF=RRColL $wCareF; $LSec=RRColL $wSec; $LGCode=RRColL $wGCode

    # =============================================================================================
    # TIERING HELPER (+19..+21) - yellow header, human reference (analysis-table.md #0)
    # =============================================================================================
    $tCol=$blockStart+19
    $tHdr = New-Object 'object[,]' 1,3
    $tHdr[0,0]='Unit Sqft'; $tHdr[0,1]='Count'; $tHdr[0,2]='Group Unit Code'
    SetBlock ($ws.Range($ws.Cells($hdrRow,$tCol),$ws.Cells($hdrRow,$tCol+2))) $tHdr
    RRStyleYellow ($ws.Range($ws.Cells($hdrRow,$tCol),$ws.Cells($hdrRow,$tCol+2)))

    $sqSeen = [ordered]@{}
    foreach($u in $units){ if([int]$u.capacity -eq 1 -and [double]$u.sqft -gt 0){ $sqSeen[[double]$u.sqft] = $u.groupUnitCode } }
    $sqList = $sqSeen.Keys | Sort-Object
    $tr=$hdrRow
    foreach($sq in $sqList){
      $tr++
      $ws.Cells($tr,$tCol).Value2=[double]$sq
      $ws.Cells($tr,$tCol+1).Formula='=COUNTIFS({0},{1}{2},{3},1)' -f $R_SQ,(RRColL $tCol),$tr,$R_CAP
      $ws.Cells($tr,$tCol+2).Value2=[string]$sqSeen[$sq]
    }
    if($tr -gt $hdrRow){
      $ws.Range($ws.Cells($hdrRow+1,$tCol),$ws.Cells($tr,$tCol)).NumberFormat='#,##0'
      $ws.Range($ws.Cells($hdrRow+1,$tCol+1),$ws.Cells($tr,$tCol+1)).NumberFormat='0'
    }

    # =============================================================================================
    # NMI Rent proposal (rule 18: ~4-5% above In-Place(Occ), rounded to end in 95) - computed here
    # from the raw source values (read directly, not via units.csv which carries no dollars).
    # =============================================================================================
    $unitsByGroup = @{}
    foreach($u in $units){ if(-not $unitsByGroup.ContainsKey($u.groupUnitCode)){ $unitsByGroup[$u.groupUnitCode]=@() }; $unitsByGroup[$u.groupUnitCode]+=$u }
    $actColIdx = RRColIdx $cAct
    function RRGroupAvgInPlace($grpKey){
      $mem = @($unitsByGroup[$grpKey] | Where-Object { [int]$_.occupied -eq 1 })
      if($mem.Count -eq 0){ return 0.0 }
      $tot=0.0
      foreach($m in $mem){
        $leaves=@($m.leafRows -split ';' | Where-Object { $_ -ne '' })
        foreach($lf in $leaves){ $tot += (RRNumOf $ws.Cells([int]$lf,$actColIdx).Value2) }
      }
      return $tot/$mem.Count
    }
    function RRProposeNMI($avgInPlace){
      if($avgInPlace -le 0){ return 0.0 }
      $cand = [math]::Round(($avgInPlace*1.045)/100.0,0)*100.0 - 5.0
      if($cand -le $avgInPlace){ $cand += 100.0 }
      return $cand
    }

    # =============================================================================================
    # MAIN SUMMARY (+23..+39) - navy header
    # =============================================================================================
    $s=$sCol
    $LUT=RRColL $s; $LAP=RRColL ($s+1); $LUN=RRColL ($s+2); $LOU=RRColL ($s+3); $LOP=RRColL ($s+4)
    $LMN=RRColL ($s+5); $LAV=RRColL ($s+6); $LMX=RRColL ($s+7); $LMA=RRColL ($s+8); $LMO=RRColL ($s+9)
    $LIP=RRColL ($s+10); $LDI=RRColL ($s+11); $LCA=RRColL ($s+12); $LNMI=RRColL ($s+13); $LCJ=RRColL ($s+14)
    $LNET=RRColL ($s+15); $LNVI=RRColL ($s+16)
    $sHdr=@('Unit Type','#Apts','#Units','Occ Units','Occ %','Min SqFt','Avg SqFt','Max SqFt','Market (All)','Market (Occ)','In-Place (Occ)','(Discount)','Care','NMI Rent','Care Adj','Net NMI Rent','NMI vs In Place')
    $sHdrArr = New-Object 'object[,]' 1,17
    for($c=0;$c -lt 17;$c++){ $sHdrArr[0,$c]=$sHdr[$c] }
    SetBlock ($ws.Range($ws.Cells($hdrRow,$s),$ws.Cells($hdrRow,$s+16))) $sHdrArr
    RRStyleNavy ($ws.Range($ws.Cells($hdrRow,$s),$ws.Cells($hdrRow,$s+16)))
    $ws.Cells($hdrRow,$s).HorizontalAlignment=-4131
    $ws.Rows.Item($hdrRow).RowHeight=30

    function RRGroupRow($row,$crit,$label,$semi,$careOnly){
      $ws.Cells($row,$s).Value2=$label
      if($careOnly -eq ""){ $rng=$R_GRP; $q='"'+$crit+'"' } else { $rng=$R_CT; $q='"'+$careOnly+'"' }
      $ws.Cells($row,$s+2).Formula = '=COUNTIFS({0},{1},{2},1)' -f $rng,$q,$R_CAP
      $ws.Cells($row,$s+3).Formula = '=COUNTIFS({0},{1},{2},1,{3},1)' -f $rng,$q,$R_CAP,$R_OCC
      # #Apts = DISTINCT-unit (apartment) count within the group, at capacity=1 - not a raw bed count.
      # A companion/semi-private pair (two beds, one door, same raw Unit code) collapses to 1 apartment
      # via the 1/COUNTIFS(...) unique-count trick; ordinary single-bed units are unaffected (count=1,
      # so 1/1=1 - identical to a plain COUNTIFS). Works uniformly for any care level/tier, including a
      # tier that mixes companion pairs with ordinary singles (output-columns.md case 2; rule 19 for MC).
      # NOTE: this nested nested COUNTIFS(unitRange,unitRange,...) unique-count trick needs true CSE
      # array evaluation - a plain .Formula assignment (no Ctrl+Shift+Enter equivalent) silently
      # returns 0 via COM, so this MUST be written via .FormulaArray, not .Formula (verified empirically).
      $ws.Cells($row,$s+1).FormulaArray = '=SUMPRODUCT(({0}={1})*({2}=1)*IFERROR(1/COUNTIFS({3},{3},{0},{1},{2},1),0))' -f $rng,$q,$R_CAP,$R_UN
      $ws.Cells($row,$s+4).Formula = '=IFERROR({0}{1}/{2}{1},0)' -f $LOU,$row,$LUN
      $ws.Cells($row,$s+5).Formula = '=IFERROR(MINIFS({0},{1},{2},{3},1),0)'     -f $R_SQ,$rng,$q,$R_CAP
      $ws.Cells($row,$s+6).Formula = '=IFERROR(AVERAGEIFS({0},{1},{2},{3},1),0)' -f $R_SQ,$rng,$q,$R_CAP
      $ws.Cells($row,$s+7).Formula = '=IFERROR(MAXIFS({0},{1},{2},{3},1),0)'     -f $R_SQ,$rng,$q,$R_CAP
      $ws.Cells($row,$s+8).Formula = '=IFERROR(AVERAGEIFS({0},{1},{2},{3},1),0)' -f $R_MK,$rng,$q,$R_CAP
      $ws.Cells($row,$s+9).Formula = '=IFERROR(AVERAGEIFS({0},{1},{2},{3},1,{4},1),0)' -f $R_MK,$rng,$q,$R_CAP,$R_OCC
      $ws.Cells($row,$s+10).Formula= '=IFERROR(AVERAGEIFS({0},{1},{2},{3},1,{4},1),0)' -f $R_RT,$rng,$q,$R_CAP,$R_OCC
      $ws.Cells($row,$s+11).Formula= '=IFERROR({0}{1}/{2}{1}-1,0)' -f $LIP,$row,$LMO
      $ws.Cells($row,$s+12).Formula= '=IFERROR(AVERAGEIFS({0},{1},{2},{3},1),0)' -f $R_CR,$rng,$q,$R_CAP
      $ws.Cells($row,$s+15).Formula= '={0}{1}+{2}{1}' -f $LNMI,$row,$LCJ
      $ws.Cells($row,$s+16).Formula= '=IFERROR({0}{1}/{2}{1}-1,0)' -f $LNET,$row,$LIP
    }

    foreach($care in @('IL','AL','MC')){
      $info=$levelInfo[$care]
      if($info.rows.Count -eq 0){
        # placeholder rows stay blank; still need a Total row (careOnly grouping, on Care Type)
      } else {
        for($ti=0;$ti -lt $info.tiers.Count;$ti++){
          $t=$info.tiers[$ti]; $row=$info.rows[$ti]
          $label = if($care -eq 'MC' -and $t.PSObject.Properties.Name -contains 'label'){ $t.label } else { $t.group }
          $semi = ($care -eq 'MC' -and $t.PSObject.Properties.Name -contains 'structure' -and $t.structure -eq 'semi')
          RRGroupRow $row $t.group $label $semi ''
          $ws.Cells($row,$s+13).Value2=[double](RRProposeNMI (RRGroupAvgInPlace $t.group))   # NMI Rent - blue judgment input
          $ws.Cells($row,$s+14).Value2=0.0                                                    # Care Adj
        }
      }
      $tRow=$info.totalRow
      RRGroupRow $tRow '' ("Total "+$care) $false $care
      if($info.rows.Count -gt 0){
        $r0=$info.rows[0]; $r1=$info.rows[$info.rows.Count-1]
        $ws.Cells($tRow,$s+1).Formula = '=SUM({0}{1}:{0}{2})' -f $LAP,$r0,$r1
        $ws.Cells($tRow,$s+13).Formula= '=IFERROR(SUMPRODUCT({0}{1}:{0}{2},{3}{1}:{3}{2})/{3}{4},0)' -f $LNMI,$r0,$r1,$LUN,$tRow
        $ws.Cells($tRow,$s+14).Formula= '=IFERROR(SUMPRODUCT({0}{1}:{0}{2},{3}{1}:{3}{2})/{3}{4},0)' -f $LCJ,$r0,$r1,$LUN,$tRow
      } else {
        $ws.Cells($tRow,$s+13).Value2=0.0; $ws.Cells($tRow,$s+14).Value2=0.0
      }
    }

    # grand Total row
    $rIL=$levelInfo['IL'].totalRow; $rAL=$levelInfo['AL'].totalRow; $rMC=$levelInfo['MC'].totalRow
    $ws.Cells($rTOT,$s).Value2='Total'
    $ws.Cells($rTOT,$s+1).Formula = '={0}{1}+{0}{2}+{0}{3}' -f $LAP,$rIL,$rAL,$rMC
    $ws.Cells($rTOT,$s+2).Formula = '=COUNTIFS({0},1)' -f $R_CAP
    $ws.Cells($rTOT,$s+3).Formula = '=COUNTIFS({0},1,{1},1)' -f $R_CAP,$R_OCC
    $ws.Cells($rTOT,$s+4).Formula = '=IFERROR({0}{1}/{2}{1},0)' -f $LOU,$rTOT,$LUN
    $ws.Cells($rTOT,$s+5).Formula = '=IFERROR(MINIFS({0},{1},1),0)'     -f $R_SQ,$R_CAP
    $ws.Cells($rTOT,$s+6).Formula = '=IFERROR(AVERAGEIFS({0},{1},1),0)' -f $R_SQ,$R_CAP
    $ws.Cells($rTOT,$s+7).Formula = '=IFERROR(MAXIFS({0},{1},1),0)'     -f $R_SQ,$R_CAP
    $ws.Cells($rTOT,$s+8).Formula = '=IFERROR(AVERAGEIFS({0},{1},1),0)' -f $R_MK,$R_CAP
    $ws.Cells($rTOT,$s+9).Formula = '=IFERROR(AVERAGEIFS({0},{1},1,{2},1),0)' -f $R_MK,$R_CAP,$R_OCC
    $ws.Cells($rTOT,$s+10).Formula= '=IFERROR(AVERAGEIFS({0},{1},1,{2},1),0)' -f $R_RT,$R_CAP,$R_OCC
    $ws.Cells($rTOT,$s+11).Formula= '=IFERROR({0}{1}/{2}{1}-1,0)' -f $LIP,$rTOT,$LMO
    $ws.Cells($rTOT,$s+12).Formula= '=IFERROR(SUMIFS({0},{1},1)/{2}{3},0)' -f $R_CR,$R_CAP,$LUN,$rTOT
    $ws.Cells($rTOT,$s+13).Formula= '=IFERROR((SUMPRODUCT({0}{1}:{0}{2},{3}{1}:{3}{2}))/{3}{4},0)' -f $LNMI,$unitFirst,$unitLast,$LUN,$rTOT
    # (the tier-weighted NMI blend above double counts absent-level placeholder rows harmlessly -
    #  guarded instead via a direct re-derivation over the populated tier rows only:)
    $tierRows=@(); foreach($care in @('IL','AL','MC')){ $tierRows += $levelInfo[$care].rows }
    if($tierRows.Count -gt 0){
      $terms = ($tierRows | ForEach-Object { "{0}{1}*{2}{1}" -f $LNMI,$_,$LUN }) -join '+'
      $ws.Cells($rTOT,$s+13).Formula = "=IFERROR(($terms)/{0}{1},0)" -f $LUN,$rTOT
      $terms2 = ($tierRows | ForEach-Object { "{0}{1}*{2}{1}" -f $LCJ,$_,$LUN }) -join '+'
      $ws.Cells($rTOT,$s+14).Formula = "=IFERROR(($terms2)/{0}{1},0)" -f $LUN,$rTOT
    } else { $ws.Cells($rTOT,$s+13).Value2=0.0; $ws.Cells($rTOT,$s+14).Value2=0.0 }
    $ws.Cells($rTOT,$s+15).Formula = '={0}{1}+{2}{1}' -f $LNMI,$rTOT,$LCJ
    $ws.Cells($rTOT,$s+16).Formula = '=IFERROR({0}{1}/{2}{1}-1,0)' -f $LNET,$rTOT,$LIP

    # summary number formats + total-row borders/labels
    $sTop=$hdrRow+1
    $ws.Range($ws.Cells($sTop,$s+1),$ws.Cells($rTOT,$s+3)).NumberFormat=$script:rrACC
    $ws.Range($ws.Cells($sTop,$s+4),$ws.Cells($rTOT,$s+4)).NumberFormat=$script:rrPCT
    $ws.Range($ws.Cells($sTop,$s+5),$ws.Cells($rTOT,$s+10)).NumberFormat=$script:rrACC
    $ws.Range($ws.Cells($sTop,$s+11),$ws.Cells($rTOT,$s+11)).NumberFormat=$script:rrPCT
    $ws.Range($ws.Cells($sTop,$s+12),$ws.Cells($rTOT,$s+15)).NumberFormat=$script:rrACC
    $ws.Range($ws.Cells($sTop,$s+16),$ws.Cells($rTOT,$s+16)).NumberFormat=$script:rrPCT
    foreach($tr2 in @($rIL,$rAL,$rMC,$rTOT)){ RRTotalBorder ($ws.Range($ws.Cells($tr2,$s),$ws.Cells($tr2,$s+16))) }
    foreach($care in @('IL','AL','MC')){ foreach($rr in $levelInfo[$care].rows){ $ws.Cells($rr,$s+13).Font.Color=$script:rrBLU } }

    # =============================================================================================
    # SECOND RESIDENTS (rTOT+2 .. +8) - four-column block: IL/AL/MC/Total
    # =============================================================================================
    $r2=$rTOT+2
    $ws.Cells($r2,$s).Value2='Second Residents'
    $ws.Cells($r2,$s+1).Value2='IL'; $ws.Cells($r2,$s+2).Value2='AL'; $ws.Cells($r2,$s+3).Value2='MC'; $ws.Cells($r2,$s+4).Value2='Total'
    RRStylePlain ($ws.Range($ws.Cells($r2,$s),$ws.Cells($r2,$s+4))) $script:rrBLK
    $ws.Cells($r2,$s).HorizontalAlignment=-4131
    $srLab=@('# 2nd Residents','Utilization','Rent','Rent / 2nd resident','Care','Care / 2nd resident')
    for($i=0;$i -lt 6;$i++){ $ws.Cells($r2+1+$i,$s).Value2=$srLab[$i] }
    $cares=@('IL','AL','MC')
    for($i=0;$i -lt 3;$i++){
      $cc=$cares[$i]; $col=$s+1+$i; $L=RRColL $col
      $ws.Cells($r2+1,$col).Formula='=COUNTIFS({0},"{1}",{2},">0")' -f $R_CT,$cc,$R_SE
      $ws.Cells($r2+2,$col).Formula='=IFERROR({0}{1}/COUNTIFS({2},"{3}",{4},1),0)' -f $L,($r2+1),$R_CT,$cc,$R_CAP
      $ws.Cells($r2+3,$col).Formula='=SUMIFS({0},{1},"{2}")' -f $R_SE,$R_CT,$cc
      $ws.Cells($r2+4,$col).Formula='=IFERROR({0}{1}/{0}{2},0)' -f $L,($r2+3),($r2+1)
      $ws.Cells($r2+5,$col).Formula='=SUMIFS({0},{1},">0",{2},"{3}")' -f $R_CR,$R_SE,$R_CT,$cc
      $ws.Cells($r2+6,$col).Formula='=IFERROR({0}{1}/{0}{2},0)' -f $L,($r2+5),($r2+1)
    }
    $LT1=RRColL ($s+1); $LT3=RRColL ($s+3); $LTT=RRColL ($s+4)
    foreach($k in @(1,3,5)){ $ws.Cells($r2+$k,$s+4).Formula='=SUM({0}{1}:{2}{1})' -f $LT1,($r2+$k),$LT3 }
    $ws.Cells($r2+2,$s+4).Formula='=IFERROR({0}{1}/COUNTIFS({2},1),0)' -f $LTT,($r2+1),$R_CAP
    $ws.Cells($r2+4,$s+4).Formula='=IFERROR({0}{1}/{0}{2},0)' -f $LTT,($r2+3),($r2+1)
    $ws.Cells($r2+6,$s+4).Formula='=IFERROR({0}{1}/{0}{2},0)' -f $LTT,($r2+5),($r2+1)
    $ws.Range($ws.Cells($r2+1,$s+1),$ws.Cells($r2+1,$s+4)).NumberFormat=$script:rrACC
    $ws.Range($ws.Cells($r2+2,$s+1),$ws.Cells($r2+2,$s+4)).NumberFormat=$script:rrPCT
    $ws.Range($ws.Cells($r2+3,$s+1),$ws.Cells($r2+6,$s+4)).NumberFormat=$script:rrACC

    # =============================================================================================
    # RENT ADJUSTMENTS (rTOT+10..+13) - orange inputs
    # =============================================================================================
    $r3=$rTOT+10
    $ws.Cells($r3,$s).Value2='Rent Adjustments'
    RRStylePlain ($ws.Range($ws.Cells($r3,$s),$ws.Cells($r3,$s+1))) $script:rrBLK
    $ws.Cells($r3,$s).HorizontalAlignment=-4131
    $adjLab=@('IL Rent Adj','AL Rent Adj','MC Rent Adj')
    for($i=0;$i -lt 3;$i++){ $ws.Cells($r3+1+$i,$s).Value2=$adjLab[$i]; $ws.Cells($r3+1+$i,$s+1).Value2=0.0 }
    $adjRng=$ws.Range($ws.Cells($r3+1,$s+1),$ws.Cells($r3+3,$s+1))
    $adjRng.Font.Color=$script:rrORG; $adjRng.NumberFormat=$script:rrACC

    # =============================================================================================
    # CHECK (rTOT+15..+20) - all red, must compute 0
    # =============================================================================================
    $r4=$rTOT+15
    $ws.Cells($r4,$s).Value2='Check'
    $ws.Cells($r4+1,$s).Value2='Units/Apartments'; $ws.Cells($r4+1,$s+1).Formula='={0}{1}-SUM({2})' -f $LUN,$rTOT,$R_CAP
    $ws.Cells($r4+2,$s).Value2='Occupied Units';   $ws.Cells($r4+2,$s+1).Formula='={0}{1}-SUM({2})' -f $LOU,$rTOT,$R_OCC
    $ws.Cells($r4+3,$s).Value2='Occupied %';       $ws.Cells($r4+3,$s+1).Formula='={0}{1}-IFERROR(SUM({2})/SUM({3}),0)' -f $LOP,$rTOT,$R_OCC,$R_CAP
    $ws.Cells($r4+4,$s).Value2='Rent';             $ws.Cells($r4+4,$s+1).Formula='={0}{1}*{2}{1}-SUMIFS({3},{4},1,{5},1)' -f $LIP,$rTOT,$LOU,$R_RT,$R_CAP,$R_OCC
    $ws.Cells($r4+5,$s).Value2='Care';             $ws.Cells($r4+5,$s+1).Formula='={0}{1}*{2}{1}-SUMIFS({3},{4},1)' -f $LCA,$rTOT,$LUN,$R_CR,$R_CAP
    $chk=$ws.Range($ws.Cells($r4,$s),$ws.Cells($r4+5,$s+1)); $chk.Font.Color=$script:rrRED
    $ws.Range($ws.Cells($r4+1,$s+1),$ws.Cells($r4+2,$s+1)).NumberFormat=$script:rrACC
    $ws.Cells($r4+3,$s+1).NumberFormat=$script:rrPCT
    $ws.Range($ws.Cells($r4+4,$s+1),$ws.Cells($r4+5,$s+1)).NumberFormat=$script:rrACC
    RRStylePlain ($ws.Range($ws.Cells($r4,$s),$ws.Cells($r4,$s+1))) $script:rrRED   # Check title = plain bold RED, no fill
    $ws.Cells($r4,$s).HorizontalAlignment=-4131

    # ---- Days in period (rTOT+22) ----
    $ws.Cells($daysRow,$s).Value2='Days in period (proration)'
    if($deprorate){
      $ws.Cells($daysRow,$s+1).Value2=[double]$daysInPeriod
      $ws.Cells($daysRow,$s+1).Font.Color=$script:rrBLU
      $ws.Cells($daysRow,$s+1).NumberFormat='0'
    } else {
      $ws.Cells($daysRow,$s+1).Value2='N/A - no de-proration (see Notes)'
    }

    # =============================================================================================
    # MOVE-IN BOXES (+41..+44, +46..+49) - navy header
    # =============================================================================================
    $mHdr=@('Min Rent','Avg Rent','Max Rent','#')
    $rowsPlan=@()
    foreach($care in @('IL','AL','MC')){
      # guard the empty case: 0..(-1) enumerates 0,-1 in PowerShell and would index a
      # zero-length tiers/rows array (fix 2026-07-29 alongside the $rows reset above).
      if($levelInfo[$care].tiers.Count -gt 0){
        foreach($ri in 0..($levelInfo[$care].tiers.Count-1)){
          $t=$levelInfo[$care].tiers[$ri]; $rowsPlan += @{ r=$levelInfo[$care].rows[$ri]; rng=$R_GRP; q=('"'+$t.group+'"') }
        }
      }
      $rowsPlan += @{ r=$levelInfo[$care].totalRow; rng=$R_CT; q=('"'+$care+'"') }
    }
    for($bi=0;$bi -lt $cfg.moveInWindows.Count;$bi++){
      $win=$cfg.moveInWindows[$bi]
      $bc=$blockStart+41+($bi*5); $LB1=RRColL $bc
      $bHdrArr = New-Object 'object[,]' 1,4
      for($i=0;$i -lt 4;$i++){ $bHdrArr[0,$i]=$mHdr[$i] }
      SetBlock ($ws.Range($ws.Cells($hdrRow,$bc),$ws.Cells($hdrRow,$bc+3))) $bHdrArr
      RRStyleNavy ($ws.Range($ws.Cells($hdrRow,$bc),$ws.Cells($hdrRow,$bc+3)))
      $d1=[datetime]::Parse($win.start,[System.Globalization.CultureInfo]::InvariantCulture)
      $d2=[datetime]::Parse($win.end,[System.Globalization.CultureInfo]::InvariantCulture)
      $ws.Cells($hdrRow-2,$bc).Value2=[double]$d1.ToOADate()
      $ws.Cells($hdrRow-1,$bc).Value2=[double]$d2.ToOADate()
      $dr=$ws.Range($ws.Cells($hdrRow-2,$bc),$ws.Cells($hdrRow-1,$bc))
      $dr.NumberFormat='m/d/yyyy'; $dr.Font.Color=$script:rrBLU; $dr.Font.Bold=$false
      $S1='$'+$LB1+'$'+($hdrRow-2); $S2='$'+$LB1+'$'+($hdrRow-1)
      foreach($p in $rowsPlan){
        $rw=$p.r
        $ws.Cells($rw,$bc+0).Formula='=IFERROR(MINIFS({0},{1},{2},{3},1,{4},1,{5},">="&{6},{5},"<="&{7}),0)'     -f $R_RT,$p.rng,$p.q,$R_CAP,$R_OCC,$R_MI,$S1,$S2
        $ws.Cells($rw,$bc+1).Formula='=IFERROR(AVERAGEIFS({0},{1},{2},{3},1,{4},1,{5},">="&{6},{5},"<="&{7}),0)' -f $R_RT,$p.rng,$p.q,$R_CAP,$R_OCC,$R_MI,$S1,$S2
        $ws.Cells($rw,$bc+2).Formula='=IFERROR(MAXIFS({0},{1},{2},{3},1,{4},1,{5},">="&{6},{5},"<="&{7}),0)'     -f $R_RT,$p.rng,$p.q,$R_CAP,$R_OCC,$R_MI,$S1,$S2
        $ws.Cells($rw,$bc+3).Formula='=COUNTIFS({0},{1},{2},1,{3},1,{4},">="&{5},{4},"<="&{6})'                  -f $p.rng,$p.q,$R_CAP,$R_OCC,$R_MI,$S1,$S2
      }
      $ws.Cells($rTOT,$bc+0).Formula='=IFERROR(MINIFS({0},{1},1,{2},1,{3},">="&{4},{3},"<="&{5}),0)'     -f $R_RT,$R_CAP,$R_OCC,$R_MI,$S1,$S2
      $ws.Cells($rTOT,$bc+1).Formula='=IFERROR(AVERAGEIFS({0},{1},1,{2},1,{3},">="&{4},{3},"<="&{5}),0)' -f $R_RT,$R_CAP,$R_OCC,$R_MI,$S1,$S2
      $ws.Cells($rTOT,$bc+2).Formula='=IFERROR(MAXIFS({0},{1},1,{2},1,{3},">="&{4},{3},"<="&{5}),0)'     -f $R_RT,$R_CAP,$R_OCC,$R_MI,$S1,$S2
      $ws.Cells($rTOT,$bc+3).Formula='=COUNTIFS({0},1,{1},1,{2},">="&{3},{2},"<="&{4})'                  -f $R_CAP,$R_OCC,$R_MI,$S1,$S2
      $ws.Range($ws.Cells($sTop,$bc),$ws.Cells($rTOT,$bc+3)).NumberFormat=$script:rrACC
      foreach($tr3 in @($rIL,$rAL,$rMC,$rTOT)){ RRTotalBorder ($ws.Range($ws.Cells($tr3,$bc),$ws.Cells($tr3,$bc+3))) }
    }

    # =============================================================================================
    # NOTES (rTOT+24+) - the anomaly/flag block (rule 20)
    # =============================================================================================
    $notesRow = $rTOT+24
    if($cfg.notes -and $cfg.notes.Count -gt 0){
      $ws.Cells($notesRow,$s).Value2='Notes'
      RRStylePlain ($ws.Range($ws.Cells($notesRow,$s),$ws.Cells($notesRow,$s+1))) $script:rrBLK
      $ws.Cells($notesRow,$s).HorizontalAlignment=-4131
      for($i=0;$i -lt $cfg.notes.Count;$i++){ $ws.Cells($notesRow+1+$i,$s).Value2=[string]$cfg.notes[$i] }
    }

    # =============================================================================================
    # RECONCILIATION (unitLast+6 ..) - built (own leaves) vs operatorTotals (from config)
    # =============================================================================================
    $rc = $unitLast+6
    $LBuilt=RRColL ($blockStart+4); $LOp=RRColL ($blockStart+5); $LDiff=RRColL ($blockStart+6)
    $ws.Cells($rc,$blockStart).Value2='RECONCILIATION (sum of own leaves vs operator section totals)'
    $ws.Cells($rc,$blockStart).Font.Bold=$true
    $ws.Cells($rc+1,$blockStart+4).Value2='Built'; $ws.Cells($rc+1,$blockStart+5).Value2='op total'; $ws.Cells($rc+1,$blockStart+6).Value2='Diff'
    $ws.Cells($rc+1,$blockStart).Font.Bold=$true
    $ws.Cells($rc+2,$blockStart).Value2='Units (rows)';  $ws.Cells($rc+2,$blockStart+4).Formula='=COUNT({0})' -f (RRAbsR $wOrder $unitFirst $unitLast)
    $ws.Cells($rc+3,$blockStart).Value2='Occupied';      $ws.Cells($rc+3,$blockStart+4).Formula='=SUM({0})' -f $R_OCC
    $ws.Cells($rc+4,$blockStart).Value2='Capacity (beds)';$ws.Cells($rc+4,$blockStart+4).Formula='=SUM({0})' -f $R_CAP

    $ot = $cfg.operatorTotals
    function RRReconLine($row,$label,$builtFormula,$opVal,$naText){
      $ws.Cells($row,$blockStart).Value2=$label
      if($null -eq $opVal){
        $ws.Cells($row,$blockStart+4).Value2='N/A'
        $ws.Cells($row,$blockStart+5).Value2=$naText
        $ws.Cells($row,$blockStart+6).Value2='N/A'
      } else {
        $ws.Cells($row,$blockStart+4).Formula=$builtFormula
        $ws.Cells($row,$blockStart+5).Value2=[double]$opVal
        $ws.Cells($row,$blockStart+6).Formula='={0}{1}-{2}{1}' -f $LBuilt,$row,$LOp
        $ws.Range($ws.Cells($row,$blockStart+4),$ws.Cells($row,$blockStart+6)).NumberFormat='#,##0.00'
      }
    }
    RRReconLine ($rc+5) 'Sum Market'        ('=SUM({0})' -f $R_MK) $ot.market $null
    RRReconLine ($rc+6) 'Sum In-Place Rent' ('=SUM({0})' -f $R_RT) $ot.actual $null
    RRReconLine ($rc+7) 'Sum In-Place Care' ('=SUM({0})' -f $R_CR) $ot.care   'N/A - no Care Fees column in source'
    RRReconLine ($rc+8) 'Sum 2nd-Occ Fee'   ('=SUM({0})' -f $R_SE) $ot.secOcc 'N/A - no 2nd-Occupant Fee column in source'

    # =============================================================================================
    # COLUMN WIDTHS + HIDDEN SPACER (formatting.md #6) + sheet-level (gridlines off, tab name)
    # =============================================================================================
    $W=@(7.9,11.0,11.1,12.7,14.0,9.6,17.6,12.9,14.7,17.1,16.3,16.7,16.6,16.6,31.6,19.7,19.7,19.7,0,19.7,19.7,19.7,8.1,23.1,8.9,7.0,8.6,7.9,6.3,6.3,6.3,11.0,11.7,12.6,9.4,6.3,8.7,7.7,12.6,13.9,8.1,10.0,8.1,8.1,8.1,8.1,9.4,8.1,8.1,8.1)
    for($c=0;$c -lt $W.Count;$c++){ if($W[$c] -gt 0){ $ws.Columns.Item($blockStart+$c).ColumnWidth=[double]$W[$c] } }
    $ws.Columns.Item($blockStart+18).Hidden=$true
    $ws.Name = "{0}.RR" -f $cfg.community
    $wb.Windows.Item(1).DisplayGridlines=$false

    # =============================================================================================
    # RECALC + GATES (rule 5)
    # =============================================================================================
    $xl.Calculation=-4105
    $xl.CalculateFull()
    Start-Sleep -Milliseconds 500
    $xl.CalculateFull()

    # gate: five-header presence
    $hdrCells = @($ws.Cells($hdrRow,$blockStart).Value2, $ws.Cells($hdrRow,$tCol).Value2, $ws.Cells($hdrRow,$s).Value2)
    for($bi=0;$bi -lt $cfg.moveInWindows.Count;$bi++){ $hdrCells += $ws.Cells($hdrRow,$blockStart+41+($bi*5)).Value2 }
    $headersPresent = -not ($hdrCells | Where-Object { -not $_ -or "$_".Trim() -eq "" })
    $result.headersPresent = [bool]$headersPresent

    # gate: formula errors via SpecialCells (Int32 error codes, NOT string scan)
    $errCount=0; $errList=@()
    $scanTop=$hdrRow-2; $scanBot=$rc+8
    try {
      $errRng = $ws.Range($ws.Cells($scanTop,$blockStart),$ws.Cells($scanBot,$blockStart+49)).SpecialCells(-4123,16)
      foreach($area in $errRng.Areas){
        foreach($cell in $area.Cells){ $errCount++; if($errList.Count -lt 20){ $errList += $cell.Address($false,$false) } }
      }
    } catch { $errCount=0 }   # SpecialCells throws when nothing matches -> clean
    $result.errorCount=$errCount; $result.errorCells=$errList

    # gate: reconcile diffs
    $recon=[ordered]@{}
    foreach($pair in @(@('market',$rc+5),@('actual',$rc+6),@('care',$rc+7),@('secOcc',$rc+8))){
      $name=$pair[0]; $row=$pair[1]
      if($null -eq $ot.$name){ $recon[$name]='N/A' } else { $recon[$name]=[double]$ws.Cells($row,$blockStart+6).Value2 }
    }
    $result.reconcile=$recon

    # gate: Check block values
    $chkVals=[ordered]@{}
    $chkLabels=@('Units/Apartments','Occupied Units','Occupied %','Rent','Care')
    for($i=0;$i -lt 5;$i++){ $chkVals[$chkLabels[$i]]=[double]$ws.Cells($r4+1+$i,$s+1).Value2 }
    $result.checkValues=$chkVals

    # ---- save (Save() in place - never SaveAs; rule word-com-saveas-hang) ----
    $saveOk=$false
    for($k=0;$k -lt 8 -and -not $saveOk;$k++){ try{ $wb.Save(); $saveOk=$true }catch{ Start-Sleep -Milliseconds 700 } }
    $result.saveOk=$saveOk
    if(-not $saveOk){ throw "workbook Save() failed after retries" }

    $success = $true
  } finally {
    if($wb){ try{ $wb.Close($false) }catch{} }
    Stop-TrackedExcel
    if($xl){ try{ $xl.Quit() }catch{}; [System.Runtime.InteropServices.Marshal]::ReleaseComObject($xl)|Out-Null }
  }

  if($success){
    Rename-Item -LiteralPath $buildingPath -NewName $finalName -Force
  } else {
    if(Test-Path -LiteralPath $buildingPath){ Remove-Item -LiteralPath $buildingPath -Force -ErrorAction SilentlyContinue }
  }

  Write-Output ($result | ConvertTo-Json -Depth 6)
  return [pscustomobject]$result
}
