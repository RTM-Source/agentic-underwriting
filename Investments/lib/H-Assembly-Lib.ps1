# .H build engine (v1, 2026-07-28) - see .claude/skills/h-underwrite/scripts/ENGINE-SPEC.md
# Dot-source, then: Invoke-HBuild -ConfigPath <config.json>
#
# Split of responsibility:
#   h-assembler sub-agent - resolves per-deal anchors (hfLast/hfRevRow/hfNoiRow/mapLast/census map),
#     emits config.json, invokes this engine, reports the gate readings it returns.
#   ENGINE (this file) - everything the h-skeleton.md contract specifies: census merge, Block A
#     occupancy, Block B detail passthrough, Block C roll-up, Block D metrics, Block E S/T/U totals +
#     the full styling contract, self-run gates, save-only-if-clean. The engine IS the golden
#     (Deal B.H Model_v1.xlsx), made executable. It never re-derives judgment - it trusts the config.
#
# House style follows lib/HF-Build-Lib.ps1 (dot-sourced below for SetV/SetBlock/tracked-Excel helpers).
# Row-geometry citations refer to the golden read-back performed 2026-07-28 (see engine ENGINE-SPEC.md).

. (Join-Path $PSScriptRoot "HF-Build-Lib.ps1")

# ---- column-letter helpers (case-collision rule: distinct, alias-safe names) -------------------
function HColL([int]$n){ $hs=""; while($n -gt 0){ $hm=($n-1)%26; $hs=[char](65+$hm)+$hs; $n=[int](($n-$hm-1)/26) }; return $hs }
function HColIdx([string]$L){ $hi=0; foreach($ch in $L.ToCharArray()){ $hi = $hi*26 + ([int][char]$ch - 64) }; return $hi }

# ---- style constants (BGR longs, matching the Deal B golden read-back) --------------------------
$script:hGREEN=32768; $script:hBLUE=16711680; $script:hBLACK=0; $script:hWHITE=16777215
$script:hGREYFILL=13882323; $script:hCHECKFILL=4697456
$script:hACCT   = '_(* #,##0_);_(* \(#,##0\);_(* "-"_);_(@_)'
$script:hINT    = '#,##0_);(#,##0)'
$script:hPCT0   = '0%'
$script:hPCT1   = '#,##0.0%_);(#,##0.0%)'
$script:hPCT2   = '0.00%'
$script:hPCT2B  = '#,##0.00%_);(#,##0.00%)'   # legacy 2-decimal variant - R column only, mgmtfee/revpct metric rows (golden read-back)
$script:hDATE   = 'm/d/yyyy'
$script:hCHKFMT = '#,##0.00'
$script:hRESFMT = '#,##0.00x'

function HBorderTop($rng){ $b=$rng.Borders.Item(8); $b.LineStyle=1; $b.Weight=2 }
function HBorderBottom($rng){ $b=$rng.Borders.Item(9); $b.LineStyle=1; $b.Weight=2 }
function HBorderBottomDouble($rng){ $b=$rng.Borders.Item(9); $b.LineStyle=-4119; $b.Weight=4 }

# =================================================================================================
# ROLL-UP CATEGORY TEMPLATE (baseline anchor = 221, per rollup-category-map.md).
# Every row is expressed as an OFFSET from rollupFirst (engine derives rollupFirst = detailLast+3).
# =================================================================================================
$script:RollupCats = @(
  @{off=0; B='AL Rent'; C='Rent Revenue - AL'}, @{off=1; B='IL Rent'; C='Rent Revenue - IL'},
  @{off=2; B='TC Rent'; C='Rent Revenue - TC'}, @{off=3; B='MC Rent'; C='Rent Revenue - MC'},
  @{off=4; B='Rent Concessions'; C='Rent Concessions'},
  @{off=5; B='Other Revenue'; C='Other Revenue'},
  @{off=6; B='Other Revenue'; C='Commercial Lease Revenue'},
  @{off=7; B='Other Revenue'; C='Rent - Second Person'},
  @{off=8; B='AL Care'; C='Care Revenue - AL'}, @{off=9; B='MC Care'; C='Care Revenue - MC'},
  @{off=10; B='Community Fees'; C='Community Fees'}
)
$script:LaborCats = @(
  @{off=15; B='Culinary'; C='Wages - Culinary'}, @{off=16; B='Activities'; C='Wages - Activities'},
  @{off=17; B='Direct Care'; C='Wages - Direct Care'}, @{off=18; B='Maintenance'; C='Wages - Maintenance'},
  @{off=19; B='Housekeeping'; C='Wages - Housekeeping'}, @{off=20; B='Marketing'; C='Wages - Marketing'},
  @{off=21; B='Admin'; C='Wages - Admin'}, @{off=22; B='Marketing'; C='Bonuses - Marketing'},
  @{off=23; B='Admin'; C='Bonuses - Admin'}
)
$script:TempCats = @(
  @{off=26; B='Culinary'; C='Contract Labor - Culinary'}, @{off=27; B='Activities'; C='Contract Labor - Activities'},
  @{off=28; B='Direct Care'; C='Contract Labor - Direct Care'}, @{off=29; B='Maintenance'; C='Contract Labor - Maintenance'},
  @{off=30; B='Housekeeping'; C='Contract Labor - Housekeeping'}, @{off=31; B='Marketing'; C='Contract Labor - Marketing'},
  @{off=32; B='Admin'; C='Contract Labor - Admin'}
)
$script:BenefitCats = @(
  @{off=35; B='Culinary'; C='Benefits - Culinary'}, @{off=36; B='Activities'; C='Benefits - Activities'},
  @{off=37; B='Direct Care'; C='Benefits - Direct Care'}, @{off=38; B='Maintenance'; C='Benefits - Maintenance'},
  @{off=39; B='Housekeeping'; C='Benefits - Housekeeping'}, @{off=40; B='Marketing'; C='Benefits - Marketing'},
  @{off=41; B='Admin'; C='Benefits - Admin'}
)
$script:OpexCats = @(
  @{off=44; B='Culinary'; C='Culinary'}, @{off=45; B='Culinary'; C='Raw Food'},
  @{off=46; B='Activities'; C='Activities'}, @{off=47; B='Direct Care'; C='Direct Care'},
  @{off=48; B='Non-Departmental'; C='COGS'}, @{off=49; B='Maintenance'; C='Maintenance'},
  @{off=50; B='Utilities'; C='Utilities'}, @{off=51; B='Housekeeping'; C='Housekeeping'},
  @{off=52; B='Marketing'; C='Marketing'}, @{off=53; B='Marketing'; C='Marketing - Contract Services'},
  @{off=54; B='Marketing'; C='Marketing - Referral Fees'}, @{off=55; B='Admin'; C='Admin'},
  @{off=56; B='Admin'; C='Admin - Legal'}, @{off=57; B='Admin'; C='Admin - Bad Debt'},
  @{off=58; B='Admin'; C='Admin - Other Expense'}, @{off=59; B='Admin'; C='Insurance - Workers Comp'},
  @{off=60; B='Admin'; C='Admin - Health & Dental'}, @{off=61; B='Admin'; C='Admin - IT'},
  @{off=62; B='Admin'; C='Admin - Telephone'}, @{off=63; B='Admin'; C='Admin - Travel'},
  @{off=64; B='Admin'; C='Admin - Bank & Payroll Fees'}, @{off=65; B='Admin'; C='Admin - Vehicle Lease'},
  @{off=66; B='Real Estate Taxes'; C='Real Estate Taxes'}, @{off=67; B='Business Insurance'; C='Business Insurance'},
  @{off=68; B='Business Insurance'; C='Insurance - Other'}, @{off=69; B='Non-Departmental'; C='Non-Departmental'},
  @{off=70; B='Non-Departmental'; C='Other Taxes'}
)
# Metrics template (baseline 310-336 + 312 out-of-order; offsets = baselineRow-221).
$script:MetricRows = @(
  @{off=89;  D='Bonuses - Marketing';       C='$ Unit '; kind='unit'}
  @{off=90;  D='Bonuses - Admin';           C='$K/Property'; kind='k1000'}
  @{off=91;  D='Benefits & PR Taxes';       C='% Labor'; kind='benpr'}
  @{off=92;  D='Culinary';                 C='$ Unit '; kind='unit'}
  @{off=93;  D='Raw Food';      C='$ PRD'; kind='prd'}
  @{off=94;  D='Activities';               C='$ Unit '; kind='unit'}
  @{off=95;  D='Direct Care';              C='$ PRD'; kind='prd'}
  @{off=96;  D='COGS';                     C='% Rent+Care'; kind='rentcare'}
  @{off=97;  D='Maintenance';              C='$ Unit '; kind='unit'}
  @{off=98;  D='Utilities';                C='$ Unit '; kind='unit'}
  @{off=99;  D='Housekeeping';             C='$ Unit '; kind='unit'}
  @{off=100; D='Admin';                    C='$K/Property'; kind='adminK'}
  @{off=101; D='Insurance - Workers Comp';           C='$ Unit '; kind='unit'}
  @{off=102; D='Admin - Health & Dental';    C='% Labor'; kind='healthdental'}
  @{off=103; D='Admin - IT';                 C='$K/Property'; kind='k1000'}
  @{off=104; D='Admin - Telephone';          C='$K/Property'; kind='k1000'}
  @{off=105; D='Admin - Travel';             C='$K/Property'; kind='k1000'}
  @{off=106; D='Admin - Bank & Payroll Fees';       C='$K/Property'; kind='k1000'}
  @{off=107; D='Admin - Vehicle Lease';      C='$K/Property'; kind='k1000'}
  @{off=108; D='Real Estate Taxes';                 C='$K/Property'; kind='k1000'}
  @{off=109; D='Insurance - Other';        C='$ Unit'; kind='unit'}
  @{off=110; D='Business Insurance';       C='$ Unit '; kind='unit'}
  @{off=111; D='Non-Departmental';         C='$K/Property'; kind='k1000'}
  @{off=112; D='Management Fee';           C='% Revenue'; kind='mgmtfee'}
  @{off=113; D='Other Taxes';              C='% Revenue'; kind='revpct'}
  @{off=114; D='Capex Reserve';            C='$ Unit'; kind='unit'}
  @{off=115; D='Marketing';                C='$ Unit'; kind='marketing'}
)

# =================================================================================================
function Invoke-HBuild {
  param([Parameter(Mandatory=$true)][string]$ConfigPath)

  [System.Threading.Thread]::CurrentThread.CurrentCulture=[System.Globalization.CultureInfo]'en-US'
  $ErrorActionPreference = "Stop"

  if(-not (Test-Path -LiteralPath $ConfigPath)){ throw "config not found: $ConfigPath" }
  $cfg = Get-Content -Raw -LiteralPath $ConfigPath | ConvertFrom-Json

  $wbPath = $cfg.workbook
  if(-not (Test-Path -LiteralPath $wbPath)){
    $leaf = Split-Path $wbPath -Leaf
    $found = Get-ChildItem -Path (Split-Path $wbPath -Parent) -Filter $leaf -ErrorAction SilentlyContinue
    if($found){ $wbPath = $found[0].FullName } else { throw "workbook not found: $($cfg.workbook)" }
  }

  $cn = [string]$cfg.cn
  $hTab = [string]$cfg.hTab
  $hfTab = [string]$cfg.hfTab
  $mapTab = [string]$cfg.mapTab
  $hfDetailStart = [int]$cfg.hfDetailStart
  $hfLast = [int]$cfg.hfLast
  $hfRevRow = [int]$cfg.hfRevRow
  $hfNoiRow = [int]$cfg.hfNoiRow
  $mapLast = [int]$cfg.mapLast
  $censusMap = $cfg.censusMap
  $headerDateRow = [int]$censusMap.headerDateRow
  $capRows = $censusMap.capacityRows
  $occRows = $censusMap.occRows
  $firstMonthCol = [string]$censusMap.firstMonthCol
  $firstMonthColIdx = HColIdx $firstMonthCol
  $windowStart = [datetime]::Parse([string]$cfg.windowStart,[System.Globalization.CultureInfo]::InvariantCulture)

  $xl=$null; $wb=$null; $success=$false
  $result = [ordered]@{
    cn=$cn; workbook=$wbPath; detailFirst=28; detailLastContent=$null; detailLast=$null
    rollupFirst=$null; rollupLast=$null; errorCount=$null; errorCells=@()
    checkRevValues=$null; checkNoiValues=$null; checkSValues=$null
    externalLinkCount=$null; censusGateOk=$null; styleGateOk=$null; saveOk=$false
  }

  try {
    $xl = New-ExcelTracked
    $wb = $null
    for($k=0;$k -lt 6 -and -not $wb;$k++){ try{ $wb=$xl.Workbooks.Open($wbPath,0,$false) }catch{ Start-Sleep -Milliseconds 500 } }
    if(-not $wb){ throw "could not open $wbPath" }
    $xl.Calculation=-4135   # manual while writing

    # =============================================================================================
    # STEP 0 - CENSUS MERGE (only if a scratch workbook was supplied)
    # Landmine: copy ALL scratch sheets in ONE grouped Worksheets(@(...)).Copy call, never one at a
    # time - a one-at-a-time copy rewrites intra-workbook cross-sheet refs into external links to the
    # scratch file's path (shipped 2228 errors in a prior run). After merge, hard-scan for "[" / "http".
    # =============================================================================================
    if($cfg.censusScratch -and "$($cfg.censusScratch)".Trim() -ne ""){
      $scratchPath = [string]$cfg.censusScratch
      if(-not (Test-Path -LiteralPath $scratchPath)){ throw "censusScratch not found: $scratchPath" }
      $scratchWb = $null
      for($k=0;$k -lt 6 -and -not $scratchWb;$k++){ try{ $scratchWb=$xl.Workbooks.Open($scratchPath,0,$true) }catch{ Start-Sleep -Milliseconds 500 } }
      if(-not $scratchWb){ throw "could not open censusScratch: $scratchPath" }
      $tabNames = @($cfg.censusTabs)
      $sheetsToCopy = @()
      foreach($tn in $tabNames){
        $found=$null
        foreach($sws in $scratchWb.Worksheets){ if($sws.Name -eq $tn){ $found=$sws; break } }
        if(-not $found){ throw "censusTabs entry not found in scratch workbook: $tn" }
        $sheetsToCopy += ,$found
      }
      if($sheetsToCopy.Count -eq 0){ throw "censusTabs is empty - nothing to merge" }
      # grouped copy - single COM call so intra-workbook refs among the copied sheets stay internal
      $hWs = $null
      foreach($existing in $wb.Worksheets){ if($existing.Name -eq $hTab){ $hWs=$existing; break } }
      $beforeSheet = if($hWs){ $hWs } else { $wb.Worksheets.Item(1) }
      $scratchWb.Worksheets.Item(@($tabNames)).Copy($beforeSheet)
      $scratchWb.Close($false)
      # "Notes" placed as FIRST sheet in the model workbook; others sit where they landed (after hTab pos)
      $notesWs=$null
      foreach($mws in $wb.Worksheets){ if($mws.Name -eq 'Notes' -and $tabNames -contains 'Notes'){ $notesWs=$mws; break } }
      if($notesWs){ $notesWs.Move($wb.Worksheets.Item(1)) }
    }

    $h = $null
    foreach($existing in $wb.Worksheets){ if($existing.Name -eq $hTab){ $h=$existing; break } }
    if(-not $h){
      $h = $wb.Worksheets.Add($wb.Worksheets.Item(1))
      $h.Name = $hTab
      $h.Move($wb.Worksheets.Item(1))
    }
    $hf = $wb.Worksheets.Item($hfTab)
    $mp = $wb.Worksheets.Item($mapTab)

    # =============================================================================================
    # BLOCK B FIRST - read HF col C (rows hfDetailStart..hfLast), skip the Margin row, write D:Q
    # =============================================================================================
    $hfRowCount = $hfLast - $hfDetailStart + 1
    $hfCRng = $hf.Range($hf.Cells($hfDetailStart,3),$hf.Cells($hfLast,3))
    $hfCVals = $hfCRng.Value2   # 2-D array [row,1] (1-based via COM's SafeArray -> PS returns object[,])

    # NOTE: only the truly-blank HF cell yields a blank .H row. A mid-block "REVENUE"/"EXPENSES"
    # banner from the HF's OWN chart of accounts is NOT special-cased here - it still gets full
    # D/B/F formulas like any other mirrored line (verified against the golden: HF's internal
    # EXPENSES banner row carries a real XLOOKUP/tag formula). The Block A row-26 "REVENUE" banner
    # is a SEPARATE fixed-template cell written later (outside this HF-row walk entirely, since the
    # walk starts at hfDetailStart) and never reaches this classification.
    $mappedRows = @()   # ordered list of hashtables: hfRow, text, kind ('blank'|'normal')
    for($i=0;$i -lt $hfRowCount;$i++){
      $hfRow = $hfDetailStart + $i
      $txt = if($hfRowCount -eq 1){ $hfCVals } else { $hfCVals[($i+1),1] }
      $txt = if($null -eq $txt){ "" } else { [string]$txt }
      $trimmed = $txt.Trim()
      if($trimmed -eq "Margin"){ continue }   # dropped - ratio row, no tag, no roll-up home
      if($trimmed -eq ""){ $mappedRows += @{hfRow=$hfRow; text=""; kind='blank'}; continue }
      $mappedRows += @{hfRow=$hfRow; text=$txt; kind='normal'}
    }
    if($mappedRows.Count -eq 0){ throw "no detail rows resolved from HF $hfTab rows $hfDetailStart..$hfLast" }

    $detailFirst = 28
    $detailLastContent = $detailFirst + $mappedRows.Count - 1
    $detailLast = $detailLastContent + 1
    $rollupFirst = $detailLast + 3
    $rollupLast = $rollupFirst + 70
    $result.detailFirst=$detailFirst; $result.detailLastContent=$detailLastContent; $result.detailLast=$detailLast
    $result.rollupFirst=$rollupFirst; $result.rollupLast=$rollupLast

    # bulk write B:Q (cols 2..17, 16 wide) for the detail block in ONE COM hop
    $nRows = $mappedRows.Count
    $blk = New-Object 'object[,]' $nRows,16
    $totalRowSet = New-Object 'System.Collections.Generic.HashSet[int]'
    $bannerBoldSet = New-Object 'System.Collections.Generic.HashSet[int]'
    for($i=0;$i -lt $nRows;$i++){ for($j=0;$j -lt 16;$j++){ $blk[$i,$j]="" } }
    for($i=0;$i -lt $nRows;$i++){
      $m = $mappedRows[$i]; $hRow = $detailFirst + $i
      if($m.kind -eq 'blank'){ continue }
      # normal row: D literal, B tag XLOOKUP, F:Q value XLOOKUP walking HF E:P
      $blk[$i,2] = $m.text
      $blk[$i,0] = "=LET(r,XLOOKUP(D{0},'{1}'!`$C`$3:`$C`${2},'{1}'!`$B`$3:`$B`${2},""""),IF(OR(r="""",r=0),"""",r))" -f $hRow,$mapTab,$mapLast
      for($cc=0;$cc -lt 12;$cc++){
        $hfValCol = HColL (5+$cc)   # HF value cols E(5)..P(16)
        $blk[$i,(4+$cc)] = "=XLOOKUP(`$D{0},'{1}'!`$C`${2}:`$C`${3},'{1}'!{4}`${2}:{4}`${3})" -f $hRow,$hfTab,$hfDetailStart,$hfLast,$hfValCol
      }
      $trimmedD = $m.text.Trim()
      if($trimmedD -match '^TOTAL ' -or $trimmedD -eq 'NET OPERATING INCOME'){ [void]$totalRowSet.Add($hRow) }
      elseif($trimmedD -eq 'REVENUE' -or $trimmedD -eq 'EXPENSES'){ [void]$bannerBoldSet.Add($hRow) }   # HF's own internal section banner - D bold only, no border, formulas intact
    }
    $rng = $h.Range($h.Cells($detailFirst,2),$h.Cells($detailFirst+$nRows-1,17))
    SetBlock $rng $blk

    # ---- Block B styling: green Calibri9 tag cells; total-row bold+border ----
    for($i=0;$i -lt $nRows;$i++){
      $m = $mappedRows[$i]; $hRow = $detailFirst + $i
      if($m.kind -eq 'normal'){
        $bc = $h.Cells($hRow,2)
        $bc.Font.Color=$script:hGREEN; $bc.Font.Name='Calibri'; $bc.Font.Size=9; $bc.Font.Bold=$false; $bc.Interior.ColorIndex=-4142
      }
    }
    foreach($hRow in $totalRowSet){
      $rr = $h.Range($h.Cells($hRow,4),$h.Cells($hRow,17))
      $rr.Font.Bold=$true; HBorderTop $rr; HBorderBottom $rr
    }
    foreach($hRow in $bannerBoldSet){ $h.Cells($hRow,4).Font.Bold=$true }
    # the ENTIRE detail block (including blank mirrored HF rows) carries the standard accounting
    # format across F:U in one bulk sweep - blank .H rows still get the format, just no value (golden
    # read-back: a blank-mirrored row like .H r34 <- HF r14 still formats F34:U34 as accounting).
    $h.Range($h.Cells($detailFirst,6),$h.Cells($detailLastContent,21)).NumberFormat=$script:hACCT

    # =============================================================================================
    # BLOCK A - OCCUPANCY (rows 5-24), links to Census
    # =============================================================================================
    SetV ($h.Cells(5,6)) ([double]$windowStart.ToOADate())
    $h.Cells(5,6).Font.Color=$script:hBLUE
    for($cc=7;$cc -le 17;$cc++){
      $prevL = HColL ($cc-1)
      $h.Cells(5,$cc).Formula = "=EOMONTH({0}5,1)" -f $prevL
      $h.Cells(5,$cc).Font.Color=$script:hBLACK
    }
    $rngHdr = $h.Range($h.Cells(5,6),$h.Cells(5,17))
    $rngHdr.Font.Bold=$true; $rngHdr.Interior.Color=$script:hGREYFILL; $rngHdr.NumberFormat=$script:hDATE
    HBorderBottom $rngHdr

    SetV ($h.Cells(7,4)) "Days"
    for($cc=6;$cc -le 17;$cc++){ $L=HColL $cc; $h.Cells(7,$cc).Formula = "=DAY(EOMONTH({0}5,0))" -f $L }

    $capLabel=@{9='IL';10='AL';11='MC'}; $capCensusRow=@{9=[int]$capRows.IL;10=[int]$capRows.AL;11=[int]$capRows.MC}
    foreach($r9 in 9,10,11){
      SetV ($h.Cells($r9,3)) "Capacity"; SetV ($h.Cells($r9,4)) $capLabel[$r9]
      for($cc=6;$cc -le 17;$cc++){
        $censusCol = HColL ($firstMonthColIdx + ($cc-6))
        $h.Cells($r9,$cc).Formula = "=Census!{0}{1}" -f $censusCol,$capCensusRow[$r9]
        $h.Cells($r9,$cc).Font.Color=$script:hGREEN
      }
    }
    SetV ($h.Cells(12,3)) "Capacity"; SetV ($h.Cells(12,4)) "Total"
    for($cc=6;$cc -le 17;$cc++){ $L=HColL $cc; $h.Cells(12,$cc).Formula="=SUM({0}9:{0}11)" -f $L; $h.Cells(12,$cc).Font.Bold=$true }

    $occLabel=@{14='IL';15='AL';16='MC'}; $occPctRow=@{14=19;15=20;16=21}
    foreach($r14 in 14,15,16){
      SetV ($h.Cells($r14,3)) "Occupied"; SetV ($h.Cells($r14,4)) $occLabel[$r14]
      for($cc=6;$cc -le 17;$cc++){ $L=HColL $cc; $h.Cells($r14,$cc).Formula="={0}{1}*{0}{2}" -f $L,$occPctRow[$r14],($r14-5) }
    }
    SetV ($h.Cells(17,3)) "Occupied"; SetV ($h.Cells(17,4)) "Total"
    for($cc=6;$cc -le 17;$cc++){ $L=HColL $cc; $h.Cells(17,$cc).Formula="=SUM({0}14:{0}16)" -f $L; $h.Cells(17,$cc).Font.Bold=$true }

    $occPctLabel=@{19='IL';20='AL';21='MC'}; $occPctCensusRow=@{19=[int]$occRows.IL;20=[int]$occRows.AL;21=[int]$occRows.MC}
    foreach($r19 in 19,20,21){
      SetV ($h.Cells($r19,3)) "Occ %"; SetV ($h.Cells($r19,4)) $occPctLabel[$r19]
      for($cc=6;$cc -le 17;$cc++){
        $censusCol = HColL ($firstMonthColIdx + ($cc-6))
        $h.Cells($r19,$cc).Formula = "=Census!{0}{1}" -f $censusCol,$occPctCensusRow[$r19]
        $h.Cells($r19,$cc).Font.Color=$script:hGREEN
      }
    }
    SetV ($h.Cells(22,3)) "Occ %"; SetV ($h.Cells(22,4)) "Total"
    for($cc=6;$cc -le 17;$cc++){ $L=HColL $cc; $h.Cells(22,$cc).Formula='=+IFERROR({0}17/{0}12,"NA")' -f $L }
    $h.Range($h.Cells(19,6),$h.Cells(21,21)).NumberFormat=$script:hPCT0
    $h.Range($h.Cells(19,6),$h.Cells(21,21)).HorizontalAlignment=-4152
    $h.Range($h.Cells(22,6),$h.Cells(22,21)).NumberFormat=$script:hPCT1

    # PRD (residents/occupied-unit denominator row is the metrics-block row - defer F24 formula
    # write until after rollupFirst/rowResidents are known; write D-label now)
    SetV ($h.Cells(24,4)) "PRD"

    # precise row list only (rows 6/8/13/18/23 are blank spacers - stay General, matching golden)
    foreach($rr in @(7,9,10,11,12,14,15,16,17,24)){ $h.Range($h.Cells($rr,6),$h.Cells($rr,21)).NumberFormat=$script:hINT }
    foreach($pair in @(@(11,12),@(16,17),@(21,22))){
      HBorderBottom ($h.Range($h.Cells($pair[0],6),$h.Cells($pair[0],17)))
      HBorderTop    ($h.Range($h.Cells($pair[1],6),$h.Cells($pair[1],17)))
    }
    $h.Cells(12,18).Font.Bold=$true; $h.Cells(17,18).Font.Bold=$true   # R-col bold follows the Total rows

    # Block B section banner + row 26/27
    SetV ($h.Cells(26,4)) "REVENUE"; $h.Cells(26,4).Font.Bold=$true

    # =============================================================================================
    # BLOCK C - ROLL-UP (rows rollupFirst..rollupFirst+79)
    # =============================================================================================
    $rowTotalRevenue=$rollupFirst+11; $rowCheckRev=$rollupFirst+13
    $rowInHouseLabor=$rollupFirst+24; $rowTotalLabor=$rollupFirst+33; $rowTotalLB=$rollupFirst+42
    $rowTotalOpex=$rollupFirst+71; $rowEBITDARM=$rollupFirst+73; $rowMgmtFee=$rollupFirst+74
    $rowEBITDAR=$rollupFirst+75; $rowMargin=$rollupFirst+76; $rowIgnore=$rollupFirst+77
    $rowCheckNOI=$rollupFirst+79
    $rowUnits=$rollupFirst+81; $rowOccUnits=$rollupFirst+82; $rowResidents=$rollupFirst+83
    $rowResidentDays=$rollupFirst+84; $rowTotalLaborDenom=$rollupFirst+85
    $rowTotalBenefitsDenom=$rollupFirst+86; $rowTotalRevenueDenom=$rollupFirst+87

    # NOW resolve PRD (F24) - depends on rowResidents
    for($cc=6;$cc -le 17;$cc++){ $L=HColL $cc; $h.Cells(24,$cc).Formula="=+{0}17*{0}7*{0}{1}" -f $L,$rowResidents }

    function HWriteCatRow($row,$B,$C){
      SetV ($h.Cells($row,2)) $B; SetV ($h.Cells($row,3)) $C
      for($cc=6;$cc -le 17;$cc++){
        $L=HColL $cc
        $h.Cells($row,$cc).Formula = "=+SUMIFS({0}`$25:{0}`${1},`$B`$25:`$B`${1},`$C{2})" -f $L,$detailLast,$row
      }
    }
    foreach($rc in $script:RollupCats){ HWriteCatRow ($rollupFirst+$rc.off) $rc.B $rc.C }
    SetV ($h.Cells($rowTotalRevenue,3)) "Total Revenue"
    for($cc=6;$cc -le 17;$cc++){ $L=HColL $cc; $h.Cells($rowTotalRevenue,$cc).Formula="=+SUM({0}{1}:{0}{2})" -f $L,($rollupFirst+0),($rollupFirst+10) }
    SetV ($h.Cells($rowCheckRev,2)) "Check"
    for($cc=6;$cc -le 17;$cc++){
      $L=HColL $cc; $hfValCol = HColL (5 + ($cc-6))
      $h.Cells($rowCheckRev,$cc).Formula = "={0}{1}-'{2}'!{3}{4}" -f $L,$rowTotalRevenue,$hfTab,$hfValCol,$hfRevRow
    }

    foreach($rc in $script:LaborCats){ HWriteCatRow ($rollupFirst+$rc.off) $rc.B $rc.C }
    SetV ($h.Cells($rowInHouseLabor,3)) "In-House Labor"
    for($cc=6;$cc -le 17;$cc++){ $L=HColL $cc; $h.Cells($rowInHouseLabor,$cc).Formula="=+SUM({0}{1}:{0}{2})" -f $L,($rollupFirst+15),($rollupFirst+23) }

    foreach($rc in $script:TempCats){ HWriteCatRow ($rollupFirst+$rc.off) $rc.B $rc.C }
    SetV ($h.Cells($rowTotalLabor,3)) "Total Labor"
    for($cc=6;$cc -le 17;$cc++){ $L=HColL $cc; $h.Cells($rowTotalLabor,$cc).Formula="=+SUM({0}{1}:{0}{2})" -f $L,$rowInHouseLabor,($rollupFirst+32) }

    foreach($rc in $script:BenefitCats){ HWriteCatRow ($rollupFirst+$rc.off) $rc.B $rc.C }
    SetV ($h.Cells($rowTotalLB,3)) "Total Labor & Benefits"
    for($cc=6;$cc -le 17;$cc++){ $L=HColL $cc; $h.Cells($rowTotalLB,$cc).Formula="=+SUM({0}{1}:{0}{2})" -f $L,$rowTotalLabor,($rollupFirst+41) }

    foreach($rc in $script:OpexCats){ HWriteCatRow ($rollupFirst+$rc.off) $rc.B $rc.C }
    SetV ($h.Cells($rowTotalOpex,3)) "Total Opex"
    for($cc=6;$cc -le 17;$cc++){ $L=HColL $cc; $h.Cells($rowTotalOpex,$cc).Formula="=+SUM({0}{1}:{0}{2})" -f $L,$rowTotalLB,($rollupFirst+70) }

    SetV ($h.Cells($rowEBITDARM,3)) "EBITDARM"
    for($cc=6;$cc -le 17;$cc++){ $L=HColL $cc; $h.Cells($rowEBITDARM,$cc).Formula="=+{0}{1}-{0}{2}" -f $L,$rowTotalRevenue,$rowTotalOpex }
    HWriteCatRow $rowMgmtFee "Management Fee" "Management Fee"
    SetV ($h.Cells($rowEBITDAR,3)) "EBITDAR"
    for($cc=6;$cc -le 17;$cc++){ $L=HColL $cc; $h.Cells($rowEBITDAR,$cc).Formula="=+{0}{1}-{0}{2}" -f $L,$rowEBITDARM,$rowMgmtFee }
    SetV ($h.Cells($rowMargin,3)) "Margin"
    for($cc=6;$cc -le 17;$cc++){ $L=HColL $cc; $h.Cells($rowMargin,$cc).Formula="=+{0}{1}/{0}{2}" -f $L,$rowEBITDAR,$rowTotalRevenue }
    SetV ($h.Cells($rowIgnore,2)) "Ignore"; SetV ($h.Cells($rowIgnore,3)) "Ignore"
    SetV ($h.Cells($rowCheckNOI,2)) "Check"
    for($cc=6;$cc -le 17;$cc++){
      $L=HColL $cc; $hfValCol = HColL (5 + ($cc-6))
      $h.Cells($rowCheckNOI,$cc).Formula = "={0}{1}-'{2}'!{3}{4}" -f $L,$rowEBITDAR,$hfTab,$hfValCol,$hfNoiRow
    }

    # ---- roll-up number formats / bold / borders ----
    # ACCT sweep covers the category rows through EBITDAR only - the Ignore label row and the blank
    # row below it (offsets 77/78) stay General (golden read-back), and Total Opex/EBITDARM/EBITDAR
    # override to the #,##0 (no-dash) int format right after.
    $h.Range($h.Cells($rollupFirst,6),$h.Cells($rowEBITDAR,21)).NumberFormat=$script:hACCT
    $h.Range($h.Cells($rowMargin,6),$h.Cells($rowMargin,21)).NumberFormat=$script:hPCT1
    foreach($chkRow in @($rowCheckRev,$rowCheckNOI)){
      $h.Cells($chkRow,2).Font.Bold=$true
      $rr = $h.Range($h.Cells($chkRow,6),$h.Cells($chkRow,21))
      $rr.NumberFormat=$script:hCHKFMT; $rr.Font.Bold=$false
      $h.Range($h.Cells($chkRow,6),$h.Cells($chkRow,17)).Interior.Color=$script:hCHECKFILL   # fill: F:Q and S only (T/U format-only, R untouched)
    }
    foreach($bp in @($rowTotalRevenue,$rowInHouseLabor,$rowTotalLabor,$rowTotalLB)){ $h.Cells($bp,3).Font.Bold=$true }
    $h.Cells($rowMargin,3).Font.Bold=$true   # Margin's label (col C) is bold like the other subtotal labels
    foreach($bv in @($rowTotalOpex,$rowEBITDARM,$rowEBITDAR)){
      $h.Range($h.Cells($bv,6),$h.Cells($bv,21)).Font.Bold=$true
      $h.Range($h.Cells($bv,6),$h.Cells($bv,21)).NumberFormat=$script:hINT   # Total Opex/EBITDARM/EBITDAR use the plain #,##0 int format, not accounting-dash
    }
    # borders span F:U (incl. R, per the h-skeleton "Borders (F:U)" contract - unlike Block A's F:Q-only rule)
    HBorderTop ($h.Range($h.Cells($rowTotalRevenue,6),$h.Cells($rowTotalRevenue,21))); HBorderBottom ($h.Range($h.Cells($rowTotalRevenue,6),$h.Cells($rowTotalRevenue,21)))
    HBorderTop ($h.Range($h.Cells($rowInHouseLabor,6),$h.Cells($rowInHouseLabor,21)))
    HBorderTop ($h.Range($h.Cells($rowTotalLabor,6),$h.Cells($rowTotalLabor,21)))
    HBorderTop ($h.Range($h.Cells($rowTotalLB,6),$h.Cells($rowTotalLB,21)))
    HBorderTop ($h.Range($h.Cells($rowTotalOpex,6),$h.Cells($rowTotalOpex,21))); HBorderBottom ($h.Range($h.Cells($rowTotalOpex,6),$h.Cells($rowTotalOpex,21)))
    HBorderTop ($h.Range($h.Cells($rowEBITDARM,6),$h.Cells($rowEBITDARM,21))); HBorderBottomDouble ($h.Range($h.Cells($rowEBITDARM,6),$h.Cells($rowEBITDARM,21)))
    HBorderTop ($h.Range($h.Cells($rowEBITDAR,6),$h.Cells($rowEBITDAR,21)));   HBorderBottomDouble ($h.Range($h.Cells($rowEBITDAR,6),$h.Cells($rowEBITDAR,21)))

    # =============================================================================================
    # BLOCK D - DENOMINATORS + METRICS
    # =============================================================================================
    SetV ($h.Cells($rowUnits,2)) "Units"
    for($cc=6;$cc -le 17;$cc++){ $L=HColL $cc; $h.Cells($rowUnits,$cc).Formula="=+{0}12" -f $L }
    SetV ($h.Cells($rowOccUnits,2)) "Occupied Units"
    for($cc=6;$cc -le 17;$cc++){ $L=HColL $cc; $h.Cells($rowOccUnits,$cc).Formula="=+{0}17" -f $L }
    SetV ($h.Cells($rowResidents,2)) "Residents/Occupied Units"
    SetV ($h.Cells($rowResidents,6)) 1.0
    for($cc=7;$cc -le 17;$cc++){ $prevL=HColL ($cc-1); $h.Cells($rowResidents,$cc).Formula="=+{0}{1}" -f $prevL,$rowResidents }
    $h.Cells($rowResidents,6).Interior.Color=49407   # orange fill signals the input; font stays black (golden)
    SetV ($h.Cells($rowResidentDays,2)) "Resident Days"
    for($cc=6;$cc -le 17;$cc++){ $L=HColL $cc; $h.Cells($rowResidentDays,$cc).Formula="=+{0}24" -f $L }
    SetV ($h.Cells($rowTotalLaborDenom,2)) "Total Labor"
    for($cc=6;$cc -le 17;$cc++){ $L=HColL $cc; $h.Cells($rowTotalLaborDenom,$cc).Formula="=+{0}{1}" -f $L,$rowTotalLabor }
    SetV ($h.Cells($rowTotalBenefitsDenom,2)) "Total Benefits"
    for($cc=6;$cc -le 17;$cc++){ $L=HColL $cc; $h.Cells($rowTotalBenefitsDenom,$cc).Formula="=+SUM({0}{1}:{0}{2})" -f $L,($rollupFirst+35),($rollupFirst+41) }
    SetV ($h.Cells($rowTotalRevenueDenom,2)) "Total Revenue"
    for($cc=6;$cc -le 17;$cc++){ $L=HColL $cc; $h.Cells($rowTotalRevenueDenom,$cc).Formula="=+{0}{1}" -f $L,$rowTotalRevenue }
    # INT format for the whole denominator block INCLUDING the R column, EXCEPT row321 (Residents)
    # which keeps its own #,##0.00x format across F:U - apply INT first, then override row321.
    $h.Range($h.Cells($rowUnits,6),$h.Cells($rowTotalRevenueDenom,21)).NumberFormat=$script:hINT
    $h.Range($h.Cells($rowResidents,6),$h.Cells($rowResidents,21)).NumberFormat=$script:hRESFMT

    foreach($mr in $script:MetricRows){
      $row = $rollupFirst + $mr.off
      SetV ($h.Cells($row,2)) $mr.D; SetV ($h.Cells($row,3)) $mr.C; SetV ($h.Cells($row,4)) $mr.D
      for($cc=6;$cc -le 17;$cc++){
        $L=HColL $cc
        $fx = switch($mr.kind){
          'unit'         { "=+SUMIFS({0}`${1}:{0}`${2},`$C`${1}:`$C`${2},`$D{3})/{0}`${4}" -f $L,$rollupFirst,$rollupLast,$row,$rowUnits }
          'k1000'        { "=+SUMIFS({0}`${1}:{0}`${2},`$C`${1}:`$C`${2},`$D{3})/1000" -f $L,$rollupFirst,$rollupLast,$row }
          'prd'          { "=+IFERROR(SUMIFS({0}`${1}:{0}`${2},`$C`${1}:`$C`${2},`$D{3})/{0}`${4},0)" -f $L,$rollupFirst,$rollupLast,$row,$rowResidentDays }
          'rentcare'     { "=+SUMIFS({0}`${1}:{0}`${2},`$C`${1}:`$C`${2},`$D{3})/SUM({0}{4}:{0}{5},{0}{6}:{0}{7})" -f $L,$rollupFirst,$rollupLast,$row,($rollupFirst+0),($rollupFirst+4),($rollupFirst+8),($rollupFirst+9) }
          'benpr'        { "=+{0}{1}/{0}{2}" -f $L,$rowTotalBenefitsDenom,$rowTotalLaborDenom }
          'healthdental' { "=+{0}{1}/{0}{2}" -f $L,($rollupFirst+60),$rowTotalLaborDenom }
          'adminK'       { "=+SUMIFS({0}`${1}:{0}`${2},`$C`${1}:`$C`${2},`$D{3})/1000+({0}{4}+{0}{5}+{0}{6})/1000" -f $L,$rollupFirst,$rollupLast,$row,($rollupFirst+56),($rollupFirst+57),($rollupFirst+58) }
          'mgmtfee'      { "=+{0}{1}/{0}{2}" -f $L,$rowMgmtFee,$rowTotalRevenueDenom }
          'revpct'       { "=+SUMIFS({0}`${1}:{0}`${2},`$C`${1}:`$C`${2},`$D{3})/{0}{4}" -f $L,$rollupFirst,$rollupLast,$row,$rowTotalRevenueDenom }
          'marketing'    { "=+SUMIFS({0}`${1}:{0}`${2},`$C`${1}:`$C`${2},`$D{3})/{0}`${4}+(({0}{5}+{0}{6}))/{0}{4}" -f $L,$rollupFirst,$rollupLast,$row,$rowUnits,($rollupFirst+54),($rollupFirst+53) }
        }
        $h.Cells($row,$cc).Formula = $fx
      }
      $fmt = if($mr.kind -in @('benpr','healthdental','rentcare')){ $script:hPCT2 }
             elseif($mr.kind -in @('mgmtfee','revpct')){ $script:hPCT2 }
             else { $script:hCHKFMT }
      $h.Range($h.Cells($row,6),$h.Cells($row,17)).NumberFormat=$fmt
    }

    # ---- R column: metric-tag validation (labor-block range) ----
    # Whether R holds "Ok" (literal, tag has no home in the labor-block range) or a real XLOOKUP
    # (tag found) is orthogonal to its NUMBER FORMAT, which follows the metric's kind - the golden
    # carries a legacy percent format on R for the percent-kind rows regardless of Ok-vs-formula
    # (1-decimal for benpr/rentcare/healthdental, 2-decimal for mgmtfee/revpct); all other kinds
    # (dollar/K-property/PRD) leave R at the workbook default (General).
    $noHomeOffsets = @(91,112,114)   # Benefits & PR Taxes / Management Fee / Capex Reserve (Deal B golden: plain "Ok")
    foreach($mr in $script:MetricRows){
      $row = $rollupFirst + $mr.off
      if($noHomeOffsets -contains $mr.off){ SetV ($h.Cells($row,18)) "Ok" }
      else { $h.Cells($row,18).Formula = "=+XLOOKUP(D{0},`$C`${1}:`$C`${2},`$C`${1}:`$C`${2})" -f $row,($rollupFirst+15),($rollupFirst+70) }
      $h.Cells($row,18).Font.Color=$script:hBLACK
      if($mr.kind -in @('benpr','rentcare','healthdental')){ $h.Cells($row,18).NumberFormat=$script:hPCT1 }
      elseif($mr.kind -in @('mgmtfee','revpct')){ $h.Cells($row,18).NumberFormat=$script:hPCT2B }
    }

    # =============================================================================================
    # BLOCK E - TOTALS COLUMNS S/T/U (T12/T6/T3)
    # =============================================================================================
    SetV ($h.Cells(2,19)) 1.0; SetV ($h.Cells(2,20)) 2.0; SetV ($h.Cells(2,21)) 4.0
    SetV ($h.Cells(3,19)) "T12"; SetV ($h.Cells(3,20)) "T6"; SetV ($h.Cells(3,21)) "T3"
    $h.Cells(4,19).Formula="=F5"; $h.Cells(4,20).Formula="=+EOMONTH(S4,6)"; $h.Cells(4,21).Formula="=+EOMONTH(T4,3)"
    $h.Cells(5,19).Formula="=Q5"; $h.Cells(5,20).Formula="=+S5"; $h.Cells(5,21).Formula="=+T5"
    $hdrBlock = $h.Range($h.Cells(2,19),$h.Cells(5,21))
    $hdrBlock.HorizontalAlignment=-4152
    $h.Range($h.Cells(2,19),$h.Cells(2,21)).Font.Color=$script:hBLUE
    $h.Range($h.Cells(4,19),$h.Cells(5,21)).NumberFormat=$script:hDATE
    $rng5 = $h.Range($h.Cells(5,19),$h.Cells(5,21)); $rng5.Font.Bold=$true; $rng5.Interior.Color=$script:hGREYFILL
    HBorderBottom $rng5
    HBorderTop ($h.Range($h.Cells(6,19),$h.Cells(6,21)))

    function HWindowFlow($row){
      foreach($p in @(@(19,'S'),@(20,'T'),@(21,'U'))){
        $cIdx=$p[0]; $L=HColL $cIdx
        $h.Cells($row,$cIdx).Formula = "=IFERROR(SUMIFS(`$E{0}:`$Q{0},`$E`$5:`$Q`$5,"">=""&{1}`$4,`$E`$5:`$Q`$5,""<=""&{1}`$5),0)*{1}`$2" -f $row,$L
      }
    }
    function HWindowStock($row){
      foreach($p in @(@(19,'S'),@(20,'T'),@(21,'U'))){
        $cIdx=$p[0]; $L=HColL $cIdx
        $h.Cells($row,$cIdx).Formula = "=+IFERROR(AVERAGEIFS(`$E{0}:`$Q{0},`$E`$5:`$Q`$5,"">=""&{1}`$4,`$E`$5:`$Q`$5,""<=""&{1}`$5),""NA"")" -f $row,$L
      }
    }
    function HWindowOccPct($row,$occRow,$capRow){
      foreach($p in @(@(19,'S'),@(20,'T'),@(21,'U'))){
        $cIdx=$p[0]; $L=HColL $cIdx
        $h.Cells($row,$cIdx).Formula = "=+IFERROR({0}{1}/{0}{2},""NA"")" -f $L,$occRow,$capRow
      }
    }
    function HTranslate($row,$fTemplate){
      # fTemplate is a scriptblock producing the formula text given a column letter, matching the
      # exact template used to build col F for this row (denominators/metrics/subtotals).
      foreach($p in @(@(19,'S'),@(20,'T'),@(21,'U'))){ $h.Cells($row,$p[0]).Formula = (& $fTemplate $p[1]) }
    }

    # Block A stock/occ%/flow rows. Rows 12/17 (the Total rows) use a plain AVERAGEIFS translate -
    # NO IFERROR wrapper, unlike the IL/AL/MC member rows (verified against the golden: S12/S17 have
    # no IFERROR while S9/S14 etc. do) - and are bold like their col-F counterparts.
    HWindowFlow 7; HWindowFlow 24
    foreach($r9 in 9,10,11,14,15,16){ HWindowStock $r9 }
    HTranslate 12 { param($c) '=+AVERAGEIFS($E12:$Q12,$E$5:$Q$5,">="&{0}$4,$E$5:$Q$5,"<="&{0}$5)' -f $c }
    HTranslate 17 { param($c) '=+AVERAGEIFS($E17:$Q17,$E$5:$Q$5,">="&{0}$4,$E$5:$Q$5,"<="&{0}$5)' -f $c }
    $h.Range($h.Cells(12,19),$h.Cells(12,21)).Font.Bold=$true; $h.Range($h.Cells(17,19),$h.Cells(17,21)).Font.Bold=$true
    HWindowOccPct 19 14 9; HWindowOccPct 20 15 10; HWindowOccPct 21 16 11
    HTranslate 22 { param($c) '=+IFERROR({0}17/{0}12,"NA")' -f $c }
    # precise row list (skip blank spacer row 13 - stays General, matching golden)
    foreach($rr in @(7,9,10,11,12,14,15,16,17,24)){ $h.Range($h.Cells($rr,19),$h.Cells($rr,21)).NumberFormat=$script:hINT }
    $h.Range($h.Cells(19,19),$h.Cells(21,21)).NumberFormat=$script:hPCT0
    $h.Range($h.Cells(22,19),$h.Cells(22,21)).NumberFormat=$script:hPCT1

    # Block B detail rows: flow (windowed) for every normal/total row
    for($i=0;$i -lt $nRows;$i++){
      $m=$mappedRows[$i]; if($m.kind -ne 'normal'){ continue }
      $hRow = $detailFirst+$i; HWindowFlow $hRow
      $h.Range($h.Cells($hRow,19),$h.Cells($hRow,21)).NumberFormat=$script:hACCT
    }
    foreach($hRow in $totalRowSet){ $rr=$h.Range($h.Cells($hRow,19),$h.Cells($hRow,21)); $rr.Font.Bold=$true; HBorderTop $rr; HBorderBottom $rr }

    # Block C category rows (flow/windowed)
    foreach($group in @($script:RollupCats,$script:LaborCats,$script:TempCats,$script:BenefitCats,$script:OpexCats)){
      foreach($rc in $group){ HWindowFlow ($rollupFirst+$rc.off) }
    }
    HWindowFlow $rowMgmtFee

    # Block C subtotal rows ("translate" - same formula, S/T/U col substituted)
    HTranslate $rowTotalRevenue { param($c) '=+SUM({0}{1}:{0}{2})' -f $c,($rollupFirst+0),($rollupFirst+10) }
    HTranslate $rowInHouseLabor { param($c) '=+SUM({0}{1}:{0}{2})' -f $c,($rollupFirst+15),($rollupFirst+23) }
    HTranslate $rowTotalLabor   { param($c) '=+SUM({0}{1}:{0}{2})' -f $c,$rowInHouseLabor,($rollupFirst+32) }
    HTranslate $rowTotalLB      { param($c) '=+SUM({0}{1}:{0}{2})' -f $c,$rowTotalLabor,($rollupFirst+41) }
    HTranslate $rowTotalOpex    { param($c) '=+SUM({0}{1}:{0}{2})' -f $c,$rowTotalLB,($rollupFirst+70) }
    HTranslate $rowEBITDARM     { param($c) '=+{0}{1}-{0}{2}' -f $c,$rowTotalRevenue,$rowTotalOpex }
    HTranslate $rowEBITDAR      { param($c) '=+{0}{1}-{0}{2}' -f $c,$rowEBITDARM,$rowMgmtFee }
    HTranslate $rowMargin       { param($c) '=+{0}{1}/{0}{2}' -f $c,$rowEBITDAR,$rowTotalRevenue }
    foreach($subRow in @($rowTotalRevenue,$rowInHouseLabor,$rowTotalLabor,$rowTotalLB,$rowTotalOpex,$rowEBITDARM,$rowEBITDAR,$rowMargin)){
      $fmt = if($subRow -eq $rowMargin){ $script:hPCT1 }
             elseif($subRow -in @($rowTotalOpex,$rowEBITDARM,$rowEBITDAR)){ $script:hINT }
             else { $script:hACCT }
      $h.Range($h.Cells($subRow,19),$h.Cells($subRow,21)).NumberFormat=$fmt
    }
    $h.Range($h.Cells($rollupFirst+72,6),$h.Cells($rollupFirst+72,21)).NumberFormat='General'   # blank row between Total Opex and EBITDARM stays General (golden read-back quirk)
    foreach($bv in @($rowTotalOpex,$rowEBITDARM,$rowEBITDAR)){ $h.Range($h.Cells($bv,19),$h.Cells($bv,21)).Font.Bold=$true }
    HBorderTop ($h.Range($h.Cells($rowTotalRevenue,19),$h.Cells($rowTotalRevenue,21))); HBorderBottom ($h.Range($h.Cells($rowTotalRevenue,19),$h.Cells($rowTotalRevenue,21)))
    HBorderTop ($h.Range($h.Cells($rowInHouseLabor,19),$h.Cells($rowInHouseLabor,21)))
    HBorderTop ($h.Range($h.Cells($rowTotalLabor,19),$h.Cells($rowTotalLabor,21)))
    HBorderTop ($h.Range($h.Cells($rowTotalLB,19),$h.Cells($rowTotalLB,21)))
    HBorderTop ($h.Range($h.Cells($rowTotalOpex,19),$h.Cells($rowTotalOpex,21))); HBorderBottom ($h.Range($h.Cells($rowTotalOpex,19),$h.Cells($rowTotalOpex,21)))
    HBorderTop ($h.Range($h.Cells($rowEBITDARM,19),$h.Cells($rowEBITDARM,21))); HBorderBottomDouble ($h.Range($h.Cells($rowEBITDARM,19),$h.Cells($rowEBITDARM,21)))
    HBorderTop ($h.Range($h.Cells($rowEBITDAR,19),$h.Cells($rowEBITDAR,21)));   HBorderBottomDouble ($h.Range($h.Cells($rowEBITDAR,19),$h.Cells($rowEBITDAR,21)))

    # Checks - S only, ties to HF annual-total column Q
    $h.Cells($rowCheckRev,19).Formula = "=S{0}-'{1}'!Q{2}" -f $rowTotalRevenue,$hfTab,$hfRevRow
    $h.Cells($rowCheckNOI,19).Formula = "=S{0}-'{1}'!Q{2}" -f $rowEBITDAR,$hfTab,$hfNoiRow
    foreach($chkRow in @($rowCheckRev,$rowCheckNOI)){
      $h.Cells($chkRow,19).Interior.Color=$script:hCHECKFILL; $h.Cells($chkRow,19).NumberFormat=$script:hCHKFMT
      $h.Range($h.Cells($chkRow,20),$h.Cells($chkRow,21)).NumberFormat=$script:hCHKFMT
    }

    # Denominators + metrics ("translate" rows)
    HTranslate $rowUnits             { param($c) '=+{0}12' -f $c }
    HTranslate $rowOccUnits           { param($c) '=+{0}17' -f $c }
    HTranslate $rowResidents          { param($c) '=$F${0}' -f $rowResidents }
    HTranslate $rowResidentDays       { param($c) '=+{0}24' -f $c }
    HTranslate $rowTotalLaborDenom    { param($c) '=+{0}{1}' -f $c,$rowTotalLabor }
    HTranslate $rowTotalBenefitsDenom { param($c) '=+SUM({0}{1}:{0}{2})' -f $c,($rollupFirst+35),($rollupFirst+41) }
    HTranslate $rowTotalRevenueDenom  { param($c) '=+{0}{1}' -f $c,$rowTotalRevenue }
    foreach($dr in @($rowUnits,$rowOccUnits,$rowResidentDays,$rowTotalLaborDenom,$rowTotalBenefitsDenom,$rowTotalRevenueDenom)){
      $h.Range($h.Cells($dr,19),$h.Cells($dr,21)).NumberFormat=$script:hINT
    }
    $h.Range($h.Cells($rowResidents,19),$h.Cells($rowResidents,21)).NumberFormat=$script:hRESFMT

    foreach($mr in $script:MetricRows){
      $row = $rollupFirst + $mr.off
      HTranslate $row {
        param($c)
        switch($mr.kind){
          'unit'         { "=+SUMIFS({0}`${1}:{0}`${2},`$C`${1}:`$C`${2},`$D{3})/{0}`${4}" -f $c,$rollupFirst,$rollupLast,$row,$rowUnits }
          'k1000'        { "=+SUMIFS({0}`${1}:{0}`${2},`$C`${1}:`$C`${2},`$D{3})/1000" -f $c,$rollupFirst,$rollupLast,$row }
          'prd'          { "=+IFERROR(SUMIFS({0}`${1}:{0}`${2},`$C`${1}:`$C`${2},`$D{3})/{0}`${4},0)" -f $c,$rollupFirst,$rollupLast,$row,$rowResidentDays }
          'rentcare'     { "=+SUMIFS({0}`${1}:{0}`${2},`$C`${1}:`$C`${2},`$D{3})/SUM({0}{4}:{0}{5},{0}{6}:{0}{7})" -f $c,$rollupFirst,$rollupLast,$row,($rollupFirst+0),($rollupFirst+4),($rollupFirst+8),($rollupFirst+9) }
          'benpr'        { "=+{0}{1}/{0}{2}" -f $c,$rowTotalBenefitsDenom,$rowTotalLaborDenom }
          'healthdental' { "=+{0}{1}/{0}{2}" -f $c,($rollupFirst+60),$rowTotalLaborDenom }
          'adminK'       { "=+SUMIFS({0}`${1}:{0}`${2},`$C`${1}:`$C`${2},`$D{3})/1000+({0}{4}+{0}{5}+{0}{6})/1000" -f $c,$rollupFirst,$rollupLast,$row,($rollupFirst+56),($rollupFirst+57),($rollupFirst+58) }
          'mgmtfee'      { "=+{0}{1}/{0}{2}" -f $c,$rowMgmtFee,$rowTotalRevenueDenom }
          'revpct'       { "=+SUMIFS({0}`${1}:{0}`${2},`$C`${1}:`$C`${2},`$D{3})/{0}{4}" -f $c,$rollupFirst,$rollupLast,$row,$rowTotalRevenueDenom }
          'marketing'    { "=+SUMIFS({0}`${1}:{0}`${2},`$C`${1}:`$C`${2},`$D{3})/{0}`${4}+(({0}{5}+{0}{6}))/{0}{4}" -f $c,$rollupFirst,$rollupLast,$row,$rowUnits,($rollupFirst+54),($rollupFirst+53) }
        }
      }
      $fmt = if($mr.kind -in @('benpr','healthdental','rentcare','mgmtfee','revpct')){ $script:hPCT2 } else { $script:hCHKFMT }
      $h.Range($h.Cells($row,19),$h.Cells($row,21)).NumberFormat=$fmt
    }

    # ---- widths + gridlines ----
    $W = @{1=8.57;2=25;3=19;4=36.57;18=30.57;19=11;20=11;21=11}
    foreach($cc in $W.Keys){ $h.Columns.Item($cc).ColumnWidth=[double]$W[$cc] }
    $h.Range($h.Cells(1,5),$h.Cells(1,17)).EntireColumn.ColumnWidth=12.14
    $h.Name = $hTab
    $wb.Windows.Item(1).DisplayGridlines=$false

    # =============================================================================================
    # SELF-RUN GATES
    # =============================================================================================
    $xl.Calculation=-4105
    $xl.CalculateFull()
    Start-Sleep -Milliseconds 400
    $xl.CalculateFull()

    # gate 1: zero formula errors workbook-wide
    $errCount=0; $errList=@()
    foreach($ws in $wb.Worksheets){
      try {
        $errRng = $ws.UsedRange.SpecialCells(-4123,16)
        foreach($area in $errRng.Areas){ foreach($cell in $area.Cells){ $errCount++; if($errList.Count -lt 30){ $errList += ($ws.Name+"!"+$cell.Address($false,$false)) } } }
      } catch { }   # throws when clean on that sheet
    }
    $result.errorCount=$errCount; $result.errorCells=$errList

    # gate 2: Check rows = 0 across F:Q, and S-column checks = 0
    $revVals=@(); $noiVals=@()
    for($cc=6;$cc -le 17;$cc++){ $revVals += [double]$h.Cells($rowCheckRev,$cc).Value2; $noiVals += [double]$h.Cells($rowCheckNOI,$cc).Value2 }
    $result.checkRevValues=$revVals; $result.checkNoiValues=$noiVals
    $result.checkSValues=@{ rev=[double]$h.Cells($rowCheckRev,19).Value2; noi=[double]$h.Cells($rowCheckNOI,19).Value2 }
    $tol=1e-6
    if(($revVals | Where-Object { [Math]::Abs($_) -gt $tol }).Count -gt 0){ throw "GATE FAIL: revenue Check row $rowCheckRev non-zero: $($revVals -join ',')" }
    if(($noiVals | Where-Object { [Math]::Abs($_) -gt $tol }).Count -gt 0){ throw "GATE FAIL: NOI Check row $rowCheckNOI non-zero: $($noiVals -join ',')" }
    if([Math]::Abs($result.checkSValues.rev) -gt $tol){ throw "GATE FAIL: S-column revenue Check non-zero: $($result.checkSValues.rev)" }
    if([Math]::Abs($result.checkSValues.noi) -gt $tol){ throw "GATE FAIL: S-column NOI Check non-zero: $($result.checkSValues.noi)" }
    if($errCount -gt 0){ throw "GATE FAIL: $errCount formula error(s): $($errList -join '; ')" }

    # gate 3: every mapping tag with non-zero dollars has a roll-up C-key home
    $rollupKeys = New-Object 'System.Collections.Generic.HashSet[string]'
    foreach($group in @($script:RollupCats,$script:LaborCats,$script:TempCats,$script:BenefitCats,$script:OpexCats)){
      foreach($rc in $group){ [void]$rollupKeys.Add($rc.C) }
    }
    [void]$rollupKeys.Add("Management Fee"); [void]$rollupKeys.Add("Ignore")
    $mapBRng = $mp.Range($mp.Cells(3,2),$mp.Cells($mapLast,2)).Value2
    $missingTags=@()
    if($mapLast -ge 3){
      for($i=1;$i -le ($mapLast-2);$i++){
        $tg = if(($mapLast-2) -eq 1){ $mapBRng } else { $mapBRng[$i,1] }
        $tg = if($null -eq $tg){ "" } else { [string]$tg }
        $tg = $tg.Trim()
        if($tg -ne "" -and -not $rollupKeys.Contains($tg) -and $missingTags -notcontains $tg){ $missingTags += $tg }
      }
    }
    if($missingTags.Count -gt 0){ throw "GATE FAIL: mapping tag(s) with no roll-up C-key home: $($missingTags -join ', ')" }

    # gate 4: external-link scan = 0 hits (census-merge landmine)
    $extCount=0
    foreach($ws in $wb.Worksheets){
      try{
        $used=$ws.UsedRange
        foreach($area in $used.Areas){
          foreach($cell in $area.Cells){
            $fx = $cell.Formula
            if($fx -and (("$fx".Contains("[")) -or ("$fx" -match "http"))){ $extCount++ }
          }
        }
      } catch {}
    }
    $result.externalLinkCount=$extCount
    if($extCount -gt 0){ throw "GATE FAIL: $extCount external-link reference(s) found post-merge" }

    # gate 5: census completeness - all 12 Block A capacity + Occ% cells non-zero; row-5 dates == Census header dates
    $censusOk=$true
    foreach($r9 in 9,10,11,19,20,21){
      for($cc=6;$cc -le 17;$cc++){ if([double]$h.Cells($r9,$cc).Value2 -eq 0){ $censusOk=$false } }
    }
    $census = $wb.Worksheets.Item("Census")
    for($cc=6;$cc -le 17;$cc++){
      $censusCol = HColL ($firstMonthColIdx + ($cc-6))
      $hd = [double]$h.Cells(5,$cc).Value2
      $cd = [double]$census.Cells($headerDateRow,(HColIdx $censusCol)).Value2
      if([Math]::Abs($hd-$cd) -gt 0){ $censusOk=$false }
    }
    $result.censusGateOk=$censusOk
    if(-not $censusOk){ throw "GATE FAIL: census completeness / date-alignment check failed" }

    # gate 6: styling spot-assertions
    $styleOk=$true
    if($h.Cells($detailFirst,2).Font.Color -ne $script:hGREEN -or $h.Cells($detailFirst,2).Font.Name -ne 'Calibri'){ $styleOk=$false }
    if($h.Cells($rowCheckRev,6).Interior.Color -ne $script:hCHECKFILL){ $styleOk=$false }
    if($h.Cells(2,19).Font.Color -ne $script:hBLUE -or $h.Cells(2,19).HorizontalAlignment -ne -4152){ $styleOk=$false }
    if(-not $h.Cells($detailLastContent,4).Font.Bold){ $styleOk=$false }
    $result.styleGateOk=$styleOk
    if(-not $styleOk){ throw "GATE FAIL: styling spot-assertion failed" }

    # ---- save (Save() in place - never SaveAs) ----
    $saveOk=$false
    for($k=0;$k -lt 8 -and -not $saveOk;$k++){ try{ $wb.Save(); $saveOk=$true }catch{ Start-Sleep -Milliseconds 700 } }
    $result.saveOk=$saveOk
    if(-not $saveOk){ throw "workbook Save() failed after retries" }

    $success = $true
  } finally {
    if($wb -and -not $success){ try{ $wb.Close($false) }catch{} }
    elseif($wb){ try{ $wb.Close($true) }catch{} }
    Stop-TrackedExcel
    if($xl){ try{ $xl.Quit() }catch{}; [System.Runtime.InteropServices.Marshal]::ReleaseComObject($xl)|Out-Null }
  }

  Write-Output ($result | ConvertTo-Json -Depth 6)
  return [pscustomobject]$result
}
