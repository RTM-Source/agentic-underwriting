# ============================================================================================
# STATUS (2026-07-14): the proven OneSite PARSER — the starting point for the engine, not the engine.
# Source-side (anchors, blocks, leaf-sum finance, reconcile-to-the-cent) is correct; reuse it. Output side
# PREDATES the v1 golden contract: 16-col AW:BL (no Group Unit Code BM, no A/B/C tiering) and the blue
# summary groups on the WRONG key (Unit Code, not Group Unit Code). Paths are hardcoded to the two OneSite
# test files it was built against. EXTEND into the new engine; do NOT run as-is. Gap list: ../RR-FORMATTING-DESIGN.md
# -> "What the next build must change".
# ============================================================================================
# rr_build_onesite_v2.ps1 — OneSite .RR build, upgraded to the Deal D operator platform's target
# conventions (see Investments/Data/Rent Roll Example.xlsx). Adds the 16-col layout (P/SP + Resident Type), the composite
# Unit Code / Basic Unit Code, 2nd Resident as a $ amount (the 2nd-occupant fee), the inactive-down
# occupancy fix, and a per-unit-type stats block. Reconciles Market/Rent/Care/2nd-Occ.
#
# SAFETY (learned the hard way 2026-06-29): builds on a COPY ("<name> .RR.xls"); never writes the
# operator's original. PowerShell var names are CASE-INSENSITIVE: write-column constants use a $w
# prefix so a loop counter ($order) can't clobber a column index ($wOrder) and scribble a diagonal.
#
# Source columns (OneSite): A=Unit E=FloorPlan I=Resident N=Move-in V=Market Z=Actual AG=2ndOccFee AI=Care
# Target block AW(49)..BL(64), header row 13, data row-aligned to each anchor.
# ============================================================================================
$ErrorActionPreference="Stop"
[System.Threading.Thread]::CurrentThread.CurrentCulture=[System.Globalization.CultureInfo]'en-US'
# Concurrency-safe cleanup: the raw windowless sweep is hook-blocked (it cross-kills a peer deal's Excel in
# another terminal). Dot-source the lib and use Clear-OrphanExcel + New-/Stop-TrackedExcel (memory never-blanket-kill-excel).
. "<repo-root>\Investments\lib\HF-Build-Lib.ps1"   # set <repo-root> to your clone path
Clear-OrphanExcel | Out-Null
$dir="<repo-root>\Investments\Data\<rent-roll-folder>"   # deal-specific input folder
$xl=New-ExcelTracked
$xl.AskToUpdateLinks=$false
function Retry($sb){ for($k=0;$k -lt 8;$k++){ try{ return & $sb }catch{ Start-Sleep -Milliseconds 500 } }; & $sb }

# target columns (new 16-col WCH layout) — $w prefix keeps them clear of loop vars (case-insensitive!)
$wOrder=49;$wUnit=50;$wUType=51;$wPSP=52;$wCare=53;$wSqFt=54;$wRType=55;$wCap=56;$wOcc=57;$wMoveIn=58;$wMkt=59;$wRent=60;$wCareF=61;$wSec=62;$wUCode=63;$wBCode=64
# source columns
$cA=1;$cE=5;$cI=9;$cN=14;$cV=22;$cZ=26;$cAG=33;$cAI=35

function AptNum($a){ $m=[regex]::Match("$a",'\d+'); if($m.Success){return [int]$m.Value}; return $null }
function NumOf($x){ if($x -eq $null){return 0.0}; try{ return [double]([string]$x -replace '[(),$]','') }catch{ return 0.0 } }

function Build-RR($srcFile){
  $src=Join-Path $dir $srcFile
  # versioned output (house standard <Community>—RR_v#.xls; bump _v# by 1, never overwrite input)
  $emdash=[char]0x2014
  $nm=[System.IO.Path]::GetFileNameWithoutExtension($srcFile)
  $ver=1; if($nm -match '_v(\d+)$'){ $ver=[int]$Matches[1]; $nm=$nm -replace '_v\d+$','' }
  if($nm.EndsWith($emdash+'RR')){ $nm=$nm.Substring(0,$nm.Length-3) }
  if($nm.EndsWith(' Rent Roll')){ $nm=$nm.Substring(0,$nm.Length-10) }
  $comm=$nm.Trim()
  $out=Join-Path $dir ("{0}{1}RR_v{2}.xls" -f $comm,$emdash,($ver+1))
  Copy-Item -LiteralPath $src -Destination $out -Force          # build on a copy; never touch the original
  $wb=Retry { $xl.Workbooks.Open($out,0,$false) }
  try {
  $ws=$wb.Worksheets.Item(1)
  $ur=$ws.UsedRange; $nrows=$ur.Rows.Count
  $v=$ur.Value2

  # guard: refuse to build if the source looks already-corrupted (anchor col A should be a unit code, not a bare 1)
  # (the try/finally below ensures this deliberate throw still releases $wb — otherwise the abort itself orphans+locks it)
  if("$($v[17,$cA])" -eq "1"){ throw "ABORT $srcFile : A17='1' -> source already corrupted; restore from OneDrive version history first." }

  $secRows=@()
  for($r=1;$r -le $nrows;$r++){ if("$($v[$r,$cI])" -match '^TOTAL (Assisted|Memory|Independent)'){ $secRows+=$r } }
  $rrMkt=0.0;$rrRent=0.0;$rrCare=0.0;$rrSec=0.0
  foreach($sr in $secRows){ $rrMkt+=NumOf $v[$sr,$cV]; $rrRent+=NumOf $v[$sr,$cZ]; $rrCare+=NumOf $v[$sr,$cAI]; $rrSec+=NumOf $v[$sr,$cAG] }
  $detailEnd= if($secRows.Count){ ($secRows|Measure-Object -Maximum).Maximum } else { $nrows }
  $propRow=$null
  for($r=$detailEnd;$r -le $nrows -and -not $propRow;$r++){ for($c=1;$c -le 40;$c++){ if("$($v[$r,$c])" -match '^TOTAL Aster Ridge'){ $propRow=$r; break } } }
  if(-not $propRow){ $propRow=$detailEnd }

  $anchors=@()
  for($r=1;$r -lt $detailEnd;$r++){
    $a=$v[$r,$cA]; $e="$($v[$r,$cE])"
    if($a -eq $null -or "$a" -eq ""){ continue }
    if("$a" -in @('Assisted Living','Memory Care','Independent Living')){ continue }
    if($e -match '.\s-\s\d+'){ $anchors+=$r }
  }

  $order=0
  $flags=@{ sline=0; vacant=0; movein=0; moveout=0; inactive=0; secfee=0 }
  $codeKeys=New-Object System.Collections.ArrayList   # distinct composite Unit Codes, computed in PS (formulas aren't calc'd yet)
  $first=$anchors[0]; $last=$anchors[$anchors.Count-1]

  for($ai=0;$ai -lt $anchors.Count;$ai++){
    $a=$anchors[$ai]
    $blockEnd= if($ai -lt $anchors.Count-1){ $anchors[$ai+1]-1 } else { $detailEnd-1 }
    $status=New-Object System.Collections.ArrayList
    $anchorI="$($v[$a,$cI])"
    if($anchorI -match '^Vacant'){ [void]$status.Add(@($a,'V')) }
    elseif($anchorI -match '(?i)inactive|down'){ [void]$status.Add(@($a,'V')) }
    elseif($anchorI -ne ""){ [void]$status.Add(@($a,'R')) }
    for($r=$a+1;$r -le $blockEnd;$r++){
      $ca=$v[$r,$cA]; if($ca -ne $null -and "$ca" -ne ""){ break }
      $ii="$($v[$r,$cI])"
      if($ii -eq ""){ continue }
      if($ii -match '^TOTAL '){ break }
      if($ii -match 'TOTAL$' -and $ii -notmatch '^TOTAL'){ continue }
      if($ii -match '^Vacant'){ [void]$status.Add(@($r,'V')); continue }
      if($ii -match '(?i)inactive|down'){ [void]$status.Add(@($r,'V')); continue }
      [void]$status.Add(@($r,'R'))
    }
    $leafRows=@(); foreach($s in $status){ $leafRows+=$s[0] }
    $residentRows=@(); foreach($s in $status){ if($s[1] -eq 'R'){ $residentRows+=$s[0] } }
    $last2=$status[$status.Count-1]
    $occupied=($last2[1] -eq 'R')
    $occRow= if($occupied -and $residentRows.Count){ $residentRows[$residentRows.Count-1] } else { $null }
    $code="$($v[$a,$cA])"
    $ev="$($v[$a,$cE])"; $sqftA= if($ev -match ' - (\d+)\s*$'){ [int]$Matches[1] } else { 1 }
    $isS= ($code -match 'S\s*$') -or ($sqftA -eq 0)   # shared 2nd-occupant: code suffix S or SqFt 0

    if($isS){ $flags.sline++ }
    if(-not $occupied){ $flags.vacant++ }
    if($occupied -and ($status|Where-Object{$_[1] -eq 'V'})){ $flags.movein++ }
    if((-not $occupied) -and $residentRows.Count){ $flags.moveout++ }
    if($status|Where-Object{ "$($v[$_[0],$cI])" -match '(?i)inactive|down' }){ $flags.inactive++ }

    $order++
    $apt=AptNum $code

    $ws.Cells($a,$wOrder).Value2=[double]$order
    $ws.Cells($a,$wUnit).Formula="=A${a}"   # Unit = raw operator unit code (matches example's raw room-id)
    $ws.Cells($a,$wUType).Formula="=TRIM(LEFT(E${a},FIND("" - "",E${a})-1))"
    $ws.Cells($a,$wPSP).Value2= if($isS){"S"}else{"P"}
    $ws.Cells($a,$wCare).Formula="=LEFT(AY${a},2)"
    $ws.Cells($a,$wSqFt).Formula="=IFERROR(VALUE(TRIM(MID(E${a},FIND("" - "",E${a})+3,20))),0)"
    $ws.Cells($a,$wRType).Value2=""
    $ws.Cells($a,$wCap).Formula="=IF(BB${a}>0,1,0)"
    if($occupied){ $ws.Cells($a,$wMoveIn).Formula="=N${occRow}" } else { $ws.Cells($a,$wMoveIn).Value2="" }
    $ws.Cells($a,$wOcc).Formula="=IF(BF${a}="""",0,1)"
    if($leafRows.Count -eq 1){
      $ws.Cells($a,$wMkt).Formula="=IFERROR(VALUE(V$($leafRows[0])),"""")"
      $ws.Cells($a,$wRent).Formula="=IFERROR(VALUE(Z$($leafRows[0])),"""")"
      $ws.Cells($a,$wCareF).Formula="=IFERROR(VALUE(AI$($leafRows[0])),"""")"
      $ws.Cells($a,$wSec).Formula="=IFERROR(VALUE(AG$($leafRows[0])),0)"
    } else {
      $ws.Cells($a,$wMkt).Formula="="+(($leafRows|%{"IFERROR(VALUE(V$_),0)"}) -join "+")
      $ws.Cells($a,$wRent).Formula="="+(($leafRows|%{"IFERROR(VALUE(Z$_),0)"}) -join "+")
      $ws.Cells($a,$wCareF).Formula="="+(($leafRows|%{"IFERROR(VALUE(AI$_),0)"}) -join "+")
      $ws.Cells($a,$wSec).Formula="="+(($leafRows|%{"IFERROR(VALUE(AG$_),0)"}) -join "+")
    }
    if(($leafRows|%{NumOf $v[$_,$cAG]}|Measure-Object -Sum).Sum -gt 0){ $flags.secfee++ }
    $ws.Cells($a,$wUCode).Formula="=IF(AX${a}="""","""",BA${a}&"" | ""&AY${a}&"" | ""&BB${a})"
    $shared= if($isS){ "&"" SHARED""" } else { "" }
    $ws.Cells($a,$wBCode).Formula="=IF(AX${a}="""","""",BA${a}&"" | ""&AY${a}$shared)"
    # mirror the Unit Code formula in PS for the stats keys (formula cells aren't calc'd until the end)
    $ut= if($ev -match '^(.*?)\s-\s'){ $Matches[1].Trim() } else { $ev.Trim() }
    $ct= if($ut.Length -ge 2){ $ut.Substring(0,2) } else { $ut }
    $key="$ct | $ut | $sqftA"
    if(-not $codeKeys.Contains($key)){ [void]$codeKeys.Add($key) }
  }

  # ---- colors / format constants (match the WCH.RR example) ----
  $cYellow = 255 + 255*256 + 204*65536    # #FFFFCC header fill
  $cNavy   = 9 + 40*256 + 76*65536         # #09284C blue-table header band
  $cWhite  = 16777215
  $acct    = '_(* #,##0_);_(* (#,##0);_(* "-"_);_(@_)'
  $acctPct = '0.0%'

  # ---- right block header (row 13): yellow, bold, centered, Calibri 11 ----
  $hdr=@('Order','Unit','Unit Type','P/SP','Care Type','SqFt','Resident Type','Capacity','Occupancy','Move in Date','Market Rent','In-Place Rent','In-Place Care','2nd Resident','Unit Code','Basic Unit Code')
  for($c=0;$c -lt $hdr.Count;$c++){ $ws.Cells(13,$wOrder+$c).Value2=$hdr[$c] }
  $hr=$ws.Range($ws.Cells(13,$wOrder),$ws.Cells(13,$wBCode))
  $hr.Interior.Color=$cYellow; $hr.Font.Bold=$true; $hr.Font.Name='Calibri'; $hr.Font.Size=11; $hr.HorizontalAlignment=-4108
  # right block column widths (from example BQ:CF)
  $rw=@(7.9,11.0,11.1,12.7,14.0,9.6,17.6,12.9,14.7,17.1,16.3,16.7,16.6,16.6,31.6,19.7)
  for($c=0;$c -lt $rw.Count;$c++){ $ws.Columns.Item($wOrder+$c).ColumnWidth=$rw[$c] }

  # ---- right block number formats ----
  $ws.Range($ws.Cells($first,$wSqFt),$ws.Cells($last,$wSqFt)).NumberFormat="#,##0"
  $ws.Range($ws.Cells($first,$wCap),$ws.Cells($last,$wOcc)).NumberFormat="0"
  $ws.Range($ws.Cells($first,$wMoveIn),$ws.Cells($last,$wMoveIn)).NumberFormat="m/d/yyyy"
  $ws.Range($ws.Cells($first,$wMkt),$ws.Cells($last,$wSec)).NumberFormat='$#,##0_);($#,##0)'

  # ============================================================================
  # BLUE TABLE - per-unit-type analytic summary (15 cols, matches WCH CI:CW)
  # cols BN(66)..CB(80); AL group then Total AL, blank, MC group then Total MC.
  # ============================================================================
  $aCol=66   # BN
  # source ranges on the right block
  $rUC="`$BK`$$first`:`$BK`$$last"; $rCT="`$BA`$$first`:`$BA`$$last"
  $rCAP="`$BD`$$first`:`$BD`$$last"; $rOCC="`$BE`$$first`:`$BE`$$last"
  $rSF="`$BB`$$first`:`$BB`$$last"; $rMK="`$BG`$$first`:`$BG`$$last"; $rRN="`$BH`$$first`:`$BH`$$last"; $rCR="`$BI`$$first`:`$BI`$$last"
  # group-label row (12) + column-header row (13), navy band
  $glab=@{ 6='Rents'; 9='Premium'; 10='In-Place' }   # 0-based col index -> label (as in example)
  foreach($k in $glab.Keys){ $ws.Cells(12,$aCol+$k).Value2=$glab[$k] }
  $ahdr=@('Unit Code','# Apts','# Units','Occ Units','Occ %','SqFt','Market (All)','Market (Occ)','In-Place (Occ)','(Discount)','Care','NMI Rent','Care Adj','Net NMI Rent','NMI vs In Place')
  for($c=0;$c -lt $ahdr.Count;$c++){ $ws.Cells(13,$aCol+$c).Value2=$ahdr[$c] }
  $aband=$ws.Range($ws.Cells(12,$aCol),$ws.Cells(13,$aCol+14))
  $aband.Interior.Color=$cNavy; $aband.Font.Color=$cWhite; $aband.Font.Bold=$true; $aband.Font.Name='Calibri'; $aband.HorizontalAlignment=-4108
  $ws.Cells(13,$aCol).HorizontalAlignment=-4131   # Unit Code header left
  # analytic column widths (from example CI:CW)
  $aw=@(23.1,8.0,6.4,8.6,7.9,6.3,11.0,11.7,12.6,9.4,6.3,8.7,7.7,12.6,13.9)
  for($c=0;$c -lt $aw.Count;$c++){ $ws.Columns.Item($aCol+$c).ColumnWidth=$aw[$c] }

  # build the row plan: AL codes, Total AL, blank, MC codes, Total MC
  $alK=@($codeKeys | Where-Object { "$_".StartsWith('AL') })
  $mcK=@($codeKeys | Where-Object { "$_".StartsWith('MC') })
  $plan=New-Object System.Collections.ArrayList
  foreach($k in $alK){ [void]$plan.Add(@{t='code';key=$k}) }
  [void]$plan.Add(@{t='total';key='Total AL';ct='AL'})
  [void]$plan.Add(@{t='blank'})
  foreach($k in $mcK){ [void]$plan.Add(@{t='code';key=$k}) }
  [void]$plan.Add(@{t='total';key='Total MC';ct='MC'})

  $sr0=14
  for($i=0;$i -lt $plan.Count;$i++){
    $rr=$sr0+$i; $row=$plan[$i]
    if($row.t -eq 'blank'){ continue }
    $ws.Cells($rr,$aCol).Value2=$row.key
    if($row.t -eq 'code'){
      $kc="`$BN`$$rr"
      $ws.Cells($rr,$aCol+1).Formula="=BP${rr}"
      $ws.Cells($rr,$aCol+2).Formula="=COUNTIFS($rUC,$kc,$rCAP,1)"
      $ws.Cells($rr,$aCol+3).Formula="=COUNTIFS($rUC,$kc,$rCAP,1,$rOCC,1)"
      $ws.Cells($rr,$aCol+5).Formula="=IFERROR(AVERAGEIFS($rSF,$rUC,$kc,$rCAP,1),"""")"
      $ws.Cells($rr,$aCol+6).Formula="=IFERROR(AVERAGEIFS($rMK,$rUC,$kc,$rCAP,1),"""")"
      $ws.Cells($rr,$aCol+7).Formula="=IFERROR(AVERAGEIFS($rMK,$rUC,$kc,$rCAP,1,$rOCC,1),"""")"
      $ws.Cells($rr,$aCol+8).Formula="=IFERROR(AVERAGEIFS($rRN,$rUC,$kc,$rCAP,1,$rOCC,1),"""")"
      $ws.Cells($rr,$aCol+10).Formula="=IFERROR(AVERAGEIFS($rCR,$rUC,$kc,$rCAP,1,$rOCC,1),"""")"
    } else {  # total
      $q="""$($row.ct)"""
      $ws.Cells($rr,$aCol+1).Formula="=BP${rr}"
      $ws.Cells($rr,$aCol+2).Formula="=COUNTIFS($rCT,$q,$rCAP,1)"
      $ws.Cells($rr,$aCol+3).Formula="=COUNTIFS($rCT,$q,$rCAP,1,$rOCC,1)"
      $ws.Cells($rr,$aCol+5).Formula="=IFERROR(AVERAGEIFS($rSF,$rCT,$q,$rCAP,1),"""")"
      $ws.Cells($rr,$aCol+6).Formula="=IFERROR(AVERAGEIFS($rMK,$rCT,$q,$rCAP,1),"""")"
      $ws.Cells($rr,$aCol+7).Formula="=IFERROR(AVERAGEIFS($rMK,$rCT,$q,$rCAP,1,$rOCC,1),"""")"
      $ws.Cells($rr,$aCol+8).Formula="=IFERROR(AVERAGEIFS($rRN,$rCT,$q,$rCAP,1,$rOCC,1),"""")"
      $ws.Cells($rr,$aCol+10).Formula="=IFERROR(AVERAGEIFS($rCR,$rCT,$q,$rCAP,1,$rOCC,1),"""")"
      $tr=$ws.Range($ws.Cells($rr,$aCol),$ws.Cells($rr,$aCol+14)); $tr.Font.Bold=$true
      $bt=$tr.Borders.Item(8); $bt.LineStyle=1; $bt.Weight=2     # top border
      $bb=$tr.Borders.Item(9); $bb.LineStyle=1; $bb.Weight=2     # bottom border
    }
    # derived cols (both code & total rows): Occ%, (Discount), Net NMI, NMI vs In-Place
    $ws.Cells($rr,$aCol+4).Formula="=IFERROR(BQ${rr}/BP${rr},0)"
    $ws.Cells($rr,$aCol+9).Formula="=IFERROR(BV${rr}/BU${rr}-1,0)"
    # Net NMI / NMI-vs-In-Place stay blank until the underwriter enters NMI Rent (NMI Rent/Care Adj left blank)
    $ws.Cells($rr,$aCol+13).Formula="=IF(BY${rr}="""","""",BY${rr}-BZ${rr})"
    $ws.Cells($rr,$aCol+14).Formula="=IF(OR(CA${rr}="""",BV${rr}=""""),"""",CA${rr}/BV${rr}-1)"
  }
  $srEnd=$sr0+$plan.Count-1
  # analytic number formats
  foreach($k in 1,2,3,5,6,7,8,10,11,12,13){ $ws.Range($ws.Cells($sr0,$aCol+$k),$ws.Cells($srEnd,$aCol+$k)).NumberFormat=$acct }
  foreach($k in 4,9,14){ $ws.Range($ws.Cells($sr0,$aCol+$k),$ws.Cells($srEnd,$aCol+$k)).NumberFormat=$acctPct }

  # reconciliation summary (below property total)
  $rs= if($propRow){$propRow+3}else{$last+4}
  $ws.Cells($rs,$wOrder).Value2="RECONCILIATION (sum of own leaves vs operator section totals)"
  $ws.Cells($rs+1,$wOrder).Value2="Units (rows)";       $ws.Cells($rs+1,$wSqFt).Value2=[double]$order
  $ws.Cells($rs+2,$wOrder).Value2="Occupied";           $ws.Cells($rs+2,$wSqFt).Formula="=SUM(BE${first}:BE${last})"
  $ws.Cells($rs+3,$wOrder).Value2="Capacity (beds)";    $ws.Cells($rs+3,$wSqFt).Formula="=SUM(BD${first}:BD${last})"
  $ws.Cells($rs,$wCare).Value2="Built"; $ws.Cells($rs,$wSqFt).Value2="op total"; $ws.Cells($rs,$wRType).Value2="diff"
  $ws.Cells($rs+4,$wOrder).Value2="Sum Market";         $ws.Cells($rs+4,$wCare).Formula="=SUM(BG${first}:BG${last})"; $ws.Cells($rs+4,$wSqFt).Value2=[double]$rrMkt;  $ws.Cells($rs+4,$wRType).Formula="=BA$($rs+4)-BB$($rs+4)"
  $ws.Cells($rs+5,$wOrder).Value2="Sum In-Place Rent";  $ws.Cells($rs+5,$wCare).Formula="=SUM(BH${first}:BH${last})"; $ws.Cells($rs+5,$wSqFt).Value2=[double]$rrRent; $ws.Cells($rs+5,$wRType).Formula="=BA$($rs+5)-BB$($rs+5)"
  $ws.Cells($rs+6,$wOrder).Value2="Sum In-Place Care";  $ws.Cells($rs+6,$wCare).Formula="=SUM(BI${first}:BI${last})"; $ws.Cells($rs+6,$wSqFt).Value2=[double]$rrCare; $ws.Cells($rs+6,$wRType).Formula="=BA$($rs+6)-BB$($rs+6)"
  $ws.Cells($rs+7,$wOrder).Value2="Sum 2nd-Occ Fee";    $ws.Cells($rs+7,$wCare).Formula="=SUM(BJ${first}:BJ${last})"; $ws.Cells($rs+7,$wSqFt).Value2=[double]$rrSec;  $ws.Cells($rs+7,$wRType).Formula="=BA$($rs+7)-BB$($rs+7)"
  $ws.Range($ws.Cells($rs+4,$wCare),$ws.Cells($rs+7,$wRType)).NumberFormat="#,##0.00"

  $wb.Windows.Item(1).DisplayGridlines=$false
  $xl.CalculateFull()
  $wb.Saved=$false
  $saveOk=$false
  for($k=0;$k -lt 6 -and -not $saveOk;$k++){ try{ $wb.Save(); $saveOk=$true }catch{ Start-Sleep -Milliseconds 600 } }

  $res=[pscustomobject]@{
    File=$out; Units=$order;
    Occ=[double]$ws.Cells($rs+2,$wSqFt).Value2; Beds=[double]$ws.Cells($rs+3,$wSqFt).Value2;
    Mkt=[double]$ws.Cells($rs+4,$wCare).Value2; MktRR=$rrMkt;
    Rent=[double]$ws.Cells($rs+5,$wCare).Value2; RentRR=$rrRent;
    Care=[double]$ws.Cells($rs+6,$wCare).Value2; CareRR=$rrCare;
    Secf=[double]$ws.Cells($rs+7,$wCare).Value2; SecRR=$rrSec;
    Flags=$flags
  }
  # Formula-error scan. The old `$x -is [string] -and $x.StartsWith("#")` test over Value2 was BROKEN
  # BOTH WAYS (2026-07-14): Excel marshals error cells as Int32 codes (=1/0 -> -2146826281), never
  # strings, so real errors were invisible; and it false-fired on headers like "#Apts". Ask Excel.
  # SpecialCells THROWS when nothing matches -- that is the clean path.
  $errs=0
  $errCells=$null
  try { $errCells=$ws.UsedRange.SpecialCells(-4123,16) } catch { $errCells=$null }   # xlFormulas, xlErrors
  if($errCells -ne $null){ $errs=$errCells.Count }
  $res|Add-Member NoteProperty Errors $errs
  $res|Add-Member NoteProperty SaveOk $saveOk
  Retry { $wb.Close($true) }; $wb=$null
  Start-Sleep -Milliseconds 500
  return $res
  } finally { if($wb){ try{ $wb.Close($false) }catch{} } }   # any throw above (incl. the ABORT guard) still releases the workbook
}

$results=@()
try {
  foreach($f in "O&O Community 1 Rent Roll_v1.xls","O&O Community 2 Rent Roll_v0.xls"){
    if(Test-Path (Join-Path $dir $f)){ $results += (Build-RR $f) } else { "  (skip $f - not present)" }
  }
} finally {
  # Always quit — a Build-RR throw (Stop) otherwise skips Quit() and orphans EXCEL holding the file lock.
  if($xl){ try{ Retry { $xl.Quit() } }catch{}; [System.Runtime.InteropServices.Marshal]::ReleaseComObject($xl)|Out-Null }
  Stop-TrackedExcel   # guarantees the tracked Excel PID is gone even if Quit() was skipped by a throw
}

foreach($r in $results){
  "==== $($r.File) ===="
  "  Units={0}  Occupied={1}  Beds={2}  FormulaErrors={3}  SaveOk={4}" -f $r.Units,$r.Occ,$r.Beds,$r.Errors,$r.SaveOk
  "  Market  : built={0:N2}  op={1:N2}  diff={2:N2}" -f $r.Mkt,$r.MktRR,($r.Mkt-$r.MktRR)
  "  Rent    : built={0:N2}  op={1:N2}  diff={2:N2}" -f $r.Rent,$r.RentRR,($r.Rent-$r.RentRR)
  "  Care    : built={0:N2}  op={1:N2}  diff={2:N2}" -f $r.Care,$r.CareRR,($r.Care-$r.CareRR)
  "  2nd-Occ : built={0:N2}  op={1:N2}  diff={2:N2}" -f $r.Secf,$r.SecRR,($r.Secf-$r.SecRR)
  "  Flags   : Sline={0} vacant={1} movein={2} moveout={3} inactive={4} secfee={5}" -f $r.Flags.sline,$r.Flags.vacant,$r.Flags.movein,$r.Flags.moveout,$r.Flags.inactive,$r.Flags.secfee
}
