# ============================================================================================
# Build-AsterRidge-HF.ps1 - HF-format stage for Aster Ridge (Demo), a FULLY SYNTHETIC deal.
#
# DEAL-SPECIFIC - do NOT run against another community. Aster Ridge's own account codes
# (4010-4510 revenue, 6010-7990 expense) and annual check constants (5,041,509 / 3,426,677 /
# 1,614,832) are hardcoded below.
#
# Input: single "Income Statement" tab (raw T12, Jul-2025..Jun-2026), renamed to "AR Unformatted HF".
# Output: new first tab "AR Historical Financials" built per hf-formatting skill / master-format.md.
# ============================================================================================
$ErrorActionPreference = "Stop"
$repoRoot = (Resolve-Path (Join-Path $PSScriptRoot "..\..")).Path   # <repo>\Investments\scripts -> <repo>
. (Join-Path $repoRoot "Investments\lib\HF-Build-Lib.ps1")
# NOTE: no Clear-OrphanExcel call here -- multi-agent .H run, orchestrator already ran it once.

$path = Join-Path $repoRoot "Investments\Data\Transactions\Aster Ridge (Demo)\Aster Ridge.H Model_v1.xlsx"

$xl = $null; $wb = $null
try {
$xl = New-ExcelTracked
for($i=1;$i -le 5 -and $wb -eq $null;$i++){ try{ $wb=$xl.Workbooks.Open($path) }catch{ Start-Sleep 2 } }
if($wb -eq $null){ throw "Could not open $path" }

$srcWs = $wb.Worksheets.Item(1)
$srcWs.Name = "AR Unformatted HF"
$v = $srcWs.UsedRange.Value2
$rows = $v.GetLength(0)   # 51
$cols = $v.GetLength(1)   # 15 (A..O)

# --- month dates: header row 5, cols C:N (3..14) ---
$dates = @()
for($c=3;$c -le 14;$c++){ $s=[double]$v[5,$c]; $dates += [datetime]::FromOADate($s) }

# --- classify + bucket leaf lines (col1=Acct code, col2=Description, col3..14=12 months, col15=Total) ---
# Section headers (REVENUE/OPERATING EXPENSES) and operator subtotals (TOTAL REVENUE / TOTAL
# OPERATING EXPENSES / NET OPERATING INCOME) all have no numeric code in col1 -- skip via that test.
$order = @("Rent","Care","Other","Culinary","Activities","Wellness","Maintenance","Utilities","Housekeeping","Marketing","Administration","PayrollTaxBen","NonDept")
$bucket = @{}; foreach($g in $order){ $bucket[$g] = New-Object System.Collections.ArrayList }

function GroupOf([int]$code){
  switch($code){
    4010 { return "Rent" }; 4020 { return "Rent" }; 4030 { return "Rent" }
    4110 { return "Care" }; 4120 { return "Care" }
    4210 { return "Other" }; 4310 { return "Other" }; 4410 { return "Other" }; 4510 { return "Other" }
    6010 { return "Wellness" }; 6020 { return "Wellness" }; 6030 { return "Wellness" }; 7210 { return "Wellness" }
    6110 { return "Culinary" }; 6120 { return "Culinary" }; 7010 { return "Culinary" }; 7020 { return "Culinary" }
    6210 { return "Housekeeping" }; 7110 { return "Housekeeping" }
    6310 { return "Maintenance" }; 7610 { return "Maintenance" }
    6410 { return "Activities" }
    6510 { return "Administration" }; 6610 { return "Administration" }; 6620 { return "Administration" }; 7710 { return "Administration" }; 7990 { return "Administration" }
    6630 { return "Marketing" }; 7410 { return "Marketing" }
    6710 { return "PayrollTaxBen" }; 6720 { return "PayrollTaxBen" }; 6730 { return "PayrollTaxBen" }
    7510 { return "Utilities" }; 7520 { return "Utilities" }; 7530 { return "Utilities" }; 7540 { return "Utilities" }
    7810 { return "NonDept" }; 7820 { return "NonDept" }; 7910 { return "NonDept" }
    default { return $null }
  }
}

for($r=1;$r -le $rows;$r++){
  $codeRaw = $v[$r,1]
  if($codeRaw -eq $null){ continue }
  $code = 0
  if(-not [int]::TryParse([string]$codeRaw, [ref]$code)){ continue }
  $g = GroupOf $code
  if($g -eq $null){ continue }
  $label = ([string]$v[$r,2]).Trim()
  $vals = New-Object 'double[]' 12
  for($i=0;$i -lt 12;$i++){ $cell=$v[$r,(3+$i)]; if($cell -ne $null){ $vals[$i]=[double]$cell } else { $vals[$i]=0 } }
  [void]$bucket[$g].Add([pscustomobject]@{ Code=[string]$code; Label=$label; Vals=$vals })
}

# --- create Formatted HF as first tab ---
$ws = $wb.Worksheets.Add($srcWs)
$ws.Name = "AR Historical Financials"
$script:ws = $ws; $script:wb = $wb
$xl.Calculation = -4135   # xlCalculationManual

HF-Header "Aster Ridge Senior Living (DEMO)" $dates
$ws.Cells(5,3).Value2 = "Book = Accrual ; Tree = wlf_cf"
HF-RevBanner $dates

$script:r = 7   # blank row; WriteGroup enters on a blank, exits on a blank

function WriteGroup($g,$totalLabel,$marker,$comment){
  $members=@()
  foreach($item in $bucket[$g]){ $row=HF-Line $item.Code $item.Label $item.Vals; $members+=$row }
  $script:r++                                       # blank before total
  $tot=HF-Total $totalLabel $members $marker
  if($comment -ne $null -and $comment -ne ""){
    $c=$ws.Cells($tot,3)
    $cm=$c.AddComment($comment)
    $cm.Shape.TextFrame.AutoSize=$true
  }
  $script:r++                                       # blank after total
  return $tot
}

$rentTot  = WriteGroup "Rent"  "TOTAL RENT REVENUE" "x" $null
$careTot  = WriteGroup "Care"  "TOTAL CARE REVENUE" "x" $null
$otherTot = WriteGroup "Other" "TOTAL OTHER REVENUE" "x" $null
$revTot   = HF-Total "TOTAL REVENUE" @($rentTot,$careTot,$otherTot) "x"

# EXPENSES banner
$script:r += 2
$er=$script:r
$ws.Cells($er,1).Value2="x"; $ws.Cells($er,3).Value2="EXPENSES"
HF-Band $er $true
$script:r = $er + 1

$cul = WriteGroup "Culinary"      "TOTAL CULINARY EXPENSES" "x" $null
$act = WriteGroup "Activities"    "TOTAL ACTIVITIES EXPENSES" "x" $null
$wel = WriteGroup "Wellness"      "TOTAL WELLNESS - NURSING EXPENSES" "x" "FLAG: No Master department home. Caregiver Wages, Med Tech Wages, Wellness Director, and Care Supplies are pooled across AL and MC by the operator (not split by care type) -- kept as one block, positioned among the care departments. Candidate: split AL/Alz via a census-based allocation in the mapping step."
$mnt = WriteGroup "Maintenance"   "TOTAL MAINTENANCE EXPENSES" "x" $null
$utl = WriteGroup "Utilities"     "TOTAL UTILITIES EXPENSES" "x" $null
$hsk = WriteGroup "Housekeeping"  "TOTAL HOUSEKEEPING EXPENSE" "x" $null
$mkt = WriteGroup "Marketing"     "TOTAL MARKETING EXPENSES" "x" $null
$adm = WriteGroup "Administration" "TOTAL ADMINISTRATION EXPENSES" "x" $null
# flag Concierge Wages row (6510) individually -- bucketed to Administration, no explicit operator dept header
foreach($item in $bucket["Administration"]){
  if($item.Code -eq "6510"){
    for($rr=$er+1;$rr -le $adm;$rr++){
      if([string]$ws.Cells($rr,19).Value2 -eq "6510"){
        $cm=$ws.Cells($rr,3).AddComment("FLAG: Operator's expense list is flat (no Administrative sub-header to confirm this against) -- bucketed here by function (front-desk / resident-services). Review in PnL mapping.")
        $cm.Shape.TextFrame.AutoSize=$true
      }
    }
  }
}
$ptb = WriteGroup "PayrollTaxBen" "TOTAL PAYROLL TAXES  and  BENEFITS EXPENSES" "x" "FLAG: No Master department home. Payroll Taxes, Employee Benefits, and Workers Comp are pooled across ALL departments (not split by function/department) -- kept as its own block rather than force-allocated into each department's payroll stub. Candidate: apply a labor-cost allocation basis in the mapping step."
$ndp = WriteGroup "NonDept"      "TOTAL NON-DEPARTMENTAL EXPENSES" "x" $null
# flag Management Fee row (7910) -- placed here per Master convention since operator list is flat (no Administrative sub-block to leave it in)
foreach($rr in (($ndp-4)..($ndp-1))){
  if([string]$ws.Cells($rr,19).Value2 -eq "7910"){
    $cm=$ws.Cells($rr,3).AddComment("FLAG: Operator's expense list is flat -- no Administrative sub-block to 'leave Management Fee in' per the usual rule. Placed in Non-Departmental per Master convention. Review in PnL mapping.")
    $cm.Shape.TextFrame.AutoSize=$true
  }
}

$opexRows = @($cul,$act,$wel,$mnt,$utl,$hsk,$mkt,$adm,$ptb,$ndp)
$opexTot  = HF-Total "TOTAL OPERATING EXPENSES" $opexRows "x"

# NET OPERATING INCOME
$script:r += 2
$noi=$script:r
$ws.Cells($noi,1).Value2="x"; $ws.Cells($noi,3).Value2="NET OPERATING INCOME"
foreach($cc in 5..17){ $L=[char](64+$cc); $ws.Cells($noi,$cc).Formula="=$L${revTot}-$L${opexTot}"; $ws.Cells($noi,$cc).NumberFormat=$ACCT }
HF-Band $noi $false

# Margin
$script:r++
$mg=$script:r
$ws.Cells($mg,3).Value2="Margin"
for($cc=5;$cc -le 16;$cc++){ $L=[char](64+$cc); $ws.Cells($mg,$cc).Formula="=IFERROR($L${noi}/$L${revTot},0)"; $ws.Cells($mg,$cc).NumberFormat="0.0%" }

HF-Finish

$ws.Activate(); $wb.Windows.Item(1).DisplayGridlines=$false
$srcWs.Activate(); $wb.Windows.Item(1).DisplayGridlines=$false
$ws.Activate()

$xl.CalculateFull(); $xl.Calculation = -4105
$wb.Save()

# ---------------- VALIDATION GATES ----------------
function OpRow($label){ for($r=1;$r -le $rows;$r++){ if("$($v[$r,2])".Trim() -eq $label){ return $r } }; return $null }
$rRev = OpRow "TOTAL REVENUE"; $rExp = OpRow "TOTAL OPERATING EXPENSES"; $rNoi = OpRow "NET OPERATING INCOME"

$revBucketM=New-Object 'double[]' 12; $expBucketM=New-Object 'double[]' 12
foreach($g in "Rent","Care","Other"){ foreach($it in $bucket[$g]){ for($i=0;$i -lt 12;$i++){ $revBucketM[$i]+=$it.Vals[$i] } } }
foreach($g in "Culinary","Activities","Wellness","Maintenance","Utilities","Housekeeping","Marketing","Administration","PayrollTaxBen","NonDept"){ foreach($it in $bucket[$g]){ for($i=0;$i -lt 12;$i++){ $expBucketM[$i]+=$it.Vals[$i] } } }

$maxRev=0.0; $maxExp=0.0; $maxNoi=0.0
for($i=0;$i -lt 12;$i++){
  $opR=[double]$v[$rRev,(3+$i)]; $opE=[double]$v[$rExp,(3+$i)]; $opN=[double]$v[$rNoi,(3+$i)]
  $dR=[math]::Abs($revBucketM[$i]-$opR); if($dR -gt $maxRev){$maxRev=$dR}
  $dE=[math]::Abs($expBucketM[$i]-$opE); if($dE -gt $maxExp){$maxExp=$dE}
  $dN=[math]::Abs(($revBucketM[$i]-$expBucketM[$i])-$opN); if($dN -gt $maxNoi){$maxNoi=$dN}
}
"GATE1 reconcile (max abs monthly diff): REV={0:N6}  EXP={1:N6}  NOI={2:N6}" -f $maxRev,$maxExp,$maxNoi

$fv=$ws.UsedRange.Value2
"REBUILT annual: REV={0:N2} (op 5,041,509)  OPEX={1:N2} (op 3,426,677)  NOI={2:N2} (op 1,614,832)" -f ([double]$fv[$revTot,17]),([double]$fv[$opexTot,17]),([double]$fv[$noi,17])

$errCells=$null
try { $errCells=$ws.UsedRange.SpecialCells(-4123,16) } catch { $errCells=$null }
if($errCells -ne $null){ "GATE2 !! FORMULA ERRORS: {0} cells at {1}" -f $errCells.Count,$errCells.Address() }
else { "GATE2 formula errors on AR Historical Financials: 0" }

"DONE rows: hfLast=$mg revTot=$revTot opexTot=$opexTot noi=$noi margin=$mg"
"Window: $($dates[0].ToString('MMM-yyyy')) - $($dates[11].ToString('MMM-yyyy'))"

$wb.Close($true); $wb=$null
}
finally {
  if($wb){ try{ $wb.Close($false) }catch{} }
  if($xl){ try{ $xl.Quit() }catch{}; [System.Runtime.InteropServices.Marshal]::ReleaseComObject($xl)|Out-Null }
  Stop-TrackedExcel
}
