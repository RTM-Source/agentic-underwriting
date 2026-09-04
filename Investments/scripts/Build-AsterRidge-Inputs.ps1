<#
.SYNOPSIS
  Generates FULLY SYNTHETIC demo input data for a fictional senior-housing community
  ("Aster Ridge Senior Living", 84 units: IL 40 / AL 30 / MC 14).

.DESCRIPTION
  Builds three fabricated source files that mimic real operator exports, for use as
  public-safe demo inputs to the underwrite pipeline:
    1. A rent-roll roster (OneSite-family raw export shape) as of 06/30/2026 -- this is
       ground truth for the community's June 2026 occupancy and in-place economics.
    2. A T12 GL income statement (Jul-2025 .. Jun-2026) whose revenue lines for June are
       derived EXACTLY from the roster, and whose leaf lines always sum to the operator
       subtotal rows to the cent.
    3. A 24-month occupancy/census report (Jul-2024 .. Jun-2026) whose June-2026 counts
       match the roster exactly.

  Every number in this script is invented for demo purposes. No real community, resident,
  or financial data is read, referenced, or reproduced anywhere in this file.

  Re-runnable: running this script again regenerates all three files in place (they are
  demo fixtures, not versioned deal deliverables).

.NOTES
  Excel COM only (no Python on this box). Uses the shared build-library helpers for safe
  scalar/array writes and tracked Excel process cleanup.
#>

param(
  [string]$DemoRoot = (Join-Path $PSScriptRoot '..\Data\Transactions\Aster Ridge (Demo)')
)

$ErrorActionPreference = 'Stop'
[System.Threading.Thread]::CurrentThread.CurrentCulture = [System.Globalization.CultureInfo]'en-US'

. (Join-Path $PSScriptRoot '..\lib\HF-Build-Lib.ps1')

# ================================================================================
# PATHS
# ================================================================================
$rrDir  = Join-Path $DemoRoot 'Rent Roll'
$hfDir  = Join-Path $DemoRoot 'Historical Financials'
New-Item -ItemType Directory -Force -Path $rrDir | Out-Null
New-Item -ItemType Directory -Force -Path $hfDir | Out-Null

$rrPathFinal      = Join-Path $rrDir '06.30.2026 Aster Ridge SeniorLivingRentRoll.xls'
$t12PathFinal     = Join-Path $hfDir '06.2026 Aster Ridge T12 Income Statement.xlsx'
$censusPathFinal  = Join-Path $hfDir 'Aster Ridge Occupancy Report Jul 2024 - Jun 2026.xlsx'

$DISCLAIMER = "SYNTHETIC DEMO DATA -- NOT A REAL COMMUNITY. All figures fabricated for demonstration/testing purposes only."

# ================================================================================
# HELPERS
# ================================================================================
function RoundInt($v){ [math]::Round([double]$v,0,[System.MidpointRounding]::AwayFromZero) }
function RoundTo5($v){ [math]::Round([double]$v/5.0,0,[System.MidpointRounding]::AwayFromZero)*5.0 }
function Wiggle([int]$m,[double]$phase){ return 0.01*[math]::Sin(($m+1)*0.9+$phase) }

function FmtMoney([double]$v){
  if([math]::Abs($v) -lt 0.0001){ return "0.00" }
  elseif($v -lt 0){ return "(" + [math]::Abs($v).ToString("N2") + ")" }
  else { return $v.ToString("N2") }
}
function ParseMoneyText($s){
  if($null -eq $s){ return 0.0 }
  $str = [string]$s
  if([string]::IsNullOrWhiteSpace($str)){ return 0.0 }
  $neg = $str.StartsWith('(')
  $clean = $str -replace '[(),]',''
  if([string]::IsNullOrWhiteSpace($clean)){ return 0.0 }
  $v = [double]$clean
  if($neg){ $v = -$v }
  return $v
}
function FmtDate($d){ if($null -eq $d){ return "" }; return $d.ToString("MM/dd/yyyy") }

function MonthLabels([int]$startYear,[int]$startMonth,[int]$n){
  $out=@(); $y=$startYear; $mo=$startMonth
  for($i=0;$i -lt $n;$i++){
    $out += (Get-Date -Year $y -Month $mo -Day 1).ToString("MMM-yy")
    $mo++; if($mo -gt 12){ $mo=1; $y++ }
  }
  return $out
}

$residentCounter = 0
function NextResidentName(){
  $script:residentCounter++
  return ("Resident {0:D3}" -f $script:residentCounter)
}

$ratioByBucket = @{ 2019=0.92; 2020=0.93; 2021=0.94; 2022=0.95; 2023=0.96; 2024=0.97; 2025=0.98; 'Q1-2026'=0.99 }
$buckets = @(2019,2020,2021,2022,2023,2024,2025,'Q1-2026')

function MoveInDate($bucket,[int]$idx){
  if($bucket -eq 'Q1-2026'){
    $mo = ($idx % 3) + 1
    $da = (($idx*17) % 27) + 1
    return Get-Date -Year 2026 -Month $mo -Day $da
  } else {
    $mo = (($idx*7) % 12) + 1
    $da = (($idx*13) % 27) + 1
    return Get-Date -Year ([int]$bucket) -Month $mo -Day $da
  }
}

# ================================================================================
# PART 1 -- ROSTER (ground truth). Fully in memory, fully fabricated.
# ================================================================================
$ilTypes = @(
  [pscustomobject]@{Type='Studio';      SqFt=420; Market=3295; Start=101; Count=8;  Vacant=@(104); Beds=1}
  [pscustomobject]@{Type='One Bedroom'; SqFt=610; Market=3895; Start=109; Count=22; Vacant=@(119); Beds=1}
  [pscustomobject]@{Type='Two Bedroom'; SqFt=850; Market=4695; Start=131; Count=10; Vacant=@(136); Beds=1}
)
$alTypes = @(
  [pscustomobject]@{Type='Studio';      SqFt=380; Market=4495; Start=201; Count=18; Vacant=@(205); Beds=1}
  [pscustomobject]@{Type='One Bedroom'; SqFt=540; Market=5295; Start=219; Count=12; Vacant=@(225); Beds=1}
)
$mcTypes = @(
  [pscustomobject]@{Type='Private';       SqFt=320; Market=7495; Start=301; Count=12; Vacant=@(308); Beds=1}
  [pscustomobject]@{Type='Semi-Private';  SqFt=400; Market=5995; Start=313; Count=2;  Vacant=@();     Beds=2}
)

$units = New-Object System.Collections.Generic.List[object]
function BuildTypeUnits($careLabel,$types){
  foreach($t in $types){
    for($i=0;$i -lt $t.Count;$i++){
      $num = $t.Start + $i
      $isVacant = $t.Vacant -contains $num
      $units.Add([pscustomobject]@{
        Care=$careLabel; UnitNum=$num; Type=$t.Type; SqFt=$t.SqFt; Market=$t.Market
        Occupied = (-not $isVacant); Beds=$t.Beds
        Residents = (New-Object System.Collections.Generic.List[object])
      }) | Out-Null
    }
  }
}
BuildTypeUnits 'IL' $ilTypes
BuildTypeUnits 'AL' $alTypes
BuildTypeUnits 'MC' $mcTypes

# couples (second-resident fee $995/mo, own charge line)
$ilCoupleUnits = @(101,103,110,122,132)   # 5 total: 2 Studio, 2 1BR, 1 2BR
$alCoupleUnits = @(202,206,221)           # 3 total: 2 Studio, 1 1BR

# 5 community-wide units with a move-in inside Apr-Jun 2026 (2 IL, 2 AL, 1 MC)
$newUnits = @{
  106 = (Get-Date -Year 2026 -Month 4 -Day 8)
  210 = (Get-Date -Year 2026 -Month 4 -Day 22)
  140 = (Get-Date -Year 2026 -Month 5 -Day 5)
  228 = (Get-Date -Year 2026 -Month 5 -Day 19)
  302 = (Get-Date -Year 2026 -Month 6 -Day 12)
}

# assign a move-in "bucket" cyclically per care type, occupied units only
$careCounters = @{IL=0; AL=0; MC=0}
foreach($u in $units){
  if(-not $u.Occupied){ continue }
  $idx = $careCounters[$u.Care]
  $u | Add-Member -NotePropertyName Bucket -NotePropertyValue $buckets[$idx % 8] -Force
  $u | Add-Member -NotePropertyName BucketIdx -NotePropertyValue $idx -Force
  $careCounters[$u.Care] = $idx + 1
}

# build residents per occupied unit
foreach($u in $units){
  if(-not $u.Occupied){ continue }
  $isNew = $newUnits.ContainsKey($u.UnitNum)
  $ratio = if($isNew){ 1.00 } else { $ratioByBucket[$u.Bucket] }
  $moveIn = if($isNew){ $newUnits[$u.UnitNum] } else { MoveInDate $u.Bucket $u.BucketIdx }
  $inPlace = RoundTo5 ($u.Market * $ratio)

  $r1 = [pscustomobject]@{Name=(NextResidentName); MoveIn=$moveIn; Market=$u.Market; InPlace=$inPlace; CareFee=0.0; SecondFee=0.0; IsSecond=$false}
  $u.Residents.Add($r1) | Out-Null

  if($u.Care -eq 'MC' -and $u.Beds -eq 2){
    # semi-private: second bed billed independently at full rate, same move-in
    $r2 = [pscustomobject]@{Name=(NextResidentName); MoveIn=$moveIn; Market=$u.Market; InPlace=$inPlace; CareFee=0.0; SecondFee=0.0; IsSecond=$false}
    $u.Residents.Add($r2) | Out-Null
  }
  elseif( ($u.Care -eq 'IL' -and $ilCoupleUnits -contains $u.UnitNum) -or ($u.Care -eq 'AL' -and $alCoupleUnits -contains $u.UnitNum) ){
    # couple: second resident pays only the $995 second-person fee
    $r2 = [pscustomobject]@{Name=(NextResidentName); MoveIn=$moveIn; Market=0.0; InPlace=0.0; CareFee=0.0; SecondFee=995.0; IsSecond=$true}
    $u.Residents.Add($r2) | Out-Null
  }
}

# AL care levels: by ascending unit#, first 12 -> L1 $595, next 10 -> L2 $1095, last 6 -> L3 $1595
$alOccUnits = $units | Where-Object { $_.Care -eq 'AL' -and $_.Occupied } | Sort-Object UnitNum
for($i=0;$i -lt $alOccUnits.Count;$i++){
  $u = $alOccUnits[$i]
  if($i -lt 12){ $fee = 595.0 } elseif($i -lt 22){ $fee = 1095.0 } else { $fee = 1595.0 }
  $u.Residents[0].CareFee = $fee
}
# MC care: flat $1,200 per resident (every MC resident, private and both semi beds)
foreach($u in ($units | Where-Object { $_.Care -eq 'MC' -and $_.Occupied })){
  foreach($res in $u.Residents){ $res.CareFee = 1200.0 }
}

# ---- roster-derived totals (June 2026 ground truth) ----
$ilUnits = $units | Where-Object { $_.Care -eq 'IL' }
$alUnits = $units | Where-Object { $_.Care -eq 'AL' }
$mcUnits = $units | Where-Object { $_.Care -eq 'MC' }

$ilOccCount = ($ilUnits | Where-Object Occupied).Count
$alOccCount = ($alUnits | Where-Object Occupied).Count
$mcOccCount = ($mcUnits | Where-Object Occupied).Count

$ilRentTotal   = 0.0; $ilSecondTotal = 0.0
foreach($u in $ilUnits){ foreach($res in $u.Residents){ $ilRentTotal += $res.InPlace; $ilSecondTotal += $res.SecondFee } }
$alRentTotal   = 0.0; $alCareTotal = 0.0; $alSecondTotal = 0.0
foreach($u in $alUnits){ foreach($res in $u.Residents){ $alRentTotal += $res.InPlace; $alCareTotal += $res.CareFee; $alSecondTotal += $res.SecondFee } }
$mcRentTotal   = 0.0; $mcCareTotal = 0.0
foreach($u in $mcUnits){ foreach($res in $u.Residents){ $mcRentTotal += $res.InPlace; $mcCareTotal += $res.CareFee } }

$avgILRent  = $ilRentTotal / $ilOccCount
$avgALRent  = $alRentTotal / $alOccCount
$avgMCRent  = $mcRentTotal / $mcOccCount
$avgALCare  = $alCareTotal / $alOccCount
$avgMCCare  = $mcCareTotal / $mcOccCount
$avg2ndFeePerILAL = ($ilSecondTotal + $alSecondTotal) / ($ilOccCount + $alOccCount)

Write-Host "=== ROSTER TOTALS (June 2026 ground truth) ==="
Write-Host ("IL occ={0} rent={1:N2} 2ndFee={2:N2}" -f $ilOccCount,$ilRentTotal,$ilSecondTotal)
Write-Host ("AL occ={0} rent={1:N2} care={2:N2} 2ndFee={3:N2}" -f $alOccCount,$alRentTotal,$alCareTotal,$alSecondTotal)
Write-Host ("MC occ={0} rent={1:N2} care={2:N2}" -f $mcOccCount,$mcRentTotal,$mcCareTotal)
$rosterGrandTotal = $ilRentTotal+$ilSecondTotal+$alRentTotal+$alCareTotal+$alSecondTotal+$mcRentTotal+$mcCareTotal
Write-Host ("ROSTER GRAND TOTAL (rent+care+fees) = {0:N2}" -f $rosterGrandTotal)

# ================================================================================
# PART 2 -- CENSUS occupancy schedule (24 months, Jul-2024 .. Jun-2026)
# ================================================================================
$occIL_1226 = @(34,34,35,35,36,36,36,35,36,36,37,37)   # Jul25..Jun26
$occAL_1226 = @(26,27,27,26,27,28,28,27,27,28,28,28)
$occMC_1226 = @(11,12,12,12,13,12,12,13,13,12,13,13)

function Ramp([double]$from,[double]$to,[int]$n){
  $out=@()
  for($i=0;$i -lt $n;$i++){ $out += (RoundInt ($from + ($to-$from)*$i/($n-1))) }
  return $out
}
$occIL_2425 = Ramp 32 34 12
$occAL_2425 = Ramp 25 26 12
$occMC_2425 = Ramp 10 11 12

# sanity: June 2026 census must equal the roster's occupied-unit counts
if($occIL_1226[11] -ne $ilOccCount -or $occAL_1226[11] -ne $alOccCount -or $occMC_1226[11] -ne $mcOccCount){
  throw "Census Jun-2026 counts do not match roster occupied counts -- fix the spec constants."
}

$occIL_full = @($occIL_2425) + @($occIL_1226)   # 24
$occAL_full = @($occAL_2425) + @($occAL_1226)
$occMC_full = @($occMC_2425) + @($occMC_1226)
$censusMonths = MonthLabels 2024 7 24
$t12Months    = MonthLabels 2025 7 12

# ================================================================================
# PART 3 -- T12 REVENUE lines (Jul-2025 .. Jun-2026), June = exact roster override
# ================================================================================
$rentIL_arr=@(); $rentAL_arr=@(); $rentMC_arr=@(); $careAL_arr=@(); $careMC_arr=@(); $secondFee_arr=@()
for($m=0;$m -lt 12;$m++){
  if($m -eq 11){
    $rentIL_arr    += $ilRentTotal
    $rentAL_arr    += $alRentTotal
    $rentMC_arr    += $mcRentTotal
    $careAL_arr    += $alCareTotal
    $careMC_arr    += $mcCareTotal
    $secondFee_arr += ($ilSecondTotal + $alSecondTotal)
  } else {
    $rentIL_arr    += (RoundInt ($occIL_1226[$m]*$avgILRent*(1+(Wiggle $m 0.0))))
    $rentAL_arr    += (RoundInt ($occAL_1226[$m]*$avgALRent*(1+(Wiggle $m 0.6))))
    $rentMC_arr    += (RoundInt ($occMC_1226[$m]*$avgMCRent*(1+(Wiggle $m 1.2))))
    $careAL_arr    += (RoundInt ($occAL_1226[$m]*$avgALCare*(1+(Wiggle $m 1.8))))
    $careMC_arr    += (RoundInt ($occMC_1226[$m]*$avgMCCare*(1+(Wiggle $m 2.4))))
    $secondFee_arr += (RoundInt (($occIL_1226[$m]+$occAL_1226[$m])*$avg2ndFeePerILAL*(1+(Wiggle $m 3.0))))
  }
}

# Community Fees (spiky, tied loosely to net move-ins), Other Resident Income, Guest Meals
$occTotal = for($m=0;$m -lt 12;$m++){ $occIL_1226[$m]+$occAL_1226[$m]+$occMC_1226[$m] }
$prevOccTotal = $occIL_2425[11]+$occAL_2425[11]+$occMC_2425[11]
$communityFee_arr=@(); $otherResident_arr=@(); $guestMeals_arr=@()
for($m=0;$m -lt 12;$m++){
  $prev = if($m -eq 0){ $prevOccTotal } else { $occTotal[$m-1] }
  $posDelta = [math]::Max(0,$occTotal[$m]-$prev)
  $cf = 3200 + $posDelta*3200 + 900*[math]::Sin(($m+1)*1.4)
  $communityFee_arr += [math]::Max(3000,[math]::Min(14000,(RoundInt $cf)))
  $orv = RoundInt (3200 + 380*[math]::Sin(($m+1)*1.3+0.4))
  $otherResident_arr += [math]::Max(2500,[math]::Min(4000,$orv))
  $gm = RoundInt (2300 + 420*[math]::Sin(($m+1)*1.7+0.9))
  $guestMeals_arr += [math]::Max(1800,[math]::Min(2900,$gm))
}

$revLines = @(
  [pscustomobject]@{Code='4010';Label='Rent Income - IL';         Arr=$rentIL_arr}
  [pscustomobject]@{Code='4020';Label='Rent Income - AL';         Arr=$rentAL_arr}
  [pscustomobject]@{Code='4030';Label='Rent Income - MC';         Arr=$rentMC_arr}
  [pscustomobject]@{Code='4110';Label='Care Level Income - AL';   Arr=$careAL_arr}
  [pscustomobject]@{Code='4120';Label='Care Level Income - MC';   Arr=$careMC_arr}
  [pscustomobject]@{Code='4210';Label='Second Person Fees';       Arr=$secondFee_arr}
  [pscustomobject]@{Code='4310';Label='Community Fees';           Arr=$communityFee_arr}
  [pscustomobject]@{Code='4410';Label='Other Resident Income';    Arr=$otherResident_arr}
  [pscustomobject]@{Code='4510';Label='Guest Meals & Ancillary';  Arr=$guestMeals_arr}
)

# ================================================================================
# PART 4 -- T12 EXPENSE lines
# ================================================================================
$wageDefs = @(
  @{Code='6010';Label='Caregiver Wages';           Base=45080;Phase=0.3}
  @{Code='6020';Label='Med Tech Wages';             Base=17020;Phase=0.9}
  @{Code='6030';Label='Wellness Director';          Base=5520; Phase=1.5}
  @{Code='6110';Label='Dining Wages';               Base=24840;Phase=2.1}
  @{Code='6120';Label='Dining Director';            Base=4784; Phase=2.7}
  @{Code='6210';Label='Housekeeping Wages';         Base=12420;Phase=3.3}
  @{Code='6310';Label='Maintenance Wages';          Base=9016; Phase=3.9}
  @{Code='6410';Label='Activities Wages';           Base=6256; Phase=4.5}
  @{Code='6510';Label='Concierge Wages';            Base=7544; Phase=5.1}
  @{Code='6610';Label='Business Office Manager';    Base=5152; Phase=5.7}
  @{Code='6620';Label='Executive Director';         Base=8464; Phase=6.3}
  @{Code='6630';Label='Sales & Marketing Wages';    Base=6992; Phase=6.9}
)
$wageArrs = @{}
foreach($wd in $wageDefs){
  $wageArrs[$wd.Code] = for($m=0;$m -lt 12;$m++){ RoundInt ($wd.Base*(1+0.012*[math]::Sin(($m+1)*1.1+$wd.Phase))) }
}
$payrollTax_arr = for($m=0;$m -lt 12;$m++){
  $wsum=0.0; foreach($wd in $wageDefs){ $wsum += $wageArrs[$wd.Code][$m] }
  RoundInt (0.095*$wsum)
}
$benefits_arr = for($m=0;$m -lt 12;$m++){ RoundInt (11224*(1+0.01*[math]::Sin(($m+1)*0.7))) }
$wc_arr       = for($m=0;$m -lt 12;$m++){ RoundInt (3036*(1+0.01*[math]::Sin(($m+1)*0.5))) }

$rawFood_arr        = for($m=0;$m -lt 12;$m++){ RoundInt (13984*(1+0.02*[math]::Sin(($m+1)*0.8+0.2))) }
$diningSupplies_arr = for($m=0;$m -lt 12;$m++){ RoundInt (2208*(1+0.02*[math]::Sin(($m+1)*1.0+0.5))) }
$hkSupplies_arr     = for($m=0;$m -lt 12;$m++){ RoundInt (2576*(1+0.02*[math]::Sin(($m+1)*1.2+0.9))) }
$careSupplies_arr   = for($m=0;$m -lt 12;$m++){ RoundInt (3312*(1+0.02*[math]::Sin(($m+1)*1.4+1.3))) }
$marketing_arr      = for($m=0;$m -lt 12;$m++){ RoundInt (5336*(1+0.05*[math]::Sin(($m+1)*0.6+1.7))) }

$elecFactor  = @(1.20,1.25,1.05,0.90,0.90,1.05,1.15,1.10,0.95,0.85,0.90,1.10)  # Jul..Jun
$gasFactor   = @(0.45,0.45,0.55,0.85,1.25,1.55,1.60,1.50,1.20,0.85,0.60,0.50)
$waterFactor = @(1.15,1.15,1.05,1.00,0.95,0.95,0.95,0.95,1.00,1.00,1.05,1.15)
$electricity_arr = for($m=0;$m -lt 12;$m++){ RoundInt (9660*$elecFactor[$m]) }
$gas_arr         = for($m=0;$m -lt 12;$m++){ RoundInt (3496*$gasFactor[$m]) }
$water_arr       = for($m=0;$m -lt 12;$m++){ RoundInt (3588*$waterFactor[$m]) }
$trash_arr       = for($m=0;$m -lt 12;$m++){ 1564 }
$rm_arr          = for($m=0;$m -lt 12;$m++){ RoundInt (6624*(1+0.04*[math]::Sin(($m+1)*0.9+2.1))) }
$it_arr          = for($m=0;$m -lt 12;$m++){ 2760 }
$insurance_arr   = for($m=0;$m -lt 12;$m++){ 10120 }
$retax_arr       = for($m=0;$m -lt 12;$m++){ 12880 }
$otherGA_arr     = for($m=0;$m -lt 12;$m++){ RoundInt (4232*(1+0.03*[math]::Sin(($m+1)*1.6+2.5))) }

# ================================================================================
# PART 5 -- assemble T12 sheet rows (values only; leaf sums drive every subtotal)
# ================================================================================
function NewT12Row(){ New-Object object[] 15 }   # A..O

$t12Rows = New-Object System.Collections.Generic.List[object]
1..4 | ForEach-Object { $t12Rows.Add((NewT12Row)) | Out-Null }
$t12Rows[0][0] = "Aster Ridge Senior Living (DEMO)"
$t12Rows[1][0] = "T12 Income Statement -- Jul 2025 through Jun 2026"
$t12Rows[2][0] = $DISCLAIMER
# row4 blank

$hdr = NewT12Row
$hdr[0]="Acct"; $hdr[1]="Description"
for($m=0;$m -lt 12;$m++){ $hdr[2+$m] = $t12Months[$m] }
$hdr[14]="Total"
$t12Rows.Add($hdr) | Out-Null   # row5

$revBand = NewT12Row; $revBand[1]="REVENUE"; $t12Rows.Add($revBand) | Out-Null   # row6

$monthlyRevTotal = New-Object double[] 12
$revLeafRowNums = New-Object System.Collections.Generic.List[int]
foreach($ln in $revLines){
  $r = NewT12Row; $r[0]=$ln.Code; $r[1]=$ln.Label
  $tot=0.0
  for($m=0;$m -lt 12;$m++){ $v=[double]$ln.Arr[$m]; $r[2+$m]=$v; $tot+=$v; $monthlyRevTotal[$m]+=$v }
  $r[14]=$tot
  $t12Rows.Add($r) | Out-Null
  $revLeafRowNums.Add($t12Rows.Count) | Out-Null
}

$rTotRev = NewT12Row; $rTotRev[1]="TOTAL REVENUE"
$grandTotRev=0.0
for($m=0;$m -lt 12;$m++){ $rTotRev[2+$m]=$monthlyRevTotal[$m]; $grandTotRev+=$monthlyRevTotal[$m] }
$rTotRev[14]=$grandTotRev
$t12Rows.Add($rTotRev) | Out-Null
$totalRevenueRowNum = $t12Rows.Count

$t12Rows.Add((NewT12Row)) | Out-Null    # blank
$expBand = NewT12Row; $expBand[1]="OPERATING EXPENSES"; $t12Rows.Add($expBand) | Out-Null

$mgmtFee_arr = for($m=0;$m -lt 12;$m++){ RoundInt (0.05*$monthlyRevTotal[$m]) }

$expLines = @()
foreach($wd in $wageDefs){ $expLines += [pscustomobject]@{Code=$wd.Code;Label=$wd.Label;Arr=$wageArrs[$wd.Code]} }
$expLines += [pscustomobject]@{Code='6710';Label='Payroll Taxes';           Arr=$payrollTax_arr}
$expLines += [pscustomobject]@{Code='6720';Label='Employee Benefits';       Arr=$benefits_arr}
$expLines += [pscustomobject]@{Code='6730';Label='Workers Comp';            Arr=$wc_arr}
$expLines += [pscustomobject]@{Code='7010';Label='Raw Food';                Arr=$rawFood_arr}
$expLines += [pscustomobject]@{Code='7020';Label='Dining Supplies';         Arr=$diningSupplies_arr}
$expLines += [pscustomobject]@{Code='7110';Label='Housekeeping Supplies';   Arr=$hkSupplies_arr}
$expLines += [pscustomobject]@{Code='7210';Label='Care Supplies';           Arr=$careSupplies_arr}
$expLines += [pscustomobject]@{Code='7410';Label='Marketing & Advertising'; Arr=$marketing_arr}
$expLines += [pscustomobject]@{Code='7510';Label='Electricity';             Arr=$electricity_arr}
$expLines += [pscustomobject]@{Code='7520';Label='Gas';                     Arr=$gas_arr}
$expLines += [pscustomobject]@{Code='7530';Label='Water/Sewer';             Arr=$water_arr}
$expLines += [pscustomobject]@{Code='7540';Label='Trash';                   Arr=$trash_arr}
$expLines += [pscustomobject]@{Code='7610';Label='Repairs & Maintenance';   Arr=$rm_arr}
$expLines += [pscustomobject]@{Code='7710';Label='IT & Telephone';          Arr=$it_arr}
$expLines += [pscustomobject]@{Code='7810';Label='Insurance';               Arr=$insurance_arr}
$expLines += [pscustomobject]@{Code='7820';Label='Real Estate Taxes';       Arr=$retax_arr}
$expLines += [pscustomobject]@{Code='7910';Label='Management Fee';          Arr=$mgmtFee_arr}
$expLines += [pscustomobject]@{Code='7990';Label='Other G&A';               Arr=$otherGA_arr}

$monthlyExpTotal = New-Object double[] 12
$expLeafRowNums = New-Object System.Collections.Generic.List[int]
foreach($ln in $expLines){
  $r = NewT12Row; $r[0]=$ln.Code; $r[1]=$ln.Label
  $tot=0.0
  for($m=0;$m -lt 12;$m++){ $v=[double]$ln.Arr[$m]; $r[2+$m]=$v; $tot+=$v; $monthlyExpTotal[$m]+=$v }
  $r[14]=$tot
  $t12Rows.Add($r) | Out-Null
  $expLeafRowNums.Add($t12Rows.Count) | Out-Null
}

$rTotExp = NewT12Row; $rTotExp[1]="TOTAL OPERATING EXPENSES"
$grandTotExp=0.0
for($m=0;$m -lt 12;$m++){ $rTotExp[2+$m]=$monthlyExpTotal[$m]; $grandTotExp+=$monthlyExpTotal[$m] }
$rTotExp[14]=$grandTotExp
$t12Rows.Add($rTotExp) | Out-Null
$totalOpexRowNum = $t12Rows.Count

$t12Rows.Add((NewT12Row)) | Out-Null   # blank
$rNOI = NewT12Row; $rNOI[1]="NET OPERATING INCOME"
$grandNOI=0.0
for($m=0;$m -lt 12;$m++){ $v=$monthlyRevTotal[$m]-$monthlyExpTotal[$m]; $rNOI[2+$m]=$v; $grandNOI+=$v }
$rNOI[14]=$grandNOI
$t12Rows.Add($rNOI) | Out-Null
$noiRowNum = $t12Rows.Count

Write-Host ("`nT12 (in-memory) Annual Revenue={0:N0} OpEx={1:N0} NOI={2:N0} Margin={3:P1}" -f $grandTotRev,$grandTotExp,$grandNOI,($grandNOI/$grandTotRev))

# ================================================================================
# PART 6 -- assemble RENT ROLL rows (OneSite raw-export shape; text-formatted $ cols)
# ================================================================================
function NewRRRow(){ New-Object object[] 43 }   # A..AQ

function AddSectionHeader($rowsList,[string]$text){
  $r = NewRRRow; $r[0]=$text; $rowsList.Add($r) | Out-Null
}

function AddUnitBlock($rowsList,$u,$leafRowNumsList,$anchorInfoList){
  $r = NewRRRow
  $r[0] = [string]$u.UnitNum
  $r[4] = "$($u.Type) - $($u.SqFt)"
  if(-not $u.Occupied){
    $r[8]  = "Vacant"
    $r[21] = FmtMoney $u.Market
    $r[25] = FmtMoney 0
    $r[26] = FmtMoney 0
    $r[28] = FmtMoney $u.Market
    $r[30] = FmtMoney $u.Market
    $r[42] = FmtMoney 0
    $rowsList.Add($r) | Out-Null
    $anchorRow = $rowsList.Count
    $leafRowNumsList.Add($anchorRow) | Out-Null
    $anchorInfoList.Add([pscustomobject]@{UnitNum=$u.UnitNum;Care=$u.Care;Occupied=$false;AnchorRow=$anchorRow}) | Out-Null
    $leafAQ = 0.0
  } else {
    $res1 = $u.Residents[0]
    $r[8]  = $res1.Name
    $r[13] = FmtDate $res1.MoveIn
    $r[17] = FmtDate $res1.MoveIn
    $r[21] = FmtMoney $u.Market
    $r[25] = FmtMoney $res1.InPlace
    $var = $res1.InPlace - $u.Market
    $r[26] = FmtMoney $var
    $r[28] = FmtMoney 0
    $r[30] = FmtMoney $var
    $r[32] = FmtMoney 0
    $r[34] = FmtMoney $res1.CareFee
    $r[36] = FmtMoney 0
    $r[39] = FmtMoney 0
    $aq1 = $res1.InPlace + $res1.CareFee
    $r[42] = FmtMoney $aq1
    $rowsList.Add($r) | Out-Null
    $anchorRow = $rowsList.Count
    $leafRowNumsList.Add($anchorRow) | Out-Null
    $anchorInfoList.Add([pscustomobject]@{UnitNum=$u.UnitNum;Care=$u.Care;Occupied=$true;AnchorRow=$anchorRow}) | Out-Null
    $leafAQ = $aq1

    if($u.Residents.Count -gt 1){
      $res2 = $u.Residents[1]
      $r2 = NewRRRow
      $r2[8]  = $res2.Name
      $r2[13] = FmtDate $res2.MoveIn
      $r2[17] = FmtDate $res2.MoveIn
      if($res2.IsSecond){
        $r2[21] = FmtMoney 0
        $r2[25] = FmtMoney 0
        $r2[32] = FmtMoney $res2.SecondFee
        $r2[34] = FmtMoney 0
        $aq2 = $res2.SecondFee
      } else {
        $r2[21] = FmtMoney $u.Market
        $r2[25] = FmtMoney $res2.InPlace
        $var2 = $res2.InPlace - $u.Market
        $r2[26] = FmtMoney $var2
        $r2[30] = FmtMoney $var2
        $r2[32] = FmtMoney 0
        $r2[34] = FmtMoney $res2.CareFee
        $aq2 = $res2.InPlace + $res2.CareFee
      }
      $r2[36] = FmtMoney 0
      $r2[39] = FmtMoney 0
      $r2[42] = FmtMoney $aq2
      $rowsList.Add($r2) | Out-Null
      $leafRowNumsList.Add($rowsList.Count) | Out-Null
      $leafAQ += $aq2
    }
  }
  $rs = NewRRRow
  $rs[8] = "$($u.UnitNum) TOTAL"
  $rs[42] = FmtMoney $leafAQ
  $rowsList.Add($rs) | Out-Null
  return $leafAQ
}

$rrRows = New-Object System.Collections.Generic.List[object]
1..12 | ForEach-Object { $rrRows.Add((NewRRRow)) | Out-Null }
$rrRows[0][0] = "OneSite Senior Living"
$rrRows[1][0] = "Rent Roll Detail Report"
$rrRows[2][0] = "Community: Aster Ridge Senior Living (DEMO)"
$rrRows[3][0] = "Report Date: 06/30/2026"
$rrRows[4][0] = $DISCLAIMER
# rows 6-12 intentionally blank (report-parameter filler in a real export)

$hdr = NewRRRow
$hdr[0]="Unit"; $hdr[4]="Floor Plan - SQ FT"; $hdr[8]="Resident"; $hdr[13]="Move-In"; $hdr[17]="Moved Onto Property"
$hdr[21]="Market Rent"; $hdr[25]="Actual Rent"; $hdr[26]="Var to Market"; $hdr[28]="Vacancy Var"; $hdr[30]="Total Var"
$hdr[32]="2nd Occupant Fee"; $hdr[34]="Care Fees"; $hdr[36]="Other Fees"; $hdr[39]="Credits"; $hdr[42]="TOTAL"
$rrRows.Add($hdr) | Out-Null    # row 13

$ilLeafRows = New-Object System.Collections.Generic.List[int]
$alLeafRows = New-Object System.Collections.Generic.List[int]
$mcLeafRows = New-Object System.Collections.Generic.List[int]
$anchorInfo = New-Object System.Collections.Generic.List[object]

AddSectionHeader $rrRows "Independent Living"
$ilSum=0.0; foreach($u in $ilUnits){ $ilSum += AddUnitBlock $rrRows $u $ilLeafRows $anchorInfo }
$r=NewRRRow; $r[8]="TOTAL Independent Living"; $r[42]=FmtMoney $ilSum; $rrRows.Add($r) | Out-Null
$ilTotalRowNum = $rrRows.Count

AddSectionHeader $rrRows "Assisted Living"
$alSum=0.0; foreach($u in $alUnits){ $alSum += AddUnitBlock $rrRows $u $alLeafRows $anchorInfo }
$r=NewRRRow; $r[8]="TOTAL Assisted Living"; $r[42]=FmtMoney $alSum; $rrRows.Add($r) | Out-Null
$alTotalRowNum = $rrRows.Count

AddSectionHeader $rrRows "Memory Care"
$mcSum=0.0; foreach($u in $mcUnits){ $mcSum += AddUnitBlock $rrRows $u $mcLeafRows $anchorInfo }
$r=NewRRRow; $r[8]="TOTAL Memory Care"; $r[42]=FmtMoney $mcSum; $rrRows.Add($r) | Out-Null
$mcTotalRowNum = $rrRows.Count

$r=NewRRRow; $r[8]="TOTAL Aster Ridge Senior Living"; $r[42]=FmtMoney ($ilSum+$alSum+$mcSum); $rrRows.Add($r) | Out-Null
$grandTotalRowNum = $rrRows.Count

Write-Host ("RR (in-memory) sections: IL={0:N2} AL={1:N2} MC={2:N2} GRAND={3:N2}" -f $ilSum,$alSum,$mcSum,($ilSum+$alSum+$mcSum))

# ================================================================================
# PART 7 -- assemble CENSUS rows
# ================================================================================
function NewCensusRow(){ New-Object object[] 25 }   # A + 24 months

$censusRows = New-Object System.Collections.Generic.List[object]
1..5 | ForEach-Object { $censusRows.Add((NewCensusRow)) | Out-Null }
$censusRows[0][0] = "Aster Ridge Senior Living (DEMO) -- Occupancy Report"
$censusRows[1][0] = "Jul 2024 - Jun 2026"
$censusRows[2][0] = $DISCLAIMER
# rows 4,5 blank

$hdr = NewCensusRow
$hdr[0] = "Month"
for($m=0;$m -lt 24;$m++){ $hdr[1+$m] = $censusMonths[$m] }
$censusRows.Add($hdr) | Out-Null   # row6

$capBand = NewCensusRow; $capBand[0]="Capacity"; $censusRows.Add($capBand) | Out-Null

$rCapIL = NewCensusRow; $rCapIL[0]="IL"; for($m=0;$m -lt 24;$m++){ $rCapIL[1+$m]=40.0 }; $censusRows.Add($rCapIL) | Out-Null
$rCapAL = NewCensusRow; $rCapAL[0]="AL"; for($m=0;$m -lt 24;$m++){ $rCapAL[1+$m]=30.0 }; $censusRows.Add($rCapAL) | Out-Null
$rCapMC = NewCensusRow; $rCapMC[0]="MC"; for($m=0;$m -lt 24;$m++){ $rCapMC[1+$m]=14.0 }; $censusRows.Add($rCapMC) | Out-Null
$rCapTot= NewCensusRow; $rCapTot[0]="Total"; for($m=0;$m -lt 24;$m++){ $rCapTot[1+$m]=84.0 }; $censusRows.Add($rCapTot) | Out-Null

$censusRows.Add((NewCensusRow)) | Out-Null   # blank

$occBand = NewCensusRow; $occBand[0]="Occupied (Month-End)"; $censusRows.Add($occBand) | Out-Null

$rOccIL = NewCensusRow; $rOccIL[0]="IL"; for($m=0;$m -lt 24;$m++){ $rOccIL[1+$m]=[double]$occIL_full[$m] }; $censusRows.Add($rOccIL) | Out-Null
$rOccILRowNum = $censusRows.Count
$rOccAL = NewCensusRow; $rOccAL[0]="AL"; for($m=0;$m -lt 24;$m++){ $rOccAL[1+$m]=[double]$occAL_full[$m] }; $censusRows.Add($rOccAL) | Out-Null
$rOccALRowNum = $censusRows.Count
$rOccMC = NewCensusRow; $rOccMC[0]="MC"; for($m=0;$m -lt 24;$m++){ $rOccMC[1+$m]=[double]$occMC_full[$m] }; $censusRows.Add($rOccMC) | Out-Null
$rOccMCRowNum = $censusRows.Count
$rOccTot= NewCensusRow; $rOccTot[0]="Total"; for($m=0;$m -lt 24;$m++){ $rOccTot[1+$m]=[double]($occIL_full[$m]+$occAL_full[$m]+$occMC_full[$m]) }; $censusRows.Add($rOccTot) | Out-Null

# ================================================================================
# PART 8 -- write the three workbooks (one Excel session, batched block writes)
# ================================================================================
$wbRR=$null; $wbT=$null; $wbC=$null
$xl = New-ExcelTracked
$rrIsXls = $true
try {
  # ---- Rent Roll ----
  $wbRR = $xl.Workbooks.Add()
  $xl.Calculation = -4135
  $wsRR = $wbRR.Worksheets.Item(1)
  $wsRR.Name = "Rent Roll"
  $totalRRRows = $rrRows.Count
  $arrRR = New-Object 'object[,]' $totalRRRows,43
  for($i=0;$i -lt $totalRRRows;$i++){ for($j=0;$j -lt 43;$j++){ $arrRR[$i,$j] = $rrRows[$i][$j] } }
  $rngRR = $wsRR.Range($wsRR.Cells(1,1), $wsRR.Cells($totalRRRows,43))
  SetBlock $rngRR $arrRR
  $wsRR.Rows.Item(13).Font.Bold = $true
  $wsRR.Cells(1,1).Font.Bold = $true
  $wsRR.Cells(5,1).Font.Italic = $true
  $wsRR.Columns.Item(9).ColumnWidth = [double]22
  $wsRR.Columns.Item(5).ColumnWidth = [double]20
  $wbRR.Windows.Item(1).DisplayGridlines = $false
  $xl.CalculateFull()

  $rrTemp = Join-Path $env:TEMP ("AsterRidge_RR_{0}.xls" -f ([guid]::NewGuid().ToString('N').Substring(0,8)))
  $saved = $false
  for($attempt=1; $attempt -le 2 -and -not $saved; $attempt++){
    try { $wbRR.SaveAs($rrTemp, 56); $saved = $true }
    catch { Start-Sleep -Milliseconds 500 }
  }
  if(-not $saved){
    $rrIsXls = $false
    $rrTemp = Join-Path $env:TEMP ("AsterRidge_RR_{0}.xlsx" -f ([guid]::NewGuid().ToString('N').Substring(0,8)))
    $wbRR.SaveAs($rrTemp, 51)
    $rrPathFinal = Join-Path $rrDir '06.30.2026 Aster Ridge SeniorLivingRentRoll.xlsx'
    Write-Host "WARNING: BIFF8 (.xls) SaveAs failed twice -- shipped .xlsx instead."
  }
  $wbRR.Close($false); $wbRR=$null
  Copy-Item -LiteralPath $rrTemp -Destination $rrPathFinal -Force
  Remove-Item -LiteralPath $rrTemp -Force -ErrorAction SilentlyContinue

  # ---- T12 Income Statement ----
  $wbT = $xl.Workbooks.Add()
  $wsT = $wbT.Worksheets.Item(1)
  $wsT.Name = "Income Statement"
  $totalTRows = $t12Rows.Count
  $arrT = New-Object 'object[,]' $totalTRows,15
  for($i=0;$i -lt $totalTRows;$i++){ for($j=0;$j -lt 15;$j++){ $arrT[$i,$j] = $t12Rows[$i][$j] } }
  $rngT = $wsT.Range($wsT.Cells(1,1), $wsT.Cells($totalTRows,15))
  SetBlock $rngT $arrT
  $wsT.Range($wsT.Cells(6,3), $wsT.Cells($totalTRows,15)).NumberFormat = "#,##0"
  $wsT.Rows.Item(5).Font.Bold = $true
  $wsT.Cells(1,1).Font.Bold = $true
  $wsT.Cells(3,1).Font.Italic = $true
  $wsT.Columns.Item(2).ColumnWidth = [double]30
  $wbT.Windows.Item(1).DisplayGridlines = $false
  $xl.CalculateFull()

  $tTemp = Join-Path $env:TEMP ("AsterRidge_T12_{0}.xlsx" -f ([guid]::NewGuid().ToString('N').Substring(0,8)))
  $wbT.SaveAs($tTemp, 51)
  $wbT.Close($false); $wbT=$null
  Copy-Item -LiteralPath $tTemp -Destination $t12PathFinal -Force
  Remove-Item -LiteralPath $tTemp -Force -ErrorAction SilentlyContinue

  # ---- Census / Occupancy ----
  $wbC = $xl.Workbooks.Add()
  $wsC = $wbC.Worksheets.Item(1)
  $wsC.Name = "Occupancy"
  $totalCRows = $censusRows.Count
  $arrC = New-Object 'object[,]' $totalCRows,25
  for($i=0;$i -lt $totalCRows;$i++){ for($j=0;$j -lt 25;$j++){ $arrC[$i,$j] = $censusRows[$i][$j] } }
  $rngC = $wsC.Range($wsC.Cells(1,1), $wsC.Cells($totalCRows,25))
  SetBlock $rngC $arrC
  $wsC.Range($wsC.Cells(6,2), $wsC.Cells($totalCRows,25)).NumberFormat = "#,##0"
  $wsC.Rows.Item(6).Font.Bold = $true
  $wsC.Cells(1,1).Font.Bold = $true
  $wsC.Cells(3,1).Font.Italic = $true
  $wsC.Columns.Item(1).ColumnWidth = [double]22
  $wbC.Windows.Item(1).DisplayGridlines = $false
  $xl.CalculateFull()

  $cTemp = Join-Path $env:TEMP ("AsterRidge_Census_{0}.xlsx" -f ([guid]::NewGuid().ToString('N').Substring(0,8)))
  $wbC.SaveAs($cTemp, 51)
  $wbC.Close($false); $wbC=$null
  Copy-Item -LiteralPath $cTemp -Destination $censusPathFinal -Force
  Remove-Item -LiteralPath $cTemp -Force -ErrorAction SilentlyContinue

  Write-Host "`nWrote:"
  Write-Host "  $rrPathFinal"
  Write-Host "  $t12PathFinal"
  Write-Host "  $censusPathFinal"
}
finally {
  if($wbRR){ try{ $wbRR.Close($false) }catch{} }
  if($wbT){  try{ $wbT.Close($false)  }catch{} }
  if($wbC){  try{ $wbC.Close($false)  }catch{} }
  $xl.Quit()
  [System.Runtime.Interopservices.Marshal]::ReleaseComObject($xl) | Out-Null
  Stop-TrackedExcel
}

# ================================================================================
# PART 9 -- GATE verification: reopen the SAVED files fresh, read-only, recompute
# ================================================================================
Write-Host "`n=== VERIFICATION (reopened from saved files) ==="
$xl2 = New-ExcelTracked
try {
  # ---- T12 ----
  $wbV = $xl2.Workbooks.Open($t12PathFinal)
  $wsV = $wbV.Worksheets.Item(1)
  $revFail=0; $expFail=0; $noiFail=0
  for($m=0;$m -lt 12;$m++){
    $col = 3+$m
    $sumRev=0.0; foreach($rn in $revLeafRowNums){ $sumRev += [double]$wsV.Cells($rn,$col).Value2 }
    $totRev = [double]$wsV.Cells($totalRevenueRowNum,$col).Value2
    if([math]::Abs($sumRev-$totRev) -gt 0.005){ $revFail++; Write-Host ("REV MISMATCH m={0} sum={1} total={2}" -f $m,$sumRev,$totRev) }

    $sumExp=0.0; foreach($rn in $expLeafRowNums){ $sumExp += [double]$wsV.Cells($rn,$col).Value2 }
    $totExp = [double]$wsV.Cells($totalOpexRowNum,$col).Value2
    if([math]::Abs($sumExp-$totExp) -gt 0.005){ $expFail++; Write-Host ("EXP MISMATCH m={0} sum={1} total={2}" -f $m,$sumExp,$totExp) }

    $noi = [double]$wsV.Cells($noiRowNum,$col).Value2
    if([math]::Abs($noi-($totRev-$totExp)) -gt 0.005){ $noiFail++; Write-Host ("NOI MISMATCH m={0}" -f $m) }
  }
  $annualRev = [double]$wsV.Cells($totalRevenueRowNum,15).Value2
  $annualExp = [double]$wsV.Cells($totalOpexRowNum,15).Value2
  $annualNOI = [double]$wsV.Cells($noiRowNum,15).Value2
  Write-Host ("T12 gate: revFail={0} expFail={1} noiFail={2} (of 12 months each)" -f $revFail,$expFail,$noiFail)
  Write-Host ("T12 ANNUAL: Revenue={0:N0}  OpEx={1:N0}  NOI={2:N0}  Margin={3:P1}" -f $annualRev,$annualExp,$annualNOI,($annualNOI/$annualRev))

  $rowByCode=@{}
  foreach($rn in $revLeafRowNums){ $rowByCode[[string]$wsV.Cells($rn,1).Value2] = $rn }
  $juneCol=14
  $tieCodes = @('4010','4020','4030','4110','4120','4210')
  $rosterVals = @{ '4010'=$ilRentTotal; '4020'=$alRentTotal; '4030'=$mcRentTotal; '4110'=$alCareTotal; '4120'=$mcCareTotal; '4210'=($ilSecondTotal+$alSecondTotal) }
  $crossFail=0
  foreach($code in $tieCodes){
    $t12Val = [double]$wsV.Cells($rowByCode[$code],$juneCol).Value2
    $rosterVal = $rosterVals[$code]
    $match = ([math]::Abs($t12Val-$rosterVal) -lt 0.005)
    if(-not $match){ $crossFail++ }
    Write-Host ("CROSS-TIE {0}: T12(Jun-26)={1:N2}  roster={2:N2}  match={3}" -f $code,$t12Val,$rosterVal,$match)
  }
  $wbV.Close($false)

  # ---- Census ----
  $wbV2 = $xl2.Workbooks.Open($censusPathFinal)
  $wsV2 = $wbV2.Worksheets.Item(1)
  $populated = 0
  for($c=2;$c -le 25;$c++){
    if([double]$wsV2.Cells($rOccILRowNum,$c).Value2 -gt 0 -and [double]$wsV2.Cells($rOccALRowNum,$c).Value2 -gt 0 -and [double]$wsV2.Cells($rOccMCRowNum,$c).Value2 -gt 0){ $populated++ }
  }
  $juneIL_c = [double]$wsV2.Cells($rOccILRowNum,25).Value2
  $juneAL_c = [double]$wsV2.Cells($rOccALRowNum,25).Value2
  $juneMC_c = [double]$wsV2.Cells($rOccMCRowNum,25).Value2
  Write-Host ("Census gate: {0}/24 months populated (non-zero all 3 care levels)" -f $populated)
  Write-Host ("Census Jun-2026: IL={0} AL={1} MC={2} (expect 37/28/13)" -f $juneIL_c,$juneAL_c,$juneMC_c)
  $wbV2.Close($false)

  # ---- Rent Roll ----
  $wbV3 = $xl2.Workbooks.Open($rrPathFinal)
  $wsV3 = $wbV3.Worksheets.Item(1)
  function SumLeafAQ($ws,$rowNums){
    $s=0.0; foreach($rn in $rowNums){ $s += (ParseMoneyText $ws.Cells($rn,43).Value2) }
    return $s
  }
  $ilCheckSum = SumLeafAQ $wsV3 $ilLeafRows
  $alCheckSum = SumLeafAQ $wsV3 $alLeafRows
  $mcCheckSum = SumLeafAQ $wsV3 $mcLeafRows
  $ilCheckTotal = ParseMoneyText $wsV3.Cells($ilTotalRowNum,43).Value2
  $alCheckTotal = ParseMoneyText $wsV3.Cells($alTotalRowNum,43).Value2
  $mcCheckTotal = ParseMoneyText $wsV3.Cells($mcTotalRowNum,43).Value2
  $grandCheckTotal = ParseMoneyText $wsV3.Cells($grandTotalRowNum,43).Value2
  Write-Host ("RR gate IL: leafSum={0:N2} sectionTotal={1:N2} match={2}" -f $ilCheckSum,$ilCheckTotal,([math]::Abs($ilCheckSum-$ilCheckTotal) -lt 0.005))
  Write-Host ("RR gate AL: leafSum={0:N2} sectionTotal={1:N2} match={2}" -f $alCheckSum,$alCheckTotal,([math]::Abs($alCheckSum-$alCheckTotal) -lt 0.005))
  Write-Host ("RR gate MC: leafSum={0:N2} sectionTotal={1:N2} match={2}" -f $mcCheckSum,$mcCheckTotal,([math]::Abs($mcCheckSum-$mcCheckTotal) -lt 0.005))
  Write-Host ("RR GRAND TOTAL (from file) = {0:N2}" -f $grandCheckTotal)

  $ilOccRead=0; $alOccRead=0; $mcOccRead=0
  foreach($a in $anchorInfo){
    $val = [string]$wsV3.Cells($a.AnchorRow,9).Value2
    $isOcc = ($val -ne "Vacant")
    if($a.Care -eq 'IL' -and $isOcc){ $ilOccRead++ }
    if($a.Care -eq 'AL' -and $isOcc){ $alOccRead++ }
    if($a.Care -eq 'MC' -and $isOcc){ $mcOccRead++ }
  }
  Write-Host ("RR occupied counts (from file): IL={0} AL={1} MC={2} (expect 37/28/13)" -f $ilOccRead,$alOccRead,$mcOccRead)
  $wbV3.Close($false)
}
finally {
  $xl2.Quit()
  [System.Runtime.Interopservices.Marshal]::ReleaseComObject($xl2) | Out-Null
  Stop-TrackedExcel
}

Write-Host "`n=== DONE ==="
