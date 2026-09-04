# Diff two .H tabs cell-by-cell over A1:U353 (or a given range): Value2 (tol 1e-6), Formula, NumberFormat,
# Font.Color, Font.Bold, Interior.Color, Borders(top/bottom) on every cell. Prints differing cells.
param(
  [Parameter(Mandatory=$true)][string]$PathA,
  [Parameter(Mandatory=$true)][string]$TabA,
  [Parameter(Mandatory=$true)][string]$PathB,
  [Parameter(Mandatory=$true)][string]$TabB,
  [int]$LastRow = 353,
  [int]$LastCol = 21,
  [switch]$SkipNumfmt,
  [switch]$SkipBorder
)
[System.Threading.Thread]::CurrentThread.CurrentCulture=[System.Globalization.CultureInfo]'en-US'
$ErrorActionPreference='Stop'
$xl = New-Object -ComObject Excel.Application
$xl.Visible=$false; $xl.DisplayAlerts=$false
$wbA = $xl.Workbooks.Open($PathA,0,$true)
$wbB = $xl.Workbooks.Open($PathB,0,$true)
$wsA = $wbA.Worksheets.Item($TabA)
$wsB = $wbB.Worksheets.Item($TabB)

$diffs = @()
$counts = @{ value=0; formula=0; numfmt=0; fontcolor=0; bold=0; fill=0; topborder=0; botborder=0 }
for($r=1;$r -le $LastRow;$r++){
  for($c=1;$c -le $LastCol;$c++){
    $ca = $wsA.Cells($r,$c); $cb = $wsB.Cells($r,$c)
    $va = $ca.Value2; $vb = $cb.Value2
    $valDiff = $false
    if(($va -is [double] -or $va -is [int]) -and ($vb -is [double] -or $vb -is [int])){
      if([Math]::Abs([double]$va - [double]$vb) -gt 1e-6){ $valDiff=$true }
    } elseif("$va" -ne "$vb"){ $valDiff=$true }
    $fa = $ca.Formula; $fb = $cb.Formula
    $fmtA = $ca.NumberFormat; $fmtB = $cb.NumberFormat
    $fca = $ca.Font.Color; $fcb = $cb.Font.Color
    $bda = $ca.Font.Bold; $bdb = $cb.Font.Bold
    $fla = $ca.Interior.Color; $flb = $cb.Interior.Color
    $topA=$ca.Borders.Item(8).LineStyle; $topB=$cb.Borders.Item(8).LineStyle
    $botA=$ca.Borders.Item(9).LineStyle; $botB=$cb.Borders.Item(9).LineStyle
    $reasons=@()
    if($valDiff){ $reasons+="value(A=$va,B=$vb)"; $counts.value++ }
    if("$fa" -ne "$fb"){ $reasons+="formula(A=$fa|B=$fb)"; $counts.formula++ }
    if(-not $SkipNumfmt -and "$fmtA" -ne "$fmtB"){ $reasons+="numfmt(A=$fmtA|B=$fmtB)"; $counts.numfmt++ }
    if($fca -ne $fcb){ $reasons+="fontcolor(A=$fca,B=$fcb)"; $counts.fontcolor++ }
    if($bda -ne $bdb){ $reasons+="bold(A=$bda,B=$bdb)"; $counts.bold++ }
    if($fla -ne $flb){ $reasons+="fill(A=$fla,B=$flb)"; $counts.fill++ }
    if(-not $SkipBorder -and $topA -ne $topB){ $reasons+="topborder(A=$topA,B=$topB)"; $counts.topborder++ }
    if(-not $SkipBorder -and $botA -ne $botB){ $reasons+="botborder(A=$botA,B=$botB)"; $counts.botborder++ }
    if($reasons.Count -gt 0){
      $addr = $ca.Address($false,$false)
      $diffs += "R$r C$c ($addr) : " + ($reasons -join '; ')
    }
  }
}
Write-Output ("TOTAL DIFFS: " + $diffs.Count)
Write-Output ("BY CATEGORY: " + ($counts | ConvertTo-Json -Compress))
$diffs

$wbA.Close($false); $wbB.Close($false); $xl.Quit(); [System.Runtime.InteropServices.Marshal]::ReleaseComObject($xl)|Out-Null
