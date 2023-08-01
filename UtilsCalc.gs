/// ---- UtilsCalc ------ RS

function calcElec(eused) {

//var lvl1 = 1.4243; // 0-500
var lvl1 = 1.5416; // 0-500
var lvl2 = 1.6346; // 501-1000
var lvl3 = 1.7552; // 1001-2000
var lvl4 = 1.8518; // 2001-3000
var lvl5 = 1.9426; // 3001->

var vat = 1.15
var cost = 0;

    if (eused > 3000)
    {
        cost = 500*lvl1 + 500*lvl2 + 1000*lvl3 + 1000*lvl4 + (eused - 3000)*lvl5;
        cost = cost*vat/100;
    }
    else if(eused > 2000)
    {
        cost = 500*lvl1 + 500*lvl2 + 1000*lvl3 + (eused - 2000)*lvl4;
        cost = cost*vat/100;
        
    }
        else if(eused > 1000)
    {
        cost = 500*lvl1 + 500*lvl2 + (eused - 1000)*lvl3;
        cost = cost*vat/100;
        
    }
    else if(eused >= 500)
    {
        cost = 500*lvl1 + (eused - 500)*lvl2;
        cost = cost*vat/100;
        
    }
    else if(eused < 500)
    {
        cost = eused*lvl1;
        cost = cost*vat/100;
        
    }
    else if (eused == 0){
      cost = 0.1;
    }
  if (cost < 0.1)
  {
    cost = 0.0;
  }
  return cost*100;
}


function calcWater($wused) {

$level1 = 9.10; // 0-6 Kilolitrers
$level2 = 18.99; // 7 -10 Kilolitrers
$level3 = 19.82; // 11-15 Kilolitrers
$level4 = 27.79; // 16-20 Kilolitrers
$level5 = 38.40; // 20-30 Kilolitrers
$level6 = 42.00; // 30-40 Kilolitrers
$level7 = 52.79; // 40 - 50
$level8 = 56.79; // 50 ->

$Limit1 = 6;
$Limit2 = 4;
$Limit3 = 5;
$Limit4 = 5;
$Limit5 = 10;
$Limit6 = 10;
$Limit7 = 10;

$LimStep1 = 6;
$LimStep2 = 10;
$LimStep3 = 15;
$LimStep4 = 20;
$LimStep5 = 30;
$LimStep6 = 40;
$LimStep7 = 50;

vat = 1.15;

$Cost = 0.00;

    if ($wused > $LimStep7)
    {
        $Cost = $Limit1*$level1 + $Limit2*$level2 + $Limit3*$level3 + $Limit4*$level4 + $Limit5*$level5 + $Limit6*$level6 + $Limit7*$level7 + ($wused - $LimStep7)*$level8;
        $Cost = $Cost*vat;
    }
    else if ($wused > $LimStep6)
    {
        $Cost = $Limit1*$level1 + $Limit2*$level2 + $Limit3*$level3 + $Limit4*$level4 + $Limit5*$level5 +$Limit6*$level6 + ($wused - $LimStep6)*$level7;
        $Cost = $Cost*vat;
    }
    else if ($wused > $LimStep5)
    {
        $Cost = $Limit1*$level1 + $Limit2*$level2 + $Limit3*$level3 + $Limit4*$level4 + $Limit5*$level5 + ($wused - $LimStep5)*$level6;
        $Cost = $Cost*vat;
    }
    else if ($wused > $LimStep4)
    {
        $Cost = $Limit1*$level1 + $Limit2*$level2 + $Limit3*$level3 + $Limit4*$level4 + ($wused - $LimStep4)*$level5;
        $Cost = $Cost*vat;
    }
    else if ($wused > $LimStep3)
    {
        $Cost = $Limit1*$level1 + $Limit2*$level2 + $Limit3*$level3 + ($wused - $LimStep3)*$level4;
        $Cost = $Cost*vat;
    }
    else if ($wused > $LimStep2)
    {
        $Cost = $Limit1*$level1 + $Limit2*$level2 + ($wused - $LimStep2)*$level3;
        $Cost = $Cost*vat;
    }
    else if ($wused >= $LimStep1)
    {
        $Cost = $Limit1*$level1 + ($wused - $LimStep1)*$level2;
        $Cost = $Cost*vat;
    }
    else if($wused < $LimStep1)
    {
        $Cost = $wused*$level1;
        $Cost = $Cost*vat;
        
    }
    if ($wused == 0)
  {
        $Cost = 0;
    }
  //
  //if ($Cost <10) && {
  //   $Cost = 10;
  //}
    return $Cost;

}



function calcGas(eused) {

//var lvl1 = 1.4243; // 0-500
var lvl1 = 1.5416; // 0-500
var lvl2 = 1.6346; // 501-1000
var lvl3 = 1.7552; // 1001-2000
var lvl4 = 1.8518; // 2001-3000
var lvl5 = 1.9426; // 3001->

var vat = 1.15
var cost = 0;

    if (eused > 3000)
    {
        cost = 500*lvl1 + 500*lvl2 + 1000*lvl3 + 1000*lvl4 + (eused - 3000)*lvl5;
        cost = cost*vat/100;
    }
    else if(eused > 2000)
    {
        cost = 500*lvl1 + 500*lvl2 + 1000*lvl3 + (eused - 2000)*lvl4;
        cost = cost*vat/100;
        
    }
        else if(eused > 1000)
    {
        cost = 500*lvl1 + 500*lvl2 + (eused - 1000)*lvl3;
        cost = cost*vat/100;
        
    }
    else if(eused >= 500)
    {
        cost = 500*lvl1 + (eused - 500)*lvl2;
        cost = cost*vat/100;
        
    }
    else if(eused < 500)
    {
        cost = eused*lvl1;
        cost = cost*vat/100;
        
    }
    else if (eused == 0){
      cost = 0.1;
    }
  if (cost < 0.1)
  {
    cost = 0.1;
  }
  return cost*100;
}
