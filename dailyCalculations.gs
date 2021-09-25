var dailyCalcFlexiSheet = ss.getSheetByName('Daily Calculation Flexi');
var dailyCalcStandaloneSheet = ss.getSheetByName('Daily Calculation Standalone');

//get daily attendance count from daily attendance cleaned tab
function getAttendanceCount(array) {
  var dailyAttdArr = dailyAttdCleanedSheet.getRange('A2:G'+dailyAttdCleanedSheet.getLastRow()).getValues();
  var countArr = [];
  var tempArr = array

  for(var i = 0 ; i < tempArr.length ; i++){

    var count = 0;

    for(var j = 0 ; j < dailyAttdArr.length ; j++){

      if(dailyAttdArr[j][5] != ''){

        var compName1 = tempArr[i][0].toLowerCase();

        var compName2 = dailyAttdArr[j][5].toLowerCase();

        if(compName1.includes(compName2)){
        
          count++;

        }

      }
    }

    countArr.push([tempArr[i][0], tempArr[i][2], count]);
  }

  return countArr;
}

// daily calculation for Flexi Teams
function dailyCalcFlexi(){

  var flexiArr = flexiSheet.getRange('A2:M'+flexiSheet.getLastRow()).getValues();
  var countArr = getAttendanceCount(flexiArr);

  var compName;
  var buffer;
  var dpQuota;
  var dpUsage;
  var dpBalance;
  var creditQuota;
  var creditUsage;
  var creditBalance;
  var charge;
  var attdExcess;

  var tempArr = [];

  for(var i = 0 ; i < flexiArr.length ; i++){
    compName = flexiArr[i][0];
    buffer = flexiArr[i][7];
    dpQuota = flexiArr[i][9];
    dpUsage = countArr[i][2];
    creditQuota = flexiArr[i][11];
    charge = 0;

    attdExcess = dpUsage - buffer;

    if(attdExcess > 0){

      dpBalance = dpQuota - attdExcess;
      creditUsage = 0;
      creditBalance = creditQuota - creditUsage;

      if(dpBalance < 0){

        creditUsage = Math.abs(dpBalance * 30);

        creditBalance = creditQuota - creditUsage;

        if(creditBalance < 0){

          charge = charge + Math.abs(creditBalance);

          creditBalance = 0;

        } 

        dpBalance = 0;
      }

    } else {
      attdExcess = 0;
      dpBalance = dpQuota - attdExcess;
      creditUsage = 0;
      creditBalance = creditQuota - creditUsage;
    }

    tempArr.push([compName, buffer, dpQuota, dpUsage, attdExcess , dpBalance, creditQuota, creditUsage, creditBalance, charge]);
  }

  for (var i = 0 ; i < tempArr.length ; i++){

    flexiArr[i][9] = tempArr[i][5];

    flexiArr[i][12] +=  tempArr[i][7];

    flexiArr[i][11] = flexiArr[i][10] - flexiArr[i][12];
  }
  
  if(dailyCalcFlexiSheet.getLastRow() == 1){

    dailyCalcFlexiSheet.getRange(dailyCalcFlexiSheet.getLastRow()+1 , 1).setValue(today.toLocaleDateString('en-MY'));

    dailyCalcFlexiSheet.getRange(dailyCalcFlexiSheet.getLastRow() , 2 , tempArr.length , tempArr[0].length).setValues(tempArr);
    
  } else {

    dailyCalcFlexiSheet.getRange(dailyCalcFlexiSheet.getLastRow()+2 , 1).setValue(today.toLocaleDateString('en-MY'));

    dailyCalcFlexiSheet.getRange(dailyCalcFlexiSheet.getLastRow() , 2 , tempArr.length , tempArr[0].length).setValues(tempArr);

  }

  flexiSheet.getRange(2,1 , flexiArr.length , flexiArr[0].length).setValues(flexiArr);

}

// daily calculation for Standalone OPPs
function dailyCalcStandalone(){

  var standaloneArr = standaloneOPPSheet.getRange('A2:J'+standaloneOPPSheet.getLastRow()).getValues();
  var countArr = getAttendanceCount(standaloneArr);
  var compName;
  var memberName;
  var attdNum;
  var dpQuota;
  var dpUsage;
  var dpBalance; 
  var creditQuota;
  var creditUsage; 
  var creditBalance; 
  var charge; 

  var tempArr = [];

  for (var i = 0 ; i < standaloneArr.length ; i++){

    compName = standaloneArr[i][0];
    memberName = standaloneArr[i][1];
    attdNum = countArr[i][2];
    dpQuota = standaloneArr[i][6];
    dpUsage = attdNum;
    creditQuota = standaloneArr[i][8];
    charge = 0;

    dpBalance = dpQuota - dpUsage;
    
    if(dpBalance < 0 ){

      creditUsage = Maths.abs(dpBalance) * 30;

      creditBalance = creditQuota - creditUsage; 

      if(creditBalance < 0){

        charge = charge + Math.abs(creditBalance);

        creditBalance = 0;

      } 

      dpBalance = 0;

    } else {

      creditUsage = 0;

      creditBalance = creditQuota - creditUsage; 

    }

    tempArr.push([compName, memberName, dpQuota, dpUsage,dpBalance, creditQuota ,creditUsage,creditBalance,charge]);
  }

  for (var i = 0 ; i < tempArr.length ; i++){

    standaloneArr[i][6] = tempArr[i][4];

    standaloneArr[i][9] +=  tempArr[i][6];

    standaloneArr[i][8] = standaloneArr[i][7] - standaloneArr[i][9];
  }

  if(dailyCalcStandaloneSheet.getLastRow() == 1){

    dailyCalcStandaloneSheet.getRange(dailyCalcStandaloneSheet.getLastRow()+1 , 1).setValue(today.toLocaleDateString('en-MY'));

    dailyCalcStandaloneSheet.getRange(dailyCalcStandaloneSheet.getLastRow() , 2 , tempArr.length , tempArr[0].length).setValues(tempArr);

  } else {

    dailyCalcStandaloneSheet.getRange(dailyCalcStandaloneSheet.getLastRow()+2 , 1).setValue(today.toLocaleDateString('en-MY'));

    dailyCalcStandaloneSheet.getRange(dailyCalcStandaloneSheet.getLastRow() , 2 , tempArr.length , tempArr[0].length).setValues(tempArr);

  }

  standaloneOPPSheet.getRange(2,1 , standaloneArr.length , standaloneArr[0].length).setValues(standaloneArr);
}



