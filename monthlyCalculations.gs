var phFlexiSheet = ss.getSheetByName('Placeholder Flexi');

//updates PlaceHolderFlexi at the start of the month
function updatePlaceholderFlexiStart(){
  var orndFlexiArr = flexiSheet.getRange('A2:M'+flexiSheet.getLastRow()).getValues();

  for (var i = 0 ; i < orndFlexiArr.length ; i++){

    orndFlexiArr[i].splice(1,1);
    orndFlexiArr[i].splice(2,4);
    orndFlexiArr[i].splice(6,2);

  }

  phFlexiSheet.getRange('A2:H').clear()
  phFlexiSheet.getRange(2 , 1 , orndFlexiArr.length , orndFlexiArr[0].length).setValues(orndFlexiArr);
}

//updates PlaceHolderFlexi at the end of the month
function updatePlaceholderFlexiEnd(){

  var orndFlexiArr = flexiSheet.getRange('A2:M'+flexiSheet.getLastRow()).getValues();

  for (var i = 0 ; i < orndFlexiArr.length ; i++){

    orndFlexiArr[i].splice(0,9);
    orndFlexiArr[i].splice(1,1);
    orndFlexiArr[i].splice(2,1);

  }

  phFlexiSheet.getRange(2 , 7 , orndFlexiArr.length , orndFlexiArr[0].length).setValues(orndFlexiArr);
}

//monthy summary calculation for flexi teams
function monthlyCalculationFlexi() {

  var dailyFlexiRawArr = dailyCalcFlexiSheet.getRange('B2:K'+dailyCalcFlexiSheet.getLastRow()).getValues();
  var phFlexiArr = phFlexiSheet.getRange('A2:H'+phFlexiSheet.getLastRow()).getValues();

  var dailyFlexiArr = dailyFlexiRawArr.filter(function (x) { 
    return !(x.every(element => element === (undefined || null || '')))
  });

  var dpUsageTotal; 
  var creditUsageTotal; 
  var charge;

  for(var i = 0 ; i < phFlexiArr.length ; i++){

    dpUsageTotal = 0;
    creditUsageTotal = 0;
    charge = 0;

    for(var j = 0 ; j < dailyFlexiArr.length ; j++){
    
      if(phFlexiArr[i][0] == dailyFlexiArr[j][0]){

        dpUsageTotal += dailyFlexiArr[j][3];
        creditUsageTotal += dailyFlexiArr[j][7];
        charge += dailyFlexiArr[j][9];

      }
    }

    phFlexiArr[i].push(dpUsageTotal, creditUsageTotal, charge);
  }

  phFlexiSheet.getRange(2 , 1 , phFlexiArr.length , phFlexiArr[0].length).setValues(phFlexiArr);
}


