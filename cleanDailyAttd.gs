var dailyAttdCleanedSheet = ss.getSheetByName('Daily Attendance Cleaned');

//generates an array by unique company names (no duplicates)
function generateUniqueCompanyArr(){

  var flexiArr = filterFlexi();
  var oppArr = filterStandaloneOPP(); 
  var compArr = [];

  for(var i = 0 ; i < flexiArr.length ; i++){
    compArr.push(flexiArr[i][0]);
  }

  for(var i = 0 ; i < oppArr.length ; i++){
    compArr.push(oppArr[i][0]);
  }

  var compSet = new Set(compArr);
  var uniqueCompArr = [...compSet];

  return uniqueCompArr;
}


// clean daily attendance tab
function cleanDailyAttendance() {

  var dailyAttdArr = dailyAttdSheet.getRange('A2:I'+dailyAttdSheet.getLastRow()).getValues();
  var uniqueCompArr = generateUniqueCompanyArr();
  var tempArr = [];

  for(var i = 0 ; i < uniqueCompArr.length ; i++){

    for(var j = 0 ; j < dailyAttdArr.length ; j++){

      if(uniqueCompArr[i].includes(dailyAttdArr[j][5])&&(dailyAttdArr[j][6]=='Y'||dailyAttdArr[j][7]=='Y'||dailyAttdArr[j][8]=='Y')){

        tempArr.push(dailyAttdArr[j]);

      }
    }
  }

  for(var i = 0 ; i < tempArr.length ; i++){
    if(tempArr[i][6] == 'Y'){
      tempArr[i][6] = "EarlyIn";
      tempArr[i].splice(7,2);
    }

    if(tempArr[i][7] == 'Y'){
      tempArr[i][6] = "LateIn";
      tempArr[i].splice(7,2);
    }

    if(tempArr[i][8] == 'Y'){
      tempArr[i][6] ="Incomplete";
      tempArr[i].splice(7,2);
    } 
  }

  dailyAttdCleanedSheet.getRange('A2:I').clear();

  if(tempArr.length != 0){

    dailyAttdCleanedSheet.getRange(2, 1, tempArr.length , tempArr[0].length).setValues(tempArr);

  }
}
