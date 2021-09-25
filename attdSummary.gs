var attdSummarySheet = ss.getSheetByName('Attendance Summary');

// sets up data validation for viewing a summary of attendances (not used)
function companyList() {

  var orndFlexiArr = flexiSheet.getRange('A2:A'+flexiSheet.getLastRow()).getValues();

  var orndOPPArr = standaloneOPPSheet.getRange('A2:A'+standaloneOPPSheet.getLastRow()).getValues();

  //attdSummarySheet.getRange(1,1).setValue('Category').setFontWeight('bold');

  var compList = [];

  for(var i = 0 ; i < orndFlexiArr.length ; i++){
    compList.push(orndFlexiArr[i][0]);


  }

  for(var i = 0 ; i < orndOPPArr.length ; i++){
    compList.push(orndOPPArr[i][0]);
  }

  var compSet = new Set(compList);

  var uniqueCompList = [...compSet];

  return uniqueCompList;
}

function setDataValidation(){

  attdSummarySheet.getRange(2, 1).setDataValidation(null);

  var rule = SpreadsheetApp.newDataValidation().requireValueInList(companyList()).build();

  attdSummarySheet.getRange(2, 1).setDataValidation(rule);
}

function generateUniqueCompanyData(){

  var masterArr = monthlyAttendanceSheet.getRange('A2:G'+monthlyAttendanceSheet.getLastRow()).getValues();

  var companyName =  attdSummarySheet.getRange(2, 1).getValue();

  var tempArr = [];

  for(var i = 0 ; i < masterArr.length ; i++){
    if(companyName.includes(masterArr[i][5])){

      tempArr.push([masterArr[i][0],masterArr[i][1],masterArr[i][2],masterArr[i][5]]);
    }
  }

  attdSummarySheet.getRange('A6:I').clear();

  if(tempArr.length != 0){
    attdSummarySheet.getRange(6 , 1 , tempArr.length , tempArr[0].length).setValues(tempArr);
  }
  
}


