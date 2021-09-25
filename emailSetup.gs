var emailSheet = SpreadsheetApp.openById('17J14IS9FJdgxpZgWF9AO3TVWkKQ9ctfgGB1DvKYWuWo');
var summaryTab = emailSheet.getSheetByName('Summary');
var rawAttdTab = emailSheet.getSheetByName('Raw Attendance');

//master script for flexi email setup
function flexiEmailSetupMaster(){

  var phArr = phFlexiSheet.getRange('A2:K'+phFlexiSheet.getLastRow()).getValues();

  for (var i = 0 ; i < phArr.length ; i++){

    var opsData = setupFlexiOpsData(phArr[i]);

    var htmlTable = setupOPPHtmlTable(phArr[i])

    var attdData = setupFlexiAttdData(phArr[i])

    var rawAttd = setupFlexiRawAttd(phArr[i])

    summaryTab.clear();

    rawAttdTab.clear();

    summaryTab.getRange(1,1,opsData.length, opsData[0].length).setValues(opsData);

    summaryTab.getRange(summaryTab.getLastRow()+2,1).setValue('Daily Usage Summary');

    summaryTab.getRange(summaryTab.getLastRow()+1,1,attdData.length, attdData[0].length).setValues(attdData);

    rawAttdTab.getRange(1,1,rawAttd.length, rawAttd[0].length).setValues(rawAttd);

    SpreadsheetApp.flush();

    sendEmail(htmlTable,phArr[i][1]);

  }

}

//setups Html table for flexi teams
function setupFlexiHtmlTable(array){

  var poopArr = array;
  var writeArr = [['Date','Company', 'Buffer', 'OPP Pax','Attendance Total','Charge']];

  var compName = poopArr[0];
  var email = poopArr[1];
  var buffer = poopArr[2];
  var pax = poopArr[3];
  var dpUsage = poopArr[8];
  var charge= poopArr[10];

  writeArr.push([today.toLocaleDateString('en-MY'),compName,buffer,pax,dpUsage,charge]);

  return writeArr;
}

//setup OPS data for flexi teams
function setupFlexiOpsData(array){

  var poopArr = array;

  var writeArr = [['Date','Company', 'Email', 'Buffer', 'OPP Pax', 'DP Quota', 'DP Usage Total' , 'DP Balance' , 'Credit Quota ', 'Credit Usage Total', 'Credit Balance' , 'Charge']];

  var compName = poopArr[0];
  var email = poopArr[1];
  var buffer = poopArr[2];
  var pax = poopArr[3];
  var dpQuota = poopArr[4];
  var dpUsage = poopArr[8];
  var dpBalance = poopArr[6];
  var creditQuota = poopArr[5];
  var creditUsage = poopArr[9];
  var creditBalance = poopArr[7];
  var charge= poopArr[10];

  writeArr.push([today.toLocaleDateString('en-MY'),compName,email,buffer,pax,dpQuota,dpUsage,dpBalance,creditQuota,creditUsage,creditBalance,charge]);

  return writeArr;
}

// setup attendance data for Flexi Teams
function setupFlexiAttdData(array){

  var poopArr = array;

  var infoArrRaw = dailyCalcFlexiSheet.getRange('B2:K'+dailyCalcFlexiSheet.getLastRow()).getValues();

  const infoArr =  infoArrRaw.filter(function (x) { 
    return !(x.every(element => element === (undefined || null || '')))
  });

  var dateRawArr = dailyCalcFlexiSheet.getRange('A2:A'+dailyCalcFlexiSheet.getLastRow()).getValues();

  const dateArr =  dateRawArr.filter(function (x) { 
    return !(x.every(element => element === (undefined || null || '')))
  });

  var tempArr = [];

  var writeArr = [['Date','Buffer','Team Member Access','Excess']];

  for(var i = 0 ; i < dateArr.length ; i++){

    writeArr.push(dateArr[i]);

  }

  for(var i = 0 ; i < infoArr.length ; i++){

    if(infoArr[i][0] === poopArr[0]){
      tempArr.push([infoArr[i][1],infoArr[i][3],infoArr[i][4]]);
    }
  }

  for(var i = 1 ; i < writeArr.length ; i++){

    writeArr[i].push(tempArr[i-1][0],tempArr[i-1][1],tempArr[i-1][2],);
  }

  return writeArr;
}

//setup raw attendance data for flexi teams
function setupFlexiRawAttd(array){

  var poopArr = array;

  var monthlyAttdArr = monthlyAttendanceSheet.getRange('A2:G'+monthlyAttendanceSheet.getLastRow()).getValues();

  var writeArr = [['Date','Card No','Staff Name']];

  for(var i = 0 ; i < monthlyAttdArr.length ; i++){

    if(poopArr[0].includes(monthlyAttdArr[i][5])){

      writeArr.push([monthlyAttdArr[i][0],monthlyAttdArr[i][1],monthlyAttdArr[i][2]]);

    }
  }

  return writeArr;
}




