var opsSheet = ss.getSheetByName('Ops Charge Sheet');


//update the OPS Charge Sheet based on other tabs mentioned in the function below
function updateOpsSheet() {

  var calculatedFlexiArr = phFlexiSheet.getRange('A2:K'+phFlexiSheet.getLastRow()).getValues();

  var calculatedStandaloneArr = phOPPSheet.getRange('A2:J'+phOPPSheet.getLastRow()).getValues();

  opsSheet.clear();

  opsSheet.getRange(1,1).setValue('Flexi').setFontWeight('bold');

  var headers = [['Date','Company', 'Email', 'Buffer', 'OPP Pax', 'DP Quota', 'Attendance Total' , 'DP Balance' , 'Credit Quota ', 'Credit Usage Total', 'Credit Balance' , 'Charge']];

  opsSheet.getRange(opsSheet.getLastRow() + 1,1, headers.length , headers[0].length).setValues(headers).setFontWeight('bold');

  opsSheet.getRange(opsSheet.getLastRow() + 1,1).setValue(new Date().toLocaleDateString('en-MY'));

  var writeFlexiArr = [];

  var compName;
  var email;
  var buffer;
  var pax;
  var dpQuota;
  var dpUsage;
  var dpBalance;
  var creditQuota;
  var creditUsage;
  var creditBalance;
  var charge;

  for(var i = 0 ; i < calculatedFlexiArr.length ; i++){

    compName = calculatedFlexiArr[i][0];
    email = calculatedFlexiArr[i][1];
    buffer = calculatedFlexiArr[i][2];
    pax = calculatedFlexiArr[i][3];
    dpQuota = calculatedFlexiArr[i][4];
    dpUsage = calculatedFlexiArr[i][8];
    dpBalance = calculatedFlexiArr[i][6];
    creditQuota = calculatedFlexiArr[i][5];
    creditUsage = calculatedFlexiArr[i][9];
    creditBalance = calculatedFlexiArr[i][7];
    charge= calculatedFlexiArr[i][10];

    writeFlexiArr.push([compName,email,buffer,pax,dpQuota,dpUsage,dpBalance,creditQuota,creditUsage,creditBalance,charge]);
  }

  opsSheet.getRange(opsSheet.getLastRow(),2,writeFlexiArr.length, writeFlexiArr[0].length).setValues(writeFlexiArr);

////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////


  opsSheet.getRange(opsSheet.getLastRow()+2,1).setValue('Standalone OPP').setFontWeight('bold');

  var headersOpp = [['Date','Company','Member','Email','DP Quota','DP Usage Total','DP Balance','Credit Quota ','Credit Usage Total','Credit Balance','Charge']];

  opsSheet.getRange(opsSheet.getLastRow() + 1,1, headersOpp.length , headersOpp[0].length).setValues(headersOpp).setFontWeight('bold');

  opsSheet.getRange(opsSheet.getLastRow() + 1,1).setValue(new Date().toLocaleDateString('en-MY'));

  var writeOPPArr = [];

  var oppCompName;
  var oppMemberName;
  var oppEmail;
  var oppDpQuota;
  var oppDpUsage;
  var oppDpBalance;
  var oppCreditQuota;
  var oppCreditUsage;
  var oppCreditBalance;
  var oppCharge;

  for(var i = 0 ; i < calculatedStandaloneArr.length ; i++){

    oppCompName = calculatedStandaloneArr[i][0];
    oppMemberName = calculatedStandaloneArr[i][1];
    oppEmail = calculatedStandaloneArr[i][2];
    oppDpQuota = calculatedStandaloneArr[i][3];
    oppDpUsage = calculatedStandaloneArr[i][7];
    oppDpBalance = calculatedStandaloneArr[i][5];
    oppCreditQuota = calculatedStandaloneArr[i][4];
    oppCreditUsage = calculatedStandaloneArr[i][8];
    oppCreditBalance = calculatedStandaloneArr[i][6];
    oppCharge = calculatedStandaloneArr[i][9];

    writeOPPArr.push([oppCompName,oppMemberName,oppEmail,oppDpQuota,oppDpUsage,oppDpBalance,oppCreditQuota,oppCreditUsage,oppCreditBalance,oppCharge]);
  }

  opsSheet.getRange(opsSheet.getLastRow(),2,writeOPPArr.length, writeOPPArr[0].length).setValues(writeOPPArr);


}

