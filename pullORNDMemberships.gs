var membershipSheet = ss.getSheetByName('ORND Membership Raw');


// Standard API call for OfficeRnD
function orndApi () {
  var postPayload = {
    "client_id": "HD9V1j6sI1aLieZ5",
    "client_secret": "iKeTMUaVXew1EwqQQSSRb2AC6vbehWQY",
    "grant_type": "client_credentials",
    "scope": "officernd.api.read"
  };

  var options = {
    "method" : "post",
    "payload" : postPayload,
  };
  
  var pre_response = UrlFetchApp.fetch('https://identity.officernd.com/oauth/token', options);
  var pre_response_json = JSON.parse(pre_response.getContentText());
  var access_token = pre_response_json.access_token; // return the access token 

  var headers = {
      "Authorization" : "Bearer " + access_token
  };
  
  var params = {
    "method": "get",
    "headers": headers
  };

  return params
}

// Imports ORnD data
function importOrndData(link){

  var post_response = UrlFetchApp.fetch(link, orndApi());
  var data_json = JSON.parse(post_response);
  var data = data_json;

  return data
}

// Change key names to uniformized the arrays of objects (key value pairs)
function changeKeyName(array,keysArr){

  let clone = JSON.parse(JSON.stringify(array));

  for(var i = 0 ; i < keysArr[0].length ; i++){
    var res = clone.map((elem) => {
      elem[keysArr[1][i]] = elem[keysArr[0][i]],
      delete elem[keysArr[0][i]]
      return elem
    });
  }

  return res

}

// Generates Ornd Array (key value pairs)
function generateOrndArray(link, keys){

  var data= importOrndData(link);

  var dataKeys = keys;

  var tempData = changeKeyName(data,dataKeys);

  return tempData
}


// Merging arrays of object according to their keys
function mergeOrndData(a1 , a2 , key){

    var res = a1.map(itm => ({
        ...a2.find((item) => (item[key] === itm[key]) && item),
        ...itm
    }));

    return res
}

// Import and process ORnD data into ORND Membership Raw Tab
function processOrndData(){
  var dataTeamKeys = [['name','_id','email'],['pic','idTeam','emailTeam']];
  
  var dataTeam = generateOrndArray('https://app.officernd.com/api/v1/organizations/worq/teams',dataTeamKeys).map(function(obj) {
    return {pic: obj.pic, idTeam: obj.idTeam , emailTeam: obj.emailTeam};
  });

  var dataMembershipKeys = [['team','plan','member','price','name'],['idTeam','idPlan','idMember','priceMembership','membershipName']];

  var dataMemberships = generateOrndArray('https://app.officernd.com/api/v1/organizations/worq/memberships',dataMembershipKeys)
  .map (function(obj) {
    return {
      idTeam: obj.idTeam, 
      idPlan: obj.idPlan, 
      idMember: obj.idMember,
      priceMembership: obj.priceMembership,
      startDate: obj.startDate,
      endDate: obj.endDate, 
      calculatedStatus: obj.calculatedStatus,
      membershipName: obj.membershipName};
  });

  var dataMemberKeys = [['team','name','_id'],['idTeam','memberName','idMember']];

  var dataMembers = generateOrndArray('https://app.officernd.com/api/v1/organizations/worq/members',dataMemberKeys)
  .map (function(obj) {
    return {
      idTeam: obj.idTeam, 
      memberName: obj.memberName, 
      idMember: obj.idMember};
  });

  var dataPlanKeys =[['_id','name'],['idPlan','planName']];
  
  var dataPlans = generateOrndArray('https://app.officernd.com/api/v1/organizations/worq/plans',dataPlanKeys)
  .map (function(obj) {
    return {
      idPlan: obj.idPlan, 
      planName: obj.planName,
      category: obj.category};
  });

  var mergedData = mergeOrndData(dataMemberships , dataMembers , 'idMember');
  mergedData = mergeOrndData(mergedData, dataTeam , 'idTeam');
  mergedData = mergeOrndData(mergedData, dataPlans , 'idPlan');

  var today = new Date();

  for (var x = 0 ; x < mergedData.length ; x++) { //This for loop is only for formatting the table

    var tempDate = new Date (mergedData[x].endDate);

    var diffInMs = tempDate.getTime() - today.getTime();

    var diff = Math.round(diffInMs / (1000 * 3600 * 24));

    if(mergedData[x].endDate != null){

      if( diff <= 30){

        if (tempDate < today){
          mergedData[x].calculatedStatus = "Expired";
        } 
        else {
          mergedData[x].calculatedStatus = "Expiring";
        }
      }

      if( diff > 30){
      mergedData[x].calculatedStatus = "Not Expiring";
      } 

    } 
    else {
      mergedData[x].calculatedStatus = "Not Expiring";
    }
  }

  return mergedData
}


//Write the data to ORND Membership Raw tab
function writeOrndData(){
  var headings = ['pic','idTeam', 'memberName', 'idMember', 'emailTeam', 'category',	'membershipName',	'priceMembership','calculatedStatus'];

  var outputRows = [];

  processOrndData().forEach(function(i) {
    outputRows.push(headings.map(function(heading) {
      return i[heading] || '';
    }));
  });

  membershipSheet.getRange('A2:F').clear();

  if (outputRows.length) {
    membershipSheet.getRange(2, 1, outputRows.length, outputRows[0].length).setValues(outputRows);
  }
}

