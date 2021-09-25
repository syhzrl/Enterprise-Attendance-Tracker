var flexiSheet = ss.getSheetByName('ORND Flexi');
var standaloneOPPSheet = ss.getSheetByName('ORND Standalone OPP');

// filters raw ORnD data to extract WFH members
function filterWFH(){

  var membershipArr = membershipSheet.getRange('A2:I'+membershipSheet.getLastRow()).getValues();
  var wfhArr = [];

  for(var i = 0 ; i < membershipArr.length ; i++){

    if(membershipArr[i][5].includes('WFH') && membershipArr[i][8] != 'Expired'){
      wfhArr.push(membershipArr[i]);
    }
  }

  return wfhArr;
}


// filter Raw ORnD data to extract Private Suite Members
function filterPS(){

  var membershipArr = membershipSheet.getRange('A2:I'+membershipSheet.getLastRow()).getValues();
  var psArr = [];

  for(var i = 0 ; i < membershipArr.length ; i++){

    if(membershipArr[i][5].includes('Private Suite') && membershipArr[i][8] != 'Expired' && membershipArr[i][6].includes('-')){
      psArr.push(membershipArr[i]);
    }
  }

  return psArr;
}

// consolidate WFH and PS data to extract Flexi Teams
function filterFlexi() {

  var wfhArr = filterWFH();
  var psArr = filterPS();
  var flexiArr = [];

  for(var i = 0 ; i < wfhArr.length ; i++){

    for(var j = 0 ; j < psArr.length ; j++){

      if(wfhArr[i][0] == psArr[j][0]){

        flexiArr.push(psArr[j].concat(wfhArr[i]));
      }
    }
  }

  for(var i = 0 ; i < flexiArr.length ; i++){

    flexiArr[i].splice(3,1);
    flexiArr[i].splice(4,1);
    flexiArr[i].splice(6,7);
    flexiArr[i].splice(8,1);
  }

  return flexiArr;
}

// consolidate WFH data to extract Standalone OPPs
function filterStandaloneOPP(){

  var wfhArr = filterWFH();
  var flexiArr = filterFlexi();
  
  for(var i = 0 ; i < wfhArr.length ; i++){

    for(var j = 0 ; j < flexiArr.length ; j++){

      if(wfhArr[i][0] == flexiArr[j][0]){

        wfhArr.splice(i,1);

      } 
    }
  }

  return wfhArr;
}


