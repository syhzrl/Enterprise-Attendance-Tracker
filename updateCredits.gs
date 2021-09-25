var today = new Date();

// As the function  name implies..it sums the value of object by their keys
function sumObjectsByKey(...objs) {
  return objs.reduce((a, b) => {
    for (let k in b) {
      if (b.hasOwnProperty(k)){
        a[k] = (a[k] || 0) + b[k];
      }
    }
    return a;
  }, {});
}

// extract company credits from ORnD
function pullCompanyCredits(array) {

  var companyIdArr = array;
  var monthlyCredit = [];
  var yesterdayDate = new Date();
  yesterdayDate.setDate(today.getDate() - 1);
  var yesterday = yesterdayDate.getDate()
 
  for (var i = 0 ; i < companyIdArr.length ; i++){

    var post_response_credits = UrlFetchApp.fetch('https://app.officernd.com/api/v1/organizations/worq/credit-accounts/stats?team='+  companyIdArr[i][1]+'&month=2021-06-'+ yesterday +'T00:00:00.000Z', orndApi());
    var credits_json = JSON.parse(post_response_credits);
    var data_credits = credits_json;
    var sum ;
 
    if(data_credits.length > 1){
      for (var j = 0 ; j < data_credits.length ; j++){
        if(j+1 < data_credits.length){
          sum = sumObjectsByKey(data_credits[j].monthly,data_credits[j+1].monthly);
          monthlyCredit.push(sum);
        }
 
      }

    } else {
      monthlyCredit.push(data_credits[0].monthly);
    }
  }
  
  var creditArr = [];

  for(var i = 0 ; i < monthlyCredit.length ; i ++){
    creditArr.push(Object.values(monthlyCredit[i]));
  }

  return creditArr;
}

//updates the Flexi Tab
function updateFlexi(){

  var flexiArr = filterFlexi();
  var creditArr = pullCompanyCredits(flexiArr);
  var tempArr = [];

  for(var i = 0 ; i < flexiArr.length ; i++){

    var int = flexiArr[i][4].replace(/[^\d]/g, '');
    var num = parseInt(int[0]);
    var buffer = Math.ceil(num = num + (num * 0.2));
    var oppPlan = flexiArr[i][6];
    var num = oppPlan.replace(/[^\d]/g, ''); 
    var pax ;

    if (num != ''){
      pax = num;
    } else {
      pax = 1;
    }

    var quota ;

    if(oppPlan.includes('Ultra-lite')){
      quota = 3 * pax;
    }

    else if(oppPlan.includes('Low')){
      quota = 5 * pax;
    }

    else if(oppPlan.includes('Medium')){
      quota = 8 * pax;
    }

    else if(oppPlan.includes('High')){
      quota = 10 * pax;
    }

    tempArr[i] = [flexiArr[i][0],flexiArr[i][2],flexiArr[i][3],flexiArr[i][4],flexiArr[i][5],flexiArr[i][6],flexiArr[i][7], buffer , pax , quota];
  }

  for(var i = 0 ; i < creditArr.length ; i++){
    tempArr[i].push(creditArr[i][0], creditArr[i][1] ,creditArr[i][2])
  }

  flexiSheet.getRange('A2:M').clear();
  flexiSheet.getRange(2 , 1 , tempArr.length , tempArr[0].length).setValues(tempArr);
}

//updates the Flexi Tab
function updateStandalone(){

  var oppArr = filterStandaloneOPP();
  var creditArr = pullCompanyCredits(oppArr);
  var tempArr = [];

  for(var i = 0 ; i < oppArr.length ; i++){

    var oppPlan = oppArr[i][6];

    var num = oppPlan.replace(/[^\d]/g, ''); 

    var pax ;

    if (num != ''){
      pax = num;
    } else {
      pax = 1;
    }

    var quota ;

    if(oppPlan.includes('Ultra-lite')){
      quota = 3 * pax;
    }

    else if(oppPlan.includes('Low')){
      quota = 5 * pax;
    }

    else if(oppPlan.includes('Medium')){
      quota = 8 * pax;
    }

    else if(oppPlan.includes('High')){
      quota = 10 * pax;
    }

    tempArr[i] = [oppArr[i][0],oppArr[i][2],oppArr[i][4],oppArr[i][6],oppArr[i][7],pax,quota];
  }

  for(var i = 0 ; i < creditArr.length ; i++){
    tempArr[i].push(creditArr[i][0], creditArr[i][1] ,creditArr[i][2])
  }

  standaloneOPPSheet.getRange('A2:K').clear();
  standaloneOPPSheet.getRange(2 , 1 , tempArr.length , tempArr[0].length).setValues(tempArr);
}



