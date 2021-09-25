var firstDay = new Date(today.getFullYear(), today.getMonth(), 1).getDate();
var lastDay = new Date(today.getFullYear(), today.getMonth() + 1, 0).getDate();

//Master Scripts that control every single functions

function masterStart() { // this will run 1st of every month

  if(today.getDate() === firstDay){

    writeOrndData();
    updateFlexi();
    updateStandalone();
    updatePlaceholderFlexiStart();
    updatePlaceholderStandaloneStart();

    monthlyAttendanceSheet.getRange('A2:G').clear();

    dailyCalcFlexiSheet.getRange('A2:K').clear();

    dailyCalcStandaloneSheet.getRange('A2:J').clear();

  }
}

function masterDaily(){ // this will run daily

  importCSV();
  cleanDailyAttendance();
  updateMonthlyAttendance();
  dailyCalcFlexi();
  dailyCalcStandalone();

}

function masterEnd(){ // this will run at the last day of the month

  if(today.getDate() === lastDay){

    updatePlaceholderFlexiEnd();
    updatePlaceholderStandaloneEnd();
    monthlyCalculationFlexi();
    monthlyCalculationStandalone();
    updateOpsSheet();
    flexiEmailSetupMaster();
    oppEmailSetupMaster();
    attendanceArchive();

  }
}

