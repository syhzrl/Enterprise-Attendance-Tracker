var emailSheet = SpreadsheetApp.openById('17J14IS9FJdgxpZgWF9AO3TVWkKQ9ctfgGB1DvKYWuWo');
var summaryTab = emailSheet.getSheetByName('Summary');

//generates a html table with given array
function getHtmlTable(array){

  var data = array;

  var TABLEFORMAT = 'cellspacing="2" cellpadding="2" dir="ltr" border="1" style="width:100%;table-layout:fixed;font-size:10pt;  font-family:arial,sans,sans-serif;border-collapse:collapse;border:1px solid #ccc;font-weight:normal;color:black;background-color:white;text-align:center;text-decoration:none;font-style:normal;'

  var htmltable = '<table ' + TABLEFORMAT +' ">';

  for (var row = 0; row<data.length; row++){

    htmltable += '<tr>';

    for (var col = 0 ;col<data[row].length; col++){

      if (data[row][col] === "" || 0) {
        htmltable += '<td>' + 'None' + '</td>';
      } 

      else if (row === 0)  {
        htmltable += '<th>' + data[row][col] + '</th>';
      }

      else {
        htmltable += '<td>' + data[row][col] + '</td>';
      }
    }   

    htmltable += '</tr>';
  }

  htmltable += '</table>';
  
  return htmltable;
}

//returns the Email Sheet datas as a blob
function getEmailSheetBlob() {

  var exportType ="xlsx";

  var params = {method:"GET",headers:{"authorization":"Bearer "+ ScriptApp.getOAuthToken()}};

  var url="https://docs.google.com/spreadsheets/d/17J14IS9FJdgxpZgWF9AO3TVWkKQ9ctfgGB1DvKYWuWo/export?format="+exportType;
          
  var fetch=UrlFetchApp.fetch(url,params);

  var blob = fetch.getBlob(); 

  return blob;
}

//send email to respective customer emails
function sendEmail(array, email) {

  var table = getHtmlTable(array);

  var htmlBody = HtmlService.createHtmlOutputFromFile('Email Body').getContent();

  var emailBody = htmlBody.replace('{table}',table);

  var blob = getEmailSheetBlob();

  MailApp.sendEmail(email,'TEST','body',{htmlBody:emailBody, attachments: [{fileName: 'filename' + ".xlsx",content: blob.getBytes(),mimeType: 'xlsx'}]});
  
}

