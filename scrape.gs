/*
Instructions: 
#1 - Visit https://script.google.com to create a new project
#2 - Create a filter for your Grabpay statements on Gmail and find the label name from the search bar on Gmail.com (e.g label:banks-grabpay-statement)
     and copy the value after label: and put into var label.
#3 - Create a new Google Sheet and get the Sheet ID via the URL and replace the value in var sheet
#4 - Find the name of the Google Sheets Tab and replace into the val tab 
#5 - Run function extractEmails to start

You can create trigger to run this daily
*/


var label = "XXXXXXX";
var sheet = SpreadsheetApp.openById("XXXXXXXXXXXXXXXXXX");
var tab = sheet.getSheetByName("XXXXXX");  


// extract emails from label in Gmail
function extractEmails() {

  // get all email threads that match label from Sheet
  var threads = GmailApp.search ("label:" + label);
  
  // get all the messages for the current batch of threads
  var messages = GmailApp.getMessagesForThreads (threads);

  for each (var message in messages)
  {
    var data = extractTable(message[0]);
    insertIntoSheets(data);
  }
}

function extractTable(message){
  var rawMessage = message.getRawContent();
  
  var bodyRegex = new RegExp('<table style="border-collapse: collapse;">[\\s\\S]*</table>');
  var messageBody = bodyRegex.exec(rawMessage);
  
  var re = new RegExp('<p style="line-height: 15px; margin-bottom: 0px; color: #666666; font-family: \'Open Sans\', \'Helvetica Neue\', Helvetica, Arial, sans-serif; font-size: 12px; font-weight: normal; text-align: left; margin: 10px 0; padding: 0; mso-line-height-rule: exactly; -ms-text-size-adjust: 100%; -webkit-text-size-adjust: 100%;">(.*?)</p>', 'g');
  var matches;     
  
  var count = 0;
  var data = [];
  var temp = [];
  
  while(matches = re.exec(messageBody)) {
    if (matches.index === re.lastIndex)
      re.lastIndex++;
    
    if (count != 0 && count % 7 != 0){
      var val = matches[1];
      
      if(val == "GrabFood Payment"){
        temp = [];
        count = 0;
        continue;
      }
      
      if((count % 7 == 5 || count % 7 == 6) && val.length != 0)
        var val = val.substring(2, val.length);

      temp.push(val);
      count++;
    }
    else{
      count = 0;
      data.push(temp);
      temp = [matches[1]];
      count++;
    }
  }
  
  return data
}

function insertIntoSheets(data){
  for each(var row in data){
    if(row.length != 0){
      var dateTime = convertStringtoDateTime(row[0])
      row.shift();
      row = dateTime.concat(row);
    
      if(findInColumn("C",row[2]) == -1){
        tab.appendRow(row);
      }
    }
  }
}

function convertStringtoDateTime(data){
  var date = data.substring(0, 11);
  var time = data.substring(12, data.length);
  
  return [date, time];
}

function findInColumn(column, data) {
  var column = sheet.getRange(column + ":" + column);  // like A:A
  
  var values = column.getValues(); 
  var row = 0;
  
  while ( values[row] && values[row][0] !== data ) {
    row++;
  }
  
  if (values[row] != null && values[row][0] == data) 
    return row+1;
  else 
    return -1;
    
}
