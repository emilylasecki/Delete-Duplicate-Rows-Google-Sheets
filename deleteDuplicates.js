//In the google sheet referrals document, check for duplicate rows in the last 4 days by searching for matching student emails. If two or more student emails match, delete the oldest one(s).


//extract the date from the google doc and search for the first row within the range (the within the last 4 days)
function extractRecentActivity() {

    var spreadsheetId = "1Bxpd4neH4HqJ1bRKU3Hp7sq7gc3JBZ12xS7USoFogUc";                                  //linked to sample document avaiable in GitHub README
    var sheet = SpreadsheetApp.openById(spreadsheetId).getActiveSheet();
    var number = sheet.getMaxRows().toString();
    var number = number.replace(".0","");
    
    var timestamp = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), "yyyy-MM-dd");         //ignore hours, mins, sec
    
    const dateObject = new Date(timestamp);                                                              //create object with todays date, ignoring hours, mins, secs
    dateObject2 = JSON.stringify(dateObject);                                                            //convert to string so dates can be compared
    
    var valuesOG = SpreadsheetApp.openById(spreadsheetId).getActiveSheet().getRange("A1:A").getValues(); //get values from the sheet and store in an object
    
    //Convert object to array to update indexing to match the Google Sheet
    var values=[];
    for (var i=0; i<=number; i++) {
      values[i+1] = valuesOG[i];
    }   
    
    
    //Calculate 4 previous days 
    var prevDays1 = new Date(timestamp);
    prevDays1.setDate(prevDays1.getDate() - 1);
    day1 = JSON.stringify(prevDays1)
    
    var prevDays2 = new Date(timestamp);
    prevDays2.setDate(prevDays2.getDate() - 2);
    day2 = JSON.stringify(prevDays2)
    
    var prevDays3 = new Date(timestamp);
    prevDays3.setDate(prevDays3.getDate() - 3);
    day3 = JSON.stringify(prevDays3)
    
    var prevDays4 = new Date(timestamp);
    prevDays4.setDate(prevDays4.getDate() - 4);
    day4 = JSON.stringify(prevDays4)
    
    //Because of the timezone configuration, this is a safety net to calc a day into the future too
    var thisDay = new Date(timestamp);
    thisDay.setDate(thisDay.getDate() +1);
    day5 = JSON.stringify(thisDay)
    
    Logger.log("Searching excel file for the following dates:");
    Logger.log(thisDay);
    Logger.log(dateObject); 
    Logger.log(prevDays1);
    Logger.log(prevDays2);
    Logger.log(prevDays3);
    //Logger.log(prevDays4);
    
    for (i=2; i<number; i++) { 
      newName = JSON.stringify(values[i]);
      newName2 = newName.slice(2,12);                                                                    //extract only yyyy-MM-dd to compare with other timestamps
      const backToObject = new Date(newName2);                                                           //turn it into an object again to add on default time
      finalConversion = JSON.stringify(backToObject);                                                    //turn back into a string for comparison
      values[i] =finalConversion; 
        if (values[i] == day5 | values[i] == dateObject2 | values[i] == day1 | values[i] == day2 | values[i] == day3 | values[i] == day4 ) {  
          Logger.log("The first date within the range was found at: " + i);                              //extract index where first match is found
          Logger.log("Searching for duplicates ...");
          checkForDuplicates(i, number);                                                           
          break;
        }
        if (i==number-1) {
          Logger.log("No rows have been added in the last four days.");
        }
      continue;
    }
    Logger.log("Finished processing.");
    }
    
    
    
    //taking the range of values to check, extract student ids and check for duplicates
    function checkForDuplicates(startIndex, endIndex) {
    
    var ender = Number(endIndex) +1;                                                                      //convert endIndex to a number and add 1 to compensate for array bound
    
    var spreadsheetId = "1Bxpd4neH4HqJ1bRKU3Hp7sq7gc3JBZ12xS7USoFogUc";
    var sheet = SpreadsheetApp.openById(spreadsheetId).getActiveSheet();
    
    var valuesOG = SpreadsheetApp.openById(spreadsheetId).getActiveSheet().getRange("B1:B").getValues();
    
    var checkThis = [];
    var values=[];
    
    for (var i=0; i<=ender; i++) {
      values[i+1] = valuesOG[i];
    } 
    
    //convert all vals in specified range from original object to a string and put into an array
    for (var i=startIndex; i<ender; i++) {
      newVal = JSON.stringify(values[i])
      newVal2 = newVal.slice(2,12);
      checkThis[i] = newVal2;
    } 
    
    //mark duplicate students as null as indicator to delete them
    for (var j = endIndex; j>=startIndex; j--) { 
    Number(j);
      for (var k = j-1; k>=startIndex; k--) {
        if (checkThis[j] == checkThis[k]) {
          checkThis[k] = null;  
        }
      }
    
    
    }
    
    //deletes the elements marked null
    Number(endIndex);
    var counter=0;
    for (i=endIndex; i>=startIndex; i--) {
      if (checkThis[i] == null) {
        sheet.deleteRow(i);
        counter++;
      }
    }
    Logger.log("Found and deleted " + counter + " duplicate rows.");
    }