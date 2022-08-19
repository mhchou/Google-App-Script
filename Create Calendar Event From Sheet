function onOpen() //add a clickable button in the sheet
{
  var ui = SpreadsheetApp.getUi();
  ui.createMenu('Sync to Calendar')
      .addItem('Update now', 'CreateCalendarEventFromSheet')
      .addToUi();
}

function CreateCalendarEventFromSheet() 
{
  var sheet = SpreadsheetApp.openById("<Id>").getSheets()[2]; // using the Automation sheet
  //Logger.log('The sheet is named "%s".', sheet.getName());

  var calendar = CalendarApp.getCalendarById('<Id>@group.calendar.google.com'); //user who runs the script must subscribe to the calendar
  //Logger.log('The calendar is named "%s".', calendar.getName());

  var range = sheet.getDataRange(); //get all the rows
  var data = range.getValues(); //pull all rows into an array
  var times = 1; //default number of working day
  
  Logger.log("Start with deleting old information before appending new ones.")

  var events = calendar.getEvents(data[0][0], new Date(data[data.length-1][0].getTime()+86400000)); //get the begging and end date in the range

  for(i in events) //delete everything in the range, start anew
    { 
      var ev = events[i];
      //Logger.log("deleting"+ev.getTitle());
      ev.deleteEvent();
    }

  for (var j = 6; j <= 9; j++) //loop through the 3 columns (H, I, J) with names plus the 1 row (G) with area
    {
      //Utilities.sleep(3000); // pause in the loop for 2000 milliseconds (2 seconds) to avoid the error "You have been creating or deleting too many calendars or calendar events in a short time. Please try again later."
      //Logger.log("pause");
      for (i in data) //loop through all the rows
      {

        if (data[parseInt(i)+1] != null) // from first to second to the last row (if it's not the last row)
        { 
          if (data[i][j] == data[parseInt(i)+1][j])  // if happens in multiple days in a row, accumulate the number of days
          {  
            times = times + 1;
            //Logger.log("["+data[i][j]+"] works "+times+" days in a row");
            //Logger.log("i is: "+i)
    
          }
          else //when the names changes, create the event with the number of days and reset the day count
          {
            var newEvent = calendar.createEvent(data[i][j], data[i-times+1][0], new Date(data[i][0].getTime()+ 86400000), {description:data[i][5]});
            newEvent.setColor(j-1);

            Logger.log("["+data[i][j]+"] works "+times+" day in a row from "+data[i-times+1][0]+" to "+data[i][0]+" and create the series here");
            times = 1; // reset reoccurrence    

            Utilities.sleep(1200); // pause in the loop for 2000 milliseconds (2 seconds) to avoid the error "You have been creating or deleting too many calendars or calendar events in a short time. Please try again later."
            //Logger.log("pause");  
          }  
        }
        else // when hit the last row
        {
          if (data[i][j] == data[parseInt(i)-1][j]) //if the last row is the same as the previous row, create the series 
          {
            var newEvent = calendar.createEvent(data[i][j], data[i-times+1][0], new Date(data[i][0].getTime()+ 86400000), {description:data[i][5]});
            newEvent.setColor(j-1)

            Logger.log("["+data[i][j]+"] finishes the on call schedule with "+times+" day and create the series here");
            times = 1; //reset reoccurrence
          }
          else //if the last row is not the same as the previous row, create its own even here
          {
            times = 1;
            var newEvent = calendar.createAllDayEvent(data[i][j], (data[i][0]), {description:data[i][5]});
            newEvent.setColor(j-1);          
            Logger.log("["+data[i][j]+"] works last "+times+" day of calendar on his own and create single event here");
            
          }
        }
      }
    }
}
