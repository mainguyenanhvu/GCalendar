// Array index starts from 0

//https://stackoverflow.com/questions/10087819/convert-date-to-another-timezone-in-javascript
function convertTZ(date, tzString) {
    return new Date((typeof date == "string" ? new Date(date) : date).toLocaleString("en-US", {timeZone: tzString}));   
}

function calculate_start_time(flight_time){
  console.log(flight_time.getHours())
  var start_time = new Date(flight_time);
  start_time.setHours(start_time.getHours() - 30);
  console.log(start_time.getHours())
  hour = start_time.getHours()
  if (hour > 22) 
  {
    start_time.setHours(start_time.getHours() - 2)
  }
  else if (hour < 6)
  {
    start_time.setHours(start_time.getHours() - start_time.getHours() - 3)
  }
  console.log(start_time.getHours())
  return start_time
}

function add_reminder(row){
  if (row[3] == ''){
    return 'No'
  }
  console.log(row)
  var calendarId = 'bi907mkoibu2vdg30cf1hu81uc@group.calendar.google.com'; //'gbmbvemaybay@gmail.com';
  var calender = CalendarApp.getCalendarById(calendarId);
  var flight_time = new Date(row[3]); //convertTZ(row[3], calender.getTimeZone())
  var start_time = calculate_start_time(flight_time)
  var event = calender.createEvent(row[2], start_time, flight_time, {description: row[2] + ' ' + flight_time.toTimeString() + ' ' + row[10]})
  console.log(event.getDescription())
  return 'Yes'
}

function check_reminder_one_sheet(sheet){
  var startRow = 3;
  var numRows = sheet.getLastRow()
  var startColumn = 1;
  var numColumns = sheet.getLastColumn();
  var dataRange = sheet.getRange(startRow, startColumn, numRows, numColumns);
  var data = dataRange.getValues();
  for (var rowI in data){
    var row = data[rowI];
    if (row[12] != 'Added'){
      status = add_reminder(row)
      if (status == 'Yes'){
        var cell = sheet.getRange(Number(rowI)+Number(startRow),13,1,1);
        cell.setValue("Added");
      }
    }
  }
}

function check_all_sheets() {
  var sheets = SpreadsheetApp.getActive().getSheets();
  for (var i in sheets){
    var sheet_i = sheets[i];
    var sheet_name = sheet_i.getName();
    if (sheet_name.startsWith('T')){
      check_reminder_one_sheet(sheet_i);
    }
  }
}

function onOpen() {
  var ui = SpreadsheetApp.getUi();
  ui.createMenu("My Menu")
      .addItem("Create Reminders", "check_all_sheets")
      .addToUi();
}
