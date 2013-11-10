function onOpen() {
  var ss = SpreadsheetApp.getActive();
  var items = [
    {name: 'Create weekly form', functionName: 'createWeeklyForm'},
    null, // Results in a line separator.
    {name: 'Create Final sheet', functionName: 'createFinalSheet'}
  ];
  ss.addMenu('Custom Menu', items);
}

function createWeeklyForm() {
  var ss = SpreadsheetApp.getActive();
  setupForm(ss); 
}

function setupForm(ss) {
  //var tempss = SpreadsheetApp.create('TEMP_STATUS_SHEET');
  var url = ss.getFormUrl();
    
  if(url) {
    var form = FormApp.openByUrl(url);
    // May change the form fields etc here
    Browser.msgBox('Existing Weekly form found');
  } 
  else {
    var form = FormApp.create('Weekly Status Form');
    form.setDestination(FormApp.DestinationType.SPREADSHEET, tempss.getId());
    form.addTextItem().setTitle('Name').setRequired(true);
    form.addTextItem().setTitle('Email').setRequired(true);
    form.addTextItem().setTitle('Team').setRequired(true);
    form.addTextItem().setTitle('Last week').setRequired(true);
    form.addTextItem().setTitle('This week').setRequired(true);
    Browser.msgBox('Weekly form set up');
  }    
}

function createOrGetSheet(ss, sname, ctr) {
  var sheet = ss.getSheetByName(sname);
  if (sheet == null) {
    sheet = ss.insertSheet(sname, ctr);
  }
  return sheet;
}

function createOrGetRowForUser(name) {
  // TODO : what if the user name is not there?
  var rownums = ScriptProperties.getProperties();
  return rownums[name];
}

function createFinalSheet() {
  var ss = SpreadsheetApp.getActive();
  var tempss = SpreadsheetApp.openByUrl("https://docs.google.com/spreadsheet/ccc?key=0Ak9HC9UFjhgydE50ZmZJRnoxQmlWRHhLVTZBNGRMenc&usp=drive_web#gid=0");
  var formResponseSheet = tempss.getSheets()[1];
  
  // sort sheet by team i.e 4th column
  formResponseSheet.sort(4);
  
  // Iterate over the entries in formResponseSheet and create sheet for each team if it doesn't already exist
  // If new sheet for team is created then add the entries - name, this week and last week
  // If, sheet already exists, then insert a colum after first column and add entries of this week and last week to it
  var lastRow = formResponseSheet.getLastRow();
  var lastColumn = formResponseSheet.getLastColumn();
  Logger.log(lastRow);
  Logger.log(lastColumn);
  var range = formResponseSheet.getRange(2, 2, lastRow - 1, lastColumn);
  var values = range.getValues();
  var teamCtr = 0;
  
  // Store the user name and their corresponding row number in script property
  ScriptProperties.setProperties({'Test1' : 2, 'Test2' : 2, 'Test3' : 2, 'Test4' : 3, 'Test5' : 3});
  
  var team = "";
  var sheet;
  var newColumn = 0;
  for (var row in values) {
    if (team != values[row][2]) {
      team = values[row][2];
      // Create or get sheet for each team
      sheet = createOrGetSheet(ss, team, teamCtr);
      // Get the column number where new data is to be added
      newColumn = sheet.getLastColumn() + 1;
      // Add header for the new column
      var headerCell = sheet.getRange(1, newColumn);
      var 
      headerCell.setValue();
      teamCtr++;
      Logger.log("Name of team %s", team); 
    }  
    
    // Get name, get its row number in the final sheet from map
    var name = values[row][0];
    var userRow = createOrGetRowForUser(name);
    Logger.log("Name %s : Row number %s", name, userRow);
    
    Logger.log("Data to be added to row %s column %s", userRow, newColumn);
    var userCell = sheet.getRange(userRow, newColumn);
    Logger.log("Setting value %s", values[row][3])
    userCell.setValue(values[row][3]);
  }
  Browser.msgBox('You clicked the second menu item!');
}

