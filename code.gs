
function onOpen() {
  var  ui=SpreadsheetApp.getUi();
  ui.createMenu('Planner').addItem('Generate Dates', 'generateDates').addToUi();
}

function generateDates() {
  var sprintSheet = SpreadsheetApp.getActive().getSheetByName("Sprint");
  var startDate = new Date(sprintSheet.getRange('B1').getValue());
  var endDate = new Date(sprintSheet.getRange('B2').getValue());
  var currentDate = new Date(startDate);
  var effectiveWorkingHours = sprintSheet.getRange('E1').getValue();
  var maxWorkingHours = sprintSheet.getRange('E2').getValue();
  
  
  
  var dates = [];
  while ( currentDate <=endDate)
  {
    dates.push(new Date(currentDate));
    currentDate.setDate(currentDate.getDate()+1);
  }
  Logger.log("Total of %s days to be created",dates.length);
  
  var teamComposition = sprintSheet.getRange('B3:4').getValues();
  Logger.log(teamComposition);
  
  var capacityTable = new Array(teamComposition.length +1);
  capacityTable[0] = ['Capacity'];
  for (var i=0;i<teamComposition[0].length;i++) {
    if (teamComposition[0][i]=='') {
       break; 
    }
    capacityTable[i+1] = [teamComposition[0][i]];
  }
  
  for (var j=0; j<dates.length;j++) {
    capacityTable[0].push(dates[j].toLocaleDateString("en-US"));
    i=1
    for (;i<capacityTable.length;i++) {
      capacityTable[i].push(getWorkingHours(teamComposition[1][i-1],dates[j], effectiveWorkingHours))
    }
  }
  var capacitySheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Capacity");
  if (capacitySheet != null) {
     capacitySheet.clear(); 
  } else {
      capacitySheet = SpreadsheetApp.getActiveSpreadsheet().insertSheet("Capacity");
  }
  
  
  generateCapacityTable(capacitySheet, capacityTable);

  var taskDetails = sprintSheet.getRange('A8:C'+sprintSheet.getLastRow()).getValues();
  createTasks(taskDetails,capacitySheet, dates);
  resizeColumns(capacitySheet, capacityTable);
}

function createTasks(taskDetails,sheet,dates) {
  
  var taskCurrentRow = sheet.getLastRow();
  for (var i=0;i<taskDetails.length;i++) {
    Logger.log("Creating Bar for "+taskDetails[i][0]);
    var taskBlockStart = -1;
    var taskBlockLength =-1;
    if (dates[0]>taskDetails[i][1]) {
      Logger.log("Start date is before the begining of Sprint: "+taskDetails[i][1]+" - sprint begins: "+dates[0]);
        taskBlockStart = 0;
    }
    for (j=0;j<dates.length; j++) { 
      if (dates[j].toLocaleDateString("en-US") == taskDetails[i][1].toLocaleDateString("en-US")) {
        Logger.log("Matched Start date: "+dates[j]+" - "+taskDetails[i][1]);
        taskBlockStart = j;
      }else if (dates[j].toLocaleDateString("en-US")==taskDetails[i][2].toLocaleDateString("en-US")) {
        Logger.log("Matched End date: "+dates[j]+" - "+taskDetails[i][2]);
        taskBlockLength = j-(taskBlockStart>0?taskBlockStart:0)+1;
        break;
      } else if (dates[j]>taskDetails[i][2]) {
        taskBlockLength = j-1;
        Logger.log("No Matching End date: "+dates[j]+" - "+taskDetails[i][2]);
        break;
      }
    }

    if (taskBlockStart>-1 && taskBlockLength==-1 && dates[j-1]<taskDetails[i][2]) {
      Logger.log("No end date found hence marking complete row ");
      taskBlockLength = dates.length-taskBlockStart;
    }
    Logger.log("Task Start: "+taskBlockStart+" length: "+taskBlockLength);
    if (taskBlockLength >0) {
      if(taskBlockStart<0)
        taskBlockStart =0;
      sheet.appendRow([taskDetails[i][0]]);
      taskBlockCells = sheet.getRange(taskCurrentRow+1, taskBlockStart+2,1,taskBlockLength);
      taskBlockCells.setBackground("#8ECE00");
      taskCurrentRow++;
    }
  }
}
 
function resizeColumns(sheet, capacityTable) {
  for (var i=1;i<=capacityTable[0].length;i++)
  {
    sheet.setColumnWidth(i, 30);
  }
  sheet.autoResizeColumn(1);
}
function generateCapacityTable(sheet, capacityTable ) {
  var rowSumFormula = "=SUM(R[0]C[-"+(capacityTable[0].length-1)+"]:R[0]C[-1])"
  var capRow =0;
  for (;capRow<capacityTable.length;capRow++) {
    sheet.appendRow(capacityTable[capRow]);
    if (capRow!=0) {
      sheet.getRange(capRow+1,capacityTable[capRow].length+1).setFormulaR1C1(rowSumFormula);
    }
  }
  sheet.appendRow(['Total Capacity']);
  
  var totalCapacityEquation = "=SUM(";
  for (var i=1; i<capacityTable.length;i++)
  {
    if (i>1)
      totalCapacityEquation += ",";
    totalCapacityEquation += "R[-"+i+"]C["+(capacityTable[0].length-1)+"]";
  }
  totalCapacityEquation += ")";
  sheet.getRange(i+1,2).setFormulaR1C1(totalCapacityEquation);
}

function getWorkingHours(location, date, effectiveHours) {
  var weekday = true;
  if (location=='Dubai') {
    if (date.getDay()>4) {
      weekday = false;
    }
  } else {
    if (date.getDay()==0 || date.getDay()==6) {
      weekday = false;
    }
  }
  
  if (weekday) {
    return effectiveHours;
  } else {
    return 0;
  }
}
