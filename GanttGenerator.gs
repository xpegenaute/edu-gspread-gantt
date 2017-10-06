var GANT_SHEET_NAME;
var COLUMNS_SHIFT;
var ROW_FOR_WEEK_NUMBERS;
var ROW_FOR_DATES;
var ROW_FOR_BODY;
var MS_A_DAY;
var WEEKEND_COLOR;
var HOLIDAY_COLOR;
var WORKING_COLOR;
var UF_COLOR;

var CELL_WIDTH;
var WEEKDAYS;
var WEEKEND;

var FIRST_DAY;
var LAST_DAY;


function initGanttGenerator() {
  GANT_SHEET_NAME = "Gantt";
  COLUMNS_SHIFT = 3;
  ROW_FOR_WEEK_NUMBERS = 1;
  ROW_FOR_DATES = 2;
  ROW_FOR_BODY = 3;
  MS_A_DAY = 1000 * 60 * 60 * 24;
  WEEKEND_COLOR = "Grey";
  HOLIDAY_COLOR = "Green";
  WORKING_COLOR = "Yellow";
  UF_COLOR = "Lime";

  CELL_WIDTH = 20;
  WEEKDAYS = ["sunday", "monday", "tuesday", "wednesday", "thursday", "friday", "saturday"];
  WEEKEND = [WEEKDAYS.indexOf("sunday"), WEEKDAYS.indexOf("saturday")];
  
  FIRST_DAY = moment.utc("2017-9-13", "YYYY-MM-DD").startOf("day");
  LAST_DAY = moment.utc("2018-4-17", "YYYY-MM-DD").startOf("day");  
}

// Warning month starts at 0 and the time is set to 9 to do not mess when changing time offse in summer and winter.
//var FIRST_DAY = moment.utc(2017, 8, 14).startOf("day");
//var LAST_DAY = moment.utc(2018, 6, 1).startOf("day");

// FIXME: Update the design to use some kind of visitor and draw it in one shot

function createTab() {
  var gantSheetName = GANT_SHEET_NAME;
  var gantSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(gantSheetName);
  if (gantSheet != null) {
    gantSheetName = [gantSheetName, SpreadsheetApp.getActiveSpreadsheet().getSheets().length].join("_");
  }
  gantSheet = SpreadsheetApp.getActiveSpreadsheet().insertSheet();
  gantSheet.setName(gantSheetName);
  gantSheet.setFrozenColumns(2);
  gantSheet.setFrozenRows(2);
  return gantSheet;
}

function createWeeksNumber(gantSheet, config) {
  var dates_row = ["",""];
  var week_number = 1;
  var n_working_days = LAST_DAY.diff(FIRST_DAY, "days");
  var currentDate = FIRST_DAY.clone();
  for(var i = 0; i < n_working_days; i++) {
    var currentDate = FIRST_DAY.clone().add(i, "days");
    if(currentDate.day() === 1) {
      dates_row.push(week_number);
      week_number = week_number + 1;
    }else {
      dates_row.push("");
    }
  }
  gantSheet.appendRow(dates_row);
}

function createScheduleColumns(gantSheet, config) {
  var dates_row = ["",""];
  var n_working_days = LAST_DAY.diff(FIRST_DAY, "days");
  for(var i = 0; i < n_working_days; i++) {
    var currentDate = FIRST_DAY.clone().add(i, "days");
    dates_row.push(currentDate.format("ddd DD/MM/YY"));
  }
  gantSheet.appendRow(dates_row);
}

function isHolidays(date, config) {
  var date_str = date.format();
  // We cannot use element in [] because it is not considered as an array. WTF.
  for(var i = 0; i < config["holidays"].length; i++) {
    if(date_str === config["holidays"][i]) {return true}
  }
  return false;
}

function isWeekend(date) {
  return WEEKEND.indexOf(date.day()) >= 0;
}

function isWorkingDay(date, config) {
  return (!isHolidays(date, config) && !isWeekend(date));
}

function formatDatesRow(gantSheet, config) {
  var n_working_days = LAST_DAY.diff(FIRST_DAY, "days");
  for(var i = 0; i < n_working_days; i++) {
    var currentDate = FIRST_DAY.clone().add(i, "days");
    var dateCell = gantSheet.getRange(ROW_FOR_DATES, COLUMNS_SHIFT + i);
    dateCell.setNumberFormat("ddd-mm");
    if (isHolidays(currentDate, config)) {
      dateCell.setBackground(HOLIDAY_COLOR);
    }else if(isWeekend(currentDate)) {
      dateCell.setBackground(WEEKEND_COLOR);
    }
    gantSheet.setColumnWidth(COLUMNS_SHIFT + i, CELL_WIDTH);
  }
}

function createTasksRows(gantSheet, config) {
  var n_working_days = LAST_DAY.diff(FIRST_DAY, "days");
  var currentDate = FIRST_DAY.clone();
  var row_number = 0;
  var current_column_number = 0;
  var accumulated_hours = 0;
  for(var uf_name in config["ufs"]) {
    var current_uf = config["ufs"][uf_name];
    var uf_row = ROW_FOR_BODY + row_number;
    gantSheet.getRange(uf_row, 1).setValue(uf_name);
    
    row_number = row_number + 1;
    
    for(var j = 0; j < current_uf["activities"].length; j++) {
      var current_activity = current_uf["activities"][j];
      
      gantSheet.getRange(ROW_FOR_BODY + row_number, 1).setValue(current_activity["name"]);
      gantSheet.getRange(ROW_FOR_BODY + row_number, 2).setFormulaR1C1("=SUM(R[0]C[1]:R[0]C[" + n_working_days + "])");
      
      while (accumulated_hours < current_activity["hours"]){
        var currentDate = FIRST_DAY.clone().add(current_column_number, "days");
        if (isWorkingDay(currentDate, config)) {
          for(var k = 0; k < config["classes"].length; k++){
            var classDay = config["classes"][k];
            var weekDayIndex = parseInt(classDay["weekday"]);
            if(currentDate.day() === weekDayIndex) {
              var date_cells = gantSheet.getRange(ROW_FOR_BODY + row_number, COLUMNS_SHIFT + current_column_number);
              date_cells.setBackground(WORKING_COLOR);
              date_cells.setValue(classDay["hours"]);
              accumulated_hours = accumulated_hours + classDay["hours"];
              break;
            }
          }
        }
        gantSheet.getRange(uf_row, COLUMNS_SHIFT + current_column_number).setBackground(UF_COLOR);
        current_column_number = current_column_number + 1;
      }
      accumulated_hours = 0;
      row_number = row_number + 1;
    }
  }
}
