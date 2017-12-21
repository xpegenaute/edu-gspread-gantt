// ################### Main Begin ##################
// Load momentjs to work with dates efficiently

var DEBUG;

function callAllInits_() {
  var keys = Object.keys(this);
  for (var i = 0; i < keys.length; i++) {
    var funcName = keys[i];
    if (funcName.indexOf("init") == 0) {
      this[funcName].call(this);
    }
  }
}

/*
(function() {
  callAllInits_();
})();
*/

function getLibraries() {
  Logger.log("######################################## Abans del fetch");
  eval(UrlFetchApp.fetch('https://cdnjs.cloudflare.com/ajax/libs/moment.js/2.18.1/moment.min.js').getContentText());
  var test = moment.utc();
  Logger.log(test.format());
  Logger.log("######################################## Després del fetch");
}

function initMain() {
  DEBUG = 2;
};


function onMyOpen() {
  if(DEBUG >= 1) Logger.log("Starting onOpen()");
  var ui = SpreadsheetApp.getUi();
  ui.createMenu("Gantt generator").addItem("Generate Gantt", "generateGantt").addToUi();
}

function generateGantt() {
  getLibraries();
  callAllInits_();
  
  if(DEBUG >= 1) Logger.log("Starting generateGantt()");
  var config = getConfig();
  
  
  // if (DEBUG >=1) Logger.log(JSON.stringify(config));

  var gantSheet = createTab();
  createWeeksNumber(gantSheet, config);
  createScheduleColumns(gantSheet, config);
  formatDatesRow(gantSheet, config);
  createTasksRows(gantSheet, config);
}

function testMomentDifference() {
  eval(UrlFetchApp.fetch('https://cdnjs.cloudflare.com/ajax/libs/moment.js/2.18.1/moment.min.js').getContentText());
  //var now = moment.utc("18/09/2017", "DD/MM/YYYY").startOf("day");
  var now = moment.utc(new Date()).startOf("day");
  var future = now.clone().add(3, "days");
  Logger.log(future - now);
}

// ################### Main End ##################
// ################### ConfigParser Begin ##################

var CONFIG_SHEET_NAME;
var WEEK_DAYS;
var TIME_OFFSET;

if (!String.prototype.startsWith) {
    String.prototype.startsWith = function(searchString, position){
      position = position || 0;
      return this.substr(position, searchString.length) === searchString;
  };
}

function initConfig() {
  if(DEBUG) {Logger.log("ConfigParser.initConfig()");}
  CONFIG_SHEET_NAME = "Config";
  WEEK_DAYS = ["sunday", "monday", "tuesday", "wednesday", "thursday", "friday", "saturday"];
  TIME_OFFSET = "GMT+2";
}


var StateMachine = {
  init: function() {
    this.currentState = Object.create(IdleState);
    this.currentState.init(this);
  },
  setCurrentState: function(newState) {
    this.currentState = Object.create(newState);
    if("init" in this.currentState) {
      this.currentState.init(this);
    }
  },
  parse: function(row) {
    return this.currentState.parse(row);
  },
  changeState: function(row) {
    return this.currentState.changeState(row);
  }
}

var State = {
  init: function(stateMachine) {
    this.stateMachine = stateMachine;
  },
  parse: function (row) {
    throw "Not Implemented";
  },
  changeState: function(row) {
    if (row[0].toString().toLowerCase() === "holidays") {
      if (DEBUG >= 1) Logger.log(["State.changeState(", row , ") -> ReadingHolidaysState"].join());
      this.stateMachine.setCurrentState(ReadingHolidaysState);
      return true;
    }else if (row[0].toString().toLowerCase() === "classes") {
      if (DEBUG >= 1) Logger.log(["State.changeState(", row , ") -> ReadingClassesState"].join());
      this.stateMachine.setCurrentState(ReadingClassesState);
      return true;
    }else if (row[0].toString().toLowerCase() === "module") {
      if (DEBUG >= 1) Logger.log(["State.changeState(", row , ") -> ReadingModuleState"].join());
      this.stateMachine.setCurrentState(ReadingModuleState);
      return true;
    }else if (row[0].toString().toLowerCase() === "ufs") {
      if (DEBUG >= 1) Logger.log(["State.changeState(", row , ") -> ReadingUFSState"].join());
      this.stateMachine.setCurrentState(ReadingUFSState);
      return true;
    }
    return false;
  }
}

var IdleState = Object.create(State);
IdleState.parse = function(row) {
  if (DEBUG >= 2) Logger.log(["IdleState.parse(", row , ")"].join())
  this.changeState(row);
}


var ReadingHolidaysState = Object.create(State);
ReadingHolidaysState.parse = function(row) {

  var retValue = {};
  if (DEBUG >= 2) Logger.log(["ReadingHolidaysState.parse(", row , ")"].join());
  var isRange = !(row[1] === "");

  var firstDate = moment.utc(Utilities.formatDate(row[0], TIME_OFFSET, "yyyy-MM-dd"));
  Logger.log(firstDate.format());
  if(isRange) {
    var lastDate = moment.utc(Utilities.formatDate(row[1], TIME_OFFSET, "yyyy-MM-dd"));
    var holidaysArray = [];
    var rangeDays = lastDate.diff(firstDate, "days");
    for (var i=0; i <= rangeDays; i++){
      holidaysArray.push(firstDate.clone().add(i, "days").startOf("day").format());
    }
    retValue = {"holidays": holidaysArray};
  }else{
    retValue = {"holidays": [firstDate.startOf("day").format()]};
  }
  return retValue;
}

var ReadingClassesState = Object.create(State);
ReadingClassesState.parse = function(row) {
  var retValue = {};
  if (DEBUG >= 2) Logger.log(["ReadingClassesState.parse(", row , ")"].join());
  var value = row[0];
  if (value) {
    var defined_weekday = row[0].toString().toLowerCase();
    var hours = row[1];
    var weekday = WEEK_DAYS.indexOf(defined_weekday);
    if (weekday < 0) {
      throw "Week day \"" + defined_weekday + "\" must be one of: " + JSON.stringify(WEEK_DAYS);
    }
    retValue["weekday"] = weekday;
    retValue["hours"] = hours;
  }
  return retValue;
}

var ReadingModuleState = Object.create(State);
ReadingModuleState.parse = function(row) {
  if (DEBUG >= 2) Logger.log(["ReadingModuleState.parse(", row , ")"].join());
}

var ReadingUFSState = Object.create(State);
ReadingUFSState.parse = function(row) {
  if (DEBUG >= 2) Logger.log(["ReadingUFSState.parse(", row , ")"].join());
  if(row[0] === "") {return {}}
  var first_value = row[0];
  var second_value = row[1];
  var uf_regexp = /^uf\d+ /i;
  if (uf_regexp.test(first_value.toString())){
    return {"uf": {"name": first_value, "hours": parseInt(second_value)}};
  }else {
    return {"activity": {"name": first_value, "hours": parseInt(second_value)}};
  }
}

function getConfig() {

  function isEmptyOrComment(row) {
    if(row[0] === "") {
      return true;
    }
    return false;
  }

  if(DEBUG >= 1) Logger.log("Starting getConfig()");
  var config = {"module": {}, "holidays": [], "classes": [], "ufs": {}};
  var stateMachine = Object.create(StateMachine);
  stateMachine.init();

  var config_sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(CONFIG_SHEET_NAME);
  var data = config_sheet.getDataRange().getValues();
  for (var i = 0; i < data.length; i++) {
    var row = data[i];
    var lastUF;
    if(isEmptyOrComment(row)){
      continue;
    }
    if(stateMachine.changeState(row)){
      continue;
    }
    var value = stateMachine.parse(row);
    if (value && "holidays" in value) {
      config["holidays"].push.apply(config["holidays"], value["holidays"]);
    }else if(value && "weekday" in value) {
      config["classes"].push({"weekday": value["weekday"], "hours": value["hours"]});
    }else if(value && "uf" in value) {
      lastUF = value["uf"]["name"];
      config["ufs"][lastUF] = {};
      config["ufs"][lastUF]["hours"] = value["uf"]["hours"];
      config["ufs"][lastUF]["activities"] = [];
    }else if(value && "activity" in value) {
      config["ufs"][lastUF]["activities"].push(value["activity"]);
    }
  }
  Logger.log(JSON.stringify(config));
  return config;
}

// ################### ConfigParser End ##################

// ################### GanttGenerator Begin ##################

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
  var n_working_days = LAST_DAY.diff(FIRST_DAY, "days") + 1;
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
  var n_working_days = LAST_DAY.diff(FIRST_DAY, "days") + 1;
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
  var n_working_days = LAST_DAY.diff(FIRST_DAY, "days") + 1;
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
  var n_working_days = LAST_DAY.diff(FIRST_DAY, "days") + 1;
  var currentDate = FIRST_DAY.clone();
  var row_number = 0;
  var current_column_number = 0;
  var accumulated_hours = 0;
  var surplus_hours_for_today = 0;
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
              var available_hours_for_today = (surplus_hours_for_today > 0) ? classDay["hours"] - surplus_hours_for_today : classDay["hours"];
              var needed_hours = current_activity["hours"] - accumulated_hours;;
              var needed_hours_for_today = 0;
              if (available_hours_for_today >= needed_hours) {
                needed_hours_for_today = needed_hours;
                surplus_hours_for_today = available_hours_for_today - needed_hours_for_today;
              }else {
                needed_hours_for_today = available_hours_for_today;
                surplus_hours_for_today = 0;
              }
              var date_cells = gantSheet.getRange(ROW_FOR_BODY + row_number, COLUMNS_SHIFT + current_column_number);
              date_cells.setBackground(WORKING_COLOR);
              date_cells.setValue(needed_hours_for_today);
              accumulated_hours = accumulated_hours + needed_hours_for_today;
              break;
            }
          }
        }
        gantSheet.getRange(uf_row, COLUMNS_SHIFT + current_column_number).setBackground(UF_COLOR);
        if(surplus_hours_for_today == 0){
          current_column_number = current_column_number + 1;
        }
      }
      accumulated_hours = 0; //surplus_hours_for_today;
      row_number = row_number + 1;
    }
  }
}
// ################### GanttGenerator End ##################
