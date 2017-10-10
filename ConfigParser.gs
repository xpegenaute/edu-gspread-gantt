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
  if (first_value.toString().toLowerCase().startsWith("uf")){
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
