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
