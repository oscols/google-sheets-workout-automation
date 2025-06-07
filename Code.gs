// Code.gs

function generateWeek() {
  var blockNumber = 1;
  var week = 1;
  var clear = false;

  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var outputSheet = ss.getSheetByName('Block ' + blockNumber);
  var inputSheet = ss.getSheetByName('Block ' + blockNumber + ' Plan');
  var dataRange = inputSheet.getDataRange();
  var inputData = dataRange.getValues();

  var rowCounter = outputSheet.getLastRow();
  if (rowCounter == 0) {
    setHeader(outputSheet);
    rowCounter++;
    setBlock(outputSheet, blockNumber);
    rowCounter++;
  }

  var weekRows = findRowsForWeek(inputData, week);

  setWeekStyle(outputSheet, week, rowCounter);
  rowCounter++;

  var workouts = findWorkouts(weekRows);
  // TODO: output workouts
  // TODO: set E1RM formulas
}

function setHeader(outputSheet) {
  outputSheet.appendRow([
    'Client Name', 'Week', 'Workout', 'Exercise', 'Exercise',
    'Set', 'Reps', 'RPE/%', 'TrueRPE', 'TrueWeight', 'E1RM', 'Comment'
  ]);
  var frozenRow = outputSheet.getRange(1, 1, 1, 12);
  frozenRow.setBackground('#5da68a');
  frozenRow.setFontColor('white');
  frozenRow.setFontWeight('Bold');
  frozenRow.setFontFamily('Roboto').setFontSize(12);

  var alignmentRange = outputSheet.getRange(1, 5, 1, 8);
  alignmentRange.setHorizontalAlignment('center');
}

function setBlock(outputSheet, blockNumber) {
  var newRow = ['', '', '', '', 'Block', '', blockNumber];
  outputSheet.appendRow(newRow);

  var outputRange = outputSheet.getRange(2, 1, 1, 12);
  outputRange.setFontFamily('Roboto').setFontSize(11);
  outputRange.setBackground('#efefef');
  
  var blockRange = outputSheet.getRange(2, 5, 1, 3);
  blockRange.setFontFamily('Roboto').setFontSize(22);
  blockRange.setFontColor('#0f6b4f');
  blockRange.setFontWeight('bold');
  blockRange.setHorizontalAlignment('center');
}

function findRowsForWeek(inputData, week) {
  var weekRows = [];
  for (var i = 1; i < inputData.length; i++) {
    if (inputData[i][1] == week) {
      weekRows.push(inputData[i]);
    }
  }
  return weekRows;
}

function setWeekStyle(outputSheet, week, rowCounter) {
  var newRow = ['', '', '', '', 'Week', '', week];
  outputSheet.appendRow(newRow);

  var outputRange = outputSheet.getRange(rowCounter, 1, 1, 12);
  outputRange.setFontFamily('Roboto').setFontSize(11);

  var weekRange = outputSheet.getRange(rowCounter, 5, 1, 3);
  weekRange.setFontFamily('Roboto').setFontSize(22);
  weekRange.setFontWeight('bold');
  weekRange.setHorizontalAlignment('center');

  outputRange.setBackground("#efefef");
  weekRange.setFontColor('#a0dbc1');
}

function findWorkouts(weekRows) {
  var workouts = [];
  var workout = [];
  var workoutNumber = 1;
  for (var i = 0; i < weekRows.length; i++) {
    if (weekRows[i][2] == workoutNumber) {
      workout.push(weekRows[i]);
    } else {
      workouts.push(workout);
      workout = [];
      workout.push(weekRows[i]);
      workoutNumber++;
    }
  }
  workouts.push(workout);
  return workouts;
}

// Other placeholder functions unchanged
function outputWorkouts(outputSheet, rowCounter, workouts) { return rowCounter; }
function estimatedMax(outputSheet, rowCounter) {}
