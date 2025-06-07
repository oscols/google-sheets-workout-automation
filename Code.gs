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

  rowCounter = outputWorkouts(outputSheet, rowCounter, workouts);

  estimatedMax(outputSheet, rowCounter);
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

function setWorkoutHeader(outputSheet, workout, rowCounter) {
  var newRow = ['', '', '', '', 'Workout ', '', workout];
  outputSheet.appendRow(newRow);

  var outputRange = outputSheet.getRange(rowCounter, 1, 1, 12);
  outputRange.setFontFamily('Roboto').setFontSize(11);
  outputRange.setBackground('#5da68a');
  outputRange.setBorder(
    true, true, true, true, false, false, 'black', SpreadsheetApp.BorderStyle.SOLID_MEDIUM
  );

  var weekRange = outputSheet.getRange(rowCounter, 5, 1, 3);
  weekRange.setFontFamily('Roboto').setFontSize(14);
  weekRange.setFontColor('white');
  weekRange.setFontWeight('bold');
  weekRange.setHorizontalAlignment('center');
}

function outputWorkouts(outputSheet, rowCounter, workouts) {
  for (var i = 0; i < workouts.length; i++) {
    setWorkoutHeader(outputSheet, i + 1, rowCounter);
    rowCounter++;

    var workout = workouts[i];
    for (var k = 0; k < workout.length; k++) {
      rowCounter = outputExercise(outputSheet, workout[k], rowCounter);
    }
  }
  rowCounter -= 1;
  return rowCounter;
}

function outputExercise(outputSheet, row, rowCounter) {
  // Ensure sets is a number
  var sets = Number(row[4]); 
  var startRow = rowCounter;

  // Process the Reps field
  var reps = [];
  if (row[5] !== '' && row[5] !== null) {
    var repsStr = row[5].toString();
    if (repsStr.indexOf(',') > -1) {
      reps = repsStr.split(',').map(function(val) { return val.trim(); });
    } else {
      reps = Array(sets).fill(repsStr.trim());
    }
  } else {
    reps = Array(sets).fill('');
  }

  // Process the RPE/% field
  var rpes = [];
  if (row[6] !== '' && row[6] !== null) {
    var rpesStr = row[6].toString();
    if (rpesStr.indexOf(',') > -1) {
      rpes = rpesStr.split(',').map(function(val) { return val.trim(); });
    } else {
      rpes = Array(sets).fill(rpesStr.trim());
    }
  } else {
    rpes = Array(sets).fill('');
  }

  // Process the Comment field
  var commentLine = 0;
  var commentText = '';
  if (row[10] != '') {
    var commentArr = row[10].split(';');
    commentText = commentArr[0];
    commentLine = commentArr[1];
  }

  // Output one row per set.
  for (var j = 1; j <= sets; j++) {
    var newRow;
    if (Number(commentLine) === j) {
      if (Number(rpes[0]) > 10) {
        newRow = [
          row[0], row[1], row[2], row[3], row[3], '', 
          reps[j-1], rpes[j-1], row[7], rpes[j-1], row[9], commentText
        ];
      } else {
        newRow = [
          row[0], row[1], row[2], row[3], row[3], '',
          reps[j-1], rpes[j-1], row[7], row[8], row[9], commentText
        ];
      }
    } else {
      if (Number(rpes[0]) > 10) {
        newRow = [
          row[0], row[1], row[2], row[3], row[3], '',
          reps[j-1], rpes[j-1], row[7], rpes[j-1], row[9], ''
        ];
      } else {
        newRow = [
          row[0], row[1], row[2], row[3], row[3], '',
          reps[j-1], rpes[j-1], row[7], row[8], row[9], ''
        ];
      }
    }
    outputSheet.appendRow(newRow);

    // Style each appended row
    var outputRange = outputSheet.getRange(rowCounter, 1, 1, 12);
    outputRange.setFontFamily('Roboto').setFontSize(11);
    outputRange.setFontWeight('bold');

    var setsRange = outputSheet.getRange(rowCounter, 6, 1, 6);
    setsRange.setHorizontalAlignment('center');

    // Highlight Exercise column background on first set
    var exerciseCell = outputSheet.getRange(rowCounter, 5);
    if (rowCounter == startRow) {
      exerciseCell.setBackground('#a0dbc1');
    } else {
      exerciseCell.setValue('');
    }
    rowCounter++;
  }

  // Apply a border around the group of rows for this exercise
  var groupRange = outputSheet.getRange(startRow, 1, sets, 12);
  groupRange.setBorder(
    true, true, true, true, true, false, 'black', SpreadsheetApp.BorderStyle.SOLID
  );

  return rowCounter;
}

function estimatedMax(outputSheet, rowCounter) {}
