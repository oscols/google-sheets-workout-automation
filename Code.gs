// Code.gs

/**
 * Adds the “Workout Tools” menu when the sheet opens
 */
function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu('Workout Tools')
    .addItem('Generate Week', 'generateWeekWizard')
    .addToUi();
}

/**
 * Step-by-step pop-ups:
 *   1. Block number
 *   2. Week number
 *   3. Clear existing output?  Yes / No
 * Then calls generateWeek(block, week, clearFlag)
 */
function generateWeekWizard() {
  const ui = SpreadsheetApp.getUi();

  const blockRes = ui.prompt(
    'Generate Week — Step 1 of 3',
    'Enter *block number* (e.g. 2):',
    ui.ButtonSet.OK_CANCEL
  );
  if (blockRes.getSelectedButton() !== ui.Button.OK) return;

  const block = parseInt(blockRes.getResponseText().trim(), 10);
  if (isNaN(block) || block <= 0) {
    ui.alert('Block number must be a positive integer.');
    return;
  }

  const weekRes = ui.prompt(
    'Generate Week — Step 2 of 3',
    'Enter *week number* (e.g. 3):',
    ui.ButtonSet.OK_CANCEL
  );
  if (weekRes.getSelectedButton() !== ui.Button.OK) return;

  const week = parseInt(weekRes.getResponseText().trim(), 10);
  if (isNaN(week) || week <= 0) {
    ui.alert('Week number must be a positive integer.');
    return;
  }

  const res = ui.alert(
    'Generate Week — Step 3 of 3',
    'Keep the existing rows in the output sheet?',
    ui.ButtonSet.YES_NO
  );
  const clearFlag = (res === ui.Button.NO);

  generateWeek(block, week, clearFlag);
}

function generateWeek(blockNumber, week, clear) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();

  var inputName  = 'Block ' + blockNumber + ' Plan';
  var inputSheet = ss.getSheetByName('Block ' + blockNumber + ' Plan');
  if (!inputSheet) {
    throw new Error("Can't find input sheet '" + inputName + "'");
  }

  var outputName  = 'Block ' + blockNumber;
  var outputSheet = ss.getSheetByName(outputName);
  if (!outputSheet) {
    outputSheet = ss.insertSheet(outputName);
  }
  
  // Clear the output sheet
  if (clear) {
    outputSheet.clear();
  }
  
  // Write the header in output sheet
  var rowCounter = outputSheet.getLastRow();
  var dataRange = inputSheet.getDataRange(); // Gets all the data
  var inputData = dataRange.getValues(); // Converts to a 2D array

  if (rowCounter == 0) {
    setHeader(outputSheet);
    rowCounter++;
    setBlock(outputSheet, blockNumber);
    rowCounter++;
  }
  rowCounter++;

  // Find rows with the correct week
  var weekRows = findRowsForWeek(inputData, week);

  // Set the style for week row
  setWeekStyle(outputSheet, week, rowCounter);
  rowCounter++;

  // 2D array with all the workouts for the week
  var workouts = findWorkouts(weekRows);

  // Output the workouts
  rowCounter = outputWorkouts(outputSheet, rowCounter, workouts);

  // Set the E1RM formula
  estimatedMax(outputSheet, rowCounter);
}

function setHeader(outputSheet) {
  outputSheet.appendRow(['Client Name', 'Week', 'Workout', 'Exercise', 'Exercise', 'Set', 'Reps', 'RPE/%', 'TrueRPE', 'TrueWeight', 'E1RM', 'Comment']);
  var frozenRow = outputSheet.getRange(1, 1, 1, 12); // Select the new row (12 columns now)
  frozenRow.setBackground('#5da68a');
  frozenRow.setFontColor('white');
  frozenRow.setFontWeight('Bold');
  frozenRow.setFontFamily('Roboto').setFontSize(12);
  frozenRow = outputSheet.getRange(1, 5, 1, 8); // Select the range for horizontal alignment (7 columns)
  frozenRow.setHorizontalAlignment('center');
}

function setBlock(outputSheet, blockNumber) {
  var newRow = ['', '', '', '', 'Block', '', blockNumber];
  outputSheet.appendRow(newRow);

  var outputRange = outputSheet.getRange(2, 1, 1, 12); // Adjust to 12 columns
  outputRange.setFontFamily('Roboto').setFontSize(11);
  outputRange.setBackground('#efefef');
  
  var blockRange = outputSheet.getRange(2, 5, 1, 3); // Select the block name
  blockRange.setFontFamily('Roboto').setFontSize(22);
  blockRange.setFontColor('#0f6b4f');
  blockRange.setFontWeight('bold');
  blockRange.setHorizontalAlignment('center');
}

function findRowsForWeek(inputData, week) {
  // Find rows with the correct week
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

function outputWorkouts(outputSheet, rowCounter, workouts) {
  for (var i = 0; i < workouts.length; i++) {
    // Set the style for new workout row
    setWorkoutHeader(outputSheet, i + 1, rowCounter);
    rowCounter++;

    var workout = workouts[i];
    for (var k = 0; k < workout.length; k++) {
      rowCounter = outputExercise(outputSheet, workout[k], rowCounter);
    }
  }

  // Account for the last rowCounter++
  rowCounter -= 1;

  return rowCounter;
}

function setWorkoutHeader(outputSheet, workout, rowCounter) {
  var newRow = ['', '', '', '', 'Workout ', '', workout];
  outputSheet.appendRow(newRow);
  var outputRange = outputSheet.getRange(rowCounter, 1, 1, 12); // Adjust to 12 columns
  outputRange.setFontFamily('Roboto').setFontSize(11);
  outputRange.setBackground('#5da68a');
  outputRange.setBorder(true, true, true, true, false, false, 'black', SpreadsheetApp.BorderStyle.SOLID_MEDIUM);
  var weekRange = outputSheet.getRange(rowCounter, 5, 1, 3);
  weekRange.setFontFamily('Roboto').setFontSize(14);
  weekRange.setFontColor('white');
  weekRange.setFontWeight('bold');
  weekRange.setHorizontalAlignment('center');
}

function outputExercise(outputSheet, row, rowCounter) {
  // Ensure sets is a number
  var sets = Number(row[4]); 
  var startRow = rowCounter;

  // Process the Reps field: if it contains a comma, split it; otherwise, replicate the single value for all sets.
  var reps = [];
  if (row[5] !== '' && row[5] !== null) {
    // Convert to string in case it's a number
    var repsStr = row[5].toString();
    if (repsStr.indexOf(',') > -1) {
      reps = repsStr.split(',').map(function(val) { return val.trim(); });
    } else {
      reps = Array(sets).fill(repsStr.trim());
    }
  } else {
    reps = Array(sets).fill('');
  }

  // Process the RPE/% field similarly.
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

  // Process the Comment field (if provided)
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
        newRow = [row[0], row[1], row[2], row[3], row[3], '', reps[j-1], rpes[j-1], row[7], rpes[j-1], row[9], commentText];
      } else {
        newRow = [row[0], row[1], row[2], row[3], row[3], '', reps[j-1], rpes[j-1], row[7], row[8], row[9], commentText];
      }
    } else {
      if (Number(rpes[0]) > 10) {
        newRow = [row[0], row[1], row[2], row[3], row[3], '', reps[j-1], rpes[j-1], row[7], rpes[j-1], row[9], ''];
      } else {
        newRow = [row[0], row[1], row[2], row[3], row[3], '', reps[j-1], rpes[j-1], row[7], row[8], row[9], ''];
      }
    }
    outputSheet.appendRow(newRow);

    // Set font style and size for each row
    var outputRange = outputSheet.getRange(rowCounter, 1, 1, 12);
    outputRange.setFontFamily('Roboto').setFontSize(11);
    outputRange.setFontWeight('bold');

    var setsRange = outputSheet.getRange(rowCounter, 6, 1, 6); // Adjusted the range for center alignment
    setsRange.setHorizontalAlignment('center');

    // Set background color and font style for the "Exercise" column
    var exerciseCell = outputSheet.getRange(rowCounter, 5);
    if (rowCounter == startRow) {
      exerciseCell.setBackground('#a0dbc1');
    } else {
      exerciseCell.setValue(''); // Clear the content of the cell
    }
    rowCounter++;
  }

  // Apply a border around the entire group of rows for the current exercise
  var groupRange = outputSheet.getRange(startRow, 1, sets, 12); // Adjust to 12 columns
  groupRange.setBorder(true, true, true, true, true, false, 'black', SpreadsheetApp.BorderStyle.SOLID);

  return rowCounter;
}

function estimatedMax(outputSheet, rowCounter) {
  for (var i = 2; i <= rowCounter; i++) {
    var formulaCell = outputSheet.getRange('K' + i);
    // ";" for Sweden locale, and "," for UK locale
    formulaCell.setFormula('=IF(AND(NOT(ISBLANK($I' + i + ')); NOT(ISBLANK($J' + i + '))); $J' + i + '/VLOOKUP($G' + i + ' + (10 - $I' + i + '); E1RM!$A$2:$B$100; 2; TRUE); "")');
  }
}