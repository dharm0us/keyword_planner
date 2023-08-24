function cpcLowRange() {
  fixValues();

  getRange().sort([
    { column: column_number("Top of page bid (low range)"), ascending: false }
  ]);
}

function cpcHighRange() {
  fixValues();

  getRange().sort([
    { column: column_number("Top of page bid (high range)"), ascending: false }
  ]);
}

function yoyChange() {
  fixValues();

  getRange().sort([
    { column: column_number("YoY change"), ascending: false }
  ]);
}

function threeMonthChange() {
  fixValues();

  // Sort order: [Column, Ascending (true/false)]
  getRange().sort([
    { column: column_number("Three month change"), ascending: false }
  ]);
}

function highVolumeLowCompetition() {
  fixValues();

  // Sort order: [Column, Ascending (true/false)]
  getRange().sort([
    { column: column_number("Avg. monthly searches"), ascending: false },
    { column: column_number("Competition (indexed value)"), ascending: true }
  ]);
}

function sortByStringLengthAndVolume() {
  fixValues();

  var keywordColNum = column_number("Keyword");
  var volumeColNum = column_number("Avg. monthly searches");
  
  // Add a helper column with string lengths
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  var dataRange = sheet.getRange(2, keywordColNum, sheet.getLastRow() - 1);
  var data = dataRange.getValues();
  var lengths = data.map(function(row) {
    return [row[0].length];
  });
  
  // Assuming the next available column after your data is where we'll place this temporary data
  var helperColumn = sheet.getLastColumn() + 1;
  sheet.getRange(2, helperColumn, lengths.length).setValues(lengths);
  
  // Sort based on the helper column (string length) and then volume
  sheet.getRange(2, 1, sheet.getLastRow() - 1, sheet.getLastColumn()).sort([
    {column: helperColumn, ascending: false},
    {column: volumeColNum, ascending: false}
  ]);
  
  // Clean up by removing the helper column
  sheet.deleteColumn(helperColumn);
}

function getRange() {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  var range = sheet.getRange(
    /*starting row*/ 2,
    /* starting column */ 1,
    /*how many rows */ sheet.getLastRow() - 1,
    /* how many columns */sheet.getLastColumn());
  return range;
}

function fixValues() {
  convertColumnTextToNumbers(["Avg. monthly searches","Top of page bid (low range)", "Top of page bid (high range)"]);

  replaceValuesInColumns(["Three month change", "YoY change"], ['--', '∞', ''], ['0%', '9999999%', '0%']);
  replaceValuesInColumns(["Avg. monthly searches", "Competition (indexed value)"], [''], ['0']);
}

function convertColumnTextToNumbers(columnNames) {
  columnNames.forEach(function (columnName) {
    var sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
    var columnIndex = column_number(columnName);

    // Get values from the specified column (excluding the header)
    var range = sheet.getRange(2, columnIndex, sheet.getLastRow() - 1);
    var values = range.getValues();

    // Iterate over each value in the column and convert text numbers to real numbers
    for (var i = 0; i < values.length; i++) {
      var cellValue = values[i][0].toString();
      // Check if it starts with an apostrophe and is followed by a number
      if (cellValue === "") {
        values[i][0] = Number("0");
      } else if (cellValue.startsWith("'") && !isNaN(cellValue.slice(1))) {
        values[i][0] = Number(cellValue.slice(1));
      }
    }

    // Set the updated values back to the range
    range.setValues(values);
  });
}

function column_number(columnName) {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();

  // Get header row values
  var headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];

  // Find the index of the column name
  var columnIndex = headers.indexOf(columnName);

  // Adjust for 0-based array vs 1-based column index
  return (columnIndex !== -1) ? columnIndex + 1 : -1;
}

function replaceValuesInColumns(columnNames, valuesToReplace, replacementValues) {
  if (valuesToReplace.length !== replacementValues.length) {
    Logger.log("Error: The length of valuesToReplace and replacementValues must be the same.");
    return;
  }

  var sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();

  if (typeof columnNames === 'string') {
    columnNames = [columnNames]; // Convert it to an array for uniform processing
  }

  // Convert column names to indices
  var columnIndices = columnNames.map(function (name) {
    return column_number(name);
  });

  columnIndices.forEach(function (columnIndex) {
    if (columnIndex === -1) {
      Logger.log("Error: Column not found.");
      return;
    }

    var range = sheet.getRange(2, columnIndex, sheet.getLastRow() - 1, 1); // Assuming data starts from row 2
    var values = range.getValues();

    for (var i = 0; i < values.length; i++) {
      for (var j = 0; j < valuesToReplace.length; j++) {
        if (values[i][0] === valuesToReplace[j]) {
          values[i][0] = replacementValues[j];
          break; // Exit inner loop once a replacement is made
        }
      }
    }

    range.setValues(values);
  });
}

