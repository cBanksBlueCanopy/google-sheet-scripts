/**
 * Replace Commas with Pipes
 * 
 * This script replaces all commas (,) with pipes (|) in the selected cells.

/**
 * Replaces all commas with pipes in the currently selected range
 */
function replaceCommasWithPipes() {
  var sheet = SpreadsheetApp.getActiveSheet();
  var range = sheet.getActiveRange();
  
  if (!range) {
    SpreadsheetApp.getUi().alert('Please select cells to modify.');
    return;
  }
  
  // Get the values from the selected range
  var values = range.getValues();
  
  // Loop through each cell and replace commas with pipes
  for (var i = 0; i < values.length; i++) {
    for (var j = 0; j < values[i].length; j++) {
      if (values[i][j] !== '') {
        // Convert to string and replace all commas with pipes
        values[i][j] = String(values[i][j]).replace(/,/g, ' |');
      }
    }
  }
  
  // Write the modified values back to the range
  range.setValues(values);
  
  // Show confirmation message
  var numCells = range.getNumRows() * range.getNumColumns();
  SpreadsheetApp.getUi().alert('Replaced commas with pipes in ' + numCells + ' cell(s).');
}


/**
 * Creates a custom menu when the spreadsheet opens
 */
function onOpen() {
  const ui = SpreadsheetApp.getUi();
  ui.createMenu('Text Tools')
    .addItem('Replace Commas with Pipes', 'replaceCommasWithPipes')
    .addToUi();
}