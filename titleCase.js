/**
 * Converts selected cells to title case
 * To use: Select cells, then run Tools > Macros > titleCaseSelected
 */
function titleCaseSelected() {
  const sheet = SpreadsheetApp.getActiveSheet();
  const range = sheet.getActiveRange();
  
  if (!range) {
    SpreadsheetApp.getUi().alert('Please select cells to convert to title case.');
    return;
  }
  
  const values = range.getValues();
  const newValues = values.map(row => 
    row.map(cell => {
      if (typeof cell === 'string') {
        return toTitleCase(cell);
      }
      return cell;
    })
  );
  
  range.setValues(newValues);
}

/**
 * Converts a string to title case
 * @param {string} str - The string to convert
 * @return {string} The title-cased string
 */
function toTitleCase(str) {
  // Words that should typically stay lowercase in title case
  const smallWords = ['a', 'an', 'and', 'as', 'at', 'but', 'by', 'for', 'in', 
                      'of', 'on', 'or', 'the', 'to', 'up', 'via', 'with'];
  
  // Acronyms that should stay uppercase
  const acronyms = ['USA', 'US'];
  
  return str.toLowerCase().split(' ').map((word, index, array) => {
    // Check if word is an acronym (case-insensitive check)
    const upperWord = word.toUpperCase();
    if (acronyms.includes(upperWord)) {
      return upperWord;
    }
    
    // Always capitalize first and last word
    if (index === 0 || index === array.length - 1) {
      return word.charAt(0).toUpperCase() + word.slice(1);
    }
    
    // Keep small words lowercase unless they're the first or last word
    if (smallWords.includes(word)) {
      return word;
    }
    
    // Capitalize other words
    return word.charAt(0).toUpperCase() + word.slice(1);
  }).join(' ');
}

/**
 * Creates a custom menu when the spreadsheet opens
 */
function onOpen() {
  const ui = SpreadsheetApp.getUi();
  ui.createMenu('Text Tools')
    .addItem('Title Case Selected', 'titleCaseSelected')
    .addToUi();
}