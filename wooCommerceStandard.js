/***
 * @description this script includes three different methods that can be used to assist with correcting Google Sheet data when preparing an import for woocommerce.
 * 
 * @method1 WordPress Image URL Replacer, replaces the file name from the google sheet with the url found in wordpress, and updates the google sheet data.
 * 
 * @medthod2 titleCaseSelected makes all selected text title cased
 * 
 * @medthod3 Attribute seperator, if commas are being used this method replaces all commas with | as this is the requried by all in one import.
 */


/**
 * WordPress Image URL Replacer for Google Sheets
 *
 * OPTIMIZED VERSION (Prefetch Mode)
 * -----------------------------------
 * • Fetches entire WordPress media library once
 * • Builds in-memory lookup map
 * • Performs all filename lookups locally
 * • Batch updates sheet values + backgrounds
 *
 * Best for media libraries under ~5,000 items
 */


// ============ CONFIGURATION ============

const WORDPRESS_SITE_URL = 'https://yourwordpress.site'; // Your WordPress site URL (no trailing slash)

// OPTIONAL: Only needed if your media library requires authentication
// Leave these empty ('') if your media is publicly accessible
const WORDPRESS_USERNAME = ''; // WordPress username (optional)
const WORDPRESS_APP_PASSWORD = ''; // WordPress Application Password (optional)


// ============ MAIN FUNCTION ============

function replaceFilenamesWithUrls() {
  const sheet = SpreadsheetApp.getActiveSheet();
  const range = sheet.getActiveRange();
  const ui = SpreadsheetApp.getUi();

  if (!range) {
    ui.alert('Please select a column containing filenames.');
    return;
  }

  if (range.getNumColumns() !== 1) {
    ui.alert(
      'Invalid Selection',
      'Please select only ONE column at a time.\n\nYou currently have ' +
        range.getNumColumns() +
        ' columns selected.',
      ui.ButtonSet.OK
    );
    return;
  }

  const values = range.getValues();
  const numRows = values.length;

  if (numRows === 0) {
    ui.alert('No data found in the selected column.');
    return;
  }

  // Build auth headers once
  const headers = { 'Content-Type': 'application/json' };
  if (WORDPRESS_USERNAME && WORDPRESS_APP_PASSWORD) {
    headers['Authorization'] =
      'Basic ' +
      Utilities.base64Encode(
        WORDPRESS_USERNAME + ':' + WORDPRESS_APP_PASSWORD
      );
  }

  ui.alert(
    'Loading Media Library',
    'Fetching entire WordPress media library into memory...\n\n' +
      'This may take a few seconds depending on library size.',
    ui.ButtonSet.OK
  );

  // 🔥 Fetch entire media library ONCE
  const mediaMap = fetchEntireMediaLibrary(headers);

  ui.alert(
    'Processing',
    `Processing ${numRows} rows locally (no further API calls)...`,
    ui.ButtonSet.OK
  );

  let successCount = 0;
  let notFoundCount = 0;
  let skippedCount = 0;
  let partialMatchCount = 0;

  const outputValues = [];
  const backgrounds = [];

  for (let row = 0; row < numRows; row++) {
    const cellValue = values[row][0];

    // Skip empty
    if (!cellValue || cellValue.toString().trim() === '') {
      skippedCount++;
      outputValues.push([cellValue]);
      backgrounds.push([null]);
      continue;
    }

    const parts = cellValue
      .toString()
      .trim()
      .split(',')
      .map(p => p.trim());

    const allUrls = parts.every(
      p => p.startsWith('http://') || p.startsWith('https://')
    );

    if (allUrls) {
      skippedCount++;
      outputValues.push([cellValue]);
      backgrounds.push([null]);
      continue;
    }

    const results = [];
    let foundAny = false;
    let notFoundAny = false;

    for (const part of parts) {
      if (part.startsWith('http://') || part.startsWith('https://')) {
        results.push(part);
        foundAny = true;
        continue;
      }

      const normalizedKey = normalizeFilename(part);
      const match = mediaMap[normalizedKey];

      if (match) {
        results.push(match);
        foundAny = true;
      } else {
        results.push(part);
        notFoundAny = true;
      }
    }

    const newValue = results.join(', ');
    outputValues.push([newValue]);

    if (foundAny && notFoundAny) {
      partialMatchCount++;
      backgrounds.push(['#ffffcc']); // Yellow
    } else if (foundAny) {
      successCount++;
      backgrounds.push([null]);
    } else {
      notFoundCount++;
      backgrounds.push(['#ffcccc']); // Red
    }
  }

  // 🚀 Batch update (massively faster than per-cell writes)
  range.setValues(outputValues);
  range.setBackgrounds(backgrounds);

  ui.alert(
    'Processing Complete',
    `Results:\n\n` +
      `✓ Successfully replaced: ${successCount}\n` +
      `◐ Partially matched: ${partialMatchCount}\n` +
      `✗ Not found: ${notFoundCount}\n` +
      `○ Skipped: ${skippedCount}\n\n` +
      `Red = no matches found\n` +
      `Yellow = partial match`,
    ui.ButtonSet.OK
  );
}


// ============ FETCH ENTIRE MEDIA LIBRARY ============

function fetchEntireMediaLibrary(headers) {
  const mediaMap = {};
  const fetchOptions = {
    method: 'get',
    headers: headers,
    muteHttpExceptions: true
  };

  let page = 1;
  const perPage = 100; // WP REST max

  while (true) {
    const url =
      `${WORDPRESS_SITE_URL}/wp-json/wp/v2/media?per_page=${perPage}&page=${page}`;

    const response = UrlFetchApp.fetch(url, fetchOptions);

    if (response.getResponseCode() !== 200) break;

    const items = JSON.parse(response.getContentText());
    if (!items || items.length === 0) break;

    for (const item of items) {
      if (!item.source_url) continue;

      const sourceUrl = item.source_url;
      const filename = sourceUrl.split('/').pop();
      const basename = filename.replace(/\.[^/.]+$/, '');

      // Store multiple lookup keys
      mediaMap[normalizeFilename(filename)] = sourceUrl;
      mediaMap[normalizeFilename(basename)] = sourceUrl;

      if (item.slug) {
        mediaMap[normalizeFilename(item.slug)] = sourceUrl;
      }

      if (item.title && item.title.rendered) {
        mediaMap[normalizeFilename(item.title.rendered)] = sourceUrl;
      }
    }

    if (items.length < perPage) break;
    page++;
  }

  return mediaMap;
}


// ============ UTILITIES ============

function normalizeFilename(str) {
  return str
    .toString()
    .split('/').pop()
    .split('\\').pop()
    .toLowerCase()
    .replace(/\.[^/.]+$/, '') // remove extension
    .replace(/[^a-z0-9]+/g, '-') // slugify
    .replace(/^-+|-+$/g, '')
    .trim();
}

// ============ CONNECTION TEST ============

function testWordPressConnection() {
  const ui = SpreadsheetApp.getUi();
  const headers = { 'Content-Type': 'application/json' };

  if (WORDPRESS_USERNAME && WORDPRESS_APP_PASSWORD) {
    headers['Authorization'] =
      'Basic ' +
      Utilities.base64Encode(
        WORDPRESS_USERNAME + ':' + WORDPRESS_APP_PASSWORD
      );
  }

  try {
    const response = UrlFetchApp.fetch(
      `${WORDPRESS_SITE_URL}/wp-json/wp/v2/media?per_page=1`,
      { method: 'get', headers: headers, muteHttpExceptions: true }
    );

    if (response.getResponseCode() === 200) {
      ui.alert('Success!', 'Connected to WordPress successfully.', ui.ButtonSet.OK);
    } else {
      ui.alert(
        'Connection Failed',
        `HTTP ${response.getResponseCode()}\n\n${response.getContentText()}`,
        ui.ButtonSet.OK
      );
    }
  } catch (e) {
    ui.alert('Connection Failed', e.toString(), ui.ButtonSet.OK);
  }
}


// ============ HELP ============

function showHelp() {
  SpreadsheetApp.getUi().alert(
    'Help',
    'WordPress Image URL Replacer (Optimized)\n\n' +
      '1. Set WORDPRESS_SITE_URL in the script\n' +
      '2. Add credentials only if your media requires authentication\n' +
      '3. Select ONE column of image filenames\n' +
      '4. WordPress Tools > Replace Filenames with URLs\n\n' +
      'This version:\n' +
      '• Loads entire media library once\n' +
      '• Performs local lookups (no per-row API calls)\n' +
      '• Uses batch updates for maximum speed\n\n' +
      'Red = no match found\n' +
      'Yellow = partial match\n' +
      'Empty/URL-only cells are skipped.',
    SpreadsheetApp.getUi().ButtonSet.OK
  );
}


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
  
  return str.toLowerCase().split(' ').map((word, index) => {
    // Check if word is an acronym (case-insensitive check)
    const upperWord = word.toUpperCase();
    if (acronyms.includes(upperWord)) {
      return upperWord;
    }

    // Always capitalize first and last word
    if (index === 0 || index === str.split(' ').length - 1) {
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
    .addItem('Title Case Selected', 'titleCaseSelected')
    .addItem('Replace Commas with Pipes', 'replaceCommasWithPipes')
    .addToUi();

   ui.createMenu('WordPress Tools')
    .addItem('Replace Filenames with URLs', 'replaceFilenamesWithUrls')
    .addItem('Test WordPress Connection', 'testWordPressConnection')
    .addSeparator()
    .addItem('Help', 'showHelp')
    .addToUi();
}
