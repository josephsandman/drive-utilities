/**
 * Retrieves and logs the types of items in a Google Form specified by its URL.
 *
 * @param {string} formUrl - The URL of the Google Form to open.
 * @return {void} This function does not return a value.
 */
function verboseForm(formUrl) {
  try {
    // Opens the Forms file by its URL.
    const form = FormApp.openByUrl(formUrl);

    // Gets the list of items in the form.
    const items = form.getItems();

    // Gets the type for each item and logs them.
    const types = items.map(item => item.getType().name());

    console.log(`DEBUG: Form item types: "${types.join(', ')}"`);

  } catch (e) {
    console.error(`ERROR: Unable to open the form at URL "${formUrl}". ${e.message}`);
  }
}

/**
 * Retrieves the names of all sheets in the active spreadsheet and logs them.
 *
 * This function iterates through all the sheets in the active spreadsheet,
 * collects their names into an array, and logs the list of sheet names 
 * to the console. It can be used for auditing or tracking spreadsheet 
 * configuration and structure.
 *
 * @param {string} sheetUrl - The URL of the spreadsheet (currently not used in the function).
 * @return {Array<string>} An array of sheet names in the active spreadsheet.
 *
 * @example
 * const sheetNames = verboseSpreadsheet("https://docs.google.com/spreadsheets/d/yourSpreadsheetId/edit");
 * Logger.log(`Retrieved sheet names: ${sheetNames.join(', ')}`);
 */
function verboseSpreadsheet(sheetUrl) {
  var out = new Array();
  var sheets = SpreadsheetApp.getActiveSpreadsheet().getSheets();
  for (var i = 0; i < sheets.length; i++) out.push([sheets[i].getName()]);
  console.log(`DEBUG: Sheet names include: "${out.join(', ')}"`);
  return out;
}