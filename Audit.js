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

function sheetnames() {
  var out = new Array()
  var sheets = SpreadsheetApp.getActiveSpreadsheet().getSheets();
  for (var i=0 ; i<sheets.length ; i++) out.push( [ sheets[i].getName() ] );
  console.log(`DEBUG: Sheet names include: "${out.join(', ')}"`)
  return out;
}