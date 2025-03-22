let fileSource
let fileDestination
let thisSheet
let thisTab
let setRange
let urlColumn

/**
 * Creates copies of a specified template file in a designated Google Drive folder.
 *
 * This function takes a template file and generates multiple copies based on a list 
 * of filenames provided in a specified range of a Google Sheet. It optionally configures 
 * Google Forms to use the new copies and logs the URLs of the newly created files in the 
 * specified column of the sheet.
 *
 * @param {string} fileSource The file URL or ID of the template file to duplicate. If invalid or missing, the user is prompted to provide it.
 * @param {string} fileDestination The parent folder URL or ID where the copies will be stored. If invalid or missing, the user is prompted to provide it.
 * @param {string} [thisSheet] The file URL or ID of the Google Sheets file containing the filenames. Defaults to the active spreadsheet if not provided.
 * @param {string} [thisTab] The name of the tab in the Google Sheets file where the filenames are located. Defaults to the currently active tab if not provided.
 * @param {string} [setRange] A1 notation range indicating where the filenames are located in the sheet. If invalid or missing, the user is prompted to provide it.
 * @param {number} urlColumn The column index (1-based) of the sheet where the file URLs will be written. If mismatched with the header, an error is logged, and the function aborts.
 * @param {string} [responseTarget] The URL or ID of the sheet to send responses to if the template is a Google Form. If not provided, no response sheet is linked.
 * @param {string} [newTabName] The desired name for the new response sheet created by the form. If not provided, the default tab name is retained.
 *
 * @returns {void} This function does not return any value.
 *
 * @throws {Error} Throws an error if any provided parameters are invalid or if there are issues with file copying.
 *
 * @example
 * createCopies('templateFileURL', 'destinationFolderURL', 'sheetURL', 'TabName', 'A2:A10', 2);
 * 
 */
function createCopies(fileSource, fileDestination, thisSheet, thisTab, setRange, urlColumn, responseTarget) {
  
  fileSource = "https://docs.google.com/spreadsheets/d/1EsfzqMSGEmTlwFy05FS80_KxsPhlpjNDCSgZm6W9ybA/edit?gid=42594969#gid=42594969";
  fileDestination  = "https://drive.google.com/drive/u/0/folders/1bwQaMsamNznAQQfBsxNBBf8je7904e2V";
  thisSheet = "https://docs.google.com/spreadsheets/d/1ISo30LLEVO4k8wnIYAkpsALfbNuVHYAEpLIXhUHXed4/edit?gid=39455103#gid=39455103";
  thisTab = "Files with link variations";
  setRange = "C5:C6";
  urlColumn = "4";

  console.log(`Start createCopies('${fileSource}', '${fileDestination}', '${thisSheet}', '${thisTab}', '${setRange}', '${urlColumn}', '${responseTarget}')`);
  console.time("createCopies() time ");

  let COPY_NAME_COL = getUserInput("Enter the header name for filenames to copy:");
  if (!COPY_NAME_COL) {
    console.error("User input for COPY_NAME_COL was invalid or canceled.");
    return;
  }

  let COPY_URL_COL = getUserInput("Enter the header name to receive new copy URLs:");
  if (!COPY_URL_COL) {
    console.error("User input for COPY_URL_COL was invalid or canceled.");
    return;
  }

  if (!fileSource) {
    console.warn(`createCopies() was run with a falsy fileSource parameter: '${fileSource}'`);
    console.info(`Prompting user for fileSource input.`);
    fileSource = getIdFromUrl(getUserInput('Enter the template file URL: '));
    console.warn(`fileSource = '${fileSource}'`);
  } else {
    console.log(`fileSource = '${fileSource}'`);
  }

  if (!fileDestination) {
    console.warn(`createCopies() was run with a falsy fileDestination parameter: '${fileDestination}'`);
    console.info(`Prompting user for fileDestination input.`);
    fileDestination = getIdFromUrl(getUserInput('Enter the destination folder URL: '));
    console.warn(`fileDestination = '${fileDestination}'`);
  } else {
    console.log(`fileDestination = '${fileDestination}'`);
  }

  if (!thisSheet) {
    console.warn(`createCopies() was run with a falsy thisSheet parameter: '${thisSheet}'`);
    console.info(`Defaulting thisSheet to the current active spreadsheet.`);
    thisSheet = SpreadsheetApp.getActiveSpreadsheet();
    console.warn(`thisSheet name = '${SpreadsheetApp.getActiveSpreadsheet().getName()}'\rthisSheet ID = '${SpreadsheetApp.getActiveSpreadsheet().getId()}'\rthisSheet URL = '${SpreadsheetApp.getActiveSpreadsheet().getUrl()}'`);
    thisSheet = SpreadsheetApp.getActiveSpreadsheet().getUrl();
  } else {
    console.log(`thisSheet name = '${SpreadsheetApp.getActiveSpreadsheet().getName()}'\rthisSheet ID = '${SpreadsheetApp.getActiveSpreadsheet().getId()}'\rthisSheet URL = '${SpreadsheetApp.getActiveSpreadsheet().getUrl()}'`);
  }

  if (!thisTab) {
    console.warn(`createCopies() was run with a falsy thisTab parameter: '${thisTab}'`);
    console.info(`Defaulting thisTab to the current open tab.`);
    thisTab = SpreadsheetApp.getActiveSheet();
    console.warn(`thisTab name = '${SpreadsheetApp.getActiveSheet().getName()}'\rthisTab ID = '${SpreadsheetApp.getActiveSheet().getSheetId()}'`);
    thisTab = SpreadsheetApp.getActiveSheet().getName();
  } else {
    console.log(`thisTab name = '${SpreadsheetApp.getActiveSheet().getName()}'\rthisTab ID = '${SpreadsheetApp.getActiveSheet().getSheetId()}'`);
  }

  if (!setRange) {
    console.warn(`createCopies() was run with a falsy setRange parameter: '${setRange}'`);
    console.info(`Prompting user for setRange input.`);
    setRange = getUserInput('Enter the filename list range in A1 notation: ');
    console.warn(`setRange = '${setRange}'`);
  } else {
    console.log(`setRange = '${setRange}'`);
  }

  let sheet = SpreadsheetApp.openByUrl(thisSheet).getSheetByName(thisTab);
  let template = DriveApp.getFileById(getIdFromUrl(fileSource)); // File you want to copy
  let folder = DriveApp.getFolderById(getIdFromUrl(fileDestination)); // Destination folder where the copies will be stored.
  // https://stackoverflow.com/a/22917499/1027723
  let dataRange = sheet.getDataRange(); // Fetch all the values in the Range.
  let data = dataRange.getDisplayValues(); // convert 2d array into object array

  let heads = data.shift();

  if (!isValidHeaderRow(heads)) {
    console.error(`Header row must contain at least two unique headers. Invalid headers: '${heads}'`);
    return;
  }

  console.log(`Headers array: '${heads}'`);

  if (!heads.includes(COPY_NAME_COL)) {
    console.error(`Abort script due to missing column header: '${COPY_NAME_COL}'`);
    return;
  }

  console.log(`urlColumn = '${urlColumn}'\rindexOf(COPY_URL_COL) = '${heads.indexOf(COPY_URL_COL)}'`);

  if ((urlColumn - 1) === heads.indexOf(COPY_URL_COL)) {
    console.log(`urlColumn parameter and COPY_URL_COL match: '${COPY_URL_COL}' === '${urlColumn}'`);
    urlColumn = heads.indexOf(COPY_URL_COL);
  } else if (urlColumn && (urlColumn - 1) !== heads.indexOf(COPY_URL_COL)) {
    console.error(`Abort script due to mismatch of urlColumn parameter and COPY_URL_COL: '${COPY_URL_COL}' !== '${urlColumn}'`);
    return;
  } else if (!urlColumn && !heads.includes(COPY_URL_COL)) {
    console.error(`Abort script due to missing column header: '${COPY_URL_COL}'`);
    return;
  } else if (!urlColumn && heads.includes(COPY_URL_COL)) {
    console.warn(`Falsy urlColumn parameter: '${urlColumn}'`);
    console.log(`Truthy heads.includes(COPY_URL_COL): '${heads.includes(COPY_URL_COL)}'`);
    console.log(`Defining urlColumn = heads.indexOf(COPY_URL_COL) = '${heads.indexOf(COPY_URL_COL)}'`);
    urlColumn = heads.indexOf(COPY_URL_COL);
  } else if (urlColumn && !heads.includes(COPY_URL_COL)) {
    console.warn(`Falsy heads.includes(COPY_URL_COL): '${urlColumn}'`);
    console.log(`Truthy urlColumn parameter: '${heads.includes(COPY_URL_COL)}'`);
    console.log(`Proceedign with urlColumn = '${urlColumn}'`);
    urlColumn = urlColumn;
  } else {
    console.error(`Unknown fatal error occurred.`);
    return;
  }

  console.log(`urlColumn = '${urlColumn}'\rindexOf(COPY_URL_COL) = '${heads.indexOf(COPY_URL_COL)}'`);

  /**
   * Converts a 2D array into an array of objects, mapping column headers to cell values.
   *
   * @param {Array<Array<string>>} data A 2D array where the first row contains column headers.
   * @returns {Array<Object>} An array of objects representing rows with key-value pairs
   * based on column headers and their corresponding cell values.
   *
   * @example
   * const data = [
   *   ['Name', 'Age', 'City'],
   *   ['Alice', '30', 'New York'],
   *   ['Bob', '25', 'San Francisco']
   * ];
   *
   * const result = convertToObjects(data);
   * console.log(result);
   * // Output:
   * // [
   * //   { Name: 'Alice', Age: '30', City: 'New York' },
   * //   { Name: 'Bob', Age: '25', City: 'San Francisco' }
   * // ]
   */
  const obj = data.map(r => (heads.reduce((o, k, i) => (o[k] = r[i] || '', o), {})));
  const out = [];
  console.time("Total row processing time");
  obj.forEach(function (row, rowIdx) {
    if (row[COPY_NAME_COL] !== '' && row[COPY_URL_COL] === '' && !sheet.isRowHiddenByFilter(rowIdx + 2)) {
      console.time(`Row '${rowIdx + 2}' processing time `);
      try {
        let newFile = template.makeCopy(row[COPY_NAME_COL].toString(), folder);
        let fileUrl = newFile.getUrl(); // Get the URL of the newly created file
        sheet.getRange(rowIdx + 2, parseInt(urlColumn)).setValue(String(fileUrl)); // Write the URL to the specified urlColumn in the corresponding row
        // formSetup(newFile, responseTarget, newTabName = row[COPY_NAME_COL].toString());

        out.push([fileUrl]);
        console.info(`Copy created for '${row[COPY_NAME_COL]}' (Row ${rowIdx + 2})`);
      } catch (e) {
        out.push([e.message] || 'Unknown error occurred');
        console.error(`Failed to create copy for '${row[COPY_NAME_COL]}' (Row ${rowIdx + 2}). Error: ${e.message}`);
      } finally {
        console.timeEnd(`Row '${rowIdx + 2}' processing time `);
      }
    } else {
      if (row[COPY_URL_COL] !== '') {
        console.log(`Skipping Row ${rowIdx + 2} - Copy already created.`);
      }
      if (sheet.isRowHiddenByFilter(rowIdx + 2)) {
        console.log(`Skipping Row ${rowIdx + 2} - Row hidden by filter.`);
      }
      out.push([row[COPY_URL_COL]] || '');
    }
  });
  console.timeEnd("Total row processing time");

  console.timeEnd("createCopies() time ");
  console.log(`End createCopies()`);
}