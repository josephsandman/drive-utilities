/**
 * @OnlyCurrentDoc
 * Indicates that this automation will only attempt to read or modify the spreadsheet this script is bound to.
 * The authorization request message presented to users reflects the limited scope.
 */

/**
 * Adds custom menus to the Google Sheets UI upon opening the spreadsheet.
 * 
 * This function is automatically triggered when the spreadsheet is opened and 
 * creates custom menus titled 'Drive utilities' and 'Gmail utilities'.
 * 
 * @param {Event} e - The event object containing information about the opening event (optional).
 * 
 * @return {void} This function does not return a value.
 * 
 * @example
 * onOpen();
 */
function onOpen(e) {
  var ui = SpreadsheetApp.getUi();
  ui.createMenu('🗂️ Drive utilities 🗂️')
    .addItem('📑 Create named files', 'copyFiles')
    .addItem('📂 Create named subfolders', 'copyFolders')
    .addSeparator()
    .addItem('🔗 Retrieve File links', 'retrieveFiles')
    .addItem('🔗 Retrieve Subfolder links', 'retrieveFolders')
    .addToUi();
  ui.createMenu('📧 Gmail utilities 📧')
      .addItem('📧 Send mail merge', 'sendEmails')
      .addToUi();
}

/**
 * Prompts the user for input and returns their response as a string.
 * 
 * Displays a prompt dialog to the user; if canceled, returns an empty string.
 * 
 * @param {string} promptMessage - The message to display in the prompt dialog.
 * @return {string} The user's input from the prompt; an empty string if canceled.
 * 
 * @example
 * const userInput = getUserInput("Please enter your name:");
 * Logger.log(`User input: ${userInput}`);
 */
function getUserInput(promptMessage) {
  return SpreadsheetApp.getUi().prompt(promptMessage).getResponseText();
}

/**
 * Extracts the folder ID from a Google Drive folder URL.
 * 
 * This function takes a Google Drive folder URL as input and uses a regular 
 * expression to extract the folder ID. If no valid ID is found, returns "undefined".
 * 
 * @param {string} url - The Google Drive folder URL.
 * @return {string|undefined} The folder ID, or undefined if not found.
 * 
 * @throws {Error} If the input URL is not a valid string.
 * 
 * @example
 * const folderId = getIdFromUrl("https://drive.google.com/drive/folders/1Cf1NbSxGq8po5fMpcwsCOq4Wcj6AwBXt");
 * Logger.log(`Extracted folder ID: ${folderId}`);
 */
function getIdFromUrl(url) {
  return url.match(/[-\w]{25,}(?!.*[-\w]{25,})/).toString();
}

/**
 * Removes the query string from a Google Drive URL.
 * 
 * This function removes any parameters following the '?' character in the URL.
 * 
 * @param {string} url - The Google Drive folder URL.
 * @return {string} The URL without the query string.
 * 
 * @example
 * const cleanUrl = removeQueryFromUrl("https://drive.google.com/drive/folders/1Cf1NbSxGq8po5fMpcwsCOq4Wcj6AwBXt?query=someParam");
 * Logger.log(`Clean URL: ${cleanUrl}`);
 */
function removeQueryFromUrl(url) {
  const myRegex = /\?.*$/;
  if (myRegex.test(String(url))) {
    return String(url).replace(myRegex, '');
  }
  return url;
}

/**
 * Creates copies of a template file with specified names in a destination folder.
 * 
 * Prompts the user for a template file URL, filenames, and a destination folder URL.
 * Creates copies for each filename and saves the URLs back to the specified column in the spreadsheet.
 * 
 * @return {void} This function does not return a value.
 * 
 * @throws {Error} If any of the inputs are invalid or if there is an issue accessing or creating files.
 * 
 * @example
 * copyFiles();
 */
function copyFiles() {
  var templateFileUrl = getUserInput('Enter the template file URL');
  var templateFile = DriveApp.getFileById(getIdFromUrl(templateFileUrl));
  var filenamesSheetUrl = getUserInput('Enter the filenames sheet URL');
  var filenamesTabName = getUserInput('Enter the filenames tab name');
  var filenamesRange = getUserInput('Enter filename list range');
  var destinationFolderUrl = getUserInput('Enter destination folder URL');
  var urlWriteColumn = getUserInput('Enter column number to write URLs');
  
  var sheet = SpreadsheetApp.openByUrl(filenamesSheetUrl).getSheetByName(filenamesTabName);
  var range = sheet.getRange(filenamesRange);
  var filenames = range.getDisplayValues();
  var destinationFolder = DriveApp.getFolderById(getIdFromUrl(destinationFolderUrl));

  var filteredFilenames = filenames.filter((filename) => filename[0] !== "");

  for (let i = 0; i < filteredFilenames.length; i++) {
    let newFile = templateFile.makeCopy(filteredFilenames[i].toString(), destinationFolder);
    let fileUrl = newFile.getUrl();
    sheet.getRange(2 + i, parseInt(urlWriteColumn)).setValue(fileUrl);
  }
}

/**
 * Creates new folders with specified names in a destination folder.
 * 
 * Prompts the user for a Google Sheet URL with folder names and creates new folders
 * in the specified destination folder.
 * 
 * @return {void} This function does not return a value.
 * 
 * @throws {Error} If any of the inputs are invalid or if there are issues creating folders.
 * 
 * @example
 * copyFolders();
 */
function copyFolders() {
  var foldersSheetUrl = getUserInput('Enter the foldernames list sheet URL');
  var foldersTabName = getUserInput('Enter the foldernames sheet tab name');
  var foldernamesRange = getUserInput('Enter filename list range');
  var urlWriteColumn = getUserInput('Enter column number to write URLs');
  var destinationFolderUrl = getUserInput('Enter destination folder URL');
  
  var sheet = SpreadsheetApp.openByUrl(foldersSheetUrl).getSheetByName(foldersTabName);
  var range = sheet.getRange(foldernamesRange);
  var foldernames = range.getDisplayValues();
  var destinationFolder = DriveApp.getFolderById(getIdFromUrl(destinationFolderUrl));

  var filteredFoldernames = foldernames.filter((foldername) => foldername[0] !== "");

  for (let i = 0; i < filteredFoldernames.length; i++) {
    let newFolder = destinationFolder.createFolder(filteredFoldernames[i].toString());
    let folderUrl = newFolder.getUrl();
    sheet.getRange(2 + i, parseInt(urlWriteColumn)).setValue(folderUrl);
  }
}

/**
 * Retrieves the names and URLs of all files within a specified Google Drive folder and writes them to the active sheet.
 * 
 * Prompts the user to enter a public link or ID of a parent folder, retrieves all files,
 * and populates the active Google Sheet with the file names and URLs starting from the second row.
 * 
 * @return {void} This function does not return a value.
 * 
 * @throws {Error} If the specified folder cannot be accessed or if there are issues retrieving files.
 * 
 * @example
 * retrieveFiles();
 */
function retrieveFiles() {
  var parentFolderUrl = getUserInput('Enter the parent folder link');
  var folder = DriveApp.getFolderById(getIdFromUrl(parentFolderUrl));
  var files = folder.getFiles();
  var fileData = [];
  
  while (files.hasNext()) {
    var file = files.next();
    var fileUrl = 'https://drive.google.com/open?id=' + file.getId();
    fileData.push([file.getName(), fileUrl]);
  }
  
  if (fileData.length > 0) {
    var sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
    sheet.getRange(2, 1, fileData.length, fileData[0].length).setValues(fileData);
  }
  
  SpreadsheetApp.getUi().alert('Files retrieved successfully.');
}

/**
 * Retrieves the names and URLs of all subfolders within a specified Google Drive folder and writes them to the active sheet.
 * 
 * Prompts the user to enter a public link or ID of a parent folder, retrieves all subfolders,
 * and populates the active Google Sheet with their names and URLs.
 * 
 * @return {void} This function does not return a value.
 * 
 * @throws {Error} If the specified folder cannot be accessed or if there are issues retrieving subfolders.
 * 
 * @example
 * retrieveFolders();
 */
function retrieveFolders() {
  var parentFolderUrl = getUserInput('Enter the parent folder link');
  var folder = DriveApp.getFolderById(getIdFromUrl(parentFolderUrl));
  var subfolders = folder.getFolders();
  var folderData = [];
  
  while (subfolders.hasNext()) {
    var subfolder = subfolders.next();
    var folderUrl = 'https://drive.google.com/drive/folders/' + subfolder.getId();
    folderData.push([subfolder.getName(), folderUrl]);
  }
  
  if (folderData.length > 0) {
    var sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
    sheet.getRange(2, 1, folderData.length, folderData[0].length).setValues(folderData);
  }
  
  SpreadsheetApp.getUi().alert('Folders retrieved successfully.');
}

/**
 * Sends emails to recipients based on data from a specified Google Sheet.
 * 
 * Retrieves recipient email addresses and the associated email sent status and 
 * uses a Gmail draft message as a template. Logs the status of sent emails directly
 * into the specified column. Prompts the user for missing header names if not provided.
 * 
 * @param {string} [subjectLine] - Subject line for the email draft message. Optional; prompts if not provided.
 * @param {string} thisSheet - The Google Sheet file URL or ID with mail merge data.
 * @param {string} thisTab - The name of the tab within the Google Sheet with mail merge data.
 * @param {string} [emailRecipients] - Header of the column containing recipient email addresses. Optional; prompts if not provided.
 * @param {string} [emailSent] - Header of the column where email sent dates are logged. Optional; prompts if not provided.
 * 
 * @return {void} This function does not return a value.
 * 
 * @throws {Error} If the specified sheet, tab, or required headers do not exist.
 * 
 * @example
 * sendEmails("Weekly Update", "1B2c...xyz", "Mail Merge");
 */
function sendEmails(subjectLine, thisSheet, thisTab, emailRecipients, emailSent) {
  console.log(`DEBUG: Parameters received by sendEmails(): \rsubjectLine: ${subjectLine}\rthisSheet: ${thisSheet}\rthisTab: ${thisTab}\remailRecipients: ${emailRecipients}\remailSent: ${emailSent}`);
  console.time('sendEmails() processing time');
  
  let RECIPIENT_COL = emailRecipients || getUserInput("Enter the header name for recipient email addresses:");
  let EMAIL_SENT_COL = emailSent || getUserInput("Enter the header name for email sent status:");

  if (!subjectLine) {
    subjectLine = Browser.inputBox( "Mail Merge",
                                    "Type or copy/paste the subject line of the Gmail " +
                                    "draft message you would like to mail merge with:",
                                    Browser.Buttons.OK_CANCEL);
    if (subjectLine === "cancel" || subjectLine === "") {
      console.error(`ERROR: Abort script due to prompt response: '${subjectLine}'`);
      return;
    }
  }

  if (!thisSheet || !thisTab) {
    sheet = SpreadsheetApp.getActiveSheet();
  } else {
    sheet = SpreadsheetApp.openById(getIdFromUrl(thisSheet)).getSheetByName(String(thisTab));
    if (!sheet) {
      console.warn(`WARNING: Sheet named '${thisTab}' was not found in '${thisSheet}'.`);
      sheet = SpreadsheetApp.getActiveSheet();
      console.log(`DEBUG: Proceeding with the current active sheet as default: '${sheet}'`);
    }
  }
  
  console.info(`INFO: Sending mail merge from '${sheet.getName()}' with subject: '${subjectLine}'`);
  
  const emailTemplate = getGmailTemplateFromDrafts_(subjectLine);
  
  const dataRange = sheet.getDataRange();
  const data = dataRange.getDisplayValues();

  const heads = data.shift();

  if (!isValidHeaderRow(heads)) {
    console.error(`ERROR: Header row must contain at least two unique headers. Invalid headers: '${heads}'`);
    return;
  }

  console.log(`DEBUG: Headers array: '${heads}'`);

  if (!heads.includes(RECIPIENT_COL)) {
    console.error(`ERROR: Abort script due to missing column header: '${RECIPIENT_COL}'`);
    return;
  }

  if (!heads.includes(EMAIL_SENT_COL)) {
    console.error(`ERROR: Abort script due to missing column header: '${EMAIL_SENT_COL}'`);
    return;
  }
  
  console.log(`DEBUG: Ready to define emailSentColIdx: '${heads.indexOf(EMAIL_SENT_COL)}'`);
  
  const emailSentColIdx = heads.indexOf(EMAIL_SENT_COL);

  console.log(`DEBUG: Email sent column: '${emailSentColIdx}'`);
  
  const obj = data.map(r => (heads.reduce((o, k, i) => (o[k] = r[i] || '', o), {})));

  const out = [];

  console.time("Total row processing time");
  obj.forEach(function(row, rowIdx) {
    if (row[RECIPIENT_COL] === '' && row[EMAIL_SENT_COL] === '' && !sheet.isRowHiddenByFilter(rowIdx + 2)) {
      console.time(`Row '${rowIdx + 2}' processing time `);
      try {
        const msgObj = fillInTemplateFromObject_(emailTemplate.message, row);

        MailApp.sendEmail(
          row[RECIPIENT_COL], 
          msgObj.subject, 
          msgObj.text, 
          {
            htmlBody: msgObj.html,
            // bcc: 'a.bbc@email.com',
            // cc: 'a.cc@email.com',
            // name: 'name of the sender',
            // replyTo: 'a.reply@email.com',
            // noReply: true, // if the email should be sent from a generic no-reply email address (not available to gmail.com users)
            attachments: emailTemplate.attachments,
            inlineImages: emailTemplate.inlineImages
          }
        );
        out.push([new Date()]);
        console.info(`INFO: Email sent to '${row[RECIPIENT_COL]}' (Row ${rowIdx + 2})`);
      } catch (e) {
        out.push([e.message]);
        console.error(`ERROR: Failed to send email to '${row[RECIPIENT_COL]}' (Row ${rowIdx + 2}). Error: ${e.message}`);
      }
      console.timeEnd(`Row '${rowIdx + 2}' processing time `);
    } else {
      if (row[EMAIL_SENT_COL] !== '') {
        console.log(`DEBUG: Skipping Row ${rowIdx + 2} - Email already sent.`);
      } 
      if (sheet.isRowHiddenByFilter(rowIdx + 2)) {
        console.log(`DEBUG: Skipping Row ${rowIdx + 2} - Row hidden by filter.`);
      }
      out.push([row[EMAIL_SENT_COL]]);
    }
  });
  console.timeEnd("Total row processing time");
  
  sheet.getRange(2, emailSentColIdx + 1, out.length).setValues(out);
  
  /**
   * Get a Gmail draft message by matching the subject line.
   * 
   * @param {string} subject_line - Subject line to search for the draft message.
   * @return {object} Contains the subject, plain and HTML message body, and attachments.
   */
  function getGmailTemplateFromDrafts_(subject_line) {
    console.info(`INFO: Searching for draft with subject line: '${subject_line}'`);
    try {
      const drafts = GmailApp.getDrafts();
      console.debug(`DEBUG: Total drafts retrieved: ${drafts.length}`);
      const draft = drafts.filter(subjectFilter_(subject_line))[0];
      if (!draft) {
        console.warn(`WARNING: No draft found matching the subject line: '${subject_line}'`);
        return;
      } else {
        console.info(`INFO: Draft found with subject: '${draft.getMessage().getSubject()}'`);
      }
      const msg = draft.getMessage();

      const allInlineImages = draft.getMessage().getAttachments({includeInlineImages: true, includeAttachments: false});
      console.debug(`DEBUG: Total inline images found: ${allInlineImages.length}`);
      const attachments = draft.getMessage().getAttachments({includeInlineImages: false});
      const htmlBody = msg.getBody();
      console.debug(`DEBUG: Draft HTML body length: ${htmlBody.length}`);

      const img_obj = allInlineImages.reduce((obj, i) => (obj[i.getName()] = i, obj), {});

      const imgexp = RegExp('<img.*?src="cid:(.*?)".*?alt="(.*?)"[^\>]+>', 'g');
      const matches = [...htmlBody.matchAll(imgexp)];
      console.debug(`DEBUG: Total inline image matches found in HTML body: ${matches.length}`);

      const inlineImagesObj = {};
      matches.forEach(match => inlineImagesObj[match[1]] = img_obj[match[2]]);
      console.info(`INFO: Returning message details for subject: '${subject_line}'`);

      return {
        message: {
          subject: subject_line,
          text: msg.getPlainBody(),
          html: htmlBody
        },
        attachments: attachments,
        inlineImages: inlineImagesObj
      };
    } catch (e) {
      console.error(`ERROR: No Gmail draft found: '${e.message}'`);
      return;
    }

    /**
     * Filter draft objects by the matching subject line.
     * 
     * @param {string} subject_line - Subject line to search for the draft message.
     * @return {function} GmailDraft object filter function.
     */
    function subjectFilter_(subject_line) {
      return function(element) {
        if (element.getMessage().getSubject() === subject_line) {
          return element;
        }
      }
    }
  }
  
  /**
   * Fills template string with data object values.
   * 
   * @param {string} template - Template string containing {{}} markers for replacement.
   * @param {object} data - Object with values to replace the {{}} markers.
   * @return {object} Message replaced with data.
   */
  function fillInTemplateFromObject_(template, data) {
    let template_string = JSON.stringify(template);

    template_string = template_string.replace(/{{[^{}]+}}/g, key => {
      const text = data[key.replace(/[{}]+/g, "")] || "";
      return escapeData_(text.replace(/\n/g, '<br>'));
    });
    return JSON.parse(template_string);
  }

  /**
   * Escapes cell data to make it JSON safe.
   * 
   * @param {string} str - String to escape special characters.
   * @return {string} Escaped string.
   */
  function escapeData_(str) {
    return str
      .replace(/[\\]/g, '\\\\')
      .replace(/[\"]/g, '\\\"')
      .replace(/[\/]/g, '\\/')
      .replace(/[\b]/g, '\\b')
      .replace(/[\f]/g, '\\f')
      .replace(/[\n]/g, '\\n')
      .replace(/[\r]/g, '\\r')
      .replace(/[\t]/g, '\\t');
  };

  function isValidHeaderRow(headers) {
    const uniqueHeaders = new Set(headers.filter(header => header.trim() !== ""));
    return uniqueHeaders.size >= 2;
  };

  console.timeEnd('sendEmails() processing time');
}

// console.time(`START: `); // start a process timer
// console.timeEnd(`END: `); // end a proceess timer
// console.log(`DEBUG: Constant message, followed by variable: '${e.message}'`); // debug
// console.info(`INFO: Constant message, followed by variable: '${e.message}'`); // info
// console.warn(`WARNING: Constant message, followed by variable: '${e.message}'`); // warning
// console.error(`ERROR: Constant message, followed by variable: '${e.message}'`); // error