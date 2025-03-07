/**
 * @OnlyCurrentDoc
 * The above comment specifies that this automation will only
 * attempt to read or modify the spreadsheet this script is bound to.
 * The authorization request message presented to users reflects the
 * limited scope.
 */

/**
 * Adds custom menus to the Google Sheets UI upon opening the spreadsheet.
 *
 * This function is automatically triggered when the spreadsheet is opened.
 * It creates a custom menu titled 'Drive utilities' that includes items for 
 * creating named files and subfolders, as well as retrieving file and 
 * subfolder links. Additionally, it adds a 'Gmail utilities' menu for sending 
 * mail merges. 
 *
 * @param {Event} e - The event object containing information about the opening
 * event (optional, for future use if needed).
 *
 * @return {void} This function does not return a value.
 *
 * @example
 * // Automatically called when the spreadsheet is opened
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
 * This function displays a prompt dialog to the user with a specified 
 * message. It waits for the user's input and returns the text 
 * response provided. If the user cancels the prompt, the returned
 * value will be an empty string.
 *
 * @param {string} promptMessage - The message to display in the prompt dialog.
 * @return {string} The user's input from the prompt; an empty string if the prompt is canceled.
 *
 * @example
 * const userInput = getUserInput("Please enter your name:");
 * Logger.log(`User input: ${userInput}`);
 */
function getUserInput(promptMessage) { // Helper function to handle user inputs
  return SpreadsheetApp.getUi().prompt(promptMessage).getResponseText();
}

/**
 * Extracts the folder ID from a Google Drive folder URL.
 *
 * This function takes a Google Drive folder URL as input and uses a regular 
 * expression to extract the folder ID portion of the URL. The folder ID is 
 * a unique identifier used by Google Drive to reference specific folders.
 *
 * If the URL is valid and contains a folder ID, the function returns the 
 * ID as a string. If the URL does not match the expected format or if 
 * no ID is found, it will return "undefined" instead of null.
 *
 * @param {string} url - The Google Drive folder URL from which to extract the ID.
 * @return {string|undefined} The folder ID, or undefined if not found.
 *
 * @throws {Error} Throws an error if the input URL is not a valid string.
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
 * This function takes a Google Drive URL as input and uses a regular expression 
 * to remove everything following the '?' character, effectively stripping 
 * the query parameters from the URL. If the URL does not contain a query string, 
 * the original URL is returned unchanged.
 *
 * @param {string} url - The Google Drive folder URL from which to remove query parameters.
 * @return {string} The URL without the query string. If no query string is present, the original URL is returned.
 *
 * @example
 * const cleanUrl = removeQueryFromUrl("https://drive.google.com/drive/folders/1Cf1NbSxGq8po5fMpcwsCOq4Wcj6AwBXt?query=someParam");
 * Logger.log(`Clean URL: ${cleanUrl}`);
 */
function removeQueryFromUrl(url) {
  const myRegex = /\?.*$/; // Regular expression to match query strings
  if (myRegex.test(String(url))) {
    return String(url).replace(myRegex, ''); // Return URL without query parameters
  }
  return url; // Return original URL if no query string is found
}

/**
 * Creates copies of a template file with specified names in a destination folder.
 *
 * This function prompts the user for a template file URL, the sheet URL, 
 * the tab name containing filenames, the range of filenames, and the 
 * destination folder URL. It then creates copies of the template file 
 * for each specified filename and saves the URLs of the new files back 
 * to the specified column in the spreadsheet.
 *
 * @return {void} This function does not return a value.
 *
 * @throws {Error} Throws an error if any of the inputs are invalid or if 
 * there is an issue accessing or creating files in Google Drive.
 *
 * @example
 * // This function will be executed when an appropriate trigger is set.
 * copyFiles();
 */
function copyFiles() {
  var templateFileUrl = getUserInput('Enter the template file URL');
  var templateFile = DriveApp.getFileById(getIdFromUrl(templateFileUrl));
  var filenamesSheetUrl = getUserInput('Enter the filenames sheet URL');
  var filenamesTabName = getUserInput('Enter the filenames tab name');
  var filenamesRange = getUserInput('Enter filename list range');
  var destinationFolderUrl = getUserInput('Enter destination folder URL');
  var urlWriteColumn = getUserInput('Enter column number to write URLs'); // Column index for writing URLs
  
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
 * This function prompts the user for a Google Sheet URL containing a list of folder names,
 * the specific tab name within that sheet, the range of folder names, and the destination
 * folder URL. It then creates new folders in the specified destination folder for each 
 * valid folder name and saves the URLs of the newly created folders back to the specified 
 * column in the Google Sheet.
 *
 * @return {void} This function does not return a value.
 *
 * @throws {Error} Throws an error if any of the inputs are invalid, if it fails to create 
 * folders in Google Drive, or if there are issues accessing the specified Google Sheet.
 *
 * @example
 * // This function will be executed when invoked by an appropriate trigger.
 * copyFolders();
 */
function copyFolders() {
  var foldersSheetUrl = getUserInput('Enter the foldernames list sheet URL');
  var foldersTabName = getUserInput('Enter the foldernames sheet tab name');
  var foldernamesRange = getUserInput('Enter filename list range');
  var urlWriteColumn = getUserInput('Enter column number to write URLs'); // Column index for writing URLs
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
 * Retrieves the names and URLs of all files within a specified Google Drive folder 
 * and writes them to the active sheet.
 *
 * This function prompts the user to enter the public link or ID of a parent folder 
 * in Google Drive. It retrieves all files within that folder, constructs their 
 * URLs, and populates the active Google Sheet with the file names and their corresponding 
 * URLs, starting from the second row.
 *
 * @return {void} This function does not return a value.
 *
 * @throws {Error} Throws an error if the specified folder cannot be accessed, 
 * if there are issues retrieving files, or if the input for the folder link is invalid.
 *
 * @example
 * // This function will be executed when invoked by a user action (button click or menu item).
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
 * Retrieves the names and URLs of all subfolders within a specified Google Drive folder 
 * and writes them to the active sheet.
 *
 * This function prompts the user to enter the public link or ID of a parent folder in 
 * Google Drive. It retrieves all subfolders within that folder, constructs their 
 * URLs, and populates the active Google Sheet with the subfolder names and their 
 * corresponding URLs, starting from the second row.
 *
 * @return {void} This function does not return a value.
 *
 * @throws {Error} Throws an error if the specified folder cannot be accessed, 
 * if there are issues retrieving subfolders, or if the input for the folder link is invalid.
 *
 * @example
 * // This function will be executed when invoked by a user action (button click or menu item).
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
 * This function retrieves recipient email addresses and the associated email sent status
 * from the provided sheet. It utilizes a Gmail draft message as a template and logs
 * the status of sent emails directly into the specified column.
 *
 * If a subject line is not provided, the function will prompt the user to enter one.
 * If the recipient email header or email sent header is not provided, the function
 * will prompt the user for these header names. If the template or necessary sheet 
 * cannot be found, the function will log an error and terminate execution gracefully.
 *
 * @param {string} [subjectLine] - The subject line for the email draft message. Optional; if not provided, will prompt the user.
 * @param {string} thisSheet - The Google Sheet file URL or ID containing the mail merge data.
 * @param {string} thisTab - The name of the tab within the Google Sheet that has the mail merge data.
 * @param {string} [emailRecipients] - The header of the column containing recipient email addresses. Optional; if not provided, will prompt the user for this header.
 * @param {string} [emailSent] - The header of the column where the date the email was sent will be logged. Optional; if not provided, will prompt the user for this header.
 *
 * @return {void} This function does not return a value.
 *
 * @throws {Error} Throws an error if the specified sheet, tab, or required headers do not exist in the provided data.
 *
 * @example
 * sendEmails("Weekly Update", "1B2c...xyz", "Mail Merge");
 */
function sendEmails(subjectLine, thisSheet, thisTab, emailRecipients, emailSent) {
  
  let RECIPIENT_COL = emailRecipients || getUserInput("Enter the header name for recipient email addresses:");
  let EMAIL_SENT_COL = emailSent || getUserInput("Enter the header name for email sent status:");

  // if triggered without proper parameters, show browser prompt
  if (!subjectLine){
    subjectLine = Browser.inputBox( "Mail Merge",
                                    "Type or copy/paste the subject line of the Gmail " +
                                    "draft message you would like to mail merge with:",
                                    Browser.Buttons.OK_CANCEL);
    if (subjectLine === "cancel" || subjectLine == ""){
      console.error(`ERROR: Abort script due to prompt input: '${subjectLine}'`);
    // if missing subject line, finish up
    return;
    }
  }

  // if parameters not provided, handle with defaults or error
  if (!thisSheet || !thisTab){
    sheet = SpreadsheetApp.getActiveSheet();
  } else {
    sheet = SpreadsheetApp.openById(getIdFromUrl(thisSheet)).getSheetByName(String(thisTab));
    if (!sheet) {
      console.warn(`WARNING: Sheet named '${thisTab}' was not found in '${thisSheet}'.`);
      sheet = SpreadsheetApp.getActiveSheet();
      console.info(`Proceeding with the current active sheet as default: '${sheet}'`);
    }
  }
  
  console.info(`INFO: Sending mail merge from '${thisSheet}' with subject: '${subjectLine}'`);
  
  // get the draft Gmail message to use as a template
  const emailTemplate = getGmailTemplateFromDrafts_(subjectLine);
  
  // get the data from the passed sheet
  const dataRange = sheet.getDataRange();
  // Fetch displayed values for each row in the Range HT Andrew Roberts 
  // https://mashe.hawksey.info/2020/04/a-bulk-email-mail-merge-with-gmail-and-google-sheets-solution-evolution-using-v8/#comment-187490
  // @see https://developers.google.com/apps-script/reference/spreadsheet/range#getdisplayvalues
  const data = dataRange.getDisplayValues();

  // assuming row 1 contains our column headings
  const heads = data.shift();

  console.log(`DEBUG: Headers array: '${heads}'`);

  // Check if the recipient column exists in headers
  if (!heads.includes(RECIPIENT_COL)) {
    console.error(`ERROR: Abort script due to missing column header: '${RECIPIENT_COL}'`);
    return;
  }

  // Check if the email sent column exists in headers
  if (!heads.includes(EMAIL_SENT_COL)) {
    console.error(`ERROR: Abort script due to missing column header: '${EMAIL_SENT_COL}'`);
    return;
  }
  
  // get the index of column named 'Email Status' (Assume header names are unique)
  // @see http://ramblings.mcpher.com/Home/excelquirks/gooscript/arrayfunctions
  const emailSentColIdx = heads.indexOf(EMAIL_SENT_COL);
  
  // convert 2d array into object array
  // @see https://stackoverflow.com/a/22917499/1027723
  // for pretty version see https://mashe.hawksey.info/?p=17869/#comment-184945
  const obj = data.map(r => (heads.reduce((o, k, i) => (o[k] = r[i] || '', o), {})));

  // used to record sent emails
  const out = [];

  // loop through all the rows of data
  obj.forEach(function(row, rowIdx){
    // only send emails is email_sent cell is blank and not hidden by filter
    if (row[EMAIL_SENT_COL] == '' && !sheet.isRowHiddenByFilter(rowIdx+2)){
      try {
        const msgObj = fillInTemplateFromObject_(emailTemplate.message, row);

        // @see https://developers.google.com/apps-script/reference/gmail/gmail-app#sendEmail(String,String,String,Object)
        // if you need to send emails with unicode/emoji characters change GmailApp for MailApp
        // there is no from parameter with MailApp
        // @see https://developers.google.com/apps-script/reference/mail/mail-app#advanced-parameters_1
        // Uncomment advanced parameters as needed (see docs for limitations)
        MailApp.sendEmail(
          row[RECIPIENT_COL], 
          msgObj.subject, 
          msgObj.text, 
          {
            htmlBody: msgObj.html,
            // bcc: 'a.bbc@email.com',
            // cc: 'a.cc@email.com',
            // from: 'an.alias@email.com', // not available when using MailApp instead of GmailApp
            // name: 'name of the sender',
            // replyTo: 'a.reply@email.com',
            // noReply: true, // if the email should be sent from a generic no-reply email address (not available to gmail.com users)
            attachments: emailTemplate.attachments,
            inlineImages: emailTemplate.inlineImages
          }
        );
        // modify cell to record email sent date
        out.push([new Date()]);
      } catch(e) {
        // modify cell to record error
        out.push([e.message]);
      }
    } else {
      out.push([row[EMAIL_SENT_COL]]);
    }
  });
  
  // updating the sheet with new data
  sheet.getRange(2, emailSentColIdx+1, out.length).setValues(out);
  
  /**
   * Get a Gmail draft message by matching the subject line.
   * @param {string} subject_line to search for draft message
   * @return {object} containing the subject, plain and html message body and attachments
  */
  function getGmailTemplateFromDrafts_(subject_line){
    try {
      // get drafts
      const drafts = GmailApp.getDrafts();
      // filter the drafts that match subject line
      const draft = drafts.filter(subjectFilter_(subject_line))[0];
      // get the message object
      const msg = draft.getMessage();

      // Handling inline images and attachments so they can be included in the merge
      // Based on https://stackoverflow.com/a/65813881/1027723
      // Get all attachments and inline image attachments
      const allInlineImages = draft.getMessage().getAttachments({includeInlineImages: true,includeAttachments:false});
      const attachments = draft.getMessage().getAttachments({includeInlineImages: false});
      const htmlBody = msg.getBody(); 

      // Create an inline image object with the image name as key 
      // (can't rely on image index as array based on insert order)
      const img_obj = allInlineImages.reduce((obj, i) => (obj[i.getName()] = i, obj) ,{});

      //Regexp to search for all img string positions with cid
      const imgexp = RegExp('<img.*?src="cid:(.*?)".*?alt="(.*?)"[^\>]+>', 'g');
      const matches = [...htmlBody.matchAll(imgexp)];

      //Initiate the allInlineImages object
      const inlineImagesObj = {};
      // built an inlineImagesObj from inline image matches
      matches.forEach(match => inlineImagesObj[match[1]] = img_obj[match[2]]);

      return {message: {subject: subject_line, text: msg.getPlainBody(), html:htmlBody}, 
              attachments: attachments, inlineImages: inlineImagesObj };
    } catch(e) {
      console.error(`ERROR: No Gmail draft found: '${e.message}'`);
      return;
    }

    /**
     * Filter draft objects with the matching subject linemessage by matching the subject line.
     * @param {string} subject_line to search for draft message
     * @return {object} GmailDraft object
    */
    function subjectFilter_(subject_line){
      return function(element) {
        if (element.getMessage().getSubject() === subject_line) {
          return element;
        }
      }
    }
  }
  
  /**
   * Fill template string with data object
   * @see https://stackoverflow.com/a/378000/1027723
   * @param {string} template string containing {{}} markers which are replaced with data
   * @param {object} data object used to replace {{}} markers
   * @return {object} message replaced with data
  */
  function fillInTemplateFromObject_(template, data) {
    // we have two templates one for plain text and the html body
    // stringifing the object means we can do a global replace
    let template_string = JSON.stringify(template);

    // token replacement
    template_string = template_string.replace(/{{[^{}]+}}/g, key => {
      return escapeData_(data[key.replace(/[{}]+/g, "")] || "");
    });
    return JSON.parse(template_string);
  }

  /**
   * Escape cell data to make JSON safe
   * @see https://stackoverflow.com/a/9204218/1027723
   * @param {string} str to escape JSON special characters from
   * @return {string} escaped string
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
}

// console.time(`START: `); // start a process timer
// console.timeEnd(`END: `); // end a proceess timer
// console.log(`DEBUG: Constant message strong here, followed by variable: '${variable}'`); // debug
// console.info(`INFO: Constant message strong here, followed by variable: '${variable}'`); // info
// console.warn(`WARNING: Constant message strong here, followed by variable: '${variable}'`); // warning
// console.error(`ERROR: Constant message strong here, followed by variable: '${variable}'`); // error