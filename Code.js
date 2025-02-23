/**
 * @OnlyCurrentDoc
 * The above comment specifies that this automation will only
 * attempt to read or modify the spreadsheet this script is bound to.
 * The authorization request message presented to users reflects the
 * limited scope.
 */

/**
 * Adds custom menus to the spreadsheet upon opening.
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
 * @param {string} promptMessage The message to display in the prompt.
 * @return {string} The user's input.
 */
function getUserInput(promptMessage) { // Helper function to handle user inputs
  return SpreadsheetApp.getUi().prompt(promptMessage).getResponseText();
}

/**
 * Extracts the folder ID from a Google Drive folder URL.
 * @param {string} url The Google Drive folder URL.
 * @return {string|null} The folder ID, or null if not found.
 */
function getIdFromUrl(url) { 
  return url.match(/[-\w]{25,}(?!.*[-\w]{25,})/).toString();
}

/**
 * Creates copies of a template file with specified names in a destination folder.
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
 * Retrieves the names and URLs of files within a specified folder and writes them to the active sheet.
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
 * Retrieves the names and URLs of subfolders within a specified folder and writes them to the active sheet.
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
 * Sends emails based on data from a sheet and a Gmail draft template.
 * @param {string} [subjectLine] The subject line of the Gmail draft template.
 * @param {Sheet} [sheet=SpreadsheetApp.getActiveSheet()] The sheet containing the data.
 */
function sendEmails(subjectLine, sheet=SpreadsheetApp.getActiveSheet()) {
  // option to skip browser prompt if you want to use this code in other projects
  if (!subjectLine){
    subjectLine = Browser.inputBox("Mail Merge", 
                                      "Type or copy/paste the subject line of the Gmail " +
                                      "draft message you would like to mail merge with:",
                                      Browser.Buttons.OK_CANCEL);

    if (subjectLine === "cancel" || subjectLine == ""){ 
    // If no subject line, finishes up
    return;
    }
  }

  // Gets the draft Gmail message to use as a template
  const emailTemplate = getGmailTemplateFromDrafts_(subjectLine);

  // Gets the data from the passed sheet
  const dataRange = sheet.getDataRange();
  // Fetches displayed values for each row in the Range HT Andrew Roberts 
  // https://mashe.hawksey.info/2020/04/a-bulk-email-mail-merge-with-gmail-and-google-sheets-solution-evolution-using-v8/#comment-187490
  // @see https://developers.google.com/apps-script/reference/spreadsheet/range#getdisplayvalues
  const data = dataRange.getDisplayValues();

  // Assumes row 1 contains our column headings
  const heads = data.shift(); 

  // Gets the index of the column named 'Email Status' (Assumes header names are unique)
  // @see http://ramblings.mcpher.com/Home/excelquirks/gooscript/arrayfunctions
  const emailSentColIdx = heads.indexOf(EMAIL_SENT_COL);

  // Converts 2d array into an object array
  // See https://stackoverflow.com/a/22917499/1027723
  // For a pretty version, see https://mashe.hawksey.info/?p=17869/#comment-184945
  const obj = data.map(r => (heads.reduce((o, k, i) => (o[k] = r[i] || '', o), {})));

  // Creates an array to record sent emails
  const out = [];

  // Loops through all the rows of data
  obj.forEach(function(row, rowIdx){
    // Only sends emails if email_sent cell is blank and not hidden by a filter
    if (row[EMAIL_SENT_COL] == '' && !sheet.isRowHiddenByFilter(rowIdx+2)){
      try {
        const msgObj = fillInTemplateFromObject_(emailTemplate.message, row);

        // See https://developers.google.com/apps-script/reference/gmail/gmail-app#sendEmail(String,String,String,Object)
        // If you need to send emails with unicode/emoji characters change GmailApp for MailApp
        // Uncomment advanced parameters as needed (see docs for limitations)
        MailApp.sendEmail(row[RECIPIENT_COL], msgObj.subject, msgObj.text, {
          htmlBody: msgObj.html,
          // bcc: 'joshmckenna+script@grace-bible.org',
          // cc: 'a.cc@email.com',
          // from: 'an.alias@email.com',
          // name: 'name of the sender',
          // replyTo: 'hr@grace-bible.org',
          // noReply: true, // if the email should be sent from a generic no-reply email address (not available to gmail.com users)
          attachments: emailTemplate.attachments,
          inlineImages: emailTemplate.inlineImages
        });
        // Edits cell to record email sent date
        out.push([new Date()]);
      } catch(e) {
        // modify cell to record error
        out.push([e.message]);
      }
    } else {
      out.push([row[EMAIL_SENT_COL]]);
    }
  });

  // Updates the sheet with new data
  sheet.getRange(2, emailSentColIdx+1, out.length).setValues(out);

  /**
   * Retrieves a Gmail draft message by matching the subject line.
   * @param {string} subject_line The subject line to search for.
   * @return {object} An object containing the subject, plain and HTML message body, and attachments.
   * @private  // This indicates it's an internal helper function
   */
  function getGmailTemplateFromDrafts_(subject_line){
    try {
      // get drafts
      const drafts = GmailApp.getDrafts();
      // filter the drafts that match subject line
      const draft = drafts.filter(subjectFilter_(subject_line))[0];
      // get the message object
      const msg = draft.getMessage();

      // Handles inline images and attachments so they can be included in the merge
      // Based on https://stackoverflow.com/a/65813881/1027723
      // Gets all attachments and inline image attachments
      const allInlineImages = draft.getMessage().getAttachments({includeInlineImages: true,includeAttachments:false});
      const attachments = draft.getMessage().getAttachments({includeInlineImages: false});
      const htmlBody = msg.getBody(); 

      // Creates an inline image object with the image name as key 
      // (can't rely on image index as array based on insert order)
      const img_obj = allInlineImages.reduce((obj, i) => (obj[i.getName()] = i, obj) ,{});

      //Regexp searches for all img string positions with cid
      const imgexp = RegExp('<img.*?src="cid:(.*?)".*?alt="(.*?)"[^\>]+>', 'g');
      const matches = [...htmlBody.matchAll(imgexp)];

      //Initiates the allInlineImages object
      const inlineImagesObj = {};
      // built an inlineImagesObj from inline image matches
      matches.forEach(match => inlineImagesObj[match[1]] = img_obj[match[2]]);

      return {message: {subject: subject_line, text: msg.getPlainBody(), html:htmlBody}, 
              attachments: attachments, inlineImages: inlineImagesObj };
    } catch(e) {
      throw new Error("Oops - can't find Gmail draft");
    }

    /**
     * Filters draft objects with the matching subject line.
     * @param {string} subject_line The subject line to search for.
     * @return {object} The GmailDraft object.
     * @private
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
   * Fills in a template string with data from an object.
   * @see https://stackoverflow.com/a/378000/1027723
   * @param {string} template The template string containing {{}} markers.
   * @param {object} data The data object used to replace the markers.
   * @return {object} The message with data filled in.
   * @private
   */
  function fillInTemplateFromObject_(template, data) {
    // We have two templates one for plain text and the html body
    // Stringifing the object means we can do a global replace
    let template_string = JSON.stringify(template);

    // Token replacement
    template_string = template_string.replace(/{{[^{}]+}}/g, key => {
      return escapeData_(data[key.replace(/[{}]+/g, "")] || "");
    });
    return JSON.parse(template_string);
  }
  
  /**
   * Escapes special characters in a string to make it JSON safe.
   * @see https://stackoverflow.com/a/9204218/1027723
   * @param {string} str The string to escape.
   * @return {string} The escaped string.
   * @private
   */
  function escapeData_(str) {
    return str
      .replace(/[\\]/g, '\\\\')
      .replace(/[\"]/g, '\\\"')
      .replace(/[\/]/g, '\\/')
      .replace(/[\b]/g, '\\b')
      .replace(/[\f]/g, '\\f')
      .replace(/[\n]/g, '<br>')
      .replace(/[\r]/g, '\\r')
      .replace(/[\t]/g, '\\t');
  };
}

/**
 * Retrieves and logs the types of items in a Google Form.
 */
function verboseForm() {
  // Opens the Forms file by its URL. If you created your script from within
  // a Google Forms file, you can use FormApp.getActiveForm() instead.
  // TODO(developer): Replace the URL with your own.
  const form = FormApp.openByUrl(
      'https://docs.google.com/forms/d/1MNyN0zGLbeZmljms04VApCxZ7ZOMT8Yomp6g5c_c3Fk/edit',
  );

  // Gets the list of items in the form.
  const items = form.getItems();

  // Gets the type for each item and logs them to the console.
  const types = items.map((item) => item.getType().name());

  console.log(types);
}

// function copyItems(itemType) { 
//   //... get user inputs...

//   for (let i = 0; i < filteredItems.length; i++) {
//     try {
//       let newItem;
//       if (itemType === "file") {
//         newItem = templateFile.makeCopy(filteredItems[i].toString(), destinationFolder);
//       } else if (itemType === "folder") {
//         newItem = destinationFolder.createFolder(filteredItems[i].toString());
//       }
//       let itemUrl = newItem.getUrl();
//       sheet.getRange(2 + i, parseInt(urlWriteColumn)).setValue(itemUrl);
//     } catch (e) {
//       // Handle errors (e.g., invalid name, permissions)
//       Browser.msgBox("Error processing item " + filteredItems[i] + ": " + e.message);
//     }
//   }
// }