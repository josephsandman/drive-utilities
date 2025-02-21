function onOpen() {
  var ui = SpreadsheetApp.getUi();
  ui.createMenu('Drive utilities')
    .addItem('Create named files', 'copyFiles')
    .addItem('Create named subfolders', 'copyFolders')
    .addItem('Retrieve File links', 'retrieveFiles')
    .addItem('Retrieve Subfolder links', 'retrieveFolders')
    .addToUi();
}

function getUserInput(promptMessage) { // Helper function to handle user inputs
  return SpreadsheetApp.getUi().prompt(promptMessage).getResponseText();
}

function copyFiles() {
  var templateFileUrl = getUserInput('Enter the template file URL');
  var templateFile = DriveApp.getFileById(getFolderIdFromLink(templateFileUrl));
  var filenamesSheetUrl = getUserInput('Enter the filenames sheet URL');
  var filenamesTabName = getUserInput('Enter the filenames tab name');
  var filenamesRange = getUserInput('Enter filename list range');
  var destinationFolderUrl = getUserInput('Enter destination folder URL');
  var urlWriteColumn = getUserInput('Enter column to write URLs'); // Column index for writing URLs
  
  var sheet = SpreadsheetApp.openByUrl(filenamesSheetUrl).getSheetByName(filenamesTabName);
  var range = sheet.getRange(filenamesRange);
  var filenames = range.getValues();
  var destinationFolder = DriveApp.getFolderById(getFolderIdFromLink(destinationFolderUrl));

  var filteredFilenames = filenames.filter((filename) => filename[0] !== "");

  for (let i = 0; i < filteredFilenames.length; i++) {
    let newFile = templateFile.makeCopy(filteredFilenames[i].toString(), destinationFolder);
    let fileUrl = newFile.getUrl();
    sheet.getRange(2 + i, parseInt(urlWriteColumn)).setValue(fileUrl);
  }
}

function copyFolders() {
  var foldersSheetUrl = getUserInput('Enter the foldernames list sheet URL');
  var foldersTabName = getUserInput('Enter the foldernames sheet tab name');
  var foldernamesRange = getUserInput('Enter filename list range');
  var destinationFolderUrl = getUserInput('Enter destination folder URL');
  var urlWriteColumn = getUserInput('Enter column to write URLs');
  
  var sheet = SpreadsheetApp.openByUrl(foldersSheetUrl).getSheetByName(foldersTabName);
  var range = sheet.getRange(foldernamesRange);
  var foldernames = range.getValues();
  var destinationFolder = DriveApp.getFolderById(getFolderIdFromLink(destinationFolderUrl));

  var filteredFoldernames = foldernames.filter((foldername) => foldername[0] !== "");

  for (let i = 0; i < filteredFoldernames.length; i++) {
    let newFolder = destinationFolder.createFolder(filteredFoldernames[i].toString());
    let folderUrl = newFolder.getUrl();
    sheet.getRange(2 + i, parseInt(urlWriteColumn)).setValue(folderUrl);
  }
}

function retrieveFiles() {
  var parentFolderUrl = getUserInput('Enter the parent folder link');
  var folder = DriveApp.getFolderById(getFolderIdFromLink(parentFolderUrl));
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

function retrieveFolders() {
  var parentFolderUrl = getUserInput('Enter the parent folder link');
  var folder = DriveApp.getFolderById(getFolderIdFromLink(parentFolderUrl));
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

function getFolderIdFromLink(link) {
  var regex = /\/folders\/([^/?]+)/;
  var match = link.match(regex);
  if (match && match[1]) {
    return match[1];
  }
  return null;
}

// Changes: 
// - Implemented getUserInput function to reduce redundancy.
// - Renamed variables for clarity and consistency.
// - Ensured filtered lists are used in loops for 'copyFiles' and 'copyFolders'.
// - Added checks to ensure data arrays are not empty before setting spreadsheet values in 'retrieveFiles' and 'retrieveFolders'.
