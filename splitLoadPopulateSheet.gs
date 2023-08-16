/* -----------------
 * Split load - Awesome Tables - Google Drive Searcher -   
 * -----------------
 * Written by: Sebastian Barry (github.com/sebastianbarry)
 * 
 * The purpose of this script is to represent the files within a Google Drive file
 * structure in a Google Sheet using Google App Script.
 * 
 * Because Google App Script has the limitation of 6 min max execution time for a 
 * script to run for non-Google Workspace accounts (30 min for Google Workspace 
 * accounts), there are a few critical time-based design decisions that are
 * explained in the README.
 * 
 * This script is made specifically for use with AwesomeTables so that the data within 
 * the sheet can be searched through using AwesomeTables. Because of this, there are 2
 * sheets that MUST exist: the "Data" sheet and the "Template" sheet which are
 * formatted specifically for AwesomeTables to be able to read and display.
 * 
 * The spreadsheet that this script will be populating data to must have 3 sheets:
 *   1. "Data" for AwesomeTables to describe the content for the table
 *   2. "Template" for AwesomeTables to describe the css styling for the table
 *   3. "Savedata" for our script to keep track of where it is at in the incremental
 *      load process
 * 
 * The way this script works is broken into 2 main steps (hence the name "split load"):
 *   1. Record All Folders
 *     - `populateSavedataFolders`: Breadth first algorithm which traverses through the 
 *       entire chosen Google Drive file structure and records every folder full path, 
 *       Id, last modified date, number of files, and the last time that folder was 
 *       updated on our sheet.
 *     - The folders are recorded in the sheet named "Savedata" because this is the sheet
 *       that will act as our queue of folders; every incremental load occurring in 
 *       step 2 will take folders off the top of this queue, load the data for those
 *       folders, and then move the folder to the bottom of this queue, updating the
 *       "last time this folder was updated on our sheet" value.
 *   2. Populate Data
 *     - Find the first x folders that contain files equal to or less than our designated
 *       "filecount". This is important because we can control how long the execution time 
 *       of the script will take based on how many files we choose to load. *Keep in mind,
 *       the way this script is designed, the script will always traverse the entire "Data"
 *       sheet which is the largest significant portion of execution time*.
 *     - `getListOfFileIds`: Collect a list of all of those files IDs and last updated in 
 *       one large object. The data in this object will be compared against the content in 
 *       the "Data" sheet in the next step, which will tell us whether we need to load the 
 *       rest of the file, delete it, skip it (if it has the same ID and last modified date),
 *       or add a new file.
 *     - `compareAndUpdateSheet`: Loop through the entire "Data" sheet.
 *       - If the file we are looking at has the same file path, then we examine it:
 *         - If ID is not in our list of current files: DELETE - old file
 *         - If ID is in our list of current files and the last modified date is newer: UPDATE
 *         - If ID is in our list of current files and the last modified date is the same: SKIP
 *       - Afterwards, the only remaining files in our list of current files will be newly added
 *         files, so we append each of these to the end of the "Data" sheet: ADD - new files
 *       - Meanwhile, we are keeping track of how many files from each folder we have 
 *         successfully "checked"; once all files from a particular folder have been checked, we
 *         pop that folder off the queue in the "Savedata" sheet and move it to the bottom, 
 *         updating the "last time this folder was updated on our sheet" value.
 * 
 */

// FUNCTIONS FOR GOOGLE APP SCRIPT TRIGGERS //

function runEvery30Minutes() {
  var maxFiles = 500;
  var folderId = "0ANEBHxLVNxEHUk9PVA";
  var spreadsheetId = "1jfQDiUfWUQCFBn_kUMB-1OsZYvr54CBrgErfYVI-MVs";

  var spreadsheet = SpreadsheetApp.openById(spreadsheetId);
  var dataSheet = spreadsheet.getSheetByName("Data");
  var savedataSheet = spreadsheet.getSheetByName("Savedata");
  
  populateDataWithFileCount(dataSheet, savedataSheet, folderId, maxFiles);
}

function runEveryMonth() {
  var folderId = "0ANEBHxLVNxEHUk9PVA";
  var spreadsheetId = "1jfQDiUfWUQCFBn_kUMB-1OsZYvr54CBrgErfYVI-MVs";

  var folder = DriveApp.getFolderById(folderId);
  var savedataSheet = SpreadsheetApp.openById(spreadsheetId).getSheetByName("Savedata");
  
  recordAllFolders(folder, savedataSheet);
}

// MAIN PROCESS FUNCTIONS //

// Step one: Record all of the folders
// Precondition: Cell B1 MUST have a value equal to the number of rows in this sheet (if the last row of the savedata sheet is 1000, B1 should be 1000)
// Precondition: The content of the folder rows should begin on row 3 and then extend to the row equal to the number of subfolders + 3 
function recordAllFolders(folder, savedataSheet) {
  // Clear current content
  var rowCount = savedataSheet.getRange("B1").getValue();
  savedataSheet.getRange(`A3:E${rowCount}`).clearContent();

  // Populate all the rows for folders and subfolders in Savedata sheet, and then record the rowCount
  rowCount = populateSavedataFolders(savedataSheet, folder);
  savedataSheet.getRange("B1").setValue(rowCount);

  console.log("Added " + rowCount + " folders to the Savedata sheet")
}

// Step two: Populate the data
// Precondition: recordAllFolders must have already been ran and the Savedata sheet must already be populated. 
//               recordAllFolders does not need to be ran before every time, but it must have been run at least once.
// Precondition: The content of the file rows should begin on row 3 and then extend to the row number equal to the number of files + 2
// Precondition: The Savedata sheet contains one row in the list of folders named "Bookend" which acts as the cleanup old files indicator
function populateDataWithFileCount(dataSheet, savedataSheet, rootFolderId, filecount = 500) {
  var lastRow = savedataSheet.getLastRow();
  var filesToLoad = 0;
  var foldersToLoad = [];

  // Initialization: Loop through the queue of folders in Savedata that equal up to the filecount and create a list of files and folders to be loaded
  for (var i = 3; filesToLoad <= filecount; i++) {
    // Gather row details
    let currentRow = savedataSheet.getRange(`A${i}:E${i}`).getValues()[0];
    let folderId = currentRow[0];
    //folderPath is at currentRow[1];
    //lastUpdated is at currentRow[2];
    let numFiles = currentRow[3];
    
    // TODO: If the folder chosen is the bookend (a bookend "folder" (not actually a folder, it is inserted by the recordAllFolders function) which acts as a flag and comes before the root drive folder), then also check to make sure there aren't any files in the sheet that have a folderpath that doesn't exist in the savedata sheet
    // Essentially completing the loop; after we reach the beginning of the loop, make sure we do any housekeeping before restarting the loop
    if(folderId == "Bookend") {
      console.log("Hit the bookend, aborting current load and cleaning up old files");
      currentRow[4] = new Date().toLocaleDateString();
      savedataSheet.deleteRow(i);
      i--;
      savedataSheet.getRange(lastRow, 1, 1, currentRow.length).setValues([currentRow]); // Use sheet.getRange().setValues() here instead of sheet.appendRow() as it is MUCH more efficient on time
    }

    // Check whether we are at the filecount limit 
    if (filesToLoad + numFiles > filecount) {
      console.log(`Total files to load: ${filesToLoad} out of ${filecount} (${foldersToLoad.length} folders)`);
      // We will move the folder to the end of the Savedata sheet queue only AFTER we have successfully loaded all of the files within that folder
      break;
    }

    // Add the folder and it's filecount to our list of files to add
    if (numFiles > 0) {
      foldersToLoad.push(currentRow);
      filesToLoad += numFiles;
    } else {
      // Skip folders that have 0 files in them; move them to the end of the Savedata sheet
      currentRow[4] = new Date().toLocaleDateString();
      savedataSheet.deleteRow(i);
      i--;
      savedataSheet.getRange(lastRow, 1, 1, currentRow.length).setValues([currentRow]); // Use sheet.getRange().setValues() here instead of sheet.appendRow() as it is MUCH more efficient on time
    }
  }

  // Gather current files: Create the object containing all current file IDs of files from the folders in the Drive
  var currentIds = getListOfFileIds(foldersToLoad);

  // Compare to existing files: Search through and traverse all existing file records from the sheet to compare -> update, delete, and add
  compareAndUpdateSheet(dataSheet, savedataSheet, currentIds, foldersToLoad, lastRow);
}

// FUNCTIONS FOR RECORD ALL FOLDERS //

function populateSavedataFolders(sheet, folder, currentPath = "Central Document Library", currentRow = 2) {
  var stack = [];
  stack.push({ folder: folder, path: currentPath });

  while (stack.length > 0) {
    var current = stack.pop();
    var folder = current.folder;
    var path = current.path;
    var files = folder.getFiles();

    // Skip .tar folders
    if (folder.getName().toLowerCase().endsWith(".tar")) {
      continue;
    }

    // Record information for the current folder
    sheet.getRange(`A${++currentRow}:E${currentRow}`).setValues([[folder.getId(), path, folder.getLastUpdated(), getFileCount(files), new Date().toLocaleString()]]);

    // Get subfolders and add them to the stack
    var subfolders = folder.getFolders();
    while (subfolders.hasNext()) {
      var subfolder = subfolders.next();
      var subfolderPath = path + "/" + subfolder.getName();
      stack.push({ folder: subfolder, path: subfolderPath });
    }
  }

  return currentRow;
}

function getFileCount(files) {
  var count = 0;
  while (files.hasNext()) {
    files.next();
    count++;
  }
  return count;
}

// FUNCTIONS FOR POPULATE DATA //

// Generates a large object to act as a dictionary for each file ID, while also containing useful data (lastUpdated and filePath) that can be passed on to getFileDetails
function getListOfFileIds(folders) {
  var fileData = {};

  for (var i = 0; i < folders.length; i++) {
    var folder = DriveApp.getFolderById(folders[i][0])
    var fileIterator = folder.getFiles();
    while (fileIterator.hasNext()) {
      var file = fileIterator.next();
      fileData[file.getId()] = [file.getLastUpdated(), folders[i][1]];
    }
  }
  
  console.log('Gathered IDs for ', Object.keys(fileData).length, ' files');
  return fileData;
}

function compareAndUpdateSheet(dataSheet, savedataSheet, currentFiles, foldersToLoad, savedataLastRow) {
  // Find the starting row of where the files in our folder exist in our sheet
  var startRow = 3; // Start reading data from row 3 (index 2)
  var lastRow = dataSheet.getLastRow();
  
  var foldersList = {};
  for (var i = 0; i < foldersToLoad.length; i++) {
    let folderPath = foldersToLoad[i][1];
    foldersList[folderPath] = foldersToLoad[i][3];
  }

  // Go through every file in the dataSheet looking for files with a matching folder
  // This loop handles all updates, skips, and deletes; adds come after this part
  console.log("Beginning search through existing sheet...");
  for(var i = startRow; i <= lastRow; i++) {
    var iPath = dataSheet.getRange(`J${i}`).getValue();
    
    // Skip any files which do not belong to the current folder list
    if (!foldersList[iPath]) {
      continue;
    }
    // Assume every file that gets to this point has a matching folder path
    var id = dataSheet.getRange(`B${i}`).getValue();
    var lastUpdated = dataSheet.getRange(`I${i}`).getValue();

    // Add 
    if (!currentFiles[id]) {
      console.log("Found file in sheet that does not exist in Drive, deleting: " + id + " " + lastUpdated);
      dataSheet.deleteRow(i);
      lastRow--;
    } else {
      // Check if the dates are not the same based on the Date property of the Date objects
      if (currentFiles[id][0].getDate() !== lastUpdated.getDate()) {
        console.log("Found outdated file in sheet, updating: " + id + " " + lastUpdated);
        // Update the file here
        let rowData = getFileDetails(id, lastUpdated, folderPath);
        dataSheet.getRange(i, 1, 1, rowData.length).setValues([rowData]);
      } else {
        console.log("Found matching file, skipping: " + id + " " + lastUpdated);
      }
      // Remove file from dictionary here
      delete currentFiles[id];
    }

    // Decrement the number of files left to load for the particular folder
    foldersList[iPath]--;
    // Once the number of files left to load for one particular folder is 0, move the folder in the Savedata sheet queue to the end
    if (foldersList[iPath] <= 0) {
      // Get the row that will be added to the end of savedataSheet 
      var row = [];
      for (var j = 0; j < foldersToLoad.length; j++) {
        if (foldersToLoad[j][1] == iPath) {
          row = foldersToLoad[j];
          row[4] = new Date() .toLocaleDateString();
        }
      }
      // Delete the row from savedataSheet and then add it to the end
      for (var j = 3; j <= Object.keys(foldersList).length + 2; j++) {
        if (savedataSheet.getRange(`B${j}`).getValue() == iPath) {
          savedataSheet.deleteRow(j);
          savedataSheet.getRange(savedataLastRow, 1, 1, row.length).setValues([row]); // Use sheet.getRange().setValues() here instead of sheet.appendRow() as it is MUCH more efficient on time
          break;
        }
      }
    }
  }

  // Iterating through the rest of the remaining files in the currentFiles dictionary to add them to the sheet
  for(var fileId in currentFiles) {
    if (currentFiles.hasOwnProperty(fileId)) {
      let lastUpdated = currentFiles[fileId][0];
      let folderPath = currentFiles[fileId][1];
      console.log("Found file that does not exist in the sheet, adding: " + fileId + " " + lastUpdated);
      
      let rowData = getFileDetails(fileId, lastUpdated, folderPath);
      dataSheet.getRange(++lastRow, 1, 1, rowData.length).setValues([rowData]); // Use dataSheet.getRange().setValues() here instead of dataSheet.appendRow() as it is MUCH more efficient on time
      
      // Decrement the number of files left to load for the particular folder
      foldersList[folderPath]--;
      // Once the number of files left to load for one particular folder is 0, move the folder in the Savedata sheet queue to the end
      if (foldersList[iPath] <= 0) {
        // Get the row that will be added to the end of savedataSheet 
        var row = [];
        for (var j = 0; j < foldersToLoad.length; j++) {
          if (foldersToLoad[j][1] == iPath) {
            row = foldersToLoad[j];
            row[4] = new Date() .toLocaleDateString();
          }
        }
        // Delete the row from savedataSheet and then add it to the end
        for (var j = 3; j <= Object.keys(foldersList).length + 2; j++) {
          if (savedataSheet.getRange(`B${j}`).getValue() == iPath) {
            savedataSheet.deleteRow(j);
            savedataSheet.getRange(savedataLastRow, 1, 1, row.length).setValues([row]); // Use sheet.getRange().setValues() here instead of sheet.appendRow() as it is MUCH more efficient on time
            break;
          }
        }
      }
    }
  }
}

function getFileDetails(id, lastUpdated, filePath) {
  var fileInfo = [];
  var file = DriveApp.getFileById(id);

  const title = file.getName();
  // id already was fetched earlier; passed in as parameter
  const url = file.getUrl();
  const [mimeType, image] = generateFileType(file.getMimeType().replace('application/', ''));
  const ownerInfo = "Undefined"; // We decided to not get the owner and ownerInfo since it always results in "Undefined" //const [ownerInfo, owner] = generateFileOwner(file.getOwner());
  const owner = "undefined";
  const creationDate = file.getDateCreated();
  // lastUpdated already was fetched earlier; passed in as parameter
  // filePath already was fetched earlier; passed in as parameter
  const description = file.getDescription();
  const fileSize = formatFileSize(file.getSize());
  // searchFile column will always be empty
  // size column is not descriptive enough for us to add anything in this column
  const contents = ''; // We decided to not get the document contents as it is not needed // const contents = generateContents(id);
  const lastModifiedBy = ''; // Get last user to update is a difficult task which we will skip for now //getLastUserToUpdate(id);
  // completeness
  // progress
  const labels = generateLabelsOnDriveItem(id);
  
  fileInfo.push(title, id, url, image, mimeType, ownerInfo, owner, creationDate, lastUpdated, filePath, description, fileSize, "", "", contents, lastModifiedBy, "", "", labels);
  console.log("Added 1 file to data: " + title + " " + id);

  return fileInfo;
}

// For each new file type we encounter, we will want to add a case here to handle it
function generateFileType(mimeType) {
  var imageURL = "https://storage.googleapis.com/aw-gapps-ressources/at/addons/file-cabinet/assets/blank-empty-file-128.png";
  if(mimeType.includes('document')) {
    mimeType = 'Document';
    imageURL = 'https://storage.googleapis.com/aw-gapps-ressources/at/addons/file-cabinet/assets/docs-128.png';
  }
  else if (mimeType.includes('spreadsheet')) {
    mimeType = 'Spreadsheet';
    imageURL = 'https://storage.googleapis.com/aw-gapps-ressources/at/addons/file-cabinet/assets/spreadsheets-128.png';
  }
  else if (mimeType.includes('presentation')) {
    mimeType = 'Presentation';
    imageURL = 'https://storage.googleapis.com/aw-gapps-ressources/at/addons/file-cabinet/assets/presentations-128.png';
  }
  else if (mimeType.includes('form')) {
    mimeType = 'Form';
  }
  else if(mimeType.includes('drawing')) {
    mimeType = 'Drawing';
    imageURL = 'https://storage.googleapis.com/aw-gapps-ressources/at/addons/file-cabinet/assets/drawings-128.png'
  }
  else if(mimeType.includes('shortcut')) {
    mimeType = 'Shortcut';
  }
  
  return [mimeType, imageURL];
}

// For the CDL purposes, this always returns "Undisclosed, undisclosed" so we will automatically return this to save time
function generateFileOwner(owner) {
  // If the file has an owner, fill in the appropriate owner email and name fields
  if (owner) {
    var email = owner.getEmail();
    var name = owner.getName();
    return [email, name];
  }
  return ["Undisclosed", "undisclosed"]; 
}

function getFolderPath(folder) {
  var path = "/" + folder.getName();
  try {
    var parent = folder.getParents().next();
  } catch (err) {
    console.error("Reached the end of the path: " + err);
  }
  if (parent) {
    path = getFolderPath(parent) + path;
  }
  
  return path;
}

function formatFileSize(bytes) {
  if (bytes === 0) return '0 Bytes';
  
  var sizes = ['Bytes', 'KB', 'MB', 'GB', 'TB'];
  var i = Math.floor(Math.log(bytes) / Math.log(1024));
  
  return parseFloat((bytes / Math.pow(1024, i)).toFixed(2)) + ' ' + sizes[i];
}

// For now we are not using generateContents since we do not want to store file contents in our sheet - takes too much time
function generateContents(fileId) {
  // Add in document contents if the file is a document
  try{
    var doc = DocumentApp.openById(fileId);
    var content = doc.getBody().getText();
    contents = content.split(/\s+/).slice(0, 5000).join(" ");
  } catch(err) {
    console.error("Error: File may not contain text: ", err);
  }
}

// For the CDL purposes, this function is unnecessary and complicated, taking more time so we will automatically return '' to save time
function getLastUserToUpdate(id) {
  var editList = [],
    revisions = Drive.Revisions.list(fileID);

  if (revisions.items && revisions.items.length > 0) {
    for (var i = 0; i < revisions.items.length; i++) {
      var revision = revisions.items[i];
      editList.push([
        revision.id,
        new Date(revision.modifiedDate).toLocaleString(),
        revision.lastModifyingUserName,
        revision.lastModifyingUser.emailAddress,
      ]);
    }
    Logger.log(editList);
  } else {
    Logger.log('No file revisions found.');
  }
  
  return '';
}

// This funciton was adapted from https://developers.google.com/apps-script/advanced/drive-labels#list_labels_for_a_item
function generateLabelsOnDriveItem(fileId) {
  var labelsString = "";
  try {
    const appliedLabels = Drive.Files.listLabels(fileId);

    //console.log('%d label(s) are applied to this file', appliedLabels.items.length);

    appliedLabels.items.forEach((appliedLabel) => {
      // Resource name of the label at the applied revision.
      const labelName = 'labels/' + appliedLabel.id + '@' + appliedLabel.revisionId;

      console.log('Fetching Label: %s', labelName);
      const label = DriveLabels.Labels.get(labelName, {view: 'LABEL_VIEW_FULL'});

      console.log('Label Title: %s', label.properties.title);

      Object.keys(appliedLabel.fields).forEach((fieldId) => {
        const fieldValue = appliedLabel.fields[fieldId];
        const field = label.fields.find((f) => f.id == fieldId);

        console.log(`Field ID: ${field.id}, Display Name: ${field.properties.displayName}`);
        switch (fieldValue.valueType) {
          case 'text':
            //console.log('Text: %s', fieldValue.text[0]);
            //labelsString += fieldValue.text[0];
            break;
          case 'integer':
            //console.log('Integer: %d', fieldValue.integer[0]);
            //labelsString += fieldValue.integer[0];
            break;
          case 'dateString':
            //console.log('Date: %s', fieldValue.dateString[0]);
            labelsString += fieldValue.dateString[0] + ", ";
            break;
          case 'user':
            const user = fieldValue.user.map((user) => {
              return `${user.emailAddress}: ${user.displayName}`;
            }).join(', ');
            //console.log(`User: ${user}`);
            labelsString += user + ", ";
            break;
          case 'selection':
            const choices = fieldValue.selection.map((choiceId) => {
              return field.selectionOptions.choices.find((choice) => choice.id === choiceId);
            });
            const selection = choices.map((choice) => {
              return `${field.properties.displayName.replace(/,/g, '')}: ${choice.properties.displayName.replace(/,/g, '')}`;
            }).join(', ');
            //console.log(`Selection: ${selection}`);
            labelsString += selection + ", ";
            break;
          default:
            //console.log('Unknown: %s', fieldValue.valueType);
            labelsString += fieldValue.value + ", ";
        }
      });
    });
  } catch (err) {
    console.log('Failed with error %s', err.message);
  }
  return labelsString.slice(0,-2); // Remove the last 2 characters ", " and return the list of labels in csv format
}
