function runEvery30Minutes() {
  var maxFiles = 500;
  populateDataWithFileCount(maxFiles);
}

function recordAllFolders() {
  // Step one: Record all of the folders
  // TODO: Replace all instances of getActiveSpreadsheet with getSpreadsheetById
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Savedata");
  var folder = DriveApp.getFolderById("0ANEBHxLVNxEHUk9PVA");

  // Clear current content
  var rowCount = sheet.getRange("B1").getValue();
  sheet.getRange(`A3:E${rowCount}`).clearContent();

  // Populate all the rows for folders and subfolders in Savedata sheet, and then record the rowCount
  rowCount = populateSavedataFolders(sheet, folder);
  sheet.getRange("B1").setValue(rowCount);

  console.log("Added " + rowCount + " folders to the Savedata sheet")
}

function populateOneFolderData() {
  // Step two: Populate the data for one folder at a time
  var spreadsheet = Spreadsheet.openById("1jfQDiUfWUQCFBn_kUMB-1OsZYvr54CBrgErfYVI-MVs");
  var savedataSheet = spreadsheet.getSheetByName("Savedata");
  var dataSheet = spreadsheet.getSheetByName("Data");
  var lastRow = savedataSheet.getLastRow();

  // Collect Folder (current) information
  do {
    var currentRow = savedataSheet.getRange("A3:E3").getValues()[0];
    var folderId = currentRow[0];
    var folderPath = currentRow[1];
    var lastUpdated = currentRow[2];
    var numFiles = currentRow[3];
    currentRow[4] = new Date().toLocaleDateString();
    savedataSheet.deleteRow(3);
    savedataSheet.getRange(lastRow, 1, 1, currentRow.length).setValues([currentRow]); // Use sheet.getRange().setValues() here instead of sheet.appendRow() as it is MUCH more efficient on time
  } while (numFiles <= 0)
  console.log("Found a folder containing files: " + folderPath);

  var folder = DriveApp.getFolderById(folderId);

  // Get the data from the savedata sheet to dive into the desired folder


  // Create the object containing all current files from the drive
  var currentIdsAndLastUpdated = getAllFilesIdAndLastUpdated(folder);

  // Search through and traverse all existing file records from the sheet to compare, update, delete, and add
  compareAndUpdateSheet(dataSheet, currentIdsAndLastUpdated, folderPath);
}

function populateDataWithFileCount(filecount = 500) {
  // Step two: Populate the data for one folder at a time
  var spreadsheet = Spreadsheet.openById("1jfQDiUfWUQCFBn_kUMB-1OsZYvr54CBrgErfYVI-MVs");
  var savedataSheet = spreadsheet.getSheetByName("Savedata");
  var dataSheet = spreadsheet.getSheetByName("Data");
  var lastRow = savedataSheet.getLastRow();
  var filesToLoad = 0;
  var foldersToLoad = [];

  // Loop through the queue of folders in Savedata that equal up to the filecount and create a list of folders to be loaded
  for (var i = 3; filesToLoad <= filecount; i++) {
    let currentRow = savedataSheet.getRange(`A${i}:E${i}`).getValues()[0];
    let folderId = currentRow[0];
    let folderPath = currentRow[1];
    let lastUpdated = currentRow[2];
    let numFiles = currentRow[3];
    if (filesToLoad + numFiles >= filecount) {
      console.log("Total files to load: " + filesToLoad + " out of " + filecount);
      break;
    }
    if (numFiles > 0) {
      foldersToLoad.push(currentRow);
      filesToLoad += numFiles;
    } else {
      currentRow[4] = new Date().toLocaleDateString();
      savedataSheet.deleteRow(i);
      i--;
      savedataSheet.getRange(lastRow, 1, 1, currentRow.length).setValues([currentRow]); // Use sheet.getRange().setValues() here instead of sheet.appendRow() as it is MUCH more efficient on time
    }
  }

  // Create the object containing all current files from the drive
  var currentIds = getListOfFileIds(foldersToLoad);

  // Search through and traverse all existing file records from the sheet to compare, update, delete, and add
  compareAndUpdateSheet2(dataSheet, savedataSheet, currentIds, foldersToLoad, lastRow);

  // TODO: If the folder chosen is the root drive folder, then also check to make sure there aren't any files in the sheet that have a folderpath that doesn't exist in the savedata sheet

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

function getAllFilesIdAndLastUpdated(folder) {
  var fileData = {};

  var fileIterator = folder.getFiles();
  while (fileIterator.hasNext()) {
    var file = fileIterator.next();
    fileData[file.getId()] = file.getLastUpdated();
  }
  
  console.log('Found ', Object.keys(fileData).length, ' files');
  return fileData;
}

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
  
  console.log('Found ', Object.keys(fileData).length, ' files');
  return fileData;
}

function compareAndUpdateSheet(sheet, currentFiles, folderPath) {
  // Find the starting row of where the files in our folder exist in our sheet
  var startRow = 3; // Start reading data from row 3 (index 2)
  var lastRow = sheet.getLastRow();

  // Go through every file in the sheet looking for files with a matching folder
  for(var i = startRow; i <= lastRow; i++) {
    var iPath = sheet.getRange(`J${i}`).getValue();
    
    // Skip any files which do not belong to the current folder
    if (folderPath != iPath) {
      continue;
    }
    // Assume every file that gets to this point has a matching folder path
    var id = sheet.getRange(`B${i}`).getValue();
    var lastUpdated = sheet.getRange(`I${i}`).getValue();

    if (!currentFiles[id]) {
      console.log("Found file in sheet that does not exist in Drive, deleting: " + id + " " + lastUpdated);
      sheet.deleteRow(i);
      lastRow--;
    } else {
      // Check if the dates are not the same based on the Date property of the Date objects
      if (currentFiles[id].getDate() !== lastUpdated.getDate()) {
        console.log("Found outdated file in sheet, updating: " + id + " " + lastUpdated);
        // Update the file here
        let rowData = getFileDetails(id, lastUpdated, folderPath);
        sheet.getRange(i, 1, 1, rowData.length).setValues([rowData]);
      } else {
        console.log("Found matching file, skipping: " + id + " " + lastUpdated);
      }
      // Remove file from dictionary here
      delete currentFiles[id];
    }
  }
  // Iterating through the rest of the remaining files in the currentFiles dictionary to add them to the sheet
  for(var fileId in currentFiles) {
    if (currentFiles.hasOwnProperty(fileId)) {
      let lastUpdated = currentFiles[fileId];
      console.log("Found file that does not exist in the sheet, adding: " + lastUpdated);
      
      let rowData = getFileDetails(fileId, lastUpdated, folderPath);
      sheet.getRange(++lastRow, 1, 1, rowData.length).setValues([rowData]); // Use sheet.getRange().setValues() here instead of sheet.appendRow() as it is MUCH more efficient on time
    }
  }
}

function compareAndUpdateSheet2(dataSheet, savedataSheet, currentFiles, foldersToLoad, savedataLastRow) {
  // Find the starting row of where the files in our folder exist in our sheet
  var startRow = 3; // Start reading data from row 3 (index 2)
  var lastRow = dataSheet.getLastRow();
  
  var foldersList = {};
  for (var i = 0; i < foldersToLoad.length; i++) {
    let folderPath = foldersToLoad[i][1];
    foldersList[folderPath] = foldersToLoad[i][3];
  }

  // Go through every file in the dataSheet looking for files with a matching folder
  console.log("Beginning search through existing sheet");
  for(var i = startRow; i <= lastRow; i++) {
    var iPath = dataSheet.getRange(`J${i}`).getValue();
    
    // Skip any files which do not belong to the current folder list
    if (!foldersList[iPath]) {
      continue;
    }
    // Assume every file that gets to this point has a matching folder path
    var id = dataSheet.getRange(`B${i}`).getValue();
    var lastUpdated = dataSheet.getRange(`I${i}`).getValue();

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
    foldersList[iPath]--;
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
      console.log("Found file that does not exist in the sheet, adding: " + lastUpdated);
      
      let rowData = getFileDetails(fileId, lastUpdated, folderPath);
      dataSheet.getRange(++lastRow, 1, 1, rowData.length).setValues([rowData]); // Use dataSheet.getRange().setValues() here instead of dataSheet.appendRow() as it is MUCH more efficient on time
      foldersList[folderPath]--;
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
  let startTime = new Date().getTime();
  const labels = generateLabelsOnDriveItem(id);
  let endTime = new Date().getTime();
  let executionTime = endTime - startTime;
  console.log("Time to get labels for 1 file: " + executionTime + " milliseconds");

  fileInfo.push(title, id, url, image, mimeType, ownerInfo, owner, creationDate, lastUpdated, filePath, description, fileSize, "", "", contents, lastModifiedBy, "", "", labels);
  
  return fileInfo;
}

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




  // var response = Drive.Files.get(fileId, { fields: "lastModifyingUser" });
  // var lastModifyingUser = response.lastModifyingUser;
  
  // var displayName = lastModifyingUser.displayName;
  // var email = lastModifyingUser.emailAddress;
  
  // return {
  //   displayName: displayName,
  //   email: email
  // };


  
  // var revisions = Drive.Revisions.list(file.getId());
  // var lastRevision = revisions.items[0];
  
  // if (lastRevision) {
  //   var lastUserEmail = lastRevision.lastModifyingUser.emailAddress;
  //   return lastUserEmail;
  // }
  
  return '';
}

function generateLabelsOnDriveItem(fileId) {
  var labelsString = "";
  try {
    const appliedLabels = Drive.Files.listLabels(fileId);

    console.log('%d label(s) are applied to this file', appliedLabels.items.length);

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
            console.log('Text: %s', fieldValue.text[0]);
            //labelsString += fieldValue.text[0];
            break;
          case 'integer':
            console.log('Integer: %d', fieldValue.integer[0]);
            //labelsString += fieldValue.integer[0];
            break;
          case 'dateString':
            console.log('Date: %s', fieldValue.dateString[0]);
            labelsString += fieldValue.dateString[0] + ", ";
            break;
          case 'user':
            const user = fieldValue.user.map((user) => {
              return `${user.emailAddress}: ${user.displayName}`;
            }).join(', ');
            console.log(`User: ${user}`);
            labelsString += user + ", ";
            break;
          case 'selection':
            const choices = fieldValue.selection.map((choiceId) => {
              return field.selectionOptions.choices.find((choice) => choice.id === choiceId);
            });
            const selection = choices.map((choice) => {
              return `${field.properties.displayName.replace(/,/g, '')}: ${choice.properties.displayName.replace(/,/g, '')}`;
            }).join(', ');
            console.log(`Selection: ${selection}`);
            labelsString += selection + ", ";
            break;
          default:
            console.log('Unknown: %s', fieldValue.valueType);
            console.log(fieldValue.value);
            labelsString += fieldValue.value + ", ";
        }
      });
    });
  } catch (err) {
    // TODO (developer) - Handle exception
    console.log('Failed with error %s', err.message);
  }
  return labelsString.slice(0,-2);
}
