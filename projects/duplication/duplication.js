/* Si vous avez des questions à propos de ce script contactez Aubin Tchoï (Directeur Qualité 022) */

// IDs of the 2 folders to duplicate
const SOURCE_ID = "",
  TARGET_ID = "";

// Copying all the files from source to target (source and target are instances of Folder class)
function copyFiles(source, target) {
  let files = source.getFiles();
  while (files.hasNext()) {
    let file = files.next();
    file.makeCopy(target);
  }
}

// Recursive function to copy all files from source to target and then iterate on each subFolder
function duplicateFolder(source, target) {
  copyFiles(source, target);
  // Looping on each subFolder (no need for a stopping condition, the hasNext method provides it)
  let folders = folder.getFolders();
  while (folders.hasNext()) {
    let subFolder = folders.next(),
      targetSubFolder = target.createFolder(subFolder.getName());
    duplicateFolder(subFolder, targetSubFolder);
  }
}

// Main function
function main() {
  const source = DriveApp.getFolderById(sourceID),
    target = DriveApp.getFolderById(targetID);
  duplicateFolder(source, target);
}