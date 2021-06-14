/** Copying all the files from source to target (source and target are instances of Folder class) */
function copyFiles(source, target) {
  let files = source.getFiles();
  while (files.hasNext()) {
    let file = files.next();
    file.makeCopy(file.getName().replace("Copie de ", ""), target);
  }
}

/** Recursive function to copy all files from source to target and then iterate on each subFolder */
function duplicateFolder(source, target) {
  copyFiles(source, target);
  // Looping on each subFolder (no need for a stopping condition, the hasNext method provides it)
  let folders = source.getFolders();
  while (folders.hasNext()) {
    let subFolder = folders.next(),
      targetSubFolder = target.createFolder(subFolder.getName());
    duplicateFolder(subFolder, targetSubFolder);
  }
}

/** Extracting an ID from an URL */
function URLtoID(url) {
  return url.replace(/\?.+/, "").split("/").slice(-1)[0];
}
