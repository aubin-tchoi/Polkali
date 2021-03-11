// Basic functions used in other scripts in the same folder

// Converting an array into a (multi+indexed)-line str
function listToQuery(arr) {
    let query = "";
    arr.forEach(function (el, idx) {
        query += `\n ${idx + 1}. ${el}`
    });
    return query;
}

// Putting the subfolder of folder that has the name folderName in the trash
function trashFolder(folder, folderName, trash) {
    let folders = folder.getFolders();
    while (folders.hasNext()) {
        let subFolder = folders.next();
        if (subFolder.getName() == folderName) {
            subFolder.moveTo(trash);
        }
    }
}

// Retrieving an array of the names of the subfolders in folder
function getSubFoldersNames(folder) {
    let subFolders = folder.getFolders(),
        names = [];
    while (subFolders.hasNext()) {
        let subFolder = folders.next();
        names.push(subFolder.getName());
    }
    return names;
}

// Retrieving a subfolder of folder with its name
function getFolderByName(folder, folderName) {
    return folder.getFoldersByName(folderName).next();
}
