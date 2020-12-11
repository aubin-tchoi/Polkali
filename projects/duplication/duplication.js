/* Si vous avez des questions à propos de ce script contactez Aubin Tchoï (Directeur Qualité 022) */

function duplicate_folder() {
    const folder = DriveApp.getFolderById("1ZnkIXGncM7g5J23rVSEMZ-nOqu80itaM"),
    destination = DriveApp.getFolderById("1qITNGvF2upvvh52z6bmq2f4G0eKbdlq5"),
    dest_list = ["Antoine", "Mai-Linh", "Loïs", "Arnaud", "Téo", "Nathan", "Caroline", "Thibaut"];
    dest_list.forEach(function(name) {
        let newFolder = destination.createFolder(name);
        let folders = folder.getFolders();
        while (folders.hasNext()) {
            let subFolder = folders.next();
            let newSubFolder = newFolder.createFolder(subFolder.getName());
            let files = subFolder.getFiles();
            while (files.hasNext()) {
              try {
                var targetFile = files.next();
                targetFile.makeCopy(targetFile.getName(), newSubFolder);
              } catch(e) {Logger.log(`Error with file ${targetFile.getName()}`);}
            }
         }
      })
  }
  