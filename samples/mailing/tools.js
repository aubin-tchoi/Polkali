// Returns a boolean function meant to be used in a filter function
function subjectFilter(template) {
    return function (element) {
        return element.getMessage().getSubject() === template;
    }
}

// Retrieving inline images inside of the html body
function getInlineImages(msg) {
    let msgHtml = msg.getBody(),
        rawc = msg.getRawContent(),
        imgTags = (msgHtml.match(/<img[^>]+>/g) || []),
        inlineImages = {};
    for (let i = 0; i < imgTags.length; i++) {
        let realattid = imgTags[i].match(/cid:(.*?)"/i);
        if (realattid) {
            let cid = realattid[1];
            imgTagNew = imgTags[i].replace(/src="[^\"]+\"/, "src=\"cid:" + cid + "\"");
            msgHtml = msgHtml.replace(imgTags[i], imgTagNew);
            let b64c1 = (rawc.lastIndexOf(cid) + cid.length + 3),
                b64cn = (rawc.substring(b64c1).indexOf("--") - 3),
                imgb64 = rawc.substring(b64c1, b64c1 + b64cn + 1),
                imgblob = Utilities.newBlob(Utilities.base64Decode(imgb64), "image/jpeg", cid);
            inlineImages[cid] = imgblob;
        }
    }
    return [msgHtml, inlineImages];
}

// Retrieving attachments inside of the html body
function getAttachments(msgHtml) {
    let attachments = [],
        attachmentIds = (msgHtml.match(/{PJ=<a href="https:\/\/drive\.google\.com\/file\/d\/([^\/]+)\//g) || []);
    for (let j = 0; j < attachmentIds.length; j++) {
        attachments.push(DriveApp.getFileById(attachmentIds[j].match(/https:\/\/drive\.google\.com\/file\/d\/([^\/]+)/)[1]).getBlob());
    }
    return [msgHtml.replace(/{PJ=[^}]+}/g, ""), attachments];
}

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
    let folders = folder.getFolders();
    while (folders.hasNext()) {
        let subFolder = folders.next();
        if (subFolder.getName() == folderName) {
            return subFolder;
        }
    }
}
