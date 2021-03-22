// Core functions

// Pushing GmailDraft draft into Folder folder
function pushDraftIntoFolder(draft, folder) {
    let msgBody = draft.getMessage().getBody(),
        msgAttachments = draft.getMessage().getAttachments();

    // Create the message PDF inside the new folder
    let htmlBodyFile = folder.createFile('body.html', msgBody, "text/html");

    // Attachments should be separated from inlineImages

    for (let j = 0; j < msgAttachments.length; j++) {
        let attachmentBlob = msgAttachments[j].copyBlob();
        folder.createFile(attachmentBlob);
    }
}

// Pulling data from folder and using it to create and save a GmailDraft
function pullDraftFromFolder(folder) {
    let blob = folder.getFilesByName("body.html").next(),
        body = blob.getDataAsString(),
        subject = folder.getName(),
        attachments = getData(folder);
    GmailApp.createDraft("", subject, body, {attachments: attachments});
}
