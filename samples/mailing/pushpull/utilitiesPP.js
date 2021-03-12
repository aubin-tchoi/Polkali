// Core functions

// Pushing GmailDraft draft into Folder folder
function pushDraftIntoFolder(draft, folder) {
    let msgBody = draft.getMessage().getBody(),
        msgAttachments = draft.getMessage().getAttachments();

    // Create the message PDF inside the new folder
    let htmlBodyFile = folder.createFile('body.html', msgBody, "text/html"),
        pdfBlob = htmlBodyFile.getAs('application/pdf');
    pdfBlob.setName(folder.getName() + ".pdf");
    folder.createFile(pdfBlob);

    // Supprimer le pdf

    for (let j = 0; j < msgAttachments.length; j++) {
        let attachmentName = msgAttachments[j].getName(),
            attachmentContentType = msgAttachments[j].getContentType(),
            attachmentBlob = msgAttachments[j].copyBlob();
        folder.createFile(attachmentBlob);
    }
}

// Pulling data from folder and using it to create and save a GmailDraft
function pullDraftFromFolder(folder) {

}
