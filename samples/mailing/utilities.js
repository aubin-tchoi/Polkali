// Core function

// Sends mails with subject subject using template template 
function sendEmails(subject, row) {
    // A mail is sent only if "Date d'envoi du mail" column is empty
    if (row["Date d'envoi du mail"] == "") {
        try {
            let msg = GmailApp.getDrafts().filter(subjectFilter(template))[0].getMessage(),
                attachments = [];

            // Inline images
            let [msgHtml, inlineImages] = getInlineImages(msg);

            // Attachments
            [msgHtml, attachments] = getAttachments(msgHtml);

            // Customizing the model with data taken from row
            Object.keys(row).forEach(function (key) {
                let regexp = new RegExp(`{${key}}|{${key.toLowerCase()}}|{${key.replace(/[éêè]/gi, "e")}}`, "gim");
                msgHtml.replace(regexp, row[key]);
            });

            // Sending the mail
            let msgPlain = msgHtml.replace(/\<br\/\>/gim, '\n').replace(/(<([^>]+)>)/gim, "");
            GmailApp.sendEmail(row["Adresse mail"], subject, msgPlain, {
                htmlBody: msgHtml,
                attachments: attachments,
                inlineImages: inlineImages
            });

            // Keeping track of sent emails
            return [new Date()];
        } catch (e) {
            Logger.log(`Issue with row ${row} : ${e}`);
            return [(e.message == "Cannot read property 'getMessage' of undefined") ? "Pas de modèle à ce nom" : e.message];
        }
    } else {
        return [row["Date d'envoi du mail"]];
    }
}
