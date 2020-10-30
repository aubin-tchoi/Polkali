// Author : Pôle Qualité 022 (Aubin Tchoï)

function mailing() {
  function sendEmails(template, subject, sheet, row, out, categories) {
    
    // Counting the number of mails sent (for the dialog box at the end)
    let c = 0;
    
    function subjectFilter_(template){
      return function(element) {
        return element.getMessage().getSubject() === template;
      }
    }
    
    if ((categories.includes(template) || categories[0].toLowerCase() == "all") && row["Date d'envoi du mail"] == "") {
      try {
        c++;
        let msg = GmailApp.getDrafts().filter(subjectFilter_(template))[0].getMessage(),
            msgHtml = msg.getBody(),
            inlineImages = {},
            imgTags = (msgHtml.match(/<img[^>]+>/g) || []),
            rawc = msg.getRawContent(),
            attachments = [];
        
        // Inline images
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
        
        // Customize with name & title
        msgHtml = msgHtml
        .replace(/{Prénom}|{prénom}|{prenom}|{Prenom}/g, row["Prénom"])
        .replace(/{Nom}|{nom}/g, row["Nom"])
        .replace(/{Titre}|{titre}/g, (row["Genre"] == "Femme") ? "Madame" : (row["Genre"] == "Homme") ? "Monsieur" : row["Sexe"]);
        
        // Attachments
        let attachment_ids = (msgHtml.match(/{PJ=<a href="https:\/\/drive\.google\.com\/file\/d\/([^\/]+)\//g) || []);
        for (let j = 0; j < attachment_ids.length; j++) {
          attachments.push(DriveApp.getFileById(attachment_ids[j].match(/https:\/\/drive\.google\.com\/file\/d\/([^\/]+)/)[1]).getBlob());
        }
        msgHtml = msgHtml.replace(/{PJ=[^}]+}/g, "");
        
        // Sending the mail
        let msgPlain = msgHtml.replace(/\<br\/\>/gi, '\n').replace(/(<([^>]+)>)/gi, "");
        GmailApp.sendEmail(row["Adresse mail"], subject, msgPlain, {htmlBody:msgHtml, attachments:attachments, inlineImages:inlineImages});
        
        // Keeping track of the mail sent data
        out.push([new Date()]);
        
      } catch(e) {
        out.push([(e.message == "Cannot read property 'getMessage' of undefined") ? "Pas de modèle à ce nom" : e.message]);
        Logger.log(e);
      }
    } else {out.push([row["Date d'envoi du mail"]]);}
    return c;
  }
  
  subject = Browser.inputBox("Envoi automatique d'email", 
                             "Entrez l'objet du mail :",
                             Browser.Buttons.OK_CANCEL);
  if (subject === "cancel" || subject == "") {
    return;
  }
  
  roles = Browser.inputBox("Envoi automatique d'email", 
                           "Entrez les différents rôles en les séparants d'une virgule (Ex : Rôle 1, Rôle 2). \\n Si vous souhaitez couvrir la totalité des rôles existants, entre 'All' (sans guillemet) :",
                           Browser.Buttons.OK_CANCEL);
  if (roles === "cancel" || roles == "") {
    return;
  }
  
  let ui = SpreadsheetApp.getUi(),
      htmlLoading = HtmlService
  .createHtmlOutput(`<img src="https://raw.githubusercontent.com/aubin-tchoi/Polkali/main/images/Kkooljem.gif" alt="Loading" width="885" height="498">`)
  .setWidth(900)
  .setHeight(520);
  
  ui.showModelessDialog(htmlLoading, "Envoi des mails..");
  
  const sheet = SpreadsheetApp.openById("1iXJalZxTTBfOfbHvAZz0e1kIGd8REmtHA55sMi0bK3Y").getSheetByName("Envoi de mail");
  let data = sheet.getRange(2, 1, (sheet.getLastRow() - 1),sheet.getLastColumn()).getValues(),
      heads = data.shift(),
      obj = data.map(r => (heads.reduce((o, k, i) => (o[k] = r[i] || '', o), {}))),
      out = [],
      categories = roles.split(", "),
      count = 0;
  
  obj.forEach(function(row) {count += sendEmails(row["Rôles"], subject, sheet, row, out, categories);})
  sheet.getRange(3, heads.indexOf("Date d'envoi du mail") + 1, out.length).setValues(out);
  
  ui.alert("Mails envoyés", `${count} mail${(count >= 2) ? "s ont été envoyés" : " a été envoyé"}.`, ui.ButtonSet.OK);
}
