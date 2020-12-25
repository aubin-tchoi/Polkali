// Author : Pôle Qualité 022 (Aubin Tchoï)

function mailing() {
  // Displays a loading screen
  function display_LoadingScreen(msg) {
    let htmlLoading = HtmlService
    .createHtmlOutput(`<img src="https://raw.githubusercontent.com/aubin-tchoi/Polkali/main/images/Kkooljem.gif" alt="Loading" width="885" height="498">`)
    .setWidth(900)
    .setHeight(520);
    ui.showModelessDialog(htmlLoading, msg);
  }

  // Sends mails with subject subject using template template 
  function send_Emails(subject, row) {
    // Returns a boolean function meant to be used in a filter function
    function subjectFilter_(template){
      return function(element) {
        return element.getMessage().getSubject() === template;
      }
    }
    
    // Retrieving inline images inside of the html body
    function get_InlineImages(msg) {
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
    function get_Attachments(msgHtml) {
      let attachments = [],
        attachment_ids = (msgHtml.match(/{PJ=<a href="https:\/\/drive\.google\.com\/file\/d\/([^\/]+)\//g) || []);
      for (let j = 0; j < attachment_ids.length; j++) {
        attachments.push(DriveApp.getFileById(attachment_ids[j].match(/https:\/\/drive\.google\.com\/file\/d\/([^\/]+)/)[1]).getBlob());
      }
      return [msgHtml.replace(/{PJ=[^}]+}/g, ""), attachments];
    }

    // A mail is sent only if "Date d'envoi du mail" column is empty
    if (row["Date d'envoi du mail"] == "") {
      try {
        let msg = GmailApp.getDrafts().filter(subjectFilter_(template))[0].getMessage(),
            attachments = [];
        
        // Inline images
        let [msgHtml, inlineImages] = get_InlineImages(msg);

        // Attachments
        [msgHtml, attachments] = get_Attachments(msgHtml);

        // Customizing the model with data taken from row
        Object.keys(row).forEach(function(key) {
          let regexp = new RegExp(`{${key}}|{${key.toLowerCase()}}|{${key.replace(/[éêè]/gi, "e")}}`, "gim");
          msgHtml.replace(regexp, row[key]);});

        // Sending the mail
        let msgPlain = msgHtml.replace(/\<br\/\>/gim, '\n').replace(/(<([^>]+)>)/gim, "");
        GmailApp.sendEmail(row["Adresse mail"], subject, msgPlain, {htmlBody:msgHtml, attachments:attachments, inlineImages:inlineImages});
          
        // Keeping track of sent emails
        return [new Date()];
      }
      catch(e) {
        Logger.log(`Issue with row ${row} : ${e}`);
        return [(e.message == "Cannot read property 'getMessage' of undefined") ? "Pas de modèle à ce nom" : e.message];
      }
    } else {return [row["Date d'envoi du mail"]];}
  }

  // -- main --
  const ui = SpreadsheetApp.getUi(),
    sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();

  // Modify this part if the first line is not the header
  let data = sheet.getRange(1, 1, (sheet.getLastRow() - 1),sheet.getLastColumn()).getValues(),
      heads = data.shift(),
      data = data.map(r => (heads.reduce((o, k, i) => (o[k] = r[i] || '', o), {}))),
      out = [];

  // User input (subject of the mails and set of used templates)
  let subject = ui.prompt("Envoi automatique d'emails", "Entrez l'objet du mail", ui.ButtonSet.OK_CANCEL);
  if (subject.getSelectedButton() == ui.Button.CANCEL || subject.getResponseText() == "") {return;}
  subject = subject.getResponseText();

  let templates = ui.prompt("Envoi automatique d'emails",
                            "Entrez les noms des templates que vous souhaitez utiliser en les séparant d'une virgule et d'un espace (ex: 'template_1, template_2'). \n Si vous souhaitez couvrir tous les templates, entrez all",
                            ui.ButtonSet.OK_CANCEL);
  if (templates.getSelectedButton() == ui.Button.CANCEL || templates.getResponseText() == "") {return;}
  templates = templates.getResponseText().split(", ")

  // Confirmation
  let count = data.filter(row => (templates.includes(row["Template"]) || templates[0].toLowerCase() == "all") && row["Date d'envoi du mail"] == "").length,
    confirm = ui.alert("Envoi automatique d'emails", `${count} mail${(count >= 2) ? "s vont être envoyés" : " va être envoyé"}.`, ui.ButtonSet.OK_CANCEL);
  if (confirm.getSelectedButton() == ui.Button.CANCEL) {return;}

  // Loading screen
  display_LoadingScreen("Envoi des mails...");
  
  // Sending the emails
  obj.forEach(function(row) {
    if (templates.includes(row["Template"]) || templates[0].toLowerCase() == "all") {out.push(send_Emails(subject, row));}});

  // Completing column "Date d'envoi du mail" with out
  sheet.getRange(2, (heads.indexOf("Date d'envoi du mail") + 1), out.length, 1).setValues(out);
  
  // Display box to notify of how many mails were sent
  ui.alert("Mails envoyés", `${count} mail${(count >= 2) ? "s ont été envoyés" : " a été envoyé"}.`, ui.ButtonSet.OK);
}
