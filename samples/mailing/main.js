/* Si vous avez des questions à propos de ce script contactez Aubin Tchoï (Directeur Qualité 022) */

// Uses function displayLoadingScreen taken from ressources.js

// This script is written for a sheet that contains a header in line 1 and data starting from line 2

// Data sheet's header
const HEADS = {
  template: "Template",
  sendDate: "Date d'envoi du mail"
}

function mailing() {
  const ui = SpreadsheetApp.getUi(),
    sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();

  // Modify this part if the first line is not the header
  let data = sheet.getRange(1, 1, (sheet.getLastRow() - 1), sheet.getLastColumn()).getValues(),
    heads = data.shift(),
    data = data.map(r => (heads.reduce((o, k, i) => (o[k] = r[i] || '', o), {}))),
    out = [];

  // User input (subject of the mails and set of used templates)
  let subject = ui.prompt("Envoi automatique d'emails", "Entrez l'objet du mail", ui.ButtonSet.OK_CANCEL);
  if (subject.getSelectedButton() == ui.Button.CANCEL || subject.getResponseText() == "") {
    return;
  }
  subject = subject.getResponseText();

  let templates = ui.prompt("Envoi automatique d'emails",
    "Entrez les noms des templates que vous souhaitez utiliser en les séparant d'une virgule et d'un espace (ex: 'template1, template2'). \n Si vous souhaitez couvrir tous les templates, entrez 'all'",
    ui.ButtonSet.OK_CANCEL);
  if (templates.getSelectedButton() == ui.Button.CANCEL || templates.getResponseText() == "") {
    return;
  }
  templates = templates.getResponseText().split(", ")

  // Confirmation
  let count = data.filter(row => (templates.includes(row[HEADS["template"]]) || templates[0].toLowerCase() == "all") && row[HEADS["sendDate"]] == "").length,
    confirm = ui.alert("Envoi automatique d'emails", `${count} mail${(count >= 2) ? "s vont être envoyés" : " va être envoyé"}.`, ui.ButtonSet.OK_CANCEL);
  if (confirm.getSelectedButton() == ui.Button.CANCEL) {
    return;
  }

  // Loading screen
  displayLoadingScreen("Envoi des mails...");

  // Sending the emails
  obj.forEach(function (row) {
    if (templates.includes(row[HEADS["template"]]) || templates[0].toLowerCase() == "all") {
      out.push(sendEmails(subject, row));
    } else {
      out.push(row[HEADS["sendDate"]]);
    }
  });

  // Completing column "Date d'envoi du mail" with out
  sheet.getRange(2, (heads.indexOf(HEADS["sendDate"]) + 1), out.length, 1).setValues(out);

  // Display box to notify of how many mails were sent
  ui.alert("Mails envoyés", `${count} mail${(count >= 2) ? "s ont été envoyés" : " a été envoyé"}.`, ui.ButtonSet.OK);
}
