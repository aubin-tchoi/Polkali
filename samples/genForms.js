/* Si vous avez des questions à propos de ce script contactez Aubin Tchoï (Directeur Qualité 022) */

// Generating a Google Forms with data taken from the Sheets
function genForms() {
  // The position of the cells that contains the informations about the description of the sheet and that of the confirmation message are hard coded
  const ui = SpreadsheetApp.getUi(),
    sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Feuille 1"),
    data = sheet.getRange(2, 2, (sheet.getLastRow() - 1), 3).getValues(),
    heads = data.shift(),
    obj = data.map(r => (heads.reduce((o, k, i) => (o[k] = r[i] || '', o), {}))),
    forms = FormApp.create(`Prêt de ${sheet.getRange(3, 9, 1, 1).getValues()[0][0].toLowerCase()}`)
    .setDescription(sheet.getRange(4, 9, 1, 1).getValues()[0][0])
    .setConfirmationMessage(sheet.getRange(5, 9, 1, 1).getValues()[0][0]);

  // First questions
  forms.addTextItem()
    .setTitle('Quel est votre prénom ?')
    .setRequired(true),
    forms.addTextItem()
    .setTitle('Quel est votre nom ?')
    .setRequired(true),
    forms.addTextItem()
    .setTitle('Quel est votre numéro de portable ?')
    .setRequired(true);

  // That is gonna be the "main" question (each answer is linked to a distinct section)
  var item = forms.addListItem()
    .setTitle('Que souhaitez vous emprunter ?')
    .setRequired(true),
    itemChoices = [];

  Logger.log(obj); // Data format is similar to a JSON

  // Each answer to the question "Que souhaitez vous emprunter" generates a distinct section
  obj.forEach(function (row) {
    // new section here + 2 additional questions before forms submission
    // You can easily add a description for each item, you just need to add a column with that information and read it here (see : https://developers.google.com/apps-script/reference/forms/page-break-item?hl=en#sethelptexttext)
    let sec = forms.addPageBreakItem()
      .setTitle(row["Bien prêté"])
      .setGoToPage(FormApp.PageNavigationType.SUBMIT),

      hor = forms.addCheckboxItem()
      .setTitle('Quand en aurez vous besoin ?')
      .setRequired(true);

    forms.addTextItem()
      .setTitle('Une remarque en particulier ?'),

      // As seen in the following regexes, the first 2 numbers in column C will be recognized as the start and the end of the time period.
      startHour = Math.floor(row["Plage horaire"].match(/([0-9]+)h/gm)[0].match(/[0-9]+/gm)[0]),
      endHour = Math.floor(row["Plage horaire"].match(/([0-9]+)h/gm)[1].match(/[0-9]+/gm)[0]),
      timeSlot = Math.floor(row["Durée d'un créneau"].match(/([0-9]+)/gm)[0]);

    Logger.log(`Bien prêté : row["Bien prêté"]`);
    Logger.log(`Heure début : ${startHour}`);
    Logger.log(`Heure fin : ${endHour}`);
    Logger.log(`Durée créneau : ${timeSlot}`);

    // Hour slots are created by splitting the time period in slots each timeSlot long
    let horChoices = [];
    for (let h_n = startHour; h_n < endHour; h_n += timeSlot) {
      horChoices.push(hor.createChoice(`${h_n}h à ${Math.min(h_n + timeSlot, endHour)}h`));
    }
    hor.setChoices(horChoices);

    // The item is added to a list that will become the list of possible choices to the question "Que souhaitez vous emprunter ?"
    itemChoices.push(item.createChoice(row["Bien prêté"], sec));
    item.setChoices(itemChoices);
  });

  // Setting up the active spreadsheet as a destination in order to be able to update the form.
  forms.setDestination(FormApp.DestinationType.SPREADSHEET, SpreadsheetApp.getActiveSpreadsheet().getId());
  sheet.getRange(2, 9, 1, 1).setValues([
    [forms.getId()]
  ]).setFontColor('white'); // That is not pretty but it will be useful each time the form is updated

  // This trigger updates the form after every submission by removing the slots chosen
  ScriptApp.newTrigger('updateForm')
    .forSpreadsheet(SpreadsheetApp.getActiveSpreadsheet())
    .onFormSubmit()
    .create();

  // Dialog box to print the links (one editor link and a reader link)
  let htmlOutput = HtmlService
    .createHtmlOutput(`<span style="font-family: 'trebuchet ms', sans-serif;">Voci le lien éditeur : <a href = "${forms.getEditUrl()}">${forms.getEditUrl()}</a><br/><br/> Voici le lien lecteur : <a href = "${forms.getPublishedUrl()}">${forms.getPublishedUrl()}</a>.<br/><br/><br/>&nbsp; La bise.<br/>`);
  ui.showModelessDialog(htmlOutput, "Un Google Forms a été créé.")
}

// Updating the Google Forms by removing occupied slots
function updateForm() {
  sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Feuille 1"),
    data = sheet.getRange(2, 2, (sheet.getLastRow() - 1), 3).getValues(),
    heads = data.shift(),
    obj = data.map(r => (heads.reduce((o, k, i) => (o[k] = r[i] || '', o), {})));
  if (sheet.getRange(2, 9, 1, 1).getValues()[0][0] == "") {
    return;
  }
  const forms = FormApp.openById(sheet.getRange(2, 9, 1, 1).getValues()[0][0]),
    responses = forms.getResponses(),
    lastResponse = responses[responses.length - 1].getItemResponses();

  lastResponse
    .filter(response => response.getItem().getTitle() == 'Quand en aurez vous besoin ?')
    .forEach(function (resp) {
      let oldChoices = forms.getItemById(resp.getItem().getId()).asCheckboxItem().getChoices(),
        newChoices = oldChoices.filter(choice => !resp.getResponse().includes(choice.getValue()));
      forms.getItemById(resp.getItem().getId()).asCheckboxItem().setChoices(newChoices);
    });
}
