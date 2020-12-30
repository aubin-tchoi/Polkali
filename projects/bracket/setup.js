/* For any further question, please contact Aubin Tchoï */

// Formats active sheet to make it useable by this sheet
const mailPos = [3, 2],
  mailHeader = [["Date d'envoi",	"Prénom", "Nom", "Adresse mail"]],
  templatePos = [1, 1],
  resultsPos = [3, 7];

function helper() {
  let sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  sheet.setName("Dashboard");
  sheet.getRange(mailPos[0], mailPos[1], 1, 4)
    .setBackground(MARKERS["mail"])
    .setValue("Mailing list")
    .merge()
    .setHorizontalAlignment("center");
  sheet.getRange((mailPos[0] + 1), mailPos[1], 1, 4)
    .setBackground(COLORS["contestants"])
    .setValues(mailHeader);
  sheet.getRange(templatePos[0], templatePos[1])
    .setBackground(MARKERS["template"])
    .setValue("INSERT THE TEMPLATE'S NAME HERE");
  sheet.getRange(resultsPos[0], resultsPos[1], 2, 1)
    .setBackgrounds([[MARKERS["groups"]], [MARKERS["currentRound"]]])
    .setValues([["Poule"], ["Numéro"]]);
}
