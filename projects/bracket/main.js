/* Si vous avez des questions à propos de ce script, contactez Aubin Tchoï (Directeur Qualité 022)   */

/* Must be bound to a spreadsheet that contains a sheet name "Dashboard" in which are given a list of mail adresses
  The script uses specific colors called markers to locate important cells */

/* There should be 1 folder for each group */

const GAME_PARAMETERS = {
    pointsPool = 10,
    winnersNumber = 2
  },
  IDS = {
    folder = "15CM8Q2M8t4PhRyjZqFLq8bCg4J0iU6ik",
    trash = "16_JRrlTut5prU1z6hpi1pSy42BOMR-t1",
    dashboard = "11gnPSgy85927z95yxoXb_iS6B2YNoTZumkQooWndjgE",
  },

  MARKERS = {
    currentRound = "#ff00ff",
    nextRound = "#d150dd",
    groups = "#f1c232",
    mail = "#9900ff",
    template = "#c27ba0"
  },
  COLORS = {
    winner: "#00ff00",
    tie: "#ea9999",
    loser: "#d9ead3",
    groups: "#ffe599",
    contestants: "#fff2cc"
  },
  FORMS_PARAMETERS = {
    description = `Bienvenue dans le bracket, pour chaque poule vous disposez de ${POINTS_POOL} points à répartir sur l'ensemble des candidats. Soyez avisés.`,
    confirmationMessage = `Merci pour votre participation, la direction du Bracket vous recontactera sous peu pour annoncer les résultats.`,
    imageSize = 300
  };

function onOpen() {
  const ui = SpreadsheetApp.getUi();
  ui.createMenu('Bracket')
    .addItem("Nouveau round", "newRound")
    .addSeparator()
    .addItem("Match nul", "tie")
    .addToUi();
}

// Creating a new Form linked to a spreadsheet that will automatically update the dashboard and sending an email to our voters
function newRound() {
  const sheet = SpreadsheetApp.openById(IDS["dashboard"]).getSheetByName("Dashboard"),
    ui = SpreadsheetApp.getUi(),
    roundNumber = findRoundNumber(sheet);

  if (roundNumber > 1) {
    removeLosers(sheet, IDS["folder"], IDS["trash"]);
  }

  const [editUrl, publishedUrl, groupNumber, groupSize] = generateForms(IDS["folder"], roundNumber, triggeredFunction),
    mailSentNumber = mailLink(sheet, publishedUrl);

  if (roundNumber == 1) {
    formatSheet(sheet, groupNumber, groupSize)
  }

  // Confirmation message (also ends the loading screen)
  let htmlOutput = HtmlService.createHtmlOutput(`<span style="font-family: 'trebuchet ms', sans-serif;">Voci le lien éditeur : <a href = "${editUrl}">éditer le Form</a>.<br/>
                                                <br/> Voici le lien lecteur : <a href = "${publishedUrl}">répondre au Form</a>.<br/>
                                                <br/> Le lien lecteur a été envoyé à ${mailSentNumber} personne${mailSentNumber >= 2 ? "s" : ""}.</span>`);
  ui.showModelessDialog(htmlOutput, "Succès de l'opération");
}

// Solve tie
function tie() {

}
