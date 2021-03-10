/* Si vous avez des questions à propos de ce script contactez Aubin Tchoï (Directeur Qualité 022) */

// Main file (triggers + main function for KPI and synchronization)

const ui = SpreadsheetApp.getUi(),
  // Colors found in PEP's graphic chart
  COLORS = {
    burgundy: "#8E3232",
    gold: "#FFBE2B",
    grey: "#404040",
    lightGrey: "#A29386"
  },
  // Sheet and spreadsheet adresses
  SHEETS = {
    id: "1lJhJuZxUt_8_mVLXe5tazXPrb2Z3wr0M49rho974sNQ",
    name: "Suivi"
  },
  // Data sheet's header
  HEADS = {
    entreprise: "Entreprise",
    premierContact: "Premier contact",
    typeContact: "Type de contact",
    état: "État",
    devis: "Devis réalisé",
    caPot: "Prix potentiel de l'étude  € (HT)",
    confiance: "Pourcentage de confiance à la conversion en réelle étude"
  },
  // Graphs' dimensions
  DIMS = {
    width: 1000,
    height: 500
  },
  // Links to every image used in this project
  IMGS = {
    loadingScreen: "https://raw.githubusercontent.com/aubin-tchoi/Polkali/main/images/loadingScreen.gif"
  },
  // Html content of what is displayed and what is sent in a mail
  HTML_CONTENT = {
    display: `<span style='font-size: 12pt;'> <span style="font-family: 'trebuchet ms', sans-serif;">KPI KPI KPI<br/></span> </span> <br/>`,
    mail: `<span style='font-size: 12pt;'> <span style="font-family: 'trebuchet ms', sans-serif;">&nbsp; &nbsp; Bonjour, <br/><br/>Voici les KPI portant sur la prospection.<br/> <br/>Bonne journée !</span> </span>`
  },
  // Drive folder's id
  DRIVE = {
    folderId: "1dPi0dht-q_rI8fUmheA1j861huYPPcAy"
  },
  // How months are spelled
  MONTH_NAMES = ["Jan", "Fév", "Mars", "Avr", "Mai", "Juin", "Juil", "Août", "Sept", "Oct", "Nov", "Déc"],
  // Different states a mission can go through
  STATES = {
    rdv: "Premier RDV réalisé",
    devis: "Devis rédigé et envoyé",
    negoc: "En négociation",
    etude: "Etude obtenue"
  },
  // States corresponding to a contact that wasn't converted into a mission
  STATES_BIS = {
    sansSuite: "Sans suite",
    aRelancer: "A relancer",
  }
// Indexes of months (/!\ LE PROCHAIN MANDAT REMMETTEZ MAI JUIN JUILLET (ou mettez mois + année pour couvrir plusieurs années)
MONTH_LIST = [1, 2, 3, 4, 5, 6, 7, 8, 9, 10, 11, 0];


/* ----- Triggers ----- */
function onOpen() {
  ui.createMenu('KPI')
    .addItem("Synchronisation de la sélection avec le suivi d'études", "syncOnSelec")
    .addItem("Affichage des KPI", "displayKPI")
    .addToUi();
}

/*
function onEdit(e) {
  let range = e.range;
  if (range.getNumColumns() == 1 && range.getNumRows() == 1 && e.value == STATES["etude"]) {
    syncOnEdit(range.getRow());
  }
} */

// Creates a monthly trigger on autoSaveKPI
function installTrigger() {
  ScriptApp.newTrigger("autoSaveKPI")
    .timeBased()
    .onMonthDay(11)
    .create();
}

/* ------ KPI ----- */
function generateKPI(mailResults, saveResults, folderId, displayResults) {

  displayLoadingScreen("Chargement des KPI..");

  // Initialization
  const sheet = SpreadsheetApp.openById(SHEETS["id"]).getSheetByName(SHEETS["name"]),
    data = sheet.getRange(4, 2, (manuallyGetLastRow(sheet) - 3), 12).getValues(),
    heads = sheet.getRange(1, 2, 1, 12).getValues().shift();

  // I chose to keep the data as an array of objects (other possibility : object of objects, the keys being the months' names and the values the data in each row)
  let obj = data.map(r => heads.reduce((o, k, i) => (o[k] = r[i] || 0, o), {})).filter(row => row["Premier contact"] != "");

  // Final outputs (displaying the charts on screen & mail content)
  let htmlOutput = HtmlService
    .createHtmlOutput(HTML_CONTENT["display"])
    .setWidth(1015)
    .setHeight(515),
    htmlMail = HtmlService.createHtmlOutput(HTML_CONTENT["mail"]),
    attachments = [];

  // KPI : Contacts par mois
  let [contactTable, conversionChart] = contacts(obj); // conversionChart is a 2D array : [month][number of contact in a given state]
  [htmlOutput, attachments] = createColumnChart(contactTable, "Contacts", htmlOutput, attachments, DIMS["width"], DIMS["height"]);

  // KPI : Taux de conversion
  let conversionTable = conversionRate(conversionChart);
  [htmlOutput, attachments] = createLineChart(conversionTable, [COLORS["burgundy"], COLORS["gold"], COLORS["grey"]], "Taux de conversion", htmlOutput, attachments, DIMS["width"], DIMS["height"]);

  // KPI : CA
  let turnoverTable = turnover(obj);
  [htmlOutput, attachments] = createColumnChart(turnoverTable, "Chiffre d'affaires", htmlOutput, attachments, DIMS["width"], DIMS["height"]);

  // KPI : Type de contact
  let contactTypeTable = contactType(obj);
  [htmlOutput, attachments] = createPieChart(contactTypeTable, "Type de contact", htmlOutput, attachments, DIMS["width"], DIMS["height"]);

  // KPI : Taux de conversion par type de contact
  let conversionTableContact = conversionRateByContact(obj);
  [htmlOutput, attachments] = createColumnChart(conversionTableContact, "Taux de conversion par type de contact", htmlOutput, attachments, DIMS["width"], DIMS["height"]);

  // Sending graphs by mail (mail adress will be prompted)
  if (mailResults) {
    sendMail(htmlMail, "KPI", attachments);
  }

  // Saving graphs in a Drive folder
  if (saveResults) {
    if (folderId != "") {
      saveOnDrive(attachments, folderId);
    } else {
      saveOnDrive(attachments);
    }
  }

  // Displaying the charts on screen
  if (displayResults) {
    ui.showModalDialog(htmlOutput, "KPI");
  }
}

// Displaying the graphs (through a menu in the sheet)
function displayKPI() {
  // Prompts and alerts should be made before any time-consuming operation (to make it so users can enter any required information and then leave)
  let mailResults = ui.alert("Envoi des diagrammes par mail", "Souhaitez vous recevoir les diagrammes par mail ?", ui.ButtonSet.YES_NO) == ui.Button.YES,
    saveResults = ui.alert("Enregistrement des images sur le Drive", "Souhaitez vous enregistrer les images sur le Drive ?", ui.ButtonSet.YES_NO) == ui.Button.YES,
    folderId = "";
  if (saveResults) {
    folderId = ui.prompt("Enregistrement des images sur le Drive", "Entrez l'id du dossier de destination :", ui.ButtonSet.OK).getResponseText();
  }
  generateKPI(mailResults, saveResults, folderId, true);
}

// Saving the graphs automatically (with a monthly trigger)
function autoSaveKPI() {
  generateKPI(false, true, "", false);
}
