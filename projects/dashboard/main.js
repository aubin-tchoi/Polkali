/* Si vous avez des questions à propos de ce script contactez Aubin Tchoï (Directeur Qualité 022) */

// Main file (triggers + main functions for KPI)

// Parameters : do not put here hardcoded values that are not parameters
const ui = SpreadsheetApp.getUi(),
  // Colors found in PEP's graphic chart
  COLORS = Object.freeze({
    plum: "#934683",
    wildOrchid: "#D66BA0",
    silverPink: "#C9ADA1",
    pine: "#72A98F",
    blueBell: "#A997DF",
    burgundy: "#8E3232",
    gold: "#FFBE2B",
    grey: "#404040",
    lightGrey: "#A29386"
  }),
  // Sheet, spreadsheet and Drive folders addresses
  ADDRESSES = Object.freeze({
    prospectionId: "1lJhJuZxUt_8_mVLXe5tazXPrb2Z3wr0M49rho974sNQ",
    prospectionName: "Suivi",
    etudesId: "1gmJLAKvUOYFeS32raOiSTYiE_ozr7YSk26Y0t0blm04",
    etudesName: "Suivi",
    driveId: "1dPi0dht-q_rI8fUmheA1j861huYPPcAy",
    slidesTemplate: "15WdicqHVF8LtOPrlwdM5iD1_qKh7YPaM15hrGGVbVzU"
  }),
  // Data sheet's header
  HEADS = Object.freeze({
    entreprise: "Entreprise",
    premierContact: "Premier contact",
    typeContact: "Type de contact",
    état: "État",
    devis: "Devis réalisé",
    caPot: "Prix potentiel de l'étude  € (HT)",
    confiance: "Pourcentage de confiance à la conversion en réelle étude",
    prix: "Prix en € (HT)",
    durée: "Durée (semaines)",
    alumni: "Alumni",
    JEH: "Nb JEH",
    concurrence: "Autres JE en concurrence"
  }),
  // Graphs' dimensions
  DIMS = {
    width: 1000,
    height: 400
  },
  // Links to every image used in this project
  IMGS = {
    loadingScreen: "https://raw.githubusercontent.com/aubin-tchoi/Polkali/main/images/loadingScreen.gif",
    thumbsUp: "https://raw.githubusercontent.com/aubin-tchoi/Polkali/main/images/thumbsUp.png"
  },
  // Html content of what is displayed and what is sent in a mail
  HTML_CONTENT = {
    display: HtmlService
      .createHtmlOutput(`<span style='font-size: 12pt;'> <span style="font-family: 'roboto', sans-serif;">KPI KPI KPI<br/></span> </span> <br/>`)
      .setWidth(1015)
      .setHeight(515),
    mail: HtmlService.createHtmlOutput(`<span style='font-size: 12pt;'> <span style="font-family: 'roboto', sans-serif;">&nbsp; &nbsp; Bonjour, <br/><br/>Voici les KPI portant sur la prospection.<br/> <br/>Bonne journée !</span> </span>`),
    saveConfirm: (url) => `<span style='font-size: 12pt;'> <span style="font-family: 'roboto', sans-serif;">Les KPIs ont été enregistrés, pouce pour ouvrir le lien (cliquez sur Boris).<br/><br/> &nbsp; &nbsp; La bise.</span></span><p style="text-align:center;"><a href=${url} target="_blank"><img src="${IMGS["thumbsUp"]}" alt="Thumbs up" width="130" height="131"></a></p>`,
    loadingScreen: HtmlService
      .createHtmlOutput(`<img src="${IMGS["loadingScreen"]}" alt="Loading" width="442" height="249">`)
      .setWidth(450)
      .setHeight(280)

  },
  // How months are spelled
  MONTH_NAMES = Object.freeze(["Jan", "Fév", "Mars", "Avr", "Mai", "Juin", "Juil", "Août", "Sept", "Oct", "Nov", "Déc"]),
  // Different states a mission can go through during the prospection phase
  ETAT_PROSP = Object.freeze({
    rdv: "Premier RDV réalisé",
    devis: "Devis rédigé et envoyé",
    negoc: "En négociation",
    etude: "Etude obtenue"
  }),
  // States corresponding to a contact that wasn't converted into a mission
  ETAT_PROSP_BIS = Object.freeze({
    sansSuite: "Sans suite",
    aRelancer: "A relancer",
  }),
  // Different states a mission can go through
  ETAT_ETUDE = Object.freeze({
    negoc: "En négociation",
    redac: "En rédaction",
    enAttente: "En attente",
    enCours: "En cours",
    terminée: "Terminée",
    cloturée: "Clôturée",
    sanSuite: "Sans suite"
  }),
  // Indexes of months
  MONTH_LIST = Object.freeze([{
      month: 1,
      year: 2020
    },
    {
      month: 2,
      year: 2020
    },
    {
      month: 3,
      year: 2020
    },
    {
      month: 4,
      year: 2020
    },
    {
      month: 5,
      year: 2020
    },
    {
      month: 6,
      year: 2020
    },
    {
      month: 7,
      year: 2020
    },
    {
      month: 8,
      year: 2020
    },
    {
      month: 9,
      year: 2020
    },
    {
      month: 10,
      year: 2020
    },
    {
      month: 11,
      year: 2020
    },
    {
      month: 0,
      year: 2021
    },
    {
      month: 1,
      year: 2021
    },
    {
      month: 2,
      year: 2021
    },
    {
      month: 3,
      year: 2021
    },
    {
      month: 4,
      year: 2021
    }
  ]);


/* ----- Triggers ----- */
function onOpen() {
  ui.createMenu('KPI')
    .addItem("Affichage des KPI", "displayKPI")
    .addItem("Enregistrement des KPI", "saveKPI")
    .addSeparator()
    .addItem("Préparation du CA", "prepCA")
    .addToUi();
}

// Creates a monthly trigger on autoSaveKPI
function installTrigger() {
  ScriptApp.newTrigger("autoSaveKPI")
    .timeBased()
    .onMonthDay(11)
    .create();
}

/* ------ KPI ----- */
function generateKPI() {

  let currentTime = new Date();
  displayLoadingScreen("Chargement des KPI..");

  // Initialization (names in franglish because why not)
  let dataProspection = extractSheetData(ADDRESSES["prospectionId"], ADDRESSES["prospectionName"], {
      data: {
        x: 4,
        y: 2
      },
      header: {
        x: 1,
        y: 2
      }
    }).filter(row => row[HEADS["premierContact"]] != ""),
    dataEtudes = extractSheetData(ADDRESSES["etudesId"], ADDRESSES["etudesName"], {
      data: {
        x: 5,
        y: 3
      },
      header: {
        x: 1,
        y: 3
      }
    }).filter(row => row[HEADS["état"]] != ETAT_ETUDE["sanSuite"]);

  currentTime = measureTime(currentTime, "extract data from the two sheets");

  // Final outputs (displaying the charts on screen & mail content)
  let htmlOutput = HTML_CONTENT["display"],
    htmlMail = HTML_CONTENT["mail"],
    attachments = [],
    charts = [];

  currentTime = measureTime(currentTime, "load the HTML content");

  // KPI : Contacts par mois
  let [contactsTable, conversionChart] = contacts(dataProspection); // conversionChart is a 2D array : [month][number of contact in a given state]
  charts.push(createColumnChart(contactsTable, "Contacts"));

  // KPI : Taux de conversion global par mois
  let conversionRateTable = conversionRate(conversionChart);
  charts.push(createLineChart(conversionRateTable, "Taux de conversion global", {
    colors: [COLORS["burgundy"]]
  }));

  // KPI : Taux de conversion entre chaque étape
  let conversionRateByTypeTable = conversionRateByType(conversionChart);
  charts.push(createColumnChart(conversionRateByTypeTable, "Taux de conversion sur chaque étape", {
    colors: [COLORS["burgundy"]],
    percent: true
  }));
  /*
  // KPI : CA
  let turnoverTable = turnover(dataProspection);
  charts.push(createColumnChart(turnoverTable, Object.values(COLORS), "Chiffre d'affaires", DIMS));
  */
  // KPI : Type de contact
  let contactTypeTable = contactType(dataProspection);
  charts.push(createPieChart(contactTypeTable, "Type de contact"));

  // KPI : Taux de conversion par type de contact
  let conversionRateByContactTable = conversionRateByContact(dataProspection);
  charts.push(createColumnChart(conversionRateByContactTable, "Taux de conversion par type de contact", {
    percent: true
  }));

  // KPI : Contacts qui sont en lien avec d'autres JE
  let [contactsConcurrenceTable, conversionConcurrenceChart] = contacts(dataProspection.filter(row => row[HEADS["concurrence"]] || false));
  charts.push(createColumnChart(contactsConcurrenceTable, "Contacts en lien avec d'autres JE"));

  // KPI : Taux de conversion en concurrence
  let conversionRateConcurrenceTable = conversionRateByType(conversionConcurrenceChart);
  charts.push(createColumnChart(conversionRateConcurrenceTable, "Taux de conversion sur chaque étape (en situation de concurrence)", {
    colors: [COLORS["burgundy"]],
    percent: true
  }));

  // KPI : nombre d'étude pour différents intervalles de prix
  let priceRangeTable = priceRange(dataEtudes.filter(row => row[HEADS["prix"]] != ""), 500, 4500, 8);
  charts.push(createColumnChart(priceRangeTable, "Nombre d'études par tranche de prix", {
    colors: [COLORS["burgundy"]]
  }));

  // KPI : nombre d'étude pour différents intervalles de prix
  let [JEHRangeTable, JEHTicks] = JEHRange(dataEtudes.filter(row => row[HEADS["JEH"]] != ""));
  charts.push(createColumnChart(JEHRangeTable, "Nombre d'études par nombre de JEHs", {
    colors: [COLORS["burgundy"]],
    hticks: JEHTicks
  }));

  // KPI : nombre d'étude pour différents intervalles de prix
  let [lengthRangeTable, lengthTicks] = lengthRange(dataEtudes.filter(row => row[HEADS["durée"]] != ""));
  charts.push(createColumnChart(lengthRangeTable, "Nombre d'études par durée d'étude (en nombre de semaines)", {
    colors: [COLORS["burgundy"]],
    hticks: lengthTicks
  }));

  // KPI : nombre d'étude pour différents intervalles de prix
  let alumniContributionTable = alumniContribution(dataEtudes);
  charts.push(createPieChart(alumniContributionTable, "Proportion du CA due aux alumni"));

  currentTime = measureTime(currentTime, "create the charts");

  // Adding the charts to the htmlOutput and the list of attachments
  charts.forEach(c => {
    convertChart(c, c.getOptions().get("title"), htmlOutput, attachments);
  });

  currentTime = measureTime(currentTime, "convert the charts");

  // Returning functions within an Object for later use, all functions are manually decorated with an execution time logger
  return {
    display: function () {
      let initialTime = new Date();
      ui.showModalDialog(htmlOutput, "KPI");
      measureTime(initialTime, "display the charts");
    },
    save: function (folderId = ADDRESSES["driveId"]) {
      let initialTime = new Date();
      saveOnDrive(attachments, folderId);
      measureTime(initialTime, "save the charts");
    },
    mail: function (adress) {
      let initialTime = new Date();
      sendMail(adress, htmlMail, "KPI", attachments);
      measureTime(initialTime, "mail the charts");
    },
    slides: function () {
      let initialTime = new Date();
      generateSlides(ADDRESSES["slidesTemplate"], attachments, folderId = ADDRESSES["driveId"]);
      measureTime(initialTime, "generate the slides");
    }
  }
}

// Displaying the graphs (through a menu in the sheet)
function displayKPI() {
  let KPI = generateKPI();
  KPI.display();
}

// Saving the graphs (through a menu in the sheet)
function saveKPI() {
  // Prompts and alerts should be made before any time-consuming operation (to make it so users can enter any required information and then leave)
  let mailResults = ui.alert("Envoi des diagrammes par mail", "Souhaitez vous recevoir les diagrammes par mail ?", ui.ButtonSet.YES_NO) == ui.Button.YES,
    saveResults = ui.alert("Enregistrement des images sur le Drive", "Souhaitez vous enregistrer les images sur le Drive ?", ui.ButtonSet.YES_NO) == ui.Button.YES,
    folderId = saveResults ? userQuery({
      title: "Enregistrement des images sur le Drive",
      query: "Entrez l'id du dossier de destination :",
      incorrectInput: "L'ID entré est invalide, veuillez recommencer."
    }) : "",
    mailAdress = mailResults ? userQuery({
      title: "Envoi des diagrammes par mail",
      query: "Entrez l'adresse mail de destination :"
    }) : "",
    KPI = generateKPI();

  if (mailResults) {
    KPI.mail(mailAdress);
  }
  if (saveResults) {
    KPI.save(folderId);
  }
  KPI.display();
}

// Saving the graphs automatically (linked to a monthly trigger)
function autoSaveKPI() {
  let KPI = generateKPI();
  KPI.save();
}

// Saving the graphs + creating slides to present them (through a menu in the sheet)
function prepCA() {
  let KPI = generateKPI();
  KPI.slides();
  KPI.save();
}
