/* Si vous avez des questions à propos de ce script contactez Aubin Tchoï (Directeur Qualité 022) */

// Main file (triggers + main functions for KPI)

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
  // Initializing the time counter
  let currentTime = new Date();

  displayLoadingScreen("Chargement des KPI..");

  // Initialization (names in franglish because why not)
  let dataProspection = extractSheetData(ADDRESSES.prospectionId, ADDRESSES.prospectionName, {
      data: {
        x: 4,
        y: 2
      },
      header: {
        x: 1,
        y: 2
      },
      trustColumn: 4
    }).filter(row => row[HEADS.premierContact] != ""),
    dataEtudes = extractSheetData(ADDRESSES.etudesId, ADDRESSES.etudesName, {
      data: {
        x: 5,
        y: 3
      },
      header: {
        x: 1,
        y: 3
      },
      trustColumn: 3
    }).filter(row => row[HEADS.état] != ETAT_ETUDE.sanSuite);

  currentTime = measureTime(currentTime, "extract data from the two sheets");

  // Final outputs (displaying the charts on screen & mail content)
  let htmlOutput = HTML_CONTENT.display,
    htmlMail = HTML_CONTENT.mail,
    attachments = [],
    charts = [];

  currentTime = measureTime(currentTime, "load the HTML content");


  // ----- BILAN PROSPECTION -----

  // KPI : Contacts par mois
  let [contactsTable, conversionChart] = contacts(dataProspection); // conversionChart is a 2D array : [month][number of contact in a given state]
  charts.push(createChart(CHART_TYPE.COLUMN, contactsTable, "Contacts"));

  // KPI : Taux de conversion global par mois
  let conversionRateTable = conversionRate(conversionChart);
  charts.push(createChart(CHART_TYPE.LINE, conversionRateTable, "Taux de conversion global", {
    colors: [COLORS.burgundy]
  }));

  // KPI : Taux de conversion entre chaque étape
  let conversionRateByTypeTable = conversionRateByType(conversionChart);
  charts.push(createChart(CHART_TYPE.COLUMN, conversionRateByTypeTable, "Taux de conversion sur chaque étape", {
    colors: [COLORS.burgundy],
    percent: true
  }));

  // KPI : CA
  /* let turnoverTable = turnover(dataProspection);
  charts.push(createColumnChart(turnoverTable, Object.values(COLORS), "Chiffre d'affaires", DIMS)); */


  // ----- TYPOLOGIE DES CONTACTS -----

  // KPI : Répartition des contacts par type
  let contactTypeTable = contactType(dataProspection);
  charts.push(createChart(CHART_TYPE.PIE, contactTypeTable, "Type de contact"));

  // KPI : Taux de conversion par type de contact
  let conversionRateByContactTable = conversionRateByContact(dataProspection);
  charts.push(createChart(CHART_TYPE.COLUMN, conversionRateByContactTable, "Taux de conversion par type de contact", {
    percent: true
  }));

  // KPI : Répartition des contacts par secteur d'activité
  let contactBySectorTable = contactBySector(dataProspection.filter(row => row[HEADS.secteur] != ""));
  charts.push(createChart(CHART_TYPE.PIE, contactBySectorTable, "Répartition des contacts par secteur"));


  // ----- CONCURRENCE AVEC D'AUTRES JE

  // KPI : Contacts qui sont en lien avec d'autres JE
  /*let [contactsConcurrenceTable, conversionConcurrenceChart] = contacts(dataProspection.filter(row => row[HEADS.concurrence] || false));
  charts.push(createChart(CHART_TYPE.COLUMN, contactsConcurrenceTable, "Contacts en lien avec d'autres JE")); */

  // KPI : Taux de conversion en concurrence
  /* let conversionRateConcurrenceTable = conversionRateByType(conversionConcurrenceChart);
  charts.push(createChart(CHART_TYPE.COLUMN, conversionRateConcurrenceTable, "Taux de conversion sur chaque étape (en situation de concurrence)", {
    colors: [COLORS.burgundy],
    percent: true
  })); */


  // ----- PERFORMANCE SELON LA TAILLE DES ETUDES -----

  // KPI : Nombre d'étude pour différents intervalles de prix
  let priceRangeTable = priceRange(dataEtudes.filter(row => row[HEADS.prix] != ""), 500, 4500, 8);
  charts.push(createChart(CHART_TYPE.COLUMN, priceRangeTable, "Nombre d'études par tranche de prix", {
    colors: [COLORS.burgundy]
  }));

  // KPI : Nombre d'étude pour différents intervalles de prix
  let [JEHRangeTable, JEHTicks] = JEHRange(dataEtudes.filter(row => row[HEADS.JEH] != ""));
  charts.push(createChart(CHART_TYPE.COLUMN, JEHRangeTable, "Nombre d'études par nombre de JEHs", {
    colors: [COLORS.burgundy],
    hticks: JEHTicks
  }));

  // KPI : Nombre d'étude pour différents intervalles de prix
  let [lengthRangeTable, lengthTicks] = lengthRange(dataEtudes.filter(row => row[HEADS.durée] != ""));
  charts.push(createChart(CHART_TYPE.COLUMN, lengthRangeTable, "Nombre d'études par durée d'étude (en nombre de semaines)", {
    colors: [COLORS.burgundy],
    hticks: lengthTicks
  }));


  // ----- MESURE DES DIFFERENTES CONTRIBUTIONS AU CA -----

  // KPI : Proportion du CA due aux alumni
  let alumniContributionTable = alumniContribution(dataEtudes);
  charts.push(createChart(CHART_TYPE.PIE, alumniContributionTable, "Proportion du CA due aux alumni", {
    colors: [COLORS.pine, COLORS.silverPink]
  }));

  // KPI : Proportion du CA venant de la prospection
  let prospectionTurnoverTable = prospectionTurnover(dataProspection.filter(row => row[HEADS.état] == ETAT_PROSP.etude));
  charts.push(createChart(CHART_TYPE.PIE, prospectionTurnoverTable, "Proportion du CA venant de la prospection", {
    colors: [COLORS.pine, COLORS.silverPink]
  }));

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
    save: function (folderId = ADDRESSES.driveId) {
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
      generateSlides(ADDRESSES.slidesTemplate, attachments, folderId = ADDRESSES.driveId);
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
