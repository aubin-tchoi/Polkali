/* Si vous avez des questions à propos de ce script contactez Aubin Tchoï (Directeur Qualité 022) */
/* Certains commentaires sont en anglais, d'autres en français.
Si vous vous contenez à une modification en surface du code (ajout d'un KPI par exemple),
vous devriez voir essentiellement des commentaires en français. */

// Main file (triggers + main functions for KPI)

/* ----- Triggers ----- */
/**
 * Fonction exécutée à l'ouverture du dashboard, permet la création du menu 'KPI'.
 * @ignore
 */
function onOpen() {
  ui.createMenu('KPI')
    .addItem("Affichage des KPI", "displayKPI")
    .addItem("Enregistrement des KPI", "saveKPI")
    .addSeparator()
    .addItem("Préparation du CA", "prepCA")
    .addToUi();
}

/**
 * Fonction qui à l'exécution crée le trigger de sauvegarde automatique (a besoin d'être exécuté une seule fois).
 */
function installTrigger() {
  ScriptApp.newTrigger("autoSaveKPI")
    .timeBased()
    .onMonthDay(11)
    .create();
}

/* ------ KPI ----- */
/**
 * Fonction main - Génération des KPI, log régulièrement le temps écoulé entre chaque étape.
 * Dans le return vous disposez des graphes dans les Objects charts (type Chart) et attachments (type Blob) ainsi que dans le htmlOutput.
 * @returns {Object} Méthodes display, save, mail, slides qui correspondent chacune à une fonctionnalité associée aux KPI.
 */
function generateKPI() {
  // Initializing the time counter
  let currentTime = new Date();

  displayLoadingScreen("Chargement des KPI..");

  // Initialization (names in franglish because why not)
  /**
   * Données sous forme d'Array d'Object
   * (1 ligne correspond à une ligne du Sheets, les clés correspondent aux champs du header indiqués dans HEADS)
   */
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
    }).filter(row => !(Object.values(ETAT_ETUDE_BIS).includes(row[HEADS.état]))),
    dataEtudesBis = extractSheetData(ADDRESSES.etudesIdBis, ADDRESSES.etudesName, {
      data: {
        x: 5,
        y: 3
      },
      header: {
        x: 1,
        y: 3
      },
      trustColumn: 3
    }).filter(row => !(Object.values(ETAT_ETUDE_BIS).includes(row[HEADS.état])));

  currentTime = measureTime(currentTime, "extract data from the two sheets");

  /**
   * Output qui sera affiché lors du display (type HtmlOutput).
   */
  let htmlOutput = HTML_CONTENT.display,
    /** 
     * Contenu Html du mail qui peut être envoyé sur demande dans le menu "Enregistrement des KPI".
     */
    htmlMail = HTML_CONTENT.mail,
    /**
     * Object regroupant les blobs des graphes (au format .png) en différentes catégories décrites par les clés de l'Object.
     */
    attachments = {},
    /**
     * Object regroupant les graphes (au type Chart) en différentes catégories décrites par les clés de l'Object.
     */
    charts = {
      summary: [],
      contactTypology: [],
      contactType: [],
      companySize: [],
      sector: [],
      missionType: [],
      competitiveness: [],
      sizeComparison: [],
      contributions: []
    },
    dataTable,
    ticks,
    conversionChart;

  currentTime = measureTime(currentTime, "load the HTML content");


  // ----- BILAN PROSPECTION -----

  // KPI : Contacts par mois
  [dataTable, conversionChart] = contacts(dataProspection); // conversionChart is a 2D array : [month][number of contact in a given state]
  charts.summary.push(createChart(CHART_TYPE.COLUMN, dataTable, "Contacts", {
    colors: COLORS_OFFICE
  }));

  // KPI : Taux de conversion global par mois
  dataTable = conversionRateOverTime(conversionChart);
  charts.summary.push(createChart(CHART_TYPE.LINE, dataTable, "Taux de conversion global", {
    colors: [COLORS.burgundy]
  }));

  // KPI : Taux de conversion entre chaque étape
  dataTable = conversionRateByType(conversionChart);
  charts.summary.push(createChart(CHART_TYPE.COLUMN, dataTable, "Taux de conversion sur chaque étape", {
    colors: [COLORS.burgundy],
    percent: true
  }));

  // KPI : CA
  /* dataTable = turnover(dataProspection);
  charts.summary.push(createColumnChart(dataTable, Object.values(COLORS), "Chiffre d'affaires", DIMS)); */


  // ----- TYPOLOGIE DES CONTACTS -----

  // KPI : Répartition des contacts par type
  dataTable = totalDistribution(HEADS.typeContact, dataProspection);
  charts.contactTypology.push(createChart(CHART_TYPE.PIE, dataTable, "Type de contact", {
    colors: COLORS_OFFICE
  }));

  // KPI : Taux de conversion par type de contact
  dataTable = conversionRate(HEADS.typeContact, dataProspection);
  charts.contactTypology.push(createChart(CHART_TYPE.COLUMN, dataTable, "Taux de conversion par type de contact", {
    colors: COLORS_DUO,
    percent: true
  }));

  // KPI : Répartition des contacts par domaine de compétence
  dataTable = totalDistribution(HEADS.domaine, dataProspection.filter(row => row[HEADS.domaine] != ""));
  charts.contactTypology.push(createChart(CHART_TYPE.PIE, dataTable, "Répartition des contacts par domaine de compétence", {
    colors: COLORS_OFFICE
  }));


  // ----- CONCURRENCE AVEC D'AUTRES JE

  // KPI : Contacts qui sont en lien avec d'autres JE
  /*let [dataTableBis, conversionConcurrenceChart] = contacts(dataProspection.filter(row => row[HEADS.concurrence] || false));
  charts.competitiveness.push(createChart(CHART_TYPE.COLUMN, dataTableBis, "Contacts en lien avec d'autres JE")); */

  // KPI : Taux de conversion en concurrence
  /* dataTable = conversionRateByType(conversionConcurrenceChart);
  charts.competitiveness.push(createChart(CHART_TYPE.COLUMN, dataTable, "Taux de conversion sur chaque étape (en situation de concurrence)", {
    colors: [COLORS.burgundy],
    percent: true
  })); */


  // ----- PERFORMANCE SELON LE TYPE DE CONTACTS -----

  // KPI : nombre d'études, CA par type de contact
  [dataTable, ticks] = performanceByContact(HEADS.typeContact, dataProspection.filter(row => row[HEADS.typeContact] != ""));
  charts.contactType.push(createChart(CHART_TYPE.COLUMN, dataTable, "Performance par type de contact", {
    vticks: ticks,
    colors: COLORS_OFFICE
  }));

  // KPI : Proportion du CA venant de chaque type de contact
  dataTable = turnoverDistribution(HEADS.typeContact, dataProspection.filter(row => row[HEADS.état] == ETAT_PROSP.etude && row[HEADS.typeContact] != ""));
  charts.contactType.push(createChart(CHART_TYPE.PIE, dataTable, "Proportion du CA venant de chaque type de contact", {
    colors: COLORS_OFFICE
  }));


  // ----- PERFORMANCE SELON LE TYPE D'ENTREPRISE -----

  // KPI : nombre d'études, CA par type d'entreprise
  [dataTable, ticks] = performance(HEADS.typeEntreprise, dataProspection.filter(row => row[HEADS.état] == ETAT_PROSP.etude && row[HEADS.secteur] != ""));
  charts.companySize.push(createChart(CHART_TYPE.COLUMN, dataTable, "Performance par type d'entreprise", {
    colors: COLORS_DUO,
    vticks: ticks
  }));

  // KPI : Proportion du CA venant de chaque type d'entreprise
  dataTable = turnoverDistribution(HEADS.typeEntreprise, dataProspection.filter(row => row[HEADS.état] == ETAT_PROSP.etude && row[HEADS.secteur] != ""));
  charts.companySize.push(createChart(CHART_TYPE.PIE, dataTable, "Proportion du CA venant de chaque type d'entreprise", {
    colors: COLORS_OFFICE
  }));
  

  // ----- PERFORMANCE PAR SECTEUR D'ACTIVITE DU CLIENT -----

  // KPI : nombre d'études, CA par secteur
  [dataTable, ticks] = performance(HEADS.secteur, dataProspection.filter(row => row[HEADS.état] == ETAT_PROSP.etude && row[HEADS.secteur] != ""));
  charts.sector.push(createChart(CHART_TYPE.COLUMN, dataTable, "Performance par secteur", {
    colors: COLORS_DUO,
    vticks: ticks
  }));

  // KPI : Proportion du CA venant de chaque secteur
  dataTable = turnoverDistribution(HEADS.secteur, dataProspection.filter(row => row[HEADS.état] == ETAT_PROSP.etude && row[HEADS.secteur] != ""));
  charts.sector.push(createChart(CHART_TYPE.PIE, dataTable, "Proportion du CA venant de chaque secteur", {
    colors: COLORS_OFFICE
  }));


  // ----- PERFORMANCE SELON LE TYPE DE PRESTATION -----

  // KPI : nombre d'études, CA par type de prestation
  [dataTable, ticks] = performance(HEADS.prestation, dataEtudesBis, HEADS.prix);
  charts.missionType.push(createChart(CHART_TYPE.COLUMN, dataTable, "Performance par type de prestation", {
    colors: COLORS_DUO,
    vticks: ticks
  }));

  // KPI : Proportion du CA obtenue sur chaque type de prestation
  dataTable = turnoverDistribution(HEADS.prestation, dataEtudesBis, HEADS.prix);
  charts.missionType.push(createChart(CHART_TYPE.PIE, dataTable, "Proportion du CA obtenue sur chaque type de prestation", {
    colors: COLORS_OFFICE
  }));

  
  // ----- PERFORMANCE SELON LA TAILLE DES ETUDES -----

  // KPI : Nombre d'étude pour différents intervalles de prix
  dataTable = priceRange(dataEtudes.filter(row => row[HEADS.prix] != ""), 500, 4500, 8);
  charts.sizeComparison.push(createChart(CHART_TYPE.COLUMN, dataTable, "Nombre d'études par tranche de prix (en €)", {
    colors: [COLORS.burgundy]
  }));

  // KPI : Nombre d'étude pour différents intervalles de prix
  [dataTable, ticks] = numberOfMissions(HEADS.JEH, dataEtudes.filter(row => row[HEADS.JEH] != ""));
  charts.sizeComparison.push(createChart(CHART_TYPE.COLUMN, dataTable, "Nombre d'études par nombre de JEHs", {
    colors: [COLORS.burgundy],
    hticks: ticks
  }));

  // KPI : Nombre d'étude pour différents intervalles de prix
  [dataTable, ticks] = numberOfMissions(HEADS.durée, dataEtudes.filter(row => row[HEADS.durée] != ""));
  charts.sizeComparison.push(createChart(CHART_TYPE.COLUMN, dataTable, "Nombre d'études par durée d'étude (en nombre de semaines)", {
    colors: [COLORS.burgundy],
    hticks: ticks
  }));


  // ----- MESURE DES DIFFERENTES CONTRIBUTIONS AU CA -----

  // KPI : Proportion du CA due aux alumni
  dataTable = turnoverDistributionBinary(HEADS.alumni, dataEtudes);
  charts.contributions.push(createChart(CHART_TYPE.PIE, dataTable, "Proportion du CA due aux alumni", {
    colors: [COLORS.pine, COLORS.silverPink]
  }));

  // KPI : Proportion du CA venant de la prospection
  dataTable = prospectionTurnover(dataProspection.filter(row => row[HEADS.état] == ETAT_PROSP.etude));
  charts.contributions.push(createChart(CHART_TYPE.PIE, dataTable, "Proportion du CA venant de la prospection", {
    colors: [COLORS.pine, COLORS.silverPink]
  }));

  currentTime = measureTime(currentTime, "create the charts");

  // Adding the charts to the htmlOutput and the list of attachments
  Object.keys(charts).forEach(key => {
    // Adding a line that describe the category of KPI
    if (charts[key].length > 0) {
      addHTMLLine(key, htmlOutput);
    }
    attachments[key] = [];
    // Converting every chart
    charts[key].forEach(c => {
      convertChart(c, c.getOptions().get("title"), htmlOutput, attachments[key]);
    });
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
    mail: function (address) {
      let initialTime = new Date();
      sendMail(address, htmlMail, "KPI", attachments);
      measureTime(initialTime, "mail the charts");
    },
    slides: function () {
      let initialTime = new Date();
      generateSlides(ADDRESSES.slidesTemplate, attachments, folderId = ADDRESSES.driveId);
      measureTime(initialTime, "generate the slides");
    }
  }
}

/**
 * Affichage des KPI dans le dashboard. Reliée à l'item "Affichage des KPI" dans le menu "KPI".
 */
function displayKPI() {
  let KPI = generateKPI();
  KPI.display();
}

/**
 * Enregistre les graphes sous forme d'images (au format .png) dans le dossier spécifié (par défaut dans Pôle Qualité -> KPI -> KPI archivés)
 * puis les envoie par mail à l'adresse spécifiée. Reliée à l'item "Enregistrement des KPI" dans le menu "KPI".
 */
function saveKPI() {
  // Prompts and alerts should be made before any time-consuming operation (to make it so users can enter any required information and then leave)
  let mailResults = ui.alert("Envoi des diagrammes par mail", "Souhaitez vous recevoir les diagrammes par mail ?", ui.ButtonSet.YES_NO) == ui.Button.YES,
    saveResults = ui.alert("Enregistrement des images sur le Drive", "Souhaitez vous enregistrer les images sur le Drive ?", ui.ButtonSet.YES_NO) == ui.Button.YES,
    folderId = saveResults ? userQuery({
      title: "Enregistrement des images sur le Drive",
      query: "Entrez l'id du dossier de destination :",
      incorrectInput: "L'ID entré est invalide, veuillez recommencer."
    }) : "",
    mailAddress = mailResults ? userQuery({
      title: "Envoi des diagrammes par mail",
      query: "Entrez l'adresse mail de destination :"
    }) : "",
    KPI = generateKPI();

  if (mailResults) {
    KPI.mail(mailAddress);
  }
  if (saveResults) {
    KPI.save(folderId);
  }
  KPI.display();
}

/**
 * Enregistre les graphes sous forme d'images (au format .png) dans Pôle Qualité -> KPI -> KPI archivés.
 * Reliée à un trigger mensuel (le 11 du mois).
 */
function autoSaveKPI() {
  let KPI = generateKPI();
  KPI.save();
}

/**
 * Crée un fichier Slides (il s'agit du Point KPI) contenant les KPI regroupés par catégorie,
 * il se trouvera dans Pôle Qualité -> KPI -> KPI archivés. Reliée à l'item "Préparation du CA" dans le menu "KPI".
 */
function prepCA() {
  let KPI = generateKPI();
  KPI.slides();
  KPI.save();
}
