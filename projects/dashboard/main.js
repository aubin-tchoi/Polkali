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
    }).filter(row => row[HEADS.état] != ETAT_ETUDE.sanSuite);

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
      competitiveness: [],
      sizeComparison: [],
      contributions: []
    };

  currentTime = measureTime(currentTime, "load the HTML content");


  // ----- BILAN PROSPECTION -----

  // KPI : Contacts par mois
  let [contactsTable, conversionChart] = contacts(dataProspection); // conversionChart is a 2D array : [month][number of contact in a given state]
  charts.summary.push(createChart(CHART_TYPE.COLUMN, contactsTable, "Contacts", {
    colors: COLORS_ANATOLE
  }));

  // KPI : Taux de conversion global par mois
  let conversionRateTable = conversionRate(conversionChart);
  charts.summary.push(createChart(CHART_TYPE.LINE, conversionRateTable, "Taux de conversion global", {
    colors: [COLORS.burgundy]
  }));

  // KPI : Taux de conversion entre chaque étape
  let conversionRateByTypeTable = conversionRateByType(conversionChart);
  charts.summary.push(createChart(CHART_TYPE.COLUMN, conversionRateByTypeTable, "Taux de conversion sur chaque étape", {
    colors: [COLORS.burgundy],
    percent: true
  }));

  // KPI : CA
  /* let turnoverTable = turnover(dataProspection);
  charts.summary.push(createColumnChart(turnoverTable, Object.values(COLORS), "Chiffre d'affaires", DIMS)); */


  // ----- TYPOLOGIE DES CONTACTS -----

  // KPI : Répartition des contacts par type
  let contactTypeTable = contactType(dataProspection);
  charts.contactTypology.push(createChart(CHART_TYPE.PIE, contactTypeTable, "Type de contact", {
    colors: COLORS_ANATOLE
  }));

  // KPI : Taux de conversion par type de contact
  let conversionRateByContactTable = conversionRateByContact(dataProspection);
  charts.contactTypology.push(createChart(CHART_TYPE.COLUMN, conversionRateByContactTable, "Taux de conversion par type de contact", {
    colors: COLORS_ANATOLE_2COLUMNS,
    percent: true
  }));

  // KPI : Répartition des contacts par domaine de compétence
  let contactByDomainTable = contactByDomain(dataProspection.filter(row => row[HEADS.domaine] != ""));
  charts.contactTypology.push(createChart(CHART_TYPE.PIE, contactByDomainTable, "Répartition des contacts par domaine de compétence", {
    colors: COLORS_ANATOLE
  }));


  // ----- CONCURRENCE AVEC D'AUTRES JE

  // KPI : Contacts qui sont en lien avec d'autres JE
  /*let [contactsConcurrenceTable, conversionConcurrenceChart] = contacts(dataProspection.filter(row => row[HEADS.concurrence] || false));
  charts.competitiveness.push(createChart(CHART_TYPE.COLUMN, contactsConcurrenceTable, "Contacts en lien avec d'autres JE")); */

  // KPI : Taux de conversion en concurrence
  /* let conversionRateConcurrenceTable = conversionRateByType(conversionConcurrenceChart);
  charts.competitiveness.push(createChart(CHART_TYPE.COLUMN, conversionRateConcurrenceTable, "Taux de conversion sur chaque étape (en situation de concurrence)", {
    colors: [COLORS.burgundy],
    percent: true
  })); */


  // ----- PERFORMANCE SELON LA TAILLE DES ETUDES -----

  // KPI : Nombre d'étude pour différents intervalles de prix
  let priceRangeTable = priceRange(dataEtudes.filter(row => row[HEADS.prix] != ""), 500, 4500, 8);
  charts.sizeComparison.push(createChart(CHART_TYPE.COLUMN, priceRangeTable, "Nombre d'études par tranche de prix", {
    colors: [COLORS.burgundy]
  }));

  // KPI : Nombre d'étude pour différents intervalles de prix
  let [JEHRangeTable, JEHTicks] = JEHRange(dataEtudes.filter(row => row[HEADS.JEH] != ""));
  charts.sizeComparison.push(createChart(CHART_TYPE.COLUMN, JEHRangeTable, "Nombre d'études par nombre de JEHs", {
    colors: [COLORS.burgundy],
    hticks: JEHTicks
  }));

  // KPI : Nombre d'étude pour différents intervalles de prix
  let [lengthRangeTable, lengthTicks] = lengthRange(dataEtudes.filter(row => row[HEADS.durée] != ""));
  charts.sizeComparison.push(createChart(CHART_TYPE.COLUMN, lengthRangeTable, "Nombre d'études par durée d'étude (en nombre de semaines)", {
    colors: [COLORS.burgundy],
    hticks: lengthTicks
  }));


  // ----- MESURE DES DIFFERENTES CONTRIBUTIONS AU CA -----

  // KPI : nombre d'études, CA par type d'entreprise
  let performanceByCompanyTypeTable = performanceByCompanyType(dataProspection.filter(row => row[HEADS.état] == ETAT_PROSP.etude).filter(row => row[HEADS.secteur] != ""));
  charts.contributions.push(createChart(CHART_TYPE.COLUMN, performanceByCompanyTypeTable, "Performance par type d'entreprise", {
    colors: COLORS_ANATOLE_2COLUMNS
  }));

  // KPI : Proportion du CA venant de chaque type d'entreprise
  let turnoverByCompanyTypeTable = turnoverByCompanyType(dataProspection.filter(row => row[HEADS.état] == ETAT_PROSP.etude).filter(row => row[HEADS.secteur] != ""));
  charts.contributions.push(createChart(CHART_TYPE.PIE, turnoverByCompanyTypeTable, "Proportion du CA venant de chaque type d'entreprise", {
    colors: COLORS_ANATOLE
  }));

  // KPI : nombre d'études, CA par secteur
  let performanceBySectorTable = performanceBySector(dataProspection.filter(row => row[HEADS.état] == ETAT_PROSP.etude).filter(row => row[HEADS.secteur] != ""));
  charts.contributions.push(createChart(CHART_TYPE.COLUMN, performanceBySectorTable, "Performance par secteur", {
    colors: COLORS_ANATOLE_2COLUMNS
  }));

  // KPI : Proportion du CA venant de chaque secteur
  let turnoverBySectorTable = turnoverBySector(dataProspection.filter(row => row[HEADS.état] == ETAT_PROSP.etude).filter(row => row[HEADS.secteur] != ""));
  charts.contributions.push(createChart(CHART_TYPE.PIE, turnoverBySectorTable, "Proportion du CA venant de chaque secteur", {
    colors: COLORS_ANATOLE
  }));

  // KPI : Proportion du CA due aux alumni
  let alumniContributionTable = alumniContribution(dataEtudes);
  charts.contributions.push(createChart(CHART_TYPE.PIE, alumniContributionTable, "Proportion du CA due aux alumni", {
    colors: [COLORS.pine, COLORS.silverPink]
  }));

  // KPI : Proportion du CA venant de la prospection
  let prospectionTurnoverTable = prospectionTurnover(dataProspection.filter(row => row[HEADS.état] == ETAT_PROSP.etude));
  charts.contributions.push(createChart(CHART_TYPE.PIE, prospectionTurnoverTable, "Proportion du CA venant de la prospection", {
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
    /**
     * Ouvre une fenêtre sur le dashboard et y affiche les graphes.
     */
    display: function () {
      let initialTime = new Date();
      ui.showModalDialog(htmlOutput, "KPI");
      measureTime(initialTime, "display the charts");
    },
    /**
     * Enregistre les graphes sous forme d'images (au format .png) dans le dossier spécifié (par défaut dans Pôle Qualité -> KPI -> KPI archivés).
     * @param {string} folderId ID du dossier de destination.
     */
    save: function (folderId = ADDRESSES.driveId) {
      let initialTime = new Date();
      saveOnDrive(attachments, folderId);
      measureTime(initialTime, "save the charts");
    },
    /**
     * Envoie par mail les graphes sous forme d'images (au format .png) à l'adresse spécifiée.
     * @param {string} address Adresse mail du destinataire.
     */
    mail: function (address) {
      let initialTime = new Date();
      sendMail(address, htmlMail, "KPI", attachments);
      measureTime(initialTime, "mail the charts");
    },
    /**
     * Crée un fichier Slides (il s'agit du Point KPI) contenant les KPI regroupés par catégorie.
     */
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
