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
  let data = {};
  Object.values(ACCESS).forEach(spreadsheet => {
    data[spreadsheet.name] = extractSheetData(spreadsheet.id, spreadsheet.sheetName, spreadsheet.pos).filter(spreadsheet.filter);
  });

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
    charts = {};

  currentTime = measureTime(currentTime, "load the HTML content");

  // Boucle sur les catégories
  Object.entries(CATEGORIES).forEach(
    category => {
      charts[category[0]] = [];
      if (Object.keys(category[1].KPIs).length > 0) {
        // Boucle sur les KPIs de la catégorie
        Object.values(category[1].KPIs).forEach(
          KPI => {
            output = KPI.extract(data[KPI.data].filter(KPI.filter || (_ => true)), KPI.options);
            charts[category[0]].push(createChart(KPI.chartType, output.dataTable, KPI.name, output.options));
          })
      }
    });

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
