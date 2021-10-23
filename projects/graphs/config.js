/* Si vous avez des questions à propos de ce script contactez Aubin Tchoï (Directeur Qualité 022) */

/** Display parameters and various hard-coded values such as values expected inside of a header or month names. */

try {
    var ui = SpreadsheetApp.getUi();
} catch (e) {
    Logger.log(e);
}
/**
 * Id du Spreadsheet contenant les graphes.
 * @constant
 * @readonly
 */
const IDSSKPI = "16EkXSEQPI7LRb45aSbjZAZFPfP4pmx0__4leyLHsDIc";

/** 
 * Enum utilisé pour identifier les différents types de graphes.
 * @enum {Symbol}
 * @constant
 * @readonly
 */
const CHART_TYPE = Object.freeze({
    COLUMN: Symbol("column"),
    PIE: Symbol("pie"),
    LINE: Symbol("line")
});

/* ----- Paramètres d'affichage ----- */
/** 
 * Nom et poste qui seront affichés sur la 1ère slide de présentation.
 * @constant
 * @readonly
 */
const qualityResp = Object.freeze({
    nameTitle: "Paul Invernizzi, Responsable Qualité 022"
}),
    /** 
     * Couleurs des graphes.
     * @constant
     * @readonly
     */
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
    COLORS_OFFICE = Object.freeze([
        "#4472C4", "#954F72", "#C00000", "#ED7D31", "#FFC000", "#70AD47", "#264478", "#592F44", "#730000",
        "#9E480E", "#997300", "#43682B", "#698ED0", "#B16C8E", "#FF0101", "#F1975A", "#FFCD33", "#8CC168"
    ]),
    COLORS_DUO = Object.freeze(["#8e3232", "#561E1E"]),
    /** 
     * Dimensions des graphes.
     * @constant
     * @readonly
     */
    DIMS = {
        width: 1000,
        height: 400
    },
    /** 
     * Liens vers différentes images/gif.
     * @constant
     * @readonly
     */
    IMGS = {
        loadingScreen: "https://raw.githubusercontent.com/aubin-tchoi/Polkali/main/images/loadingScreen.gif",
        thumbsUp: "https://raw.githubusercontent.com/aubin-tchoi/Polkali/main/images/thumbsUp.png"
    },
    /** 
     * Contenu HTML (mail, affichage des KPI et écran de chargement).
     * @constant
     * @readonly
     */
    HTML_CONTENT = {
        display: HtmlService
            .createHtmlOutput(`<span style='font-size: 12pt;'> <span style="font-family: 'roboto', sans-serif;">&nbsp; &nbsp; Voici quelques KPI représentant l'activité de la JE :</span> </span>`)
            .setWidth(1015)
            .setHeight(515),
        mail: HtmlService.createHtmlOutput(`<span style='font-size: 12pt;'> <span style="font-family: 'roboto', sans-serif;">&nbsp; &nbsp; Bonjour, <br/><br/>Voici les KPI portant sur la prospection.<br/> <br/>Bonne journée !</span> </span>`),
        saveConfirm: (url) => HtmlService
            .createHtmlOutput(`<span style='font-size: 12pt;'> <span style="font-family: 'roboto', sans-serif;">Les KPIs ont été enregistrés, pouce pour ouvrir le lien (cliquez sur Boris).<br/><br/> &nbsp; &nbsp; La bise.</span></span><p style="text-align:center;"><a href=${url} target="_blank"><img src="${IMGS.thumbsUp}" alt="Thumbs up" width="130" height="131"></a></p>`)
            .setHeight(235)
            .setWidth(600),
        loadingScreen: HtmlService
            .createHtmlOutput(`<img src="${IMGS.loadingScreen}" alt="Loading" width="442" height="249">`)
            .setWidth(450)
            .setHeight(280)
    };

/* ----- Paramètres de lecture des sheets ----- */
/**  
 * Header des fichiers de données.
 * @constant
 * @readonly
 */
const HEADS = Object.freeze({
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
    concurrence: "Autres JE en concurrence",
    domaine: "Domaine de compétence",
    secteur: "Secteur du Client",
    typeEntreprise: "Type d'entreprise",
    prestation: "Prestation"
}),
    /**  
     * Différents états possible dans la prospection (dans le suivi de la prospection).
     * @constant
     * @readonly
     */
    ETAT_PROSP = Object.freeze({
        rdv: "Premier RDV réalisé",
        devis: "Devis rédigé et envoyé",
        negoc: "En négociation",
        etude: "Etude obtenue"
    }),
    /** 
     * États correspondant à un contact qui n'a pas abouti sur une étude (dans le suivi de la prospection).
     * @constant
     * @readonly
     */
    ETAT_PROSP_BIS = Object.freeze({
        sansSuite: "Sans suite",
        aRelancer: "A relancer",
    }),
    /** 
     * Différents types de contact.
     * @constant
     * @readonly
     */
    CONTACT_TYPE = Object.freeze({
        quali: "Prospection quali ",
        spontané: "Contact spontané",
        classique: "Prospection classique ",
        recommandé: "Recommandé ",
        ao: "Appel d'offre",
        site: "Site",
        redirectionQuali: "Redirection suite à la prospection quali ",
        redirectionClassique: "Redirection suite à la prospection classique "
    }),
    /** 
     * Différents états possible d'une étude (dans le suivi des études).
     * @constant
     * @readonly
     */
    ETAT_ETUDE = Object.freeze({
        negoc: "En négociation",
        redac: "En rédaction",
        enAttente: "En attente",
        enCours: "En cours",
        terminée: "Terminée",
        cloturée: "Clôturée",
        sanSuite: "Sans suite"
    }),
    /** 
     * Différents états possible d'une étude non aboutie (dans le suivi des études).
     * @constant
     * @readonly
     */
    ETAT_ETUDE_BIS = Object.freeze({
        enAttente: "En attente",
        sanSuite: "Sans suite",
        aRelancer: "A relancer",
        negoc: "En négociation",
        redac: "En rédaction"
    }),
    /** 
     * Questions présentes dans le questionnaire satisfaction des élèves.
     * @constant
     * @readonly
     */
    LISTE_QUESTION_RATING_QS_E = Object.freeze([
        "Cela correspondait-il à tes attentes ? (adéquation en matière de compétences, etc.)",
        "Les échanges avec le suiveur étaient-ils de qualité ?",
        "La charge de travail était-elle celle à laquelle tu t’attendais ?",
        "Note la clarté des PEP recrute"
    ]) 
    /** 
    * Questions présentes dans le questionnaire satisfaction des Clients.
    * @constant
    * @readonly
    */
    LISTE_QUESTION_RATING_QS_C = Object.freeze([
        
    ]);

/* ----- Paramètres portant sur les mois ----- */
/** 
 * Index des mois à représenter.
 * @constant
 * @readonly
 */
const MONTH_LIST = Object.freeze([{
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
},
{
    month: 5,
    year: 2021
},
{
    month: 6,
    year: 2021
},
{
    month: 7,
    year: 2021
},
{
    month: 8,
    year: 2021
}
]),
    /**
     * Manière dont les mois sont écrits.
     * @constant
     * @readonly
     */
    MONTH_NAMES = Object.freeze(["Jan", "Fév", "Mars", "Avr", "Mai", "Juin", "Juil", "Août", "Sept", "Oct", "Nov", "Déc"]);

/* ----- Paramètres des Charts ----- */
/**
 * Valeurs par défaut des options utilisées dans la création des Charts.
 * @constant
 * @readonly
 */
const DEFAULT_PARAMS = Object.freeze({
    colors: Object.values(COLORS),
    percent: false,
    hticks: 'auto',
    vticks: 'auto',
    width: DIMS.width,
    height: DIMS.height
});