/* Si vous avez des questions à propos de ce script contactez Aubin Tchoï (Directeur Qualité 022) */

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
        concurrence: "Autres JE en concurrence",
        secteur: "Secteur"
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
        saveConfirm: (url) => HtmlService
            .createHtmlOutput(`<span style='font-size: 12pt;'> <span style="font-family: 'roboto', sans-serif;">Les KPIs ont été enregistrés, pouce pour ouvrir le lien (cliquez sur Boris).<br/><br/> &nbsp; &nbsp; La bise.</span></span><p style="text-align:center;"><a href=${url} target="_blank"><img src="${IMGS.thumbsUp}" alt="Thumbs up" width="130" height="131"></a></p>`)
            .setHeight(235)
            .setWidth(600),
        loadingScreen: HtmlService
            .createHtmlOutput(`<img src="${IMGS.loadingScreen}" alt="Loading" width="442" height="249">`)
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
    ]),
    CONTACT_TYPE = Object.freeze({
        quali: "Prospection quali ",
        spontané: "Contact spontané",
        classique: "Prospection classique ",
        recommandé: "Recommandé ",
        ao: "Appel d'offre",
        site: "Site",
        redirectionQuali: "Redirection suite à la prospection quali ",
        redirectionClassique: "Redirection suite à la prospection classique "
    });

// Enum used to choose the type of chart chosen
const CHART_TYPE = Object.freeze({
    COLUMN: Symbol("column"),
    PIE: Symbol("pie"),
    LINE: Symbol("line")
});
