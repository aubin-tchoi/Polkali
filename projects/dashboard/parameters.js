/* Si vous avez des questions à propos de ce script contactez Aubin Tchoï (Directeur Qualité 022) */

// Parameters : do not put here hardcoded values that are not parameters

try {
    var ui = SpreadsheetApp.getUi();
} catch (e) {
    Logger.log(e);
}

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
 * Couleurs des graphes.
 * @constant
 * @readonly
 */
const qualityResp = Object.freeze({
    nameTitle: "Paul Invernizzi, Responsable Qualité 022"
})
const COLORS = Object.freeze({
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
    },
    /** 
     * Titres des slides de transition.
     * @constant
     * @readonly
     */
    TITLES = {
        summary: "Bilan",
        contactTypology: "Typologie des contacts",
        competitiveness: "Compétitivité face aux autres JE",
        sizeComparison: "Performance sur différentes tailles d'étude",
        contributions: "Mesure des différentes contributions au CA",
        sector: "Performance sur différents secteurs d'activité du Client",
        companySize: "Performance sur différentes tailles d'entreprises",
        contactType: "Performance sur différents types de contact",
        missionType: "Performance sur différentes prestations"
    },

    /* ----- Paramètres d'accès ----- */
    /**  
     * IDs et noms de sheets.
     * @constant
     * @readonly
     */
    ADDRESSES = Object.freeze({
        prospectionId: "1lJhJuZxUt_8_mVLXe5tazXPrb2Z3wr0M49rho974sNQ",
        prospectionName: "Suivi",
        etudesId: "1gmJLAKvUOYFeS32raOiSTYiE_ozr7YSk26Y0t0blm04",
        etudesName: "Suivi",
        etudesIdBis: "1XY8CVuvPpscU2BkkpX0HZ4TqQo0SDWMIh3kzDSvvhcM",
        driveId: "1dPi0dht-q_rI8fUmheA1j861huYPPcAy",
        slidesTemplate: "15WdicqHVF8LtOPrlwdM5iD1_qKh7YPaM15hrGGVbVzU",
        devisSiteId: "1gYRsgfM86D0dw1lsrbEhsIIUxBo-o0bA93vEBHyfCHM",
        devisSiteName: "Réponses au formulaire 1",
        bddProspectionId: "1dKE_5-Yoi_ACJeq0R7nDY0U1Vz3ccvQo",
    }),
    /* ----- Marqueurs ----- */
    /**  
     * Marqueurs des sheets importants.
     * @constant
     * @readonly
     */
    ACCESS = Object.freeze({
        etude2021: {
            name: "etude2021",
            id: "1gmJLAKvUOYFeS32raOiSTYiE_ozr7YSk26Y0t0blm04",
            sheetName: "Suivi",
            pos: {
                data: {
                    x: 5,
                    y: 3
                },
                header: {
                    x: 1,
                    y: 3
                },
                trustColumn: 3
            },
            filter: row => !(Object.values(ETAT_ETUDE_BIS).includes(row[HEADS.état]))
        },
        etude2122: {
            name: "etude2122",
            id: "1zdomLLcx2M5tAo1KWmfUA5YbEMHTQcRYObObvHLdJAI",
            sheetName: "Suivi",
            pos: {
                data: {
                    x: 5,
                    y: 3
                },
                header: {
                    x: 1,
                    y: 3
                },
                trustColumn: 3
            },
            filter: row => !(Object.values(ETAT_ETUDE_BIS).includes(row[HEADS.état]))
        },
        prosp2021: {
            name: "prosp2021",
            id: "1lJhJuZxUt_8_mVLXe5tazXPrb2Z3wr0M49rho974sNQ",
            sheetName: "Suivi",
            pos: {
                data: {
                    x: 4,
                    y: 2
                },
                header: {
                    x: 1,
                    y: 2
                },
                trustColumn: 4
            },
            filter: row => (row[HEADS.premierContact] != "")
        },
        prosp2122: {
            name: "prosp2122",
            id: "1zdomLLcx2M5tAo1KWmfUA5YbEMHTQcRYObObvHLdJAI",
            sheetName: "Suivi",
            pos: {
                data: {
                    x: 4,
                    y: 2
                },
                header: {
                    x: 1,
                    y: 2
                },
                trustColumn: 4
            },
            filter: row => (row[HEADS.premierContact] != "")
        },
        devisSite: {
            name: "devisSite",
            id: "1gYRsgfM86D0dw1lsrbEhsIIUxBo-o0bA93vEBHyfCHM",
            sheetName: "Réponses au formulaire 1",
            pos: {
                data: {
                    x: 2,
                    y: 1
                },
                header: {
                    x: 1,
                    y: 1
                },
                trustColumn: 1
            },
            filter: _ => true
        }
    }),
    /* ----- Paramètres d'accès bdd ----- */
    /**  
     * IDs et noms de sheets.
     * @constant
     * @readonly
     */
    BDDPROSP = Object.freeze({
        bddFC023Id: "1LeAwWXSPEYQu-m24mdjyBjvs7DnecHDatpIyJugB43w",
        bddT023Id: "1AGNmN3qJeS4M2SwKFgPiad6L-SesxJpzc2ANROO71jU",
        bddIng023Id: "1fbvkVGqTUohNY0mbtCXNUEBXL-X2qSPz03EKzfXqHH0",
        bddInd023: "1jP3UMeBRBGXaFbVxA8NuGu7o1a4zm-xLpe-G9j0QuYE",
        bddBTP023: "1wbnP5qAHuQBizKMOXYsoXjQTuXtivpoNZXevBGaBDLU",
        bddBA023: "1SiZy8T7ZyvWvq6KD2cujuWTtHKWMnrEUpx9kGMtL-Dw"
    }),

    /**  
     * Header des fichiers de données.
     * @constant
     * @readonly
     */
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

    /* ----- Liste des KPIs groupés par catégorie ----- */
    /**  
     * Liste des KPIs.
     * @constant
     * @readonly
     */
    CATEGORIES = {
        summary: {
            id: "",
            slideTitle: "Bilan",
            KPIs: {
                contacts: {
                    name: "Contacts",
                    extract: contacts,
                    data: "prosp2021",
                    options: {
                        colors: COLORS_OFFICE
                    },
                    chartType: CHART_TYPE.COLUMN
                },
                conversionMensuel: {
                    name: "Taux de conversion global",
                    extract: conversionRateOverTime,
                    data: "prosp2021",
                    options: {
                        colors: [COLORS.burgundy]
                    },
                    chartType: CHART_TYPE.LINE
                },
                site: {
                    name: "Nombre de contact venant du site",
                    extract: contactBySite,
                    data: "devisSite",
                    options: {
                        colors: [COLORS.pine, COLORS.silverPink]
                    },
                    chartType: CHART_TYPE.COLUMN
                },
                conversionEtapes: {
                    name: "Taux de conversion sur chaque étape",
                    extract: conversionRateByType,
                    data: "prosp2021",
                    options: {
                        colors: [COLORS.burgundy],
                        percent: true
                    },
                    chartType: CHART_TYPE.COLUMN
                },
            }
        },
        contactTypology: {
            id: "",
            slideTitle: "Typologie des contacts",
            KPIs: {
                repartitionDesContactsParType: {
                    name: "Type de contact",
                    extract: (data, options) => totalDistribution(HEADS.typeContact, data, options),
                    data: "prosp2021",
                    options: {
                        colors: COLORS_OFFICE
                    },
                    chartType: CHART_TYPE.PIE
                },
                tauxDeConversionParTypeDeContact: {
                    name: "Taux de conversion par type de contact",
                    extract: (data, options) => conversionRate(HEADS.typeContact, data, options),
                    data: "prosp2021",
                    options: {
                        colors: COLORS_DUO,
                        percent: true
                    },
                    chartType: CHART_TYPE.COLUMN
                },
                repartitionDesContactsParDomaineDeCompétence: {
                    name: "Répartition des contacts par domaine de compétence",
                    extract: (data, options) => totalDistribution(HEADS.domaine, data, options),
                    data: "prosp2021",
                    filter: row => row[HEADS.domaine] != "",
                    options: {
                        colors: COLORS_OFFICE
                    },
                    chartType: CHART_TYPE.PIE
                },
            }
        },
        competitiveness: {
            id: "",
            slideTitle: "Compétitivité face aux autres JE",
            KPIs: {}
        },
        contactType: {
            id: "",
            slideTitle: "Performance sur différents types de contact",
            KPIs: {
                etudeTypeContact: {
                    name: "Performance par type de contact",
                    extract: (data, options) => performanceByContact(HEADS.typeContact, data, options),
                    data: "prosp2021",
                    filter: row => row[HEADS.typeContact] != "",
                    options: {
                        colors: COLORS_OFFICE
                    },
                    chartType: CHART_TYPE.COLUMN
                },
                CATypeContact: {
                    name: "Proportion du CA venant de chaque type de contact",
                    extract: (data, options) => turnoverDistribution(HEADS.typeContact, data, options),
                    data: "prosp2021",
                    filter: row => row[HEADS.état] == ETAT_PROSP.etude && row[HEADS.typeContact] != "",
                    options: {
                        colors: COLORS_OFFICE
                    },
                    chartType: CHART_TYPE.PIE
                }
            }
        },
        companyType: {
            id: "",
            slideTitle: "Performance sur différents types d'entreprises",
            KPIs: {
                etudesTypeEntreprise: {
                    name: "Performance par type d'entreprise",
                    extract: (data, options) => performance(HEADS.typeEntreprise, data, options),
                    data: "prosp2021",
                    filter: row => row[HEADS.état] == ETAT_PROSP.etude && row[HEADS.secteur] != "",
                    options: {
                        colors: COLORS_DUO
                    },
                    chartType: CHART_TYPE.COLUMN
                },
                CATypeEntreprise: {
                    name: "Proportion du CA venant de chaque type d'entreprise",
                    extract: (data, options) => turnoverDistribution(HEADS.typeEntreprise, data, options),
                    data: "prosp2021",
                    filter: row => row[HEADS.état] == ETAT_PROSP.etude && row[HEADS.secteur] != "",
                    options: {
                        colors: COLORS_OFFICE
                    },
                    chartType: CHART_TYPE.PIE
                }
            }
        },
        sector: {
            id: "",
            slideTitle: "Performance sur différents secteurs d'activité du Client",
            KPIs: {
                etudesSecteur: {
                    name: "Performance par secteur",
                    extract: (data, options) => performance(HEADS.secteur, data, options),
                    data: "prosp2021",
                    filter: row => row[HEADS.état] == ETAT_PROSP.etude && row[HEADS.secteur] != "",
                    options: {
                        colors: COLORS_DUO
                    },
                    chartType: CHART_TYPE.COLUMN
                },
                CASecteur: {
                    name: "Proportion du CA venant de chaque secteur",
                    extract: (data, options) => turnoverDistribution(HEADS.secteur, data, options),
                    data: "prosp2021",
                    filter: row => row[HEADS.état] == ETAT_PROSP.etude && row[HEADS.secteur] != "",
                    options: {
                        colors: COLORS_OFFICE
                    },
                    chartType: CHART_TYPE.PIE
                }
            }
        },
        sizeComparison: {
            id: "",
            slideTitle: "Performance sur différentes tailles d'étude",
            KPIs: {
                perfPrix: {
                    name: "Nombre d'études par tranche de prix (en €)",
                    extract: priceRange,
                    data: "etude2021",
                    filter: row => row[HEADS.prix] != "",
                    options: {
                        colors: [COLORS.burgundy],
                        lowerBound: 500,
                        higherBound: 4500,
                        nbrRanges: 8
                    },
                    chartType: CHART_TYPE.COLUMN
                },
                perfNombreJEH: {
                    name: "Nombre d'études par nombre de JEHs",
                    extract: (data, options) => numberOfMissions(HEADS.JEH, data, options),
                    data: "etude2021",
                    filter: row => row[HEADS.JEH] != "",
                    options: {
                        colors: [COLORS.burgundy]
                    },
                    chartType: CHART_TYPE.COLUMN
                },
                perfDureeEtude: {
                    name: "Nombre d'études par durée d'étude (en nombre de semaines)",
                    extract: (data, options) => numberOfMissions(HEADS.durée, data, options),
                    data: "etude2021",
                    filter: row => row[HEADS.durée] != "",
                    options: {
                        colors: [COLORS.burgundy]
                    },
                    chartType: CHART_TYPE.COLUMN
                }
            }
        },
        contributions: {
            id: "",
            slideTitle: "Mesure des différentes contributions au CA",
            KPIs: {
                CAAlumni: {
                    name: "Proportion du CA due aux alumni",
                    extract: (data, options) => turnoverDistributionBinary(HEADS.alumni, data, options),
                    data: "etude2021",
                    options: {
                        colors: [COLORS.pine, COLORS.silverPink]
                    },
                    chartType: CHART_TYPE.PIE
                },
                CAProsp: {
                    name: "Proportion du CA venant de la prospection",
                    extract: prospectionTurnover,
                    data: "prosp2021",
                    filter: row => row[HEADS.état] == ETAT_PROSP.etude,
                    options: {
                        colors: [COLORS.pine, COLORS.silverPink]
                    },
                    chartType: CHART_TYPE.PIE
                }
            }
        },
    },


    /* ----- Paramètres portant sur les mois ----- */
    /** 
     * Index des mois à représenter.
     * @constant
     * @readonly
     */
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
    /**
     * Manière dont les mois sont écrits.
     * @constant
     * @readonly
     */
    MONTH_NAMES = Object.freeze(["Jan", "Fév", "Mars", "Avr", "Mai", "Juin", "Juil", "Août", "Sept", "Oct", "Nov", "Déc"]);
