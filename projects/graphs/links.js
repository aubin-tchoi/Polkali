/* Si vous avez des questions à propos de ce script contactez Aubin Tchoï (Directeur Qualité 022) */

/** Hardcoded values of IDs and access parameters. */

/**
 * IDs et noms de sheets.
 * @constant
 * @readonly
 */
 const ADDRESSES = Object.freeze({
  driveId: "1dPi0dht-q_rI8fUmheA1j861huYPPcAy",
  slidesTemplate: "15WdicqHVF8LtOPrlwdM5iD1_qKh7YPaM15hrGGVbVzU",
  bddProspectionId: "1dKE_5-Yoi_ACJeq0R7nDY0U1Vz3ccvQo",
}),
  /**
   * IDs et noms de sheet contenant des données.
   * @constant
   * @readonly
   */
  DATA_LINKS = Object.freeze({
    etude: {
      etude2021: {
        id: "1gmJLAKvUOYFeS32raOiSTYiE_ozr7YSk26Y0t0blm04",
        sheetName: "Suivi",
        pos: {
          data: {
            x: 5,
            y: 1
          },
          header: {
            x: 1,
            y: 1
          },
          trustColumn: 3
        },
        filter: row => (Object.values(ETAT_ETUDE).includes(row[HEADS.état]))
      },
      etude2122: {
        id: "1h1rObRvdb2GKxdzgTZD3Kvwde_TUTlTOLC3l8BKzOBQ",
        sheetName: "Suivi",
        pos: {
          data: {
            x: 5,
            y: 1
          },
          header: {
            x: 1,
            y: 1
          },
          trustColumn: 3
        },
        filter: row => (Object.values(ETAT_ETUDE).includes(row[HEADS.état]))
      }
    },
    etudeMandat: {
      etude2122: {
        id: "1h1rObRvdb2GKxdzgTZD3Kvwde_TUTlTOLC3l8BKzOBQ",
        sheetName: "Suivi",
        pos: {
          data: {
            x: 5,
            y: 1
          },
          header: {
            x: 1,
            y: 1
          },
          trustColumn: 3
        },
        filter: row => (Object.values(ETAT_ETUDE).includes(row[HEADS.état]))
      }
    },
    prosp: {
      prosp2021: {
        id: "1lJhJuZxUt_8_mVLXe5tazXPrb2Z3wr0M49rho974sNQ",
        sheetName: "Suivi",
        pos: {
          data: {
            x: 4,
            y: 1
          },
          header: {
            x: 1,
            y: 1
          },
          trustColumn: 4
        },
        filter: row => (row[HEADS.premierContact] != "")
      },
      prosp2122: {
        id: "1zdomLLcx2M5tAo1KWmfUA5YbEMHTQcRYObObvHLdJAI",
        sheetName: "Suivi",
        pos: {
          data: {
            x: 4,
            y: 1
          },
          header: {
            x: 1,
            y: 1
          },
          trustColumn: 4
        },
        filter: row => (row[HEADS.premierContact] != "")
      }
    },
    prospMandat: {
      prosp2122: {
        id: "1zdomLLcx2M5tAo1KWmfUA5YbEMHTQcRYObObvHLdJAI",
        sheetName: "Suivi",
        pos: {
          data: {
            x: 4,
            y: 1
          },
          header: {
            x: 1,
            y: 1
          },
          trustColumn: 4
        },
        filter: row => (row[HEADS.premierContact] != "")
      }
    },
    prospPrec: {
      prosp2021: {
        id: "1lJhJuZxUt_8_mVLXe5tazXPrb2Z3wr0M49rho974sNQ",
        sheetName: "Suivi",
        pos: {
          data: {
            x: 4,
            y: 1
          },
          header: {
            x: 1,
            y: 1
          },
          trustColumn: 4
        },
        filter: row => (row[HEADS.premierContact] != "")
      }
    },
    infoSite: {
      devisSite: {
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
    },
    bdd: {
      banque: {
        id: "1SiZy8T7ZyvWvq6KD2cujuWTtHKWMnrEUpx9kGMtL-Dw",
        sheetName: "BDD",
        pos: {
          data: {
            x: 2,
            y: 1
          },
          header: {
            x: 1,
            y: 1
          },
          trustColumn: 2
        },
        filter: _ => true
      },
      industrie: {
        id: "1jP3UMeBRBGXaFbVxA8NuGu7o1a4zm-xLpe-G9j0QuYE",
        sheetName: "BDD",
        pos: {
          data: {
            x: 2,
            y: 1
          },
          header: {
            x: 1,
            y: 1
          },
          trustColumn: 2
        },
        filter: _ => true
      },
      conseil: {
        id: "1LeAwWXSPEYQu-m24mdjyBjvs7DnecHDatpIyJugB43w",
        sheetName: "BDD",
        pos: {
          data: {
            x: 2,
            y: 1
          },
          header: {
            x: 1,
            y: 1
          },
          trustColumn: 2
        },
        filter: row => row["Finance et conseil"] < 2
      },
      btp: {
        id: "1wbnP5qAHuQBizKMOXYsoXjQTuXtivpoNZXevBGaBDLU",
        sheetName: "BDD",
        pos: {
          data: {
            x: 2,
            y: 1
          },
          header: {
            x: 1,
            y: 1
          },
          trustColumn: 2
        },
        filter: _ => true
      },
      ingenierie: {
        id: "1fbvkVGqTUohNY0mbtCXNUEBXL-X2qSPz03EKzfXqHH0",
        sheetName: "BDD",
        pos: {
          data: {
            x: 2,
            y: 1
          },
          header: {
            x: 1,
            y: 1
          },
          trustColumn: 2
        },
        filter: _ => true
      },
      vet: {
        id: "1AGNmN3qJeS4M2SwKFgPiad6L-SesxJpzc2ANROO71jU",
        sheetName: "BDD",
        pos: {
          data: {
            x: 2,
            y: 1
          },
          header: {
            x: 1,
            y: 1
          },
          trustColumn: 2
        },
        filter: _ => true
      }
    },
    questionnaireClients: {
      questionnaire:{
        id: "198pDrAWgrXHoho5mzl0wix178UraCBrOW-iOXKIy9vA",
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
    },
    questionnaireEtudiants: {
      questionnaire:{
        id: "1bVdfG-KBsuZTJ486Kh00nqOj9InilGKT_T_vkAoINsU",
        sheetName: "Feuille 1",
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
    },
    communicationFacebook: {
      communication: {
        id: "1bVdfG-KBsuZTJ486Kh00nqOj9InilGKT_T_vkAoINsU",
        sheetName: "Facebook",
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
    },
    communicationLinkedIn: {
      communication: {
        id: "1bVdfG-KBsuZTJ486Kh00nqOj9InilGKT_T_vkAoINsU",
        sheetName: "LinkedIn",
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
    }
  }),

  /* ----- Liste des KPIs groupés par catégorie ----- */
  /**
   * Liste des KPIs.
   * @constant
   * @readonly
   */
  CATEGORIES = {
    summary: {
      id: "16EkXSEQPI7LRb45aSbjZAZFPfP4pmx0__4leyLHsDIc",
      slideTitle: "Bilan",
      KPIs: {
        contacts: {
          name: "Contacts",
          extract: contacts,
          data: "prosp",
          options: {
            colors: COLORS_OFFICE,
            series: {0:{labelInLegend: "Premier RDV réalisé"},
                    1:{labelInLegend: "Devis rédigé et envoyé"},
                    2:{labelInLegend: "En négociation"},
                    3:{labelInLegend: "Etude obtenue"},
            },
            htitle : "Date du premier contact (Mois Année)",
            vtitle : "Nombre d'études ayant passé chaque stade"
          },
          chartType: CHART_TYPE.COLUMN
        },
        totalConversionRate: {
          name: "Taux de conversion total",
          extract: conversionTotal,
          data: "prospMandat",
          dataHelp: "etudeMandat",
          options: {
            colors: COLORS_OFFICE,
            htitle : "",
            vtitle : "Pourcentage de conversion (%)"
          },
          chartType: CHART_TYPE.COLUMN
        },
        conversionMensuel: {
          name: "Taux de conversion de devis global",
          extract: conversionRateOverTime,
          data: "prosp",
          dataHelp: "etude",
          options: {
            colors: [COLORS.burgundy],
            htitle : "Date du premier contact (Mois Année)",
            vtitle : "Taux de conversion (%)"
          },
          chartType: CHART_TYPE.COLUMN
        },
        site: {
          name: "Nombre de contact venant du site",
          extract: contactBySite,
          data: "infoSite",
          options: {
            colors: COLORS.silverPink,
            htitle : "Date du contact (Mois Année)",
            vtitle : "Nombre de demande de contact faites sur le site"
          },
          chartType: CHART_TYPE.COLUMN
        },
        nombresImportants: {
          name : "Nombres d'Etudes potentielles, d'Etudes signées et de CA signé",
          extract: keyNumbers,
          data : "etudeMandat",
          options: {
            colors : COLORS.burgundy,
            htitle : "",
            vtitle : ""
            },
          chartType : CHART_TYPE.COLUMN
        },
        conversionEtapes: {
          name: "Taux de conversion sur chaque étape 20-22",
          extract: conversionRateByType,
          data: "prosp",
          options: {
            colors: [COLORS.burgundy],
            htitle : "",
            vtitle : "Pourcentage de conversion (%)",
            percent: true
          },
          chartType: CHART_TYPE.COLUMN
        },
      }
    },
    prospection: {
      id: "1GIfUFsYQ4IZOx8etbQA2lzAgGErBmor_Ve9nS9sd_rY",
      slideTitle: "Nombre de prospection dans le mandat",
      KPIs: {
        nbMailSecteur: {
          name: "Mail par secteur",
          extract: (data,options) => prospectionParSecteur(data,options),
          data: "bdd",
          options: {
            colors: COLORS_OFFICE,
            htitle : "",
            vtitle : ""
          },
          chartType: CHART_TYPE.PIE
        },
        nbMailTotal: {
          name: 'Nb Mail Total',
          extract: (data,options) => prospectionTotalNumber(data,options),
          data: "bdd",
          options: {
            colors: COLORS_OFFICE,
            htitle : "",
            vtitle : ""
          },
          chartType: CHART_TYPE.PIE
        },
      }
    },
    com: {
      id: "1xRghUQsXRLP_kBYqYCT9yGJoU_yzDzh-8AtuXOMtSYM",
      slideTitle: 'Point sur la com',
      KPIs: {
        comFacebook: {
          name: "Com Facebook",
          extract: (data,options) => suiviCom(data,options),
          filter: row => row["Message important"] == true,
          data: "communicationFacebook",
          options: {
            colors: COLORS_OFFICE,
            htitle : "Date de chaque post avec le nombre de post (Mois Année)",
            vtitle : "Nombre de réactions cummulées sur le mois"
          },
          chartType: CHART_TYPE.LINE
        },
        comLinkedIn: {
          name: "Com LinkedIn",
          extract: (data,options) => suiviCom(data,options),
          data: "communicationLinkedIn",
          options: {
            colors: COLORS_OFFICE,
            htitle : "Date de chaque post avec le nombre de post (Mois Année)",
            vtitle : "Nombre de réactions cummulées sur le mois"
          },
          chartType: CHART_TYPE.LINE
        }
      }
    },
    questionnaire: {
      id: "1w_twJAGcbWzMV2kGE27vmA2-Bi48KqY9ca0Gq1q67t0",
      slideTitle: "Réponses aux questionnaires de satisfaction",
      KPIs: {
        satisfactionGeneraleClient: {
          name : "Satisfaction générale Client",
          extract: (data, options) => answerAnalyze(data,options,"Degré de satisfaction général"),
          data: "questionnaireClients",
          options: {
            colors: COLORS_OFFICE,
            htitle : "",
            vtitle : ""
          },
          chartType: CHART_TYPE.PIE
        },
        satisfactionRapiditeClient: {
          name : "Rapidité de traitement de la demande",
          extract: (data, options) => answerAnalyze(data,options,"Êtes-vous satisfait de notre accueil et de nos réponses ? [Rapidité de traitement de la demande]"),
          data: "questionnaireClients",
          options: {
            colors: COLORS_OFFICE,
            htitle : "",
            vtitle : ""
          },
          chartType: CHART_TYPE.PIE
        },
        satisfactionReactiviteClient: {
          name : "Réactivité à vos questions",
          extract: (data, options) => answerAnalyze(data,options,"Êtes-vous satisfait de notre accueil et de nos réponses ? [Réactivité à vos questions]"),
          data: "questionnaireClients",
          options: {
            colors: COLORS_OFFICE,
            htitle : "",
            vtitle : ""
          },
          chartType: CHART_TYPE.PIE
        },
        satisfactionDevisClient: {
          name : "Satisfaction Devis Client",
          extract: (data, options) => answerAnalyze(data,options,"Globalement, êtes-vous satisfait de notre devis ? [Clarté et simplicité du devis]"),
          data: "questionnaireClients",
          options: {
            colors: COLORS_OFFICE,
            htitle : "",
            vtitle : ""
          },
          chartType: CHART_TYPE.PIE
        },
        satisfactionZimbraEtudiant: {
          name : "Satisfaction Zimbra",
          extract: (data, options) => answerAnalyze(data,options,"Clarté du message de présentation de la mission"),
          data: "questionnaireEtudiants",
          options: {
            colors: COLORS_OFFICE,
            htitle : "",
            vtitle : ""
          },
          chartType: CHART_TYPE.PIE
        },
        satisfactionSuiveurEtudiant: {
          name: "Qualité suiveur",
          extract: (data, options) => answerAnalyze(data,options,"Les échanges avec le suiveur étaient-ils de qualité ?"),
          data: "questionnaireEtudiants",
          options: {
            colors: COLORS_OFFICE,
            htitle : "",
            vtitle : ""
          },
          chartType: CHART_TYPE.PIE
        }
      }
    },
    contactTypology: {
      id: "1BRUkav3iEDMtR7p_Gc9KQ1rHXNyavNnXDvAyyDYg4B8",
      slideTitle: "Typologie des contacts",
      KPIs: {
        repartitionDesContactsParTypeActuel: {
          name: "Type de contact (Actuel)",
          extract: (data, options) => totalDistribution(HEADS.typeContact, data, options),
          data: "prospMandat",
          options: {
            colors: COLORS_OFFICE,
            htitle : "",
            vtitle : ""
          },
          chartType: CHART_TYPE.PIE
        },
        tauxDeConversionParTypeDeContact: {
          name: "Taux de conversion par type de contact",
          extract: (data, options) => conversionRate(HEADS.typeContact, data, options),
          data: "prospMandat",
          options: {
            colors: COLORS_DUO,
            percent: true,
            series: {0:{labelInLegend: "Nombre de contact"},
                    1:{labelInLegend: "Taux de conversion global"},
            },
            htitle : "Type de contact",
            vtitle : "Nombre de contact et taux de conversion"
          },
          chartType: CHART_TYPE.COLUMN
        },
        repartitionDesContactsParDomaineDeCompétenceActuel: {
          name: "Contacts par domaine de compétence (Actuel)",
          extract: (data, options) => totalDistribution(HEADS.domaine, data, options),
          data: "prospMandat",
          filter: row => row[HEADS.domaine] != "",
          options: {
            colors: COLORS_OFFICE,
            htitle : "Domaine de compétence",
            vtitle : "Nombre de contacts reçus"
          },
          chartType: CHART_TYPE.PIE
        },
        etudeDomaine:{
          name: "Performance par domaine de compétence",
          extract: (data, options, dataAux) => performanceByContact(HEADS.domaine, data, options, dataAux),
          data: "prospMandat",
          dataHelp: "etudeMandat",
          filter: row => row[HEADS.domaine] != "",
          options: {
            colors: COLORS_OFFICE,
            series: {0:{labelInLegend: "Nombre de devis envoyés"},
                    1:{labelInLegend: "Nombre d'études signées"},
                    2:{labelInLegend: "CA en milliers d'euros"}
            },
            htitle : "Domaine de compétence",
            vtitle : ""
          },
          chartType: CHART_TYPE.COLUMN
        },
        etudeTypeContact: {
          name: "Performance par type de contact",
          extract: (data, options, dataAux) => performanceByContact(HEADS.typeContact, data, options, dataAux),
          data: "prospMandat",
          dataHelp: "etudeMandat",
          filter: row => row[HEADS.typeContact] != "",
          options: {
            colors: COLORS_OFFICE,
            series: {0:{labelInLegend: "Nombre de devis envoyés"},
                    1:{labelInLegend: "Nombre d'études signées"},
                    2:{labelInLegend: "CA en milliers d'euros"}
            },
            htitle : "Type de contact",
            vtitle : ""
          },
          chartType: CHART_TYPE.COLUMN
        },
        CATypeContactActuel: {
          name: "CA par type de contact (Notre Mandat)",
          extract: (data, options, dataAux) => turnoverDistribution2(HEADS.typeContact, data, options, dataAux),
          data: "prospMandat",
          dataHelp: "etudeMandat",
          filter: row => row[HEADS.typeContact] != "",
          options: {
            colors: COLORS_OFFICE,
            htitle : "",
            vtitle : ""
          },
          chartType: CHART_TYPE.PIE
        },
        CADomaineCompetence: {
          name: "CA par domaine de compétence (Notre Mandat)",
          extract: (data, options, dataAux) => turnoverDistribution2(HEADS.domaine, data, options, dataAux),
          data: "prospMandat",
          dataHelp: "etudeMandat",
          filter: row => row[HEADS.domaine] != "",
          optons: {
            colors: COLORS_OFFICE,
            htitle : "",
            vtitle : ""
          },
          chartType: CHART_TYPE.PIE
        }
      }
    },
    companyType: {
      id: "1c0jFZz2z4naCrGzmabfo6nsMDgtKWsAGB9jaA9udxVs",
      slideTitle: "Performance sur différents types d'entreprises",
      KPIs: {
        etudesTypeEntreprise: {
          name: "Performance par type d'entreprise",
          extract: (data, options, dataAux) => performance(HEADS.typeEntreprise, data, options, dataAux),
          data: "prosp",
          dataHelp: "etude",
          filter: row => row[HEADS.secteur] != "",
          options: {
            colors: COLORS_DUO,
            series: {0:{labelInLegend: "Nombre d'études"},
                    1:{labelInLegend: "CA en milliers d'euros"}
            },
            htitle : "",
            vtitle : ""
          },
          chartType: CHART_TYPE.COLUMN
        },
        CATypeEntreprise: {
          name: "Proportion du CA venant de chaque type d'entreprise",
          extract: (data, options) => turnoverDistribution(HEADS.typeEntreprise, data, options),
          data: "prosp",
          filter: row => row[HEADS.état] == ETAT_PROSP.etude && row[HEADS.secteur] != "",
          options: {
            colors: COLORS_OFFICE,
            is3D: true,
            enableInteractivity: true,
            htitle : "",
            vtitle : ""
          },
          chartType: CHART_TYPE.PIE
        }
      }
    },
    sector: {
      id: "1V-WipYoxxPouwbQ2BxGhwvTidAJddtAXAA8Xw5Fp0jE",
      slideTitle: "Performance sur différents secteurs d'activité du Client",
      KPIs: {
        etudesSecteurActuel: {
          name: "Performance par secteur",
          extract: (data, options, dataAux) => performanceByContact(HEADS.secteur, data, options, dataAux),
          data: "prospMandat",
          dataHelp: "etudeMandat",
          filter: row => row[HEADS.secteur] != "",
          options: {
            colors: COLORS_DUO,
          series: {0:{labelInLegend: "Nombre de devis envoyés"},
                    1:{labelInLegend: "Nombre d'études signées"},
                    2:{labelInLegend: "CA en milliers d'euros"}
            },
            htitle : "Secteurs d'entreprises",
            vtitle : ""
          },
          chartType: CHART_TYPE.COLUMN
        },
        etudesFiliereActuel: {
          name: "Performance par filiere",
          extract: (data, options,dataAux) => performanceByContact(HEADS.filiere, data, options, dataAux),
          data: "prospMandat",
          dataHelp: 'etudeMandat',
          filter: row => row[HEADS.secteur] != "",
          options: {
            colors: COLORS_DUO,
            series: {0:{labelInLegend: "Nombre de devis envoyés"},
                    1:{labelInLegend: "Nombre d'études signées"},
                    2:{labelInLegend: "CA en milliers d'euros"}
            },
            htitle : "Filières des Ponts",
            vtitle : ""
          },
          chartType: CHART_TYPE.COLUMN
        },
        CASecteurActuel: {
          name: "Proportion du CA par secteur (Notre Mandat)",
          extract: (data, options,dataAux) => turnoverDistribution2(HEADS.secteur, data, options,dataAux),
          data: "prospMandat",
          dataHelp: "etudeMandat",
          filter: row => row[HEADS.état] == ETAT_PROSP.etude && row[HEADS.secteur] != "",
          options: {
            colors: COLORS_OFFICE,
            htitle : "",
            vtitle : ""
          },
          chartType: CHART_TYPE.PIE
        },
        CASecteurAncien: {
          name: "Proportion du CA par secteur (Mandat précédent)",
          extract: (data, options) => turnoverDistribution(HEADS.secteur, data, options),
          data: "prospPrec",
          filter: row => row[HEADS.état] == ETAT_PROSP.etude && row[HEADS.secteur] != "",
          options: {
            colors: COLORS_OFFICE,
            htitle : "",
            vtitle : ""
          },
          chartType: CHART_TYPE.PIE
        },
        CADomaineActuel: {
          name: "CA par filiere (Notre Mandat)",
          extract: (data, options,dataAux) => turnoverDistribution2(HEADS.filiere, data, options,dataAux),
          data: "prospMandat",
          dataHelp: "etudeMandat",
          filter: row => row[HEADS.état] == ETAT_PROSP.etude && row[HEADS.secteur] != "",
          options: {
            colors: COLORS_OFFICE,
            htitle : "",
            vtitle : ""
          },
          chartType: CHART_TYPE.PIE
        }
      }
    }
  };