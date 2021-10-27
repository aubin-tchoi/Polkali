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
            y: 3
          },
          header: {
            x: 1,
            y: 3
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
            y: 3
          },
          header: {
            x: 1,
            y: 3
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
            y: 3
          },
          header: {
            x: 1,
            y: 3
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
      }
    },
    prospMandat: {
      prosp2122: {
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
    }
  }),

  /* ----- Paramètres d'accès bdd ----- */
  /**  
   * IDs et noms des BDD de prospection.
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
          data: "prosp",
          options: {
            colors: COLORS_OFFICE,
            series: {0:{labelInLegend: "Premier RDV réalisé"},
                    1:{labelInLegend: "Devis rédigé et envoyé"},
                    2:{labelInLegend: "En négociation"},
                    3:{labelInLegend: "Etude obtenue"}
          }},
          chartType: CHART_TYPE.COLUMN
        },
        totalConversionRate: {
          name: "Taux de conversion total",
          extract: conversionTotal,
          data: "prospMandat",
          options: {
            colors: COLORS_OFFICE
          },
          chartType: CHART_TYPE.COLUMN
        },
        conversionMensuel: {
          name: "Taux de conversion global",
          extract: conversionRateOverTime,
          data: "prosp",
          options: {
            colors: [COLORS.burgundy]
          },
          chartType: CHART_TYPE.COLUMN
        },
        site: {
          name: "Nombre de contact venant du site",
          extract: contactBySite,
          data: "infoSite",
          options: {
            colors: [COLORS.pine, COLORS.silverPink]
          },
          chartType: CHART_TYPE.COLUMN
        },
        nombresImportants: {
          name : "Nombres d'Etudes potentielles, d'Etudes signées et de CA signé",
          extract: keyNumbers,
          data : "etudeMandat",
          options: {
            colors : [COLORS.burgundy]
          },
          chartType : CHART_TYPE.COLUMN
        },
        conversionEtapes: {
          name: "Taux de conversion sur chaque étape",
          extract: conversionRateByType,
          data: "prosp",
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
          data: "prosp",
          options: {
            colors: COLORS_OFFICE
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
            }
          },
          chartType: CHART_TYPE.COLUMN
        },
        repartitionDesContactsParDomaineDeCompétence: {
          name: "Répartition des contacts par domaine de compétence",
          extract: (data, options) => totalDistribution(HEADS.domaine, data, options),
          data: "prosp",
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
          data: "prosp",
          filter: row => row[HEADS.typeContact] != "",
          options: {
            colors: COLORS_OFFICE,
            series: {0:{labelInLegend: "Nombre de devis envoyés"},
                    1:{labelInLegend: "Nombre d'études signées"},
                    2:{labelInLegend: "CA en milliers d'euros"}
          }},
          chartType: CHART_TYPE.COLUMN
        },
        CATypeContact: {
          name: "Proportion du CA venant de chaque type de contact",
          extract: (data, options) => turnoverDistribution(HEADS.typeContact, data, options),
          data: "prosp",
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
          data: "prosp",
          filter: row => row[HEADS.état] == ETAT_PROSP.etude && row[HEADS.secteur] != "",
          options: {
            colors: COLORS_DUO,
            series: {0:{labelInLegend: "Nombre d'études"},
                    1:{labelInLegend: "CA en milliers d'euros"}
            }
          },
          chartType: CHART_TYPE.COLUMN
        },
        CATypeEntreprise: {
          name: "Proportion du CA venant de chaque type d'entreprise",
          extract: (data, options) => turnoverDistribution(HEADS.typeEntreprise, data, options),
          data: "prosp",
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
          data: "prosp",
          filter: row => row[HEADS.état] == ETAT_PROSP.etude && row[HEADS.secteur] != "",
          options: {
            colors: COLORS_DUO,
          series: {0:{labelInLegend: "Nombre d'études"},
                    1:{labelInLegend: "CA en milliers d'euros"}
            },
          },
          chartType: CHART_TYPE.COLUMN
        },
        CASecteur: {
          name: "Proportion du CA venant de chaque secteur",
          extract: (data, options) => turnoverDistribution(HEADS.secteur, data, options),
          data: "prosp",
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
          data: "etude",
          filter: row => row[HEADS.prix] != "" && !(Object.values(ETAT_ETUDE_BIS).includes(row[HEADS.état])),
          options: {
            colors: [COLORS.burgundy],
            lowerBound: 500,
            higherBound: 4500,
            nbrRanges: 8
          },
          chartType: CHART_TYPE.COLUMN
        }//,
        //perfNombreJEH: {
        //  name: "Nombre d'études par nombre de JEHs",
        //  extract: (data, options) => numberOfMissions(HEADS.JEH, data, options),
        //  data: "etude",
        //  filter: row => row[HEADS.JEH] != "",
        //  options: {
        //    colors: [COLORS.burgundy]
        //  },
        //  chartType: CHART_TYPE.COLUMN
        //},
        //perfDureeEtude: {
        //  name: "Nombre d'études par durée d'étude (en nombre de semaines)",
        //  extract: (data, options) => numberOfMissions(HEADS.durée, data, options),
        //  data: "etude",
        //  filter: row => row[HEADS.durée] != "",
        //  options: {
        //    colors: [COLORS.burgundy]
        //  },
        //  chartType: CHART_TYPE.COLUMN
        //}
      }
    },
    contributions: {
      id: "",
      slideTitle: "Mesure des différentes contributions au CA",
      KPIs: {
        CAAlumni: {
          name: "Proportion du CA due aux alumni",
          extract: (data, options) => turnoverDistributionBinary(HEADS.alumni, data, options),
          data: "etude",
          filter: row => !(Object.values(ETAT_ETUDE_BIS).includes(row[HEADS.état])),
          options: {
            colors: [COLORS.pine, COLORS.silverPink]
          },
          chartType: CHART_TYPE.PIE
        },
        CAProsp: {
          name: "Proportion du CA venant de la prospection",
          extract: prospectionTurnover,
          data: "prosp",
          filter: row => row[HEADS.état] == ETAT_PROSP.etude,
          options: {
            colors: [COLORS.pine, COLORS.silverPink]
          },
          chartType: CHART_TYPE.PIE
        }
      }
    },
  };