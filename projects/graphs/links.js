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
    prospPrec: {
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
      id: "1BRUkav3iEDMtR7p_Gc9KQ1rHXNyavNnXDvAyyDYg4B8",
      slideTitle: "Typologie des contacts",
      KPIs: {
        repartitionDesContactsParTypeActuel: {
          name: "Type de contact (Actuel)",
          extract: (data, options) => totalDistribution(HEADS.typeContact, data, options),
          data: "prospMandat",
          options: {
            colors: COLORS_OFFICE
          },
          chartType: CHART_TYPE.PIE
        },
        repartitionDesContactsParTypeAncien: {
          name: "Type de contact (Ancien)",
          extract: (data, options) => totalDistribution(HEADS.typeContact, data, options),
          data: "prospPrec",
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
        repartitionDesContactsParDomaineDeCompétenceActuel: {
          name: "Contacts par domaine de compétence (Actuel)",
          extract: (data, options) => totalDistribution(HEADS.domaine, data, options),
          data: "prospMandat",
          filter: row => row[HEADS.domaine] != "",
          options: {
            colors: COLORS_OFFICE
          },
          chartType: CHART_TYPE.PIE
        },
        repartitionDesContactsParDomaineDeCompétenceAncien: {
          name: "Contacts par domaine de compétence (Ancien)",
          extract: (data, options) => totalDistribution(HEADS.domaine, data, options),
          data: "prospPrec",
          filter: row => row[HEADS.domaine] != "",
          options: {
            colors: COLORS_OFFICE
          },
          chartType: CHART_TYPE.PIE
        },
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
        CATypeContactActuel: {
          name: "CA par type de contact (Notre Mandat)",
          extract: (data, options) => turnoverDistribution(HEADS.typeContact, data, options),
          data: "prospMandat",
          filter: row => row[HEADS.état] == ETAT_PROSP.etude && row[HEADS.typeContact] != "",
          options: {
            colors: COLORS_OFFICE
          },
          chartType: CHART_TYPE.PIE
        },
        CATypeContactAncien: {
          name: "CA par type de contact (Ancien Mandat)",
          extract: (data, options) => turnoverDistribution(HEADS.typeContact, data, options),
          data: "prospPrec",
          filter: row => row[HEADS.état] == ETAT_PROSP.etude && row[HEADS.typeContact] != "",
          options: {
            colors: COLORS_OFFICE
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
      id: "1V-WipYoxxPouwbQ2BxGhwvTidAJddtAXAA8Xw5Fp0jE",
      slideTitle: "Performance sur différents secteurs d'activité du Client",
      KPIs: {
        etudesSecteurActuel: {
          name: "Performance par secteur (Actuel)",
          extract: (data, options) => performance(HEADS.secteur, data, options),
          data: "prospMandat",
          filter: row => row[HEADS.état] == ETAT_PROSP.etude && row[HEADS.secteur] != "",
          options: {
            colors: COLORS_DUO,
          series: {0:{labelInLegend: "Nombre d'études"},
                    1:{labelInLegend: "CA en milliers d'euros"}
            },
          },
          chartType: CHART_TYPE.COLUMN
        },
        etudesSecteurAncien: {
          name: "Performance par secteur (Ancien)",
          extract: (data, options) => performance(HEADS.secteur, data, options),
          data: "prospPrec",
          filter: row => row[HEADS.état] == ETAT_PROSP.etude && row[HEADS.secteur] != "",
          options: {
            colors: COLORS_DUO,
          series: {0:{labelInLegend: "Nombre d'études"},
                    1:{labelInLegend: "CA en milliers d'euros"}
            },
          },
          chartType: CHART_TYPE.COLUMN
        },
        etudesFiliereActuel: {
          name: "Performance par filiere (Actuel)",
          extract: (data, options) => performance(HEADS.filiere, data, options),
          data: "prospMandat",
          filter: row => row[HEADS.état] == ETAT_PROSP.etude && row[HEADS.secteur] != "",
          options: {
            colors: COLORS_DUO,
          series: {0:{labelInLegend: "Nombre d'études"},
                    1:{labelInLegend: "CA en milliers d'euros"}
            },
          },
          chartType: CHART_TYPE.COLUMN
        },
        etudesFiliereAncien: {
          name: "Performance par filiere (Ancien)",
          extract: (data, options) => performance(HEADS.filiere, data, options),
          data: "prospPrec",
          filter: row => row[HEADS.état] == ETAT_PROSP.etude && row[HEADS.secteur] != "",
          options: {
            colors: COLORS_DUO,
          series: {0:{labelInLegend: "Nombre d'études"},
                    1:{labelInLegend: "CA en milliers d'euros"}
            },
          },
          chartType: CHART_TYPE.COLUMN
        },
        CASecteurActuel: {
          name: "Proportion du CA par secteur (Notre Mandat)",
          extract: (data, options) => turnoverDistribution(HEADS.secteur, data, options),
          data: "prospMandat",
          filter: row => row[HEADS.état] == ETAT_PROSP.etude && row[HEADS.secteur] != "",
          options: {
            colors: COLORS_OFFICE
          },
          chartType: CHART_TYPE.PIE
        },
        CASecteurAncien: {
          name: "Proportion du CA par secteur (Mandat précédent)",
          extract: (data, options) => turnoverDistribution(HEADS.secteur, data, options),
          data: "prospPrec",
          filter: row => row[HEADS.état] == ETAT_PROSP.etude && row[HEADS.secteur] != "",
          options: {
            colors: COLORS_OFFICE
          },
          chartType: CHART_TYPE.PIE
        },
        CADomaineActuel: {
          name: "CA par filiere (Notre Mandat)",
          extract: (data, options) => turnoverDistribution(HEADS.filiere, data, options),
          data: "prospMandat",
          filter: row => row[HEADS.état] == ETAT_PROSP.etude && row[HEADS.secteur] != "",
          options: {
            colors: COLORS_OFFICE
          },
          chartType: CHART_TYPE.PIE
        },
        CADomaineAncien: {
          name: "CA par filiere (Mandat précédent)",
          extract: (data, options) => turnoverDistribution(HEADS.filiere, data, options),
          data: "prospPrec",
          filter: row => row[HEADS.état] == ETAT_PROSP.etude && row[HEADS.secteur] != "",
          options: {
            colors: COLORS_OFFICE
          },
          chartType: CHART_TYPE.PIE
        }
      }
    }
  };