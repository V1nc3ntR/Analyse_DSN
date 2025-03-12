// DSNService.gs - Service principal pour l'importation et le traitement des DSN

const DSNService = (function() {
  /**
   * Importe et analyse des DSN à partir de l'API
   * @param {object} formData - Données du formulaire d'import
   * @return {object} - Résultat de l'import et analyse
   */
  function importAndAnalyzeDSN(formData) {
    // Paramètres
    const apiIdentifiant = formData.apiIdentifiant;
    const apiMotDePasse = formData.apiMotDePasse;
    const codeDossier = formData.codeDossier;
    const codeEtablissement = formData.codeEtablissement;
    const dateDebut = new Date(formData.dateDebut);
    const dateFin = new Date(formData.dateFin);
    const settings = formData.settings || {};
    
    // Vérification des paramètres obligatoires
    if (!apiIdentifiant || !apiMotDePasse || !codeDossier || !codeEtablissement || !dateDebut || !dateFin) {
      throw new Error("Paramètres d'API incomplets");
    }
    
    // Importer les DSN via l'API
    let apiResult;
    try {
      apiResult = APIHandler.importerDSNForAnalysis({
        apiIdentifiant: apiIdentifiant,
        apiMotDePasse: apiMotDePasse,
        codeDossier: codeDossier,
        codeEtablissement: codeEtablissement,
        dateDebut: dateDebut,
        dateFin: dateFin
      });
    } catch (e) {
      throw new Error("Erreur lors de l'import API: " + e.message);
    }
    
    if (!apiResult.apiDSNData || apiResult.apiDSNData.length === 0) {
      return {
        success: false,
        message: "Aucune DSN trouvée pour la période spécifiée"
      };
    }
    
    // Parser les DSN au nouveau format 2025
    const parsedDSNs = [];
    
    apiResult.apiDSNData.forEach(dsnItem => {
      try {
        const parsedModel = DSNParser2025.parseDSNContent(dsnItem.data);
        parsedDSNs.push({
          annee: dsnItem.annee,
          mois: dsnItem.mois,
          data: parsedModel
        });
      } catch (e) {
        Logger.log("Erreur lors du parsing DSN %s-%s: %s", dsnItem.annee, dsnItem.mois, e);
      }
    });
    
    // Analyser les données DSN
    let analysisReport = null;
    
    if (parsedDSNs.length > 0) {
      try {
        analysisReport = DSNAnalysisService.analyzeAllDSN(parsedDSNs.map(d => d.data));
      } catch (e) {
        Logger.log("Erreur lors de l'analyse des DSN: %s", e);
      }
    }
    
    // Stocker les données si demandé
    if (settings.storeInSheets !== false) {
      storeParsedDSNToSheets(parsedDSNs, settings);
    }
    
    // Générer un rapport si demandé
    let report = null;
    if (settings.generateReport && analysisReport) {
      try {
        if (settings.reportFormat === "pdf") {
          report = DSNReportService.generatePDFReport(analysisReport);
        } else {
          report = DSNReportService.generateSpreadsheetReport(analysisReport);
        }
      } catch (e) {
        Logger.log("Erreur lors de la génération du rapport: %s", e);
      }
    }
    
    return {
      success: true,
      parsedDSNs: parsedDSNs,
      totalDSN: parsedDSNs.length,
      analysisReport: analysisReport,
      report: report
    };
  }
  
  /**
   * Importe et analyse des DSN à partir de fichiers locaux
   * @param {array} dsnFiles - Contenu des fichiers DSN
   * @param {object} settings - Paramètres de traitement
   * @return {object} - Résultat de l'import et analyse
   */
  function importAndAnalyzeLocalDSN(dsnFiles, settings) {
    if (!dsnFiles || dsnFiles.length === 0) {
      return {
        success: false,
        message: "Aucun fichier DSN fourni"
      };
    }
    
    // Parser les DSN au nouveau format 2025
    const parsedDSNs = [];
    
    dsnFiles.forEach((dsnText, index) => {
      try {
        const parsedModel = DSNParser2025.parseDSNContent(dsnText);
        
        // Extraire le mois et l'année à partir du moisPrincipal
        let annee = new Date().getFullYear(); // Par défaut
        let mois = new Date().getMonth() + 1; // Par défaut
        
        if (parsedModel.moisPrincipal) {
          const parts = parsedModel.moisPrincipal.split("-");
          if (parts.length === 2) {
            annee = parseInt(parts[0], 10);
            mois = parseInt(parts[1], 10);
          }
        }
        
        parsedDSNs.push({
          annee: annee,
          mois: mois,
          data: parsedModel
        });
      } catch (e) {
        Logger.log("Erreur lors du parsing du fichier DSN #%s: %s", index + 1, e);
      }
    });
    
    if (parsedDSNs.length === 0) {
      return {
        success: false,
        message: "Aucun fichier DSN n'a pu être analysé correctement"
      };
    }
    
    // Analyser les données DSN
    let analysisReport = null;
    
    try {
      analysisReport = DSNAnalysisService.analyzeAllDSN(parsedDSNs.map(d => d.data));
    } catch (e) {
      Logger.log("Erreur lors de l'analyse des DSN: %s", e);
    }
    
    // Stocker les données si demandé
    if (settings.storeInSheets !== false) {
      storeParsedDSNToSheets(parsedDSNs, settings);
    }
    
    // Générer un rapport si demandé
    let report = null;
    if (settings.generateReport && analysisReport) {
      try {
        if (settings.reportFormat === "pdf") {
          report = DSNReportService.generatePDFReport(analysisReport);
        } else {
          report = DSNReportService.generateSpreadsheetReport(analysisReport);
        }
      } catch (e) {
        Logger.log("Erreur lors de la génération du rapport: %s", e);
      }
    }
    
    return {
      success: true,
      parsedDSNs: parsedDSNs,
      totalDSN: parsedDSNs.length,
      analysisReport: analysisReport,
      report: report
    };
  }
  
  /**
   * Stocke les DSN parsées dans les feuilles de calcul
   * @param {array} parsedDSNs - DSN parsées au format 2025
   * @param {object} settings - Paramètres de stockage
   */
  function storeParsedDSNToSheets(parsedDSNs, settings) {
    if (!parsedDSNs || parsedDSNs.length === 0) {
      return;
    }
    
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    
    // Créer une feuille de suivi si elle n'existe pas
    let trackingSheet = ss.getSheetByName("DSN_Tracking");
    if (!trackingSheet) {
      trackingSheet = ss.insertSheet("DSN_Tracking");
      
      // Initialiser l'en-tête
      trackingSheet.getRange(1, 1, 1, 7).setValues([
        ["Mois", "Année", "Date d'import", "Nb Salariés", "Nb Contrats", "Nb Arrêts", "Chemin vers données"]
      ]);
      
      trackingSheet.getRange(1, 1, 1, 7).setFontWeight("bold").setBackground("#E0E0E0");
    }
    
    parsedDSNs.forEach(dsn => {
      const mm = dsn.mois < 10 ? "0" + dsn.mois : dsn.mois;
      const sheetName = settings.sheetNameFormat === "YYYY-MM" ? 
                        dsn.annee + "-" + mm : 
                        mm + "-" + dsn.annee;
      
      // Vérifier si la feuille existe déjà
      let sheet = ss.getSheetByName(sheetName);
      
      // Supprimer la feuille existante si demandé
      if (sheet && settings.overwriteData) {
        ss.deleteSheet(sheet);
        sheet = null;
      }
      
      if (!sheet) {
        sheet = ss.insertSheet(sheetName);
        
        // Initialiser la structure de la feuille selon le nouveau format
        _initSheetForNewFormat(sheet);
      }
      
      // Stocker les données
      _storeDataInSheet(sheet, dsn.data, settings);
      
      // Mettre à jour la feuille de suivi
      const lastRow = trackingSheet.getLastRow() + 1;
      trackingSheet.getRange(lastRow, 1, 1, 7).setValues([
        [
          mm,
          dsn.annee,
          new Date(),
          dsn.data.salaries ? dsn.data.salaries.length : 0,
          dsn.data.contrats ? dsn.data.contrats.length : 0,
          dsn.data.arrets ? dsn.data.arrets.length : 0,
          "=" + sheetName + "!A1" // Lien vers la feuille des données
        ]
      ]);
      
      // Formater la date
      trackingSheet.getRange(lastRow, 3).setNumberFormat("dd/MM/yyyy HH:mm:ss");
    });
    
    // Trier la feuille de suivi par année et mois
    const dataRange = trackingSheet.getRange(2, 1, trackingSheet.getLastRow() - 1, 7);
    if (dataRange.getNumRows() > 0) {
      dataRange.sort([{column: 2, ascending: true}, {column: 1, ascending: true}]);
    }
  }
  
  /**
   * Initialise la structure d'une feuille pour le nouveau format DSN 2025
   * @private
   * @param {Sheet} sheet - Feuille à initialiser
   */
  function _initSheetForNewFormat(sheet) {
    // Configurer la mise en page
    sheet.setColumnWidth(1, 250);  // Col A
    sheet.setColumnWidth(2, 150);  // Col B
    sheet.setColumnWidth(3, 150);  // Col C
    sheet.setColumnWidth(4, 150);  // Col D
    sheet.setColumnWidth(5, 150);  // Col E
    sheet.setColumnWidth(6, 150);  // Col F
    sheet.setColumnWidth(7, 200);  // Col G
    
    // Créer une structure à onglets en utilisant des sections repliables
    
    // En-tête général
    sheet.getRange("A1:G1").merge();
    sheet.getRange("A1").setValue("DONNÉES DSN DÉTAILLÉES (Format 2025)")
      .setFontSize(14)
      .setFontWeight("bold")
      .setHorizontalAlignment("center")
      .setBackground("#4285F4")
      .setFontColor("white");
    
    // Date d'analyse
    sheet.getRange("A2:G2").merge();
    sheet.getRange("A2").setValue("Importé le " + 
                                 Utilities.formatDate(new Date(), 
                                                     Session.getScriptTimeZone(), 
                                                     "dd/MM/yyyy à HH:mm"))
      .setHorizontalAlignment("center");
    
    // Section 1: Information établissement
    sheet.getRange("A4:G4").merge();
    sheet.getRange("A4").setValue("ÉTABLISSEMENT")
      .setBackground("#E0E0E0")
      .setFontWeight("bold");
    
    sheet.getRange("A5").setValue("SIREN");
    sheet.getRange("A6").setValue("NIC");
    sheet.getRange("A7").setValue("Raison sociale");
    sheet.getRange("A8").setValue("Adresse");
    
    // Section 2: Salariés
    sheet.getRange("A10:G10").merge();
    sheet.getRange("A10").setValue("SALARIÉS")
      .setBackground("#E0E0E0")
      .setFontWeight("bold");
    
    // En-têtes colonnes salariés
    sheet.getRange("A11:G11").setValues([
      ["NIR", "Nom", "Prénom", "Sexe", "Date Naissance", "Contrats", "Actions"]
    ]);
    sheet.getRange("A11:G11").setFontWeight("bold").setBackground("#F3F3F3");
    
    // Section 3: Contrats
    sheet.getRange("A50:G50").merge();
    sheet.getRange("A50").setValue("CONTRATS")
      .setBackground("#E0E0E0")
      .setFontWeight("bold");
    
    // En-têtes colonnes contrats
    sheet.getRange("A51:G51").setValues([
      ["Numéro Contrat", "NIR Salarié", "Nature", "Date Début", "Date Fin", "Statut", "Actions"]
    ]);
    sheet.getRange("A51:G51").setFontWeight("bold").setBackground("#F3F3F3");
    
    // Section 4: Rémunérations
    sheet.getRange("A100:G100").merge();
    sheet.getRange("A100").setValue("RÉMUNÉRATIONS")
      .setBackground("#E0E0E0")
      .setFontWeight("bold");
    
    // En-têtes colonnes rémunérations
    sheet.getRange("A101:G101").setValues([
      ["Numéro Contrat", "Type", "Montant", "Début Période", "Fin Période", "Nb Heures", "Actions"]
    ]);
    sheet.getRange("A101:G101").setFontWeight("bold").setBackground("#F3F3F3");
    
    // Section 5: Arrêts de travail
    sheet.getRange("A150:G150").merge();
    sheet.getRange("A150").setValue("ARRÊTS DE TRAVAIL")
      .setBackground("#E0E0E0")
      .setFontWeight("bold");
    
    // En-têtes colonnes arrêts
    sheet.getRange("A151:G151").setValues([
      ["NIR Salarié", "Motif", "Date Début", "Date Fin", "Date Reprise", "Nb Jours", "Actions"]
    ]);
    sheet.getRange("A151:G151").setFontWeight("bold").setBackground("#F3F3F3");
  }
  
  /**
   * Stocke les données dans une feuille initialisée
   * @private
   * @param {Sheet} sheet - Feuille cible
   * @param {object} data - Données DSN au format 2025
   * @param {object} settings - Paramètres de stockage
   */
  function _storeDataInSheet(sheet, data, settings) {
    // Informations établissement
    if (data.etablissement) {
      sheet.getRange("B5").setValue(data.etablissement.siren);
      sheet.getRange("B6").setValue(data.etablissement.nic);
      sheet.getRange("B7").setValue(data.etablissement.raisonSociale);
      
      if (data.etablissement.adresse) {
        const adresse = [
          data.etablissement.adresse.voie,
          data.etablissement.adresse.codePostal,
          data.etablissement.adresse.commune,
          data.etablissement.adresse.pays
        ].filter(Boolean).join(", ");
        
        sheet.getRange("B8").setValue(adresse);
      }
    }
    
    // Salariés
    if (data.salaries && data.salaries.length > 0) {
      const salarieRows = [];
      
      data.salaries.forEach(salarie => {
        const contratsList = salarie.contrats.join(", ");
        
        salarieRows.push([
          salarie.identite.nir,
          salarie.identite.nom,
          salarie.identite.prenoms,
          salarie.identite.sexe === "01" ? "H" : 
          salarie.identite.sexe === "02" ? "F" : salarie.identite.sexe,
          salarie.identite.dateNaissance,
          contratsList,
          "Voir détails" // Action (pourrait être un lien ou un bouton dans une IHM avancée)
        ]);
      });
      
      if (salarieRows.length > 0) {
        const startRow = 12;
        sheet.getRange(startRow, 1, salarieRows.length, 7).setValues(salarieRows);
        
        // Formater les dates
        sheet.getRange(startRow, 5, salarieRows.length, 1).setNumberFormat("dd/MM/yyyy");
      }
    }
    
    // Contrats
    if (data.contrats && data.contrats.length > 0) {
      const contratRows = [];
      
      data.contrats.forEach(contrat => {
        // Déterminer le statut du contrat
        let statut = "En cours";
        
        if (contrat.id.dateFin) {
          const now = new Date();
          if (contrat.id.dateFin < now) {
            statut = "Terminé";
          } else {
            statut = "À échéance";
          }
        }
        
        contratRows.push([
          contrat.id.numeroContrat,
          contrat.id.nirSalarie,
          _getNatureContratLabel(contrat.caracteristiques.nature),
          contrat.id.dateDebut,
          contrat.id.dateFin,
          statut,
          "Voir détails" // Action
        ]);
      });
      
      if (contratRows.length > 0) {
        const startRow = 52;
        sheet.getRange(startRow, 1, contratRows.length, 7).setValues(contratRows);
        
        // Formater les dates
        sheet.getRange(startRow, 4, contratRows.length, 2).setNumberFormat("dd/MM/yyyy");
      }
    }
    
    // Rémunérations
    const remuRows = [];
    
    if (data.contrats) {
      data.contrats.forEach(contrat => {
        if (contrat.remunerations && contrat.remunerations.length > 0) {
          contrat.remunerations.forEach(remu => {
            remuRows.push([
              contrat.id.numeroContrat,
              _getTypeRemunerationLabel(remu.type),
              remu.montant,
              remu.periode.debut,
              remu.periode.fin,
              remu.nbHeures,
              "" // Action
            ]);
          });
        }
      });
    }
    
    if (remuRows.length > 0) {
      const startRow = 102;
      sheet.getRange(startRow, 1, remuRows.length, 7).setValues(remuRows);
      
      // Formater les dates et montants
      sheet.getRange(startRow, 3, remuRows.length, 1).setNumberFormat("#,##0.00 €");
      sheet.getRange(startRow, 4, remuRows.length, 2).setNumberFormat("dd/MM/yyyy");
      sheet.getRange(startRow, 6, remuRows.length, 1).setNumberFormat("0.00");
    }
    
    // Arrêts de travail
    if (data.arrets && data.arrets.length > 0) {
      const arretRows = [];
      
      data.arrets.forEach(arret => {
        let nbJours = "";
        if (arret.dateDebut && arret.dateFin) {
          const diffTime = Math.abs(arret.dateFin - arret.dateDebut);
          nbJours = Math.ceil(diffTime / (1000 * 60 * 60 * 24)) + 1; // +1 pour inclure le premier jour
        }
        
        arretRows.push([
          arret.nirSalarie,
          _getMotifArretLabel(arret.motif),
          arret.dateDebut,
          arret.dateFin,
          arret.dateReprise,
          nbJours,
          "" // Action
        ]);
      });
      
      if (arretRows.length > 0) {
        const startRow = 152;
        sheet.getRange(startRow, 1, arretRows.length, 7).setValues(arretRows);
        
        // Formater les dates
        sheet.getRange(startRow, 3, arretRows.length, 3).setNumberFormat("dd/MM/yyyy");
      }
    }
  }
  
  /**
   * Récupère le libellé pour une nature de contrat
   * @private
   * @param {string} code - Code nature de contrat
   * @return {string} - Libellé
   */
  function _getNatureContratLabel(code) {
    const natures = {
      "01": "CDI",
      "02": "CDD",
      "03": "Intérim",
      "07": "CDI intermittent",
      "08": "CDD intermittent",
      "09": "Contrat de travail temporaire",
      "10": "Contrat de travail saisonnier",
      "20": "Contrat de mission d'un collaborateur occasionnel du service public",
      "21": "Contrat de mission d'un VRP multicartes",
      "29": "Convention de stage",
      "32": "Contrat d'appui au projet d'entreprise",
      "50": "Nomination",
      "51": "Contrat de mission d'un non-titulaire de la fonction publique",
      "52": "Contrat d'engagement maritime à durée indéterminée",
      "53": "Contrat d'engagement maritime à durée déterminée",
      "60": "Contrat d'engagement éducatif",
      "70": "Contrat de soutien et d'aide par le travail",
      "80": "Mandat social",
      "81": "Mandat d'élu",
      "82": "Contrat de travail à durée indéterminée de chantier",
      "89": "Volontariat de service civique",
      "90": "Sans contrat de travail",
      "91": "Contrat d'engagement maritime",
      "92": "Stage (au sens de la formation professionnelle)",
      "93": "CDD sans terme précis certain"
    };
    
    return natures[code] || code || "Non spécifié";
  }
  
  /**
   * Récupère le libellé pour un type de rémunération
   * @private
   * @param {string} code - Code type de rémunération
   * @return {string} - Libellé
   */
  function _getTypeRemunerationLabel(code) {
    const types = {
      "001": "Rémunération brute non plafonnée",
      "002": "Salaire brut soumis à contributions d'Assurance chômage",
      "003": "Salaire rétabli – reconstitué",
      "010": "Salaire de base",
      "011": "Heures supplémentaires",
      "012": "Heures d'équivalence",
      "013": "Heures d'habillage, déshabillage, pause",
      "014": "Bonus, primes except.",
      "015": "Prime liée à l'activité",
      "016": "Prime d'ancienneté",
      "017": "Prime d'assiduité",
      "018": "Avantages en nature",
      "025": "Indemnité légale de fin de CDD",
      "026": "Indemnité légale de fin de mission",
      "027": "Indemnité légale de licenciement",
      "028": "Indemnité légale de retraite",
      "029": "Indemnité non imposable de licenciement",
      "031": "Indemnités de congés payés",
      "039": "Complément RP",
      "040": "Rémunération intermittent",
      "902": "Prime ou indemnité de transport"
    };
    
    return types[code] || code || "Non spécifié";
  }
  
  /**
   * Récupère le libellé pour un motif d'arrêt
   * @private
   * @param {string} code - Code motif d'arrêt
   * @return {string} - Libellé
   */
  function _getMotifArretLabel(code) {
    const motifs = {
      "01": "Maladie",
      "02": "Maternité",
      "03": "Paternité / accueil de l'enfant",
      "04": "Accident du travail",
      "05": "Accident de trajet",
      "06": "Maladie professionnelle",
      "07": "Temps partiel thérapeutique",
      "08": "Activité partielle",
      "09": "Adoption",
      "10": "Congé suite à accident ou maladie vie privée",
      "11": "Congé de formation professionnelle",
      "12": "Congé de formation de cadres et animateurs pour la jeunesse",
      "13": "Autres"
    };
    
    return motifs[code] || code || "Non spécifié";
  }
  
  // API publique
  return {
    importAndAnalyzeDSN,
    importAndAnalyzeLocalDSN,
    storeParsedDSNToSheets
  };
})();
