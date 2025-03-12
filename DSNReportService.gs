// DSNReportService.gs - Service de génération de rapports DSN

const DSNReportService = (function() {
  /**
   * Génère un rapport complet au format texte riche dans une feuille de calcul
   * @param {object} analysisData - Données d'analyse DSN
   * @param {string} [title="Rapport d'analyse DSN"] - Titre du rapport
   * @return {object} - Feuille de calcul générée et URL du rapport
   */
  function generateSpreadsheetReport(analysisData, title = "Rapport d'analyse DSN") {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    
    // Créer une nouvelle feuille pour le rapport
    const today = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), "yyyy-MM-dd");
    const reportName = `Rapport_DSN_${today}`;
    
    // Vérifier si la feuille existe déjà et la supprimer le cas échéant
    let reportSheet = ss.getSheetByName(reportName);
    if (reportSheet) {
      ss.deleteSheet(reportSheet);
    }
    
    // Créer une nouvelle feuille
    reportSheet = ss.insertSheet(reportName);
    
    // Formatage global de la feuille
    reportSheet.setColumnWidth(1, 250);  // Colonne A
    reportSheet.setColumnWidth(2, 400);  // Colonne B
    reportSheet.setColumnWidth(3, 200);  // Colonne C
    reportSheet.setColumnWidth(4, 200);  // Colonne D
    reportSheet.setColumnWidth(5, 200);  // Colonne E
    
    // Générer l'en-tête du rapport
    _generateReportHeader(reportSheet, title, analysisData);
    
    // Section 1: Vue générale
    let row = 10;
    row = _generateOverviewSection(reportSheet, row, analysisData);
    
    // Section 2: Rémunérations
    row = _generateSalarySection(reportSheet, row, analysisData);
    
    // Section 3: Contrats
    row = _generateContractSection(reportSheet, row, analysisData);
    
    // Section 4: Alertes et anomalies
    row = _generateAnomalySection(reportSheet, row, analysisData);
    
    // Section 5: Tendances
    row = _generateTrendsSection(reportSheet, row, analysisData);
    
    // NOUVEAUTÉ: Section 6: Index d'égalité professionnelle
    if (analysisData.indexEgalite) {
      row = _generateEqualityIndexSection(reportSheet, row, analysisData.indexEgalite);
    }
    
    // Ajouter des graphiques
    _addCharts(reportSheet, analysisData);
    
    // Retourner la feuille générée et son URL
    return {
      sheet: reportSheet,
      url: ss.getUrl() + "#gid=" + reportSheet.getSheetId()
    };
  }
  
  /**
   * Génère un rapport au format PDF à partir des données d'analyse
   * @param {object} analysisData - Données d'analyse DSN
   * @param {string} [filename] - Nom du fichier PDF (optional)
   * @return {object} - Informations sur le fichier généré
   */
  function generatePDFReport(analysisData, filename) {
    // D'abord générer le rapport dans une feuille de calcul
    const report = generateSpreadsheetReport(analysisData);
    const sheet = report.sheet;
    
    // Nom du fichier par défaut
    if (!filename) {
      const today = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), "yyyy-MM-dd");
      filename = `Rapport_DSN_${today}`;
    }
    
    // Exporter en PDF en utilisant ExportManager si disponible
    if (typeof ExportManager !== 'undefined' && ExportManager.exportSheetToPdf) {
      try {
        const pdfFile = ExportManager.exportSheetToPdf(sheet, filename);
        return {
          success: true,
          file: pdfFile,
          message: "Rapport PDF généré avec succès"
        };
      } catch (e) {
        return {
          success: false,
          message: "Erreur lors de la génération du PDF: " + e.toString()
        };
      }
    } else {
      // Méthode alternative si ExportManager n'est pas disponible
      const spreadsheetId = sheet.getParent().getId();
      const sheetId = sheet.getSheetId();
      
      // Créer l'URL pour l'export PDF
      const url = `https://docs.google.com/spreadsheets/d/${spreadsheetId}/export?` +
                 `format=pdf&` +
                 `size=a4&` +
                 `portrait=true&` +
                 `fitw=true&` +
                 `gid=${sheetId}`;
      
      // Retourner l'URL pour téléchargement manuel
      return {
        success: true,
        downloadUrl: url,
        message: "Utiliser cette URL pour télécharger le rapport PDF"
      };
    }
  }
  
  /**
   * Génère un rapport au format CSV avec les principales données
   * @param {object} analysisData - Données d'analyse DSN
   * @return {string} - Contenu CSV
   */
  function generateCSVReport(analysisData) {
    // Préparer les données pour l'export CSV
    const data = [];
    
    // 1. Informations générales
    data.push({
      section: "Information générales",
      libelle: "Période d'analyse",
      valeur: _formatPeriode(analysisData.global.periodeCouverte)
    });
    
    data.push({
      section: "Information générales",
      libelle: "Nombre de DSN analysées",
      valeur: analysisData.global.totalDSN.toString()
    });
    
    // 2. Effectifs
    data.push({
      section: "Effectifs",
      libelle: "Total des effectifs",
      valeur: analysisData.global.effectifs.total.toString()
    });
    
    data.push({
      section: "Effectifs",
      libelle: "Hommes",
      valeur: analysisData.global.effectifs.hommes.toString()
    });
    
    data.push({
      section: "Effectifs",
      libelle: "Femmes",
      valeur: analysisData.global.effectifs.femmes.toString()
    });
    
    // 3. Répartition par âge
    Object.keys(analysisData.global.effectifs.repartitionAge).forEach(tranche => {
      data.push({
        section: "Répartition par âge",
        libelle: tranche,
        valeur: analysisData.global.effectifs.repartitionAge[tranche].toString()
      });
    });
    
    // 4. Rémunérations
    data.push({
      section: "Rémunérations",
      libelle: "Masse salariale totale",
      valeur: analysisData.global.remuneration.masseSalarialeTotale.toFixed(2) + " €"
    });
    
    data.push({
      section: "Rémunérations",
      libelle: "Rémunération moyenne",
      valeur: analysisData.global.remuneration.remunMoyenne.toFixed(2) + " €"
    });
    
    data.push({
      section: "Rémunérations",
      libelle: "Rémunération médiane",
      valeur: analysisData.global.remuneration.remunMediane.toFixed(2) + " €"
    });
    
    data.push({
      section: "Rémunérations",
      libelle: "Écart salarial H/F",
      valeur: analysisData.global.remuneration.ecartHommesFemmes.toFixed(2) + "%"
    });
    
    // 5. Contrats
    data.push({
      section: "Contrats",
      libelle: "Nombre total de contrats",
      valeur: analysisData.global.contrats.total.toString()
    });
    
    Object.keys(analysisData.global.contrats.typesContrat).forEach(type => {
      let libelle = type;
      // Transformer les codes en libellés plus lisibles
      if (type === "01") libelle = "CDI";
      else if (type === "02") libelle = "CDD";
      else if (type === "03") libelle = "Intérim";
      
      data.push({
        section: "Types de contrats",
        libelle: libelle,
        valeur: analysisData.global.contrats.typesContrat[type].toString()
      });
    });
    
    // 6. Anomalies
    data.push({
      section: "Anomalies",
      libelle: "Écarts salariaux significatifs",
      valeur: analysisData.anomalies.ecartsSalariaux.length.toString()
    });
    
    data.push({
      section: "Anomalies",
      libelle: "Rémunérations anormales",
      valeur: analysisData.anomalies.remunerationAnormales.length.toString()
    });
    
    data.push({
      section: "Anomalies",
      libelle: "Données incomplètes",
      valeur: analysisData.anomalies.incompletes.length.toString()
    });
    
    // NOUVEAUTÉ: 7. Index d'égalité professionnelle
    if (analysisData.indexEgalite && analysisData.indexEgalite.resultat.total !== null) {
      data.push({
        section: "Index d'égalité professionnelle",
        libelle: "Score global",
        valeur: analysisData.indexEgalite.resultat.total + "/100"
      });
      
      // Ajouter le détail des indicateurs
      if (!analysisData.indexEgalite.indicateur1.nonCalculable) {
        data.push({
          section: "Index d'égalité professionnelle",
          libelle: "Indicateur 1: Écart de rémunération",
          valeur: analysisData.indexEgalite.indicateur1.score + "/40"
        });
      }
      
      if (!analysisData.indexEgalite.indicateur2.nonCalculable) {
        data.push({
          section: "Index d'égalité professionnelle",
          libelle: "Indicateur 2: Écart d'augmentations",
          valeur: analysisData.indexEgalite.indicateur2.score + "/20"
        });
      }
      
      if (!analysisData.indexEgalite.indicateur3.nonCalculable) {
        data.push({
          section: "Index d'égalité professionnelle",
          libelle: "Indicateur 3: Écart de promotions",
          valeur: analysisData.indexEgalite.indicateur3.score + "/15"
        });
      }
      
      if (!analysisData.indexEgalite.indicateur4.nonCalculable) {
        data.push({
          section: "Index d'égalité professionnelle",
          libelle: "Indicateur 4: Retour congé maternité",
          valeur: analysisData.indexEgalite.indicateur4.score + "/15"
        });
      }
      
      if (!analysisData.indexEgalite.indicateur5.nonCalculable) {
        data.push({
          section: "Index d'égalité professionnelle",
          libelle: "Indicateur 5: 10 plus hautes rémunérations",
          valeur: analysisData.indexEgalite.indicateur5.score + "/10"
        });
      }
    }
    
    // Générer le CSV en utilisant ExportManager si disponible
    if (typeof ExportManager !== 'undefined' && ExportManager.exportToCsv) {
      return ExportManager.exportToCsv(data);
    } else {
      // Méthode manuelle
      return _manualExportToCsv(data);
    }
  }
  
  /**
   * NOUVEAUTÉ: Génère un rapport dédié à l'Index d'Égalité Professionnelle
   * @param {object} indexEgaliteData - Données de l'Index d'Égalité
   * @param {object} options - Options de rapport
   * @return {object} - Feuille de calcul générée et URL du rapport
   */
  function generateEqualityIndexReport(indexEgaliteData, options = {}) {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    
    // Créer une nouvelle feuille pour le rapport
    const today = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), "yyyy-MM-dd");
    const reportName = options.sheetName || `Index_Egalite_Professionnelle_${today}`;
    
    // Vérifier si la feuille existe déjà et la supprimer le cas échéant
    let reportSheet = ss.getSheetByName(reportName);
    if (reportSheet) {
      ss.deleteSheet(reportSheet);
    }
    
    // Créer une nouvelle feuille
    reportSheet = ss.insertSheet(reportName);
    
    // Formatage global de la feuille
    reportSheet.setColumnWidth(1, 250);  // Colonne A
    reportSheet.setColumnWidth(2, 200);  // Colonne B
    reportSheet.setColumnWidth(3, 200);  // Colonne C
    reportSheet.setColumnWidth(4, 200);  // Colonne D
    reportSheet.setColumnWidth(5, 200);  // Colonne E
    
    // En-tête du rapport
    reportSheet.getRange("A1:E1").merge();
    reportSheet.getRange("A1").setValue("INDEX DE L'ÉGALITÉ PROFESSIONNELLE FEMMES-HOMMES")
      .setFontWeight("bold")
      .setFontSize(16)
      .setHorizontalAlignment("center")
      .setBackground("#4285F4")
      .setFontColor("white");
    
    // Date du rapport
    reportSheet.getRange("A2:E2").merge();
    reportSheet.getRange("A2").setValue("Généré le " + 
                                       Utilities.formatDate(new Date(), 
                                                          Session.getScriptTimeZone(), 
                                                          "dd/MM/yyyy"))
      .setFontStyle("italic")
      .setHorizontalAlignment("center");
    
    // Informations générales
    reportSheet.getRange("A4:B4").merge();
    reportSheet.getRange("A4").setValue("INFORMATIONS GÉNÉRALES")
      .setFontWeight("bold")
      .setBackground("#E0E0E0");
    
    let row = 5;
    
    // Période de référence
    reportSheet.getRange(`A${row}`).setValue("Période de référence:");
    if (indexEgaliteData.metadata.periodeReference.debut && indexEgaliteData.metadata.periodeReference.fin) {
      const debut = Utilities.formatDate(indexEgaliteData.metadata.periodeReference.debut, 
                                        Session.getScriptTimeZone(), "MMMM yyyy");
      const fin = Utilities.formatDate(indexEgaliteData.metadata.periodeReference.fin, 
                                      Session.getScriptTimeZone(), "MMMM yyyy");
      reportSheet.getRange(`B${row}`).setValue(`${debut} à ${fin}`);
    } else {
      reportSheet.getRange(`B${row}`).setValue("Non spécifiée");
    }
    
    row++;
    
    // Effectifs
    reportSheet.getRange(`A${row}`).setValue("Effectif total:");
    reportSheet.getRange(`B${row}`).setValue(indexEgaliteData.metadata.effectifTotal);
    
    row++;
    
    reportSheet.getRange(`A${row}`).setValue("Effectif hommes:");
    reportSheet.getRange(`B${row}`).setValue(indexEgaliteData.metadata.effectifHommes);
    
    row++;
    
    reportSheet.getRange(`A${row}`).setValue("Effectif femmes:");
    reportSheet.getRange(`B${row}`).setValue(indexEgaliteData.metadata.effectifFemmes);
    
    row++;
    
    // Taux de féminisation
    const tauxFeminisation = indexEgaliteData.metadata.effectifTotal > 0 ? 
      (indexEgaliteData.metadata.effectifFemmes / indexEgaliteData.metadata.effectifTotal) * 100 : 0;
    
    reportSheet.getRange(`A${row}`).setValue("Taux de féminisation:");
    reportSheet.getRange(`B${row}`).setValue(tauxFeminisation / 100).setNumberFormat("0.0%");
    
    row += 2;
    
    // Résultat global
    reportSheet.getRange(`A${row}:E${row}`).merge();
    reportSheet.getRange(`A${row}`).setValue("RÉSULTAT GLOBAL")
      .setFontWeight("bold")
      .setBackground("#E0E0E0");
    
    row++;
    
    if (indexEgaliteData.resultat.total !== null) {
      reportSheet.getRange(`A${row}:B${row}`).merge();
      reportSheet.getRange(`A${row}`).setValue(`Index égalité professionnelle: ${indexEgaliteData.resultat.total}/100 points`)
        .setFontWeight("bold")
        .setFontSize(14)
        .setHorizontalAlignment("center");
      
      // Coloriser selon le score
      const score = indexEgaliteData.resultat.total;
      let backgroundColor = "#C8E6C9"; // Vert clair (≥ 85)
      
      if (score < 75) {
        backgroundColor = "#FFCDD2"; // Rouge clair (< 75) - Non conforme
      } else if (score < 85) {
        backgroundColor = "#FFF9C4"; // Jaune clair (75-84)
      }
      
      reportSheet.getRange(`A${row}:B${row}`).setBackground(backgroundColor);
    } else {
      reportSheet.getRange(`A${row}:B${row}`).merge();
      reportSheet.getRange(`A${row}`).setValue("Index non calculable")
        .setFontWeight("bold")
        .setFontSize(14)
        .setHorizontalAlignment("center")
        .setBackground("#F5F5F5");
      
      row++;
      
      reportSheet.getRange(`A${row}:B${row}`).merge();
      reportSheet.getRange(`A${row}`).setValue(
        indexEgaliteData.resultat.publication.methodologie || 
        "Nombre de points calculables insuffisant (< 75 points)"
      )
        .setFontStyle("italic")
        .setHorizontalAlignment("center");
    }
    
    row += 2;
    
    // Détail des indicateurs
    reportSheet.getRange(`A${row}:E${row}`).merge();
    reportSheet.getRange(`A${row}`).setValue("DÉTAIL DES INDICATEURS")
      .setFontWeight("bold")
      .setBackground("#E0E0E0");
    
    row++;
    
    // En-têtes du tableau des indicateurs
    reportSheet.getRange(`A${row}`).setValue("Indicateur");
    reportSheet.getRange(`B${row}`).setValue("Résultat");
    reportSheet.getRange(`C${row}`).setValue("Points obtenus");
    reportSheet.getRange(`D${row}`).setValue("Points possibles");
    reportSheet.getRange(`E${row}`).setValue("Calculable");
    
    reportSheet.getRange(`A${row}:E${row}`).setFontWeight("bold").setBackground("#F3F3F3");
    
    row++;
    
    // Indicateur 1: Écart de rémunération
    _addIndicatorRow(reportSheet, row, 
                   "1 - Écart de rémunération F/H", 
                   indexEgaliteData.indicateur1.ecartGlobal ? 
                     indexEgaliteData.indicateur1.ecartGlobal.toFixed(1) + "%" : "N/A",
                   indexEgaliteData.indicateur1.score,
                   40,
                   !indexEgaliteData.indicateur1.nonCalculable);
    
    row++;
    
    // Indicateur 2: Écart de taux d'augmentation
    _addIndicatorRow(reportSheet, row, 
                   "2 - Écart de taux d'augmentation", 
                   indexEgaliteData.indicateur2.ecartTauxAugmentation ? 
                     indexEgaliteData.indicateur2.ecartTauxAugmentation.toFixed(1) + " pts" : "N/A",
                   indexEgaliteData.indicateur2.score,
                   20,
                   !indexEgaliteData.indicateur2.nonCalculable);
    
    row++;
    
    // Indicateur 3: Écart de taux de promotion
    _addIndicatorRow(reportSheet, row, 
                   "3 - Écart de taux de promotion", 
                   indexEgaliteData.indicateur3.ecartTauxPromotion ? 
                     indexEgaliteData.indicateur3.ecartTauxPromotion.toFixed(1) + " pts" : "N/A",
                   indexEgaliteData.indicateur3.score,
                   15,
                   !indexEgaliteData.indicateur3.nonCalculable);
    
    row++;
    
    // Indicateur 4: Retour de congé maternité
    _addIndicatorRow(reportSheet, row, 
                   "4 - Retour de congé maternité", 
                   indexEgaliteData.indicateur4.tauxRespect ? 
                     indexEgaliteData.indicateur4.tauxRespect.toFixed(1) + "%" : "N/A",
                   indexEgaliteData.indicateur4.score,
                   15,
                   !indexEgaliteData.indicateur4.nonCalculable);
    
    row++;
    
    // Indicateur 5: 10 plus hautes rémunérations
    _addIndicatorRow(reportSheet, row, 
                   "5 - 10 plus hautes rémunérations", 
                   indexEgaliteData.indicateur5.nombreFemmesPlusHautesRemunerations ? 
                     indexEgaliteData.indicateur5.nombreFemmesPlusHautesRemunerations + " femmes" : "N/A",
                   indexEgaliteData.indicateur5.score,
                   10,
                   !indexEgaliteData.indicateur5.nonCalculable);
    
    row += 2;
    
    // Détails par indicateur
    _addDetailedIndicatorsSections(reportSheet, row, indexEgaliteData);
    
    // Retourner la feuille générée et son URL
    return {
      sheet: reportSheet,
      url: ss.getUrl() + "#gid=" + reportSheet.getSheetId()
    };
  }
  
  /**
   * NOUVEAUTÉ: Génère une attestation de conformité pour l'Index d'Égalité Professionnelle
   * @param {object} indexEgaliteData - Données de l'Index d'Égalité
   * @param {object} options - Options de génération
   * @return {object} - Document généré
   */
  function generateEqualityIndexCertificate(indexEgaliteData, options = {}) {
    // Vérifier si l'index est calculable
    if (!indexEgaliteData.resultat.total) {
      return {
        success: false,
        message: "Impossible de générer l'attestation: index non calculable"
      };
    }
    
    // Si l'entreprise n'est pas conforme (score < 75), on ne peut pas générer d'attestation
    if (indexEgaliteData.resultat.total < 75) {
      return {
        success: false,
        message: "Impossible de générer l'attestation: index inférieur à 75 points (non conforme)"
      };
    }
    
    try {
      // Créer un Document Google Docs
      const docTitle = options.title || "Attestation Index Égalité Professionnelle";
      const doc = DocumentApp.create(docTitle);
      const body = doc.getBody();
      
      // En-tête
      const heading = body.appendParagraph("ATTESTATION DE CONFORMITÉ");
      heading.setHeading(DocumentApp.ParagraphHeading.HEADING1);
      heading.setAlignment(DocumentApp.HorizontalAlignment.CENTER);
      
      // Sous-titre
      const subheading = body.appendParagraph("Index de l'égalité professionnelle femmes-hommes");
      subheading.setHeading(DocumentApp.ParagraphHeading.HEADING2);
      subheading.setAlignment(DocumentApp.HorizontalAlignment.CENTER);
      
      // Espace
      body.appendParagraph("");
      
      // Corps du document
      const intro = body.appendParagraph("Je soussigné(e), ......................................., " +
                                        "agissant en qualité de ........................................, " +
                                        `atteste sur l'honneur que l'entreprise ${options.companyName || "..........................."} ` +
                                        `(SIREN: ${options.siren || "..........................."}) ` +
                                        "est en conformité avec ses obligations en matière d'égalité professionnelle entre les femmes et les hommes.");
      intro.setAlignment(DocumentApp.HorizontalAlignment.JUSTIFY);
      
      body.appendParagraph("");
      
      // Résultat
      const scoreText = body.appendParagraph(`Pour la période du ${_formatDateFr(indexEgaliteData.metadata.periodeReference.debut)} ` +
                                           `au ${_formatDateFr(indexEgaliteData.metadata.periodeReference.fin)}, ` +
                                           `l'Index de l'égalité professionnelle obtenu est de ${indexEgaliteData.resultat.total}/100 points.`);
      scoreText.setAlignment(DocumentApp.HorizontalAlignment.JUSTIFY);
      scoreText.setBold(true);
      
      body.appendParagraph("");
      
      // Détail des indicateurs
      const detailIntro = body.appendParagraph("Détail des indicateurs:");
      detailIntro.setAlignment(DocumentApp.HorizontalAlignment.LEFT);
      
      // Indicateur 1
      _addIndicatorTextToCertificate(body, 
                                   "Écart de rémunération entre les femmes et les hommes", 
                                   indexEgaliteData.indicateur1, 40);
      
      // Indicateur 2
      _addIndicatorTextToCertificate(body, 
                                   "Écarts de taux d'augmentations individuelles", 
                                   indexEgaliteData.indicateur2, 20);
      
      // Indicateur 3
      _addIndicatorTextToCertificate(body, 
                                   "Écarts de taux de promotions", 
                                   indexEgaliteData.indicateur3, 15);
      
      // Indicateur 4
      _addIndicatorTextToCertificate(body, 
                                   "Pourcentage de salariées augmentées au retour de congé maternité", 
                                   indexEgaliteData.indicateur4, 15);
      
      // Indicateur 5
      _addIndicatorTextToCertificate(body, 
                                   "Nombre de salariés du sexe sous-représenté parmi les 10 plus hautes rémunérations", 
                                   indexEgaliteData.indicateur5, 10);
      
      body.appendParagraph("");
      body.appendParagraph("");
      
      // Signature
      const signatureDate = body.appendParagraph("Fait à ........................., le " + 
                                               Utilities.formatDate(new Date(), Session.getScriptTimeZone(), "dd/MM/yyyy"));
      signatureDate.setAlignment(DocumentApp.HorizontalAlignment.RIGHT);
      
      body.appendParagraph("");
      body.appendParagraph("");
      body.appendParagraph("");
      
      const signature = body.appendParagraph("Signature et cachet");
      signature.setAlignment(DocumentApp.HorizontalAlignment.RIGHT);
      
      // Enregistrer et fermer le document
      doc.saveAndClose();
      
      return {
        success: true,
        document: doc,
        url: doc.getUrl(),
        message: "Attestation générée avec succès"
      };
    } catch (e) {
      return {
        success: false,
        message: "Erreur lors de la génération de l'attestation: " + e.toString()
      };
    }
  }
  
  /**
   * Génère un tableau de bord détaillé pour l'ensemble des données
   * @param {object} dashboard - Données du tableau de bord
   * @return {object} - Données du tableau de bord
   */
  function generateGenderPayGapReport(dashboard) {
    // Statistiques générales
    const stats = dashboard.stats;
    
    // Analyse des arrêts de travail
    const arrets = dashboard.arrets;
    
    // Rapports de conformité
    const compliance = dashboard.compliance;
    
    // Détection des anomalies salariales
    const anomalies = dashboard.anomalies;
    
    // NOUVEAUTÉ: Index d'égalité professionnelle
    const indexEgalite = dashboard.indexEgalite;
    
    // Agréger toutes les données dans un objet de tableau de bord
    return {
      stats: stats,
      arrets: arrets,
      compliance: compliance,
      anomalies: anomalies,
      indexEgalite: indexEgalite,
      lastUpdate: new Date()
    };
  }
  
  /**
   * Génère un rapport au format Excel pour exportation
   * @param {object} dashboard - Données du tableau de bord
   * @return {Spreadsheet} - Feuille Excel générée
   */
  function exportDashboardToSheet(dashboard) {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    let reportSheet = ss.getSheetByName("Tableau de bord");
    
    if (reportSheet) {
      // Supprimer la feuille existante pour la recréer
      ss.deleteSheet(reportSheet);
    }
    
    reportSheet = ss.insertSheet("Tableau de bord");
    
    // Titre du rapport
    reportSheet.getRange("A1:F1").merge();
    reportSheet.getRange("A1").setValue("TABLEAU DE BORD DSN")
      .setFontWeight("bold")
      .setFontSize(16)
      .setHorizontalAlignment("center");
    
    // Date du rapport
    reportSheet.getRange("A2:F2").merge();
    reportSheet.getRange("A2").setValue("Généré le " + Utilities.formatDate(new Date(), Session.getScriptTimeZone(), "dd/MM/yyyy à HH:mm"))
      .setFontStyle("italic")
      .setHorizontalAlignment("center");
    
    // Section 1: Statistiques globales
    reportSheet.getRange("A4").setValue("STATISTIQUES GLOBALES")
      .setFontWeight("bold")
      .setFontSize(14);
    
    reportSheet.getRange("A5").setValue("Nombre total d'individus:");
    reportSheet.getRange("B5").setValue(dashboard.stats.global.totalIndividus);
    
    reportSheet.getRange("A6").setValue("Hommes:");
    reportSheet.getRange("B6").setValue(dashboard.stats.global.hommes);
    
    reportSheet.getRange("A7").setValue("Femmes:");
    reportSheet.getRange("B7").setValue(dashboard.stats.global.femmes);
    
    reportSheet.getRange("A8").setValue("Salaire moyen:");
    reportSheet.getRange("B8").setValue(dashboard.stats.global.salaireMoyen)
      .setNumberFormat("0.00 €");
    
    reportSheet.getRange("A9").setValue("Salaire médian:");
    reportSheet.getRange("B9").setValue(dashboard.stats.global.salaireMedian)
      .setNumberFormat("0.00 €");
    
    reportSheet.getRange("A10").setValue("Écart salarial H/F:");
    reportSheet.getRange("B10").setValue(dashboard.stats.global.ecartHommesFemmes + "%")
      .setNumberFormat("0.00%");
    
    // Section 2: Répartition par tranche d'âge
    reportSheet.getRange("A12").setValue("RÉPARTITION PAR TRANCHE D'ÂGE")
      .setFontWeight("bold")
      .setFontSize(14);
    
    reportSheet.getRange("A13:E13").setValues([["Tranche d'âge", "Total", "Hommes", "Femmes", "Salaire moyen"]]);
    
    let row = 14;
    Object.keys(dashboard.stats.tranchesAge).forEach(tranche => {
      const t = dashboard.stats.tranchesAge[tranche];
      reportSheet.getRange(`A${row}:E${row}`).setValues([
        [tranche, t.total, t.hommes, t.femmes, t.salaireMoyen]
      ]);
      reportSheet.getRange(`E${row}`).setNumberFormat("0.00 €");
      row++;
    });
    
    // Section 3: Anomalies salariales
    row += 2;
    reportSheet.getRange(`A${row}`).setValue("ANOMALIES SALARIALES")
      .setFontWeight("bold")
      .setFontSize(14);
    
    if (dashboard.anomalies.length > 0) {
      row++;
      reportSheet.getRange(`A${row}:H${row}`).setValues([
        ["Mois", "NIR", "Nom", "Prénom", "Contrat", "Ancien salaire", "Nouveau salaire", "Variation"]
      ]);
      
      dashboard.anomalies.forEach(a => {
        row++;
        reportSheet.getRange(`A${row}:H${row}`).setValues([
          [a.mois, a.nir, a.nom, a.prenom, a.contrat, a.ancienSalaire, a.nouveauSalaire, a.variation]
        ]);
        reportSheet.getRange(`F${row}:G${row}`).setNumberFormat("0.00 €");
      });
    } else {
      row++;
      reportSheet.getRange(`A${row}`).setValue("Aucune anomalie salariale détectée.")
        .setFontStyle("italic");
    }
    
    // Section 4: Arrêts de travail
    row += 2;
    reportSheet.getRange(`A${row}`).setValue("ARRÊTS DE TRAVAIL")
      .setFontWeight("bold")
      .setFontSize(14);
    
    row++;
    reportSheet.getRange(`A${row}`).setValue("Nombre total d'arrêts:");
    reportSheet.getRange(`B${row}`).setValue(dashboard.arrets.total);
    
    row++;
    reportSheet.getRange(`A${row}`).setValue("Durée moyenne des arrêts:");
    reportSheet.getRange(`B${row}`).setValue(dashboard.arrets.dureeMoyenne.toFixed(1) + " jours");
    
    row += 2;
    reportSheet.getRange(`A${row}`).setValue("Répartition par motif");
    
    row++;
    reportSheet.getRange(`A${row}:D${row}`).setValues([
      ["Motif", "Nombre", "% du total", "Durée moyenne"]
    ]);
    
    Object.keys(dashboard.arrets.parMotif).forEach(motif => {
      if (dashboard.arrets.parMotif[motif].count > 0) {
        row++;
        const pct = (dashboard.arrets.parMotif[motif].count / dashboard.arrets.total) * 100;
        reportSheet.getRange(`A${row}:D${row}`).setValues([
          [
            dashboard.arrets.parMotif[motif].libelle,
            dashboard.arrets.parMotif[motif].count,
            pct.toFixed(1) + "%",
            dashboard.arrets.parMotif[motif].dureeMoyenne.toFixed(1) + " jours"
          ]
        ]);
      }
    });
    
    // NOUVEAUTÉ: Section 5: Index d'égalité professionnelle 
    if (dashboard.indexEgalite && dashboard.indexEgalite.resultat) {
      row += 3;
      reportSheet.getRange(`A${row}`).setValue("INDEX D'ÉGALITÉ PROFESSIONNELLE")
        .setFontWeight("bold")
        .setFontSize(14);
      
      row++;
      if (dashboard.indexEgalite.resultat.total !== null) {
        reportSheet.getRange(`A${row}`).setValue("Score global:");
        reportSheet.getRange(`B${row}`).setValue(dashboard.indexEgalite.resultat.total + "/100")
          .setFontWeight("bold");
        
        // Mettre en évidence si le score est non conforme (< 75)
        if (dashboard.indexEgalite.resultat.total < 75) {
          reportSheet.getRange(`B${row}`).setBackground("#FFCDD2"); // Rouge clair
        }
      } else {
        reportSheet.getRange(`A${row}`).setValue("Index non calculable")
          .setFontWeight("bold");
      }
      
      row += 2;
      reportSheet.getRange(`A${row}`).setValue("Détail des indicateurs:");
      
      row++;
      reportSheet.getRange(`A${row}:C${row}`).setValues([
        ["Indicateur", "Score obtenu", "Score maximum"]
      ]);
      
      // Indicateur 1
      if (!dashboard.indexEgalite.indicateur1.nonCalculable) {
        row++;
        reportSheet.getRange(`A${row}:C${row}`).setValues([
          ["1 - Écart de rémunération", dashboard.indexEgalite.indicateur1.score, 40]
        ]);
      }
      
      // Indicateur 2
      if (!dashboard.indexEgalite.indicateur2.nonCalculable) {
        row++;
        reportSheet.getRange(`A${row}:C${row}`).setValues([
          ["2 - Écart de taux d'augmentation", dashboard.indexEgalite.indicateur2.score, 20]
        ]);
      }
      
      // Indicateur 3
      if (!dashboard.indexEgalite.indicateur3.nonCalculable) {
        row++;
        reportSheet.getRange(`A${row}:C${row}`).setValues([
          ["3 - Écart de taux de promotion", dashboard.indexEgalite.indicateur3.score, 15]
        ]);
      }
      
      // Indicateur 4
      if (!dashboard.indexEgalite.indicateur4.nonCalculable) {
        row++;
        reportSheet.getRange(`A${row}:C${row}`).setValues([
          ["4 - Retour de congé maternité", dashboard.indexEgalite.indicateur4.score, 15]
        ]);
      }
      
      // Indicateur 5
      if (!dashboard.indexEgalite.indicateur5.nonCalculable) {
        row++;
        reportSheet.getRange(`A${row}:C${row}`).setValues([
          ["5 - 10 plus hautes rémunérations", dashboard.indexEgalite.indicateur5.score, 10]
        ]);
      }
    }
    
    // Appliquer un formatage global
    reportSheet.autoResizeColumns(1, 8);
    
    return reportSheet;
  }
  
  // ===== NOUVEAUTÉ: Fonctions privées pour les rapports d'égalité =====
  
  /**
   * Ajoute une ligne d'indicateur dans le rapport d'Index d'Égalité
   * @private
   * @param {Sheet} sheet - Feuille de rapport
   * @param {number} row - Numéro de ligne
   * @param {string} label - Libellé de l'indicateur
   * @param {string} result - Résultat de l'indicateur
   * @param {number} score - Score obtenu
   * @param {number} maxScore - Score maximum
   * @param {boolean} calculable - Indicateur calculable
   */
  function _addIndicatorRow(sheet, row, label, result, score, maxScore, calculable) {
    sheet.getRange(`A${row}`).setValue(label);
    sheet.getRange(`B${row}`).setValue(result);
    
    if (calculable) {
      sheet.getRange(`C${row}`).setValue(score);
      sheet.getRange(`D${row}`).setValue(maxScore);
      sheet.getRange(`E${row}`).setValue("Oui");
    } else {
      sheet.getRange(`C${row}`).setValue("N/A");
      sheet.getRange(`D${row}`).setValue(maxScore);
      sheet.getRange(`E${row}`).setValue("Non");
    }
  }
  
  /**
   * Ajoute un indicateur à l'attestation
   * @private
   * @param {Body} body - Corps du document
   * @param {string} title - Titre de l'indicateur
   * @param {object} indicator - Données de l'indicateur
   * @param {number} maxScore - Score maximum
   */
  function _addIndicatorTextToCertificate(body, title, indicator, maxScore) {
    let text;
    
    if (indicator.nonCalculable) {
      text = `• ${title}: Non calculable (${indicator.motifNonCalculable || "effectifs insuffisants"})`;
    } else {
      text = `• ${title}: ${indicator.score}/${maxScore} points`;
    }
    
    const paragraph = body.appendParagraph(text);
    paragraph.setAlignment(DocumentApp.HorizontalAlignment.LEFT);
  }
  
  /**
   * Ajoute les sections détaillées des indicateurs
   * @private
   * @param {Sheet} sheet - Feuille de rapport
   * @param {number} startRow - Ligne de début
   * @param {object} indexEgaliteData - Données de l'Index d'Égalité
   * @return {number} - Nouvelle ligne courante
   */
  function _addDetailedIndicatorsSections(sheet, startRow, indexEgaliteData) {
    let row = startRow;
    
    // Détail de l'indicateur 1 - Écart de rémunération
    if (!indexEgaliteData.indicateur1.nonCalculable && indexEgaliteData.indicateur1.parCategorie.length > 0) {
      sheet.getRange(`A${row}:E${row}`).merge();
      sheet.getRange(`A${row}`).setValue("DÉTAIL DE L'INDICATEUR 1: ÉCART DE RÉMUNÉRATION")
        .setFontWeight("bold")
        .setBackground("#E0E0E0");
      
      row++;
      
      // En-têtes du tableau détaillé
      sheet.getRange(`A${row}`).setValue("Catégorie");
      sheet.getRange(`B${row}`).setValue("Tranche d'âge");
      sheet.getRange(`C${row}`).setValue("Rému. Hommes");
      sheet.getRange(`D${row}`).setValue("Rému. Femmes");
      sheet.getRange(`E${row}`).setValue("Écart (%)");
      
      sheet.getRange(`A${row}:E${row}`).setFontWeight("bold").setBackground("#F3F3F3");
      
      row++;
      
      // Données détaillées par catégorie
      indexEgaliteData.indicateur1.parCategorie.forEach(cat => {
        sheet.getRange(`A${row}`).setValue(cat.categorie);
        sheet.getRange(`B${row}`).setValue(cat.trancheAge);
        sheet.getRange(`C${row}`).setValue(cat.remuHommes).setNumberFormat("#,##0.00 €");
        sheet.getRange(`D${row}`).setValue(cat.remuFemmes).setNumberFormat("#,##0.00 €");
        sheet.getRange(`E${row}`).setValue(cat.ecart / 100).setNumberFormat("0.0%");
        
        // Colorer l'écart selon son importance
        if (cat.ecart > 5) {
          sheet.getRange(`E${row}`).setBackground("#FFCDD2"); // Rouge clair (défavorable aux femmes)
        } else if (cat.ecart < -5) {
          sheet.getRange(`E${row}`).setBackground("#C8E6C9"); // Vert clair (défavorable aux hommes)
        }
        
        row++;
      });
      
      // Écart global pondéré
      row++;
      sheet.getRange(`A${row}:D${row}`).merge();
      sheet.getRange(`A${row}`).setValue("Écart global pondéré:").setFontWeight("bold");
      sheet.getRange(`E${row}`).setValue(indexEgaliteData.indicateur1.ecartGlobal / 100)
        .setNumberFormat("0.0%")
        .setFontWeight("bold");
      
      row += 2;
    }
    
    // Détail de l'indicateur 2 - Écart de taux d'augmentation
    if (!indexEgaliteData.indicateur2.nonCalculable) {
      sheet.getRange(`A${row}:E${row}`).merge();
      sheet.getRange(`A${row}`).setValue("DÉTAIL DE L'INDICATEUR 2: ÉCART DE TAUX D'AUGMENTATION")
        .setFontWeight("bold")
        .setBackground("#E0E0E0");
      
      row++;
      
      sheet.getRange(`A${row}`).setValue("Pourcentage d'hommes augmentés:");
      sheet.getRange(`B${row}`).setValue(indexEgaliteData.indicateur2.tauxAugmentationHommes / 100)
        .setNumberFormat("0.0%");
      
      row++;
      
      sheet.getRange(`A${row}`).setValue("Pourcentage de femmes augmentées:");
      sheet.getRange(`B${row}`).setValue(indexEgaliteData.indicateur2.tauxAugmentationFemmes / 100)
        .setNumberFormat("0.0%");
      
      row++;
      
      sheet.getRange(`A${row}`).setValue("Écart en points de %:");
      sheet.getRange(`B${row}`).setValue(indexEgaliteData.indicateur2.ecartTauxAugmentation);
      
      row++;
      
      sheet.getRange(`A${row}`).setValue("Nombre d'hommes pris en compte:");
      sheet.getRange(`B${row}`).setValue(indexEgaliteData.indicateur2.nombreHommes);
      
      row++;
      
      sheet.getRange(`A${row}`).setValue("Nombre de femmes prises en compte:");
      sheet.getRange(`B${row}`).setValue(indexEgaliteData.indicateur2.nombreFemmes);
      
      row += 2;
    }
    
    // Détail de l'indicateur 3 - Écart de taux de promotion
    if (!indexEgaliteData.indicateur3.nonCalculable) {
      sheet.getRange(`A${row}:E${row}`).merge();
      sheet.getRange(`A${row}`).setValue("DÉTAIL DE L'INDICATEUR 3: ÉCART DE TAUX DE PROMOTION")
        .setFontWeight("bold")
        .setBackground("#E0E0E0");
      
      row++;
      
      sheet.getRange(`A${row}`).setValue("Pourcentage d'hommes promus:");
      sheet.getRange(`B${row}`).setValue(indexEgaliteData.indicateur3.tauxPromotionHommes / 100)
        .setNumberFormat("0.0%");
      
      row++;
      
      sheet.getRange(`A${row}`).setValue("Pourcentage de femmes promues:");
      sheet.getRange(`B${row}`).setValue(indexEgaliteData.indicateur3.tauxPromotionFemmes / 100)
        .setNumberFormat("0.0%");
      
      row++;
      
      sheet.getRange(`A${row}`).setValue("Écart en points de %:");
      sheet.getRange(`B${row}`).setValue(indexEgaliteData.indicateur3.ecartTauxPromotion);
      
      row++;
      
      sheet.getRange(`A${row}`).setValue("Nombre d'hommes pris en compte:");
      sheet.getRange(`B${row}`).setValue(indexEgaliteData.indicateur3.nombreHommes);
      
      row++;
      
      sheet.getRange(`A${row}`).setValue("Nombre de femmes prises en compte:");
      sheet.getRange(`B${row}`).setValue(indexEgaliteData.indicateur3.nombreFemmes);
      
      row += 2;
    }
    
    // Détail de l'indicateur 4 - Retour de congé maternité
    if (!indexEgaliteData.indicateur4.nonCalculable) {
      sheet.getRange(`A${row}:E${row}`).merge();
      sheet.getRange(`A${row}`).setValue("DÉTAIL DE L'INDICATEUR 4: RETOUR DE CONGÉ MATERNITÉ")
        .setFontWeight("bold")
        .setBackground("#E0E0E0");
      
      row++;
      
      sheet.getRange(`A${row}`).setValue("Nombre de retours de congé maternité:");
      sheet.getRange(`B${row}`).setValue(indexEgaliteData.indicateur4.nombreRetourCongeMaternite);
      
      row++;
      
      sheet.getRange(`A${row}`).setValue("Nombre de femmes augmentées au retour:");
      sheet.getRange(`B${row}`).setValue(indexEgaliteData.indicateur4.nombreAugmentees);
      
      row++;
      
      sheet.getRange(`A${row}`).setValue("Taux de respect (%):");
      sheet.getRange(`B${row}`).setValue(indexEgaliteData.indicateur4.tauxRespect / 100)
        .setNumberFormat("0.0%");
      
      row += 2;
    }
    
    // Détail de l'indicateur 5 - 10 plus hautes rémunérations
    if (!indexEgaliteData.indicateur5.nonCalculable) {
      sheet.getRange(`A${row}:E${row}`).merge();
      sheet.getRange(`A${row}`).setValue("DÉTAIL DE L'INDICATEUR 5: 10 PLUS HAUTES RÉMUNÉRATIONS")
        .setFontWeight("bold")
        .setBackground("#E0E0E0");
      
      row++;
      
      sheet.getRange(`A${row}`).setValue("Nombre de femmes parmi les 10+ hautes rémunérations:");
      sheet.getRange(`B${row}`).setValue(indexEgaliteData.indicateur5.nombreFemmesPlusHautesRemunerations);
      
      row++;
      
      sheet.getRange(`A${row}`).setValue("Nombre d'hommes parmi les 10+ hautes rémunérations:");
      sheet.getRange(`B${row}`).setValue(indexEgaliteData.indicateur5.nombreHommesPlusHautesRemunerations);
    }
    
    return row;
  }
  
  /**
   * Génère la section d'index d'égalité professionnelle dans le rapport principal
   * @private
   * @param {Sheet} sheet - Feuille de calcul
   * @param {number} startRow - Ligne de début
   * @param {object} indexEgalite - Données de l'Index d'Égalité
   * @return {number} - Nouvelle ligne courante
   */
  function _generateEqualityIndexSection(sheet, startRow, indexEgalite) {
    let row = startRow;
    
    // Titre de section
    sheet.getRange(`A${row}:E${row}`).merge();
    sheet.getRange(`A${row}`).setValue("INDEX D'ÉGALITÉ PROFESSIONNELLE")
      .setFontWeight("bold")
      .setBackground("#F1F3F4");
    
    row += 2;
    
    // Résultat global
    if (indexEgalite.resultat.total !== null) {
      sheet.getRange(`A${row}:B${row}`).merge();
      sheet.getRange(`A${row}`).setValue(`Index égalité professionnelle: ${indexEgalite.resultat.total}/100 points`)
        .setFontWeight("bold")
        .setFontSize(14)
        .setHorizontalAlignment("center");
      
      // Coloriser selon le score
      const score = indexEgalite.resultat.total;
      let backgroundColor = "#C8E6C9"; // Vert clair (≥ 85)
      
      if (score < 75) {
        backgroundColor = "#FFCDD2"; // Rouge clair (< 75) - Non conforme
      } else if (score < 85) {
        backgroundColor = "#FFF9C4"; // Jaune clair (75-84)
      }
      
      sheet.getRange(`A${row}:B${row}`).setBackground(backgroundColor);
    } else {
      sheet.getRange(`A${row}:B${row}`).merge();
      sheet.getRange(`A${row}`).setValue("Index non calculable")
        .setFontWeight("bold")
        .setFontSize(14)
        .setHorizontalAlignment("center")
        .setBackground("#F5F5F5");
      
      row++;
      
      sheet.getRange(`A${row}:B${row}`).merge();
      sheet.getRange(`A${row}`).setValue(
        indexEgalite.resultat.publication.methodologie || 
        "Nombre de points calculables insuffisant (< 75 points)"
      )
        .setFontStyle("italic")
        .setHorizontalAlignment("center");
    }
    
    row += 2;
    
    // Détail des indicateurs
    sheet.getRange(`A${row}`).setValue("Détail des indicateurs");
    sheet.getRange(`A${row}`).setFontWeight("bold");
    
    row++;
    
    // En-têtes du tableau des indicateurs
    sheet.getRange(`A${row}:D${row}`).setValues([
      ["Indicateur", "Score", "Maximum", "Calculable"]
    ]);
    sheet.getRange(`A${row}:D${row}`).setFontWeight("bold").setBackground("#F3F3F3");
    
    row++;
    
    // Indicateur 1
    sheet.getRange(`A${row}`).setValue("1 - Écart de rémunération");
    sheet.getRange(`B${row}`).setValue(indexEgalite.indicateur1.nonCalculable ? "N/A" : indexEgalite.indicateur1.score);
    sheet.getRange(`C${row}`).setValue(40);
    sheet.getRange(`D${row}`).setValue(indexEgalite.indicateur1.nonCalculable ? "Non" : "Oui");
    
    row++;
    
    // Indicateur 2
    sheet.getRange(`A${row}`).setValue("2 - Écart de taux d'augmentation");
    sheet.getRange(`B${row}`).setValue(indexEgalite.indicateur2.nonCalculable ? "N/A" : indexEgalite.indicateur2.score);
    sheet.getRange(`C${row}`).setValue(20);
    sheet.getRange(`D${row}`).setValue(indexEgalite.indicateur2.nonCalculable ? "Non" : "Oui");
    
    row++;
    
    // Indicateur 3
    sheet.getRange(`A${row}`).setValue("3 - Écart de taux de promotion");
    sheet.getRange(`B${row}`).setValue(indexEgalite.indicateur3.nonCalculable ? "N/A" : indexEgalite.indicateur3.score);
    sheet.getRange(`C${row}`).setValue(15);
    sheet.getRange(`D${row}`).setValue(indexEgalite.indicateur3.nonCalculable ? "Non" : "Oui");
    
    row++;
    
    // Indicateur 4
    sheet.getRange(`A${row}`).setValue("4 - Retour de congé maternité");
    sheet.getRange(`B${row}`).setValue(indexEgalite.indicateur4.nonCalculable ? "N/A" : indexEgalite.indicateur4.score);
    sheet.getRange(`C${row}`).setValue(15);
    sheet.getRange(`D${row}`).setValue(indexEgalite.indicateur4.nonCalculable ? "Non" : "Oui");
    
    row++;
    
    // Indicateur 5
    sheet.getRange(`A${row}`).setValue("5 - 10 plus hautes rémunérations");
    sheet.getRange(`B${row}`).setValue(indexEgalite.indicateur5.nonCalculable ? "N/A" : indexEgalite.indicateur5.score);
    sheet.getRange(`C${row}`).setValue(10);
    sheet.getRange(`D${row}`).setValue(indexEgalite.indicateur5.nonCalculable ? "Non" : "Oui");
    
    // Ajouter un espace après la section
    row += 2;
    
    return row;
  }
  
  /**
   * Format une date au format français (dd/MM/yyyy)
   * @private
   * @param {Date} date - Date à formater
   * @return {string} - Date formatée
   */
  function _formatDateFr(date) {
    if (!date) return "";
    return Utilities.formatDate(date, Session.getScriptTimeZone(), "dd/MM/yyyy");
  }
  
  /**
   * Génère l'en-tête du rapport
   * @private
   * @param {Sheet} sheet - Feuille de calcul
   * @param {string} title - Titre du rapport
   * @param {object} data - Données d'analyse
   */
  function _generateReportHeader(sheet, title, data) {
    // Titre principal
    sheet.getRange("A1:E1").merge();
    sheet.getRange("A1").setValue(title)
      .setFontSize(16)
      .setFontWeight("bold")
      .setHorizontalAlignment("center")
      .setBackground("#4285F4")
      .setFontColor("white");
    
    // Date du rapport
    sheet.getRange("A2:E2").merge();
    sheet.getRange("A2").setValue("Généré le " + 
                                 Utilities.formatDate(new Date(), 
                                                    Session.getScriptTimeZone(), 
                                                    "dd/MM/yyyy à HH:mm"))
      .setFontStyle("italic")
      .setHorizontalAlignment("center");
    
    // Période d'analyse
    sheet.getRange("A3:E3").merge();
    sheet.getRange("A3").setValue("Période analysée : " + _formatPeriode(data.global.periodeCouverte))
      .setFontWeight("bold")
      .setHorizontalAlignment("center");
    
    // Nombre de DSN analysées
    sheet.getRange("A4:E4").merge();
    sheet.getRange("A4").setValue(`${data.global.totalDSN} DSN analysées`)
      .setHorizontalAlignment("center");
    
    // Séparateur
    sheet.getRange("A6:E6").merge();
    sheet.getRange("A6").setValue("SYNTHÈSE")
      .setBackground("#E0E0E0")
      .setFontWeight("bold")
      .setHorizontalAlignment("center");
    
    // Insérer les 4 principaux KPIs
    const kpis = [
      {
        title: "Effectif total",
        value: data.global.effectifs.total,
        format: "0"
      },
      {
        title: "Salaire moyen",
        value: data.global.remuneration.remunMoyenne,
        format: "0.00 €"
      },
      {
        title: "Contrats",
        value: data.global.contrats.total,
        format: "0"
      },
      {
        title: "Écart H/F",
        value: data.global.remuneration.ecartHommesFemmes,
        format: "0.00%"
      }
    ];
    
    sheet.getRange("A7:A8").merge();
    sheet.getRange("B7:B8").merge();
    sheet.getRange("C7:C8").merge();
    sheet.getRange("D7:D8").merge();
    sheet.getRange("E7:E8").merge();
    
    sheet.getRange("A7").setValue(kpis[0].title).setHorizontalAlignment("center").setFontWeight("bold");
    sheet.getRange("B7").setValue(kpis[0].value).setHorizontalAlignment("center").setFontSize(16);
    
    sheet.getRange("C7").setValue(kpis[1].title).setHorizontalAlignment("center").setFontWeight("bold");
    sheet.getRange("D7").setValue(kpis[1].value).setHorizontalAlignment("center").setFontSize(16)
      .setNumberFormat(kpis[1].format);
    
    sheet.getRange("A9:A10").merge();
    sheet.getRange("B9:B10").merge();
    sheet.getRange("C9:C10").merge();
    sheet.getRange("D9:D10").merge();
    sheet.getRange("E9:E10").merge();
    
    sheet.getRange("A9").setValue(kpis[2].title).setHorizontalAlignment("center").setFontWeight("bold");
    sheet.getRange("B9").setValue(kpis[2].value).setHorizontalAlignment("center").setFontSize(16);
    
    sheet.getRange("C9").setValue(kpis[3].title).setHorizontalAlignment("center").setFontWeight("bold");
    sheet.getRange("D9").setValue(kpis[3].value / 100).setHorizontalAlignment("center").setFontSize(16)
      .setNumberFormat(kpis[3].format);
  }
  
  /**
   * Génère la section vue d'ensemble
   * @private
   * @param {Sheet} sheet - Feuille de calcul
   * @param {number} startRow - Ligne de début
   * @param {object} data - Données d'analyse
   * @return {number} - Nouvelle ligne courante
   */
  function _generateOverviewSection(sheet, startRow, data) {
    let row = startRow;
    
    // Titre de section
    sheet.getRange(`A${row}:E${row}`).merge();
    sheet.getRange(`A${row}`).setValue("VUE D'ENSEMBLE DES EFFECTIFS")
      .setBackground("#F1F3F4")
      .setFontWeight("bold");
    
    row += 2;
    
    // Répartition par sexe
    sheet.getRange(`A${row}`).setValue("Répartition par sexe");
    sheet.getRange(`A${row}`).setFontWeight("bold");
    
    row++;
    sheet.getRange(`A${row}`).setValue("Hommes");
    sheet.getRange(`B${row}`).setValue(data.global.effectifs.hommes);
    sheet.getRange(`C${row}`).setValue(
      data.global.effectifs.total > 0 ? 
      (data.global.effectifs.hommes / data.global.effectifs.total) * 100 : 0
    ).setNumberFormat("0.0%");
    
    row++;
    sheet.getRange(`A${row}`).setValue("Femmes");
    sheet.getRange(`B${row}`).setValue(data.global.effectifs.femmes);
    sheet.getRange(`C${row}`).setValue(
      data.global.effectifs.total > 0 ? 
      (data.global.effectifs.femmes / data.global.effectifs.total) * 100 : 0
    ).setNumberFormat("0.0%");
    
    row += 2;
    
    // Répartition par âge
    sheet.getRange(`A${row}`).setValue("Répartition par tranche d'âge");
    sheet.getRange(`A${row}`).setFontWeight("bold");
    
    row++;
    
    for (const tranche in data.global.effectifs.repartitionAge) {
      sheet.getRange(`A${row}`).setValue(tranche);
      sheet.getRange(`B${row}`).setValue(data.global.effectifs.repartitionAge[tranche]);
      sheet.getRange(`C${row}`).setValue(
        data.global.effectifs.total > 0 ? 
        (data.global.effectifs.repartitionAge[tranche] / data.global.effectifs.total) * 100 : 0
      ).setNumberFormat("0.0%");
      row++;
    }
    
    // Ajouter un espace après la section
    row += 2;
    
    return row;
  }
  
  /**
   * Génère la section des rémunérations
   * @private
   * @param {Sheet} sheet - Feuille de calcul
   * @param {number} startRow - Ligne de début
   * @param {object} data - Données d'analyse
   * @return {number} - Nouvelle ligne courante
   */
  function _generateSalarySection(sheet, startRow, data) {
    let row = startRow;
    
    // Titre de section
    sheet.getRange(`A${row}:E${row}`).merge();
    sheet.getRange(`A${row}`).setValue("ANALYSE DES RÉMUNÉRATIONS")
      .setBackground("#F1F3F4")
      .setFontWeight("bold");
    
    row += 2;
    
    // Statistiques générales
    sheet.getRange(`A${row}`).setValue("Statistiques générales");
    sheet.getRange(`A${row}`).setFontWeight("bold");
    
    row++;
    sheet.getRange(`A${row}`).setValue("Masse salariale totale");
    sheet.getRange(`B${row}`).setValue(data.global.remuneration.masseSalarialeTotale)
      .setNumberFormat("#,##0.00 €");
    
    row++;
    sheet.getRange(`A${row}`).setValue("Rémunération moyenne");
    sheet.getRange(`B${row}`).setValue(data.global.remuneration.remunMoyenne)
      .setNumberFormat("#,##0.00 €");
    
    row++;
    sheet.getRange(`A${row}`).setValue("Rémunération médiane");
    sheet.getRange(`B${row}`).setValue(data.global.remuneration.remunMediane)
      .setNumberFormat("#,##0.00 €");
    
    row++;
    sheet.getRange(`A${row}`).setValue("Rémunération minimum");
    sheet.getRange(`B${row}`).setValue(data.global.remuneration.remunMin)
      .setNumberFormat("#,##0.00 €");
    
    row++;
    sheet.getRange(`A${row}`).setValue("Rémunération maximum");
    sheet.getRange(`B${row}`).setValue(data.global.remuneration.remunMax)
      .setNumberFormat("#,##0.00 €");
    
    row += 2;
    
    // Écart salarial hommes/femmes
    sheet.getRange(`A${row}`).setValue("Écart salarial hommes/femmes");
    sheet.getRange(`A${row}`).setFontWeight("bold");
    
    row++;
    sheet.getRange(`A${row}`).setValue("Écart global");
    sheet.getRange(`B${row}`).setValue(data.global.remuneration.ecartHommesFemmes / 100)
      .setNumberFormat("0.00%");
    
    // Colorer l'écart selon son importance
    if (Math.abs(data.global.remuneration.ecartHommesFemmes) > 20) {
      sheet.getRange(`B${row}`).setBackground("#FFCDD2"); // Rouge clair (écart important)
    } else if (Math.abs(data.global.remuneration.ecartHommesFemmes) > 10) {
      sheet.getRange(`B${row}`).setBackground("#FFF9C4"); // Jaune clair (écart moyen)
    } else {
      sheet.getRange(`B${row}`).setBackground("#C8E6C9"); // Vert clair (écart faible)
    }
    
    row += 2;
    
    // Rémunération par tranche d'âge
    sheet.getRange(`A${row}`).setValue("Rémunération par tranche d'âge");
    sheet.getRange(`A${row}`).setFontWeight("bold");
    
    row++;
    sheet.getRange(`A${row}:C${row}`).setValues([
      ["Tranche d'âge", "Rémunération moyenne", "Écart avec la moyenne globale"]
    ]);
    sheet.getRange(`A${row}:C${row}`).setFontWeight("bold").setBackground("#F3F3F3");
    
    row++;
    for (const tranche in data.global.remuneration.remunParTrancheAge) {
      const donnees = data.global.remuneration.remunParTrancheAge[tranche];
      const ecart = donnees.moyenne && data.global.remuneration.remunMoyenne ?
        ((donnees.moyenne - data.global.remuneration.remunMoyenne) / data.global.remuneration.remunMoyenne) * 100 : 0;
      
      sheet.getRange(`A${row}`).setValue(tranche);
      sheet.getRange(`B${row}`).setValue(donnees.moyenne).setNumberFormat("#,##0.00 €");
      sheet.getRange(`C${row}`).setValue(ecart / 100).setNumberFormat("0.00%");
      
      row++;
    }
    
    // Ajouter un espace après la section
    row += 2;
    
    return row;
  }
  
  /**
   * Génère la section des contrats
   * @private
   * @param {Sheet} sheet - Feuille de calcul
   * @param {number} startRow - Ligne de début
   * @param {object} data - Données d'analyse
   * @return {number} - Nouvelle ligne courante
   */
  function _generateContractSection(sheet, startRow, data) {
    let row = startRow;
    
    // Titre de section
    sheet.getRange(`A${row}:E${row}`).merge();
    sheet.getRange(`A${row}`).setValue("ANALYSE DES CONTRATS")
      .setBackground("#F1F3F4")
      .setFontWeight("bold");
    
    row += 2;
    
    // Répartition par type de contrat
    sheet.getRange(`A${row}`).setValue("Répartition par type de contrat");
    sheet.getRange(`A${row}`).setFontWeight("bold");
    
    row++;
    sheet.getRange(`A${row}:C${row}`).setValues([
      ["Type de contrat", "Nombre", "Pourcentage"]
    ]);
    sheet.getRange(`A${row}:C${row}`).setFontWeight("bold").setBackground("#F3F3F3");
    
    row++;
    const nbContrats = data.global.contrats.total;
    
    for (const type in data.global.contrats.typesContrat) {
      let libelle = type;
      // Transformer les codes en libellés plus lisibles
      if (type === "01") libelle = "CDI";
      else if (type === "02") libelle = "CDD";
      else if (type === "03") libelle = "Intérim";
      else if (type === "07") libelle = "Contrat à durée indéterminée intermittent";
      else if (type === "08") libelle = "Contrat à durée déterminée intermittent";
      else if (type === "09") libelle = "Contrat de travail temporaire";
      else if (type === "10") libelle = "Contrat de travail saisonnier";
      
      const nombre = data.global.contrats.typesContrat[type];
      const pourcentage = nbContrats > 0 ? (nombre / nbContrats) * 100 : 0;
      
      sheet.getRange(`A${row}`).setValue(libelle);
      sheet.getRange(`B${row}`).setValue(nombre);
      sheet.getRange(`C${row}`).setValue(pourcentage / 100).setNumberFormat("0.0%");
      
      row++;
    }
    
    row++;
    
    // Temps plein vs temps partiel
    sheet.getRange(`A${row}`).setValue("Temps plein vs temps partiel");
    sheet.getRange(`A${row}`).setFontWeight("bold");
    
    row++;
    sheet.getRange(`A${row}:C${row}`).setValues([
      ["Type", "Nombre", "Pourcentage"]
    ]);
    sheet.getRange(`A${row}:C${row}`).setFontWeight("bold").setBackground("#F3F3F3");
    
    row++;
    const tempPlein = data.global.contrats.tempPlein;
    const tempsPartiel = data.global.contrats.tempsPartiel;
    
    sheet.getRange(`A${row}`).setValue("Temps plein");
    sheet.getRange(`B${row}`).setValue(tempPlein);
    sheet.getRange(`C${row}`).setValue(nbContrats > 0 ? (tempPlein / nbContrats) * 100 : 0).setNumberFormat("0.0%");
    
    row++;
    sheet.getRange(`A${row}`).setValue("Temps partiel");
    sheet.getRange(`B${row}`).setValue(tempsPartiel);
    sheet.getRange(`C${row}`).setValue(nbContrats > 0 ? (tempsPartiel / nbContrats) * 100 : 0).setNumberFormat("0.0%");
    
    // Ajouter un espace après la section
    row += 2;
    
    return row;
  }
  
  /**
   * Génère la section des anomalies
   * @private
   * @param {Sheet} sheet - Feuille de calcul
   * @param {number} startRow - Ligne de début
   * @param {object} data - Données d'analyse
   * @return {number} - Nouvelle ligne courante
   */
  function _generateAnomalySection(sheet, startRow, data) {
    let row = startRow;
    
    // Titre de section
    sheet.getRange(`A${row}:E${row}`).merge();
    sheet.getRange(`A${row}`).setValue("ALERTES ET ANOMALIES")
      .setBackground("#F1F3F4")
      .setFontWeight("bold");
    
    row += 2;
    
    // Écarts salariaux significatifs
    sheet.getRange(`A${row}`).setValue("Écarts salariaux significatifs entre hommes et femmes");
    sheet.getRange(`A${row}`).setFontWeight("bold");
    
    if (data.anomalies.ecartsSalariaux.length > 0) {
      row++;
      sheet.getRange(`A${row}:E${row}`).setValues([
        ["Profil", "Moyenne Hommes", "Moyenne Femmes", "Écart (%)", "Effectifs"]
      ]);
      sheet.getRange(`A${row}:E${row}`).setFontWeight("bold").setBackground("#F3F3F3");
      
      row++;
      data.anomalies.ecartsSalariaux.forEach(ecart => {
        sheet.getRange(`A${row}`).setValue(ecart.profil);
        sheet.getRange(`B${row}`).setValue(ecart.moyenneHommes).setNumberFormat("#,##0.00 €");
        sheet.getRange(`C${row}`).setValue(ecart.moyenneFemmes).setNumberFormat("#,##0.00 €");
        sheet.getRange(`D${row}`).setValue(ecart.ecartPourcentage / 100).setNumberFormat("0.00%");
        sheet.getRange(`E${row}`).setValue(`H: ${ecart.nbHommes}, F: ${ecart.nbFemmes}`);
        
        // Colorer l'écart selon son importance
        if (ecart.ecartPourcentage > 20) {
          sheet.getRange(`D${row}`).setBackground("#FFCDD2"); // Rouge clair (écart important)
        } else if (ecart.ecartPourcentage > 10) {
          sheet.getRange(`D${row}`).setBackground("#FFF9C4"); // Jaune clair (écart moyen)
        }
        
        row++;
      });
    } else {
      row++;
      sheet.getRange(`A${row}`).setValue("Aucun écart significatif détecté");
      row++;
    }
    
    row++;
    
    // Rémunérations anormales
    sheet.getRange(`A${row}`).setValue("Rémunérations anormales par rapport à la moyenne du profil");
    sheet.getRange(`A${row}`).setFontWeight("bold");
    
    if (data.anomalies.remunerationAnormales.length > 0) {
      row++;
      sheet.getRange(`A${row}:E${row}`).setValues([
        ["Nom", "Prénom", "Profil", "Rémunération", "Écart (%)"]
      ]);
      sheet.getRange(`A${row}:E${row}`).setFontWeight("bold").setBackground("#F3F3F3");
      
      row++;
      data.anomalies.remunerationAnormales.forEach(anomalie => {
        sheet.getRange(`A${row}`).setValue(anomalie.nom);
        sheet.getRange(`B${row}`).setValue(anomalie.prenom);
        sheet.getRange(`C${row}`).setValue(anomalie.profil);
        sheet.getRange(`D${row}`).setValue(anomalie.remuneration).setNumberFormat("#,##0.00 €");
        sheet.getRange(`E${row}`).setValue(anomalie.ecartPourcentage / 100).setNumberFormat("0.00%");
        
        // Colorer l'écart selon son importance
        if (Math.abs(anomalie.ecartPourcentage) > 50) {
          sheet.getRange(`E${row}`).setBackground("#FFCDD2"); // Rouge clair (écart très important)
        } else if (Math.abs(anomalie.ecartPourcentage) > 30) {
          sheet.getRange(`E${row}`).setBackground("#FFF9C4"); // Jaune clair (écart important)
        }
        
        row++;
      });
    } else {
      row++;
      sheet.getRange(`A${row}`).setValue("Aucune rémunération anormale détectée");
      row++;
    }
    
    row++;
    
    // Données incomplètes
    sheet.getRange(`A${row}`).setValue("Données incomplètes");
    sheet.getRange(`A${row}`).setFontWeight("bold");
    
    if (data.anomalies.incompletes.length > 0) {
      row++;
      sheet.getRange(`A${row}:C${row}`).setValues([
        ["Nom", "Prénom", "Champs manquants"]
      ]);
      sheet.getRange(`A${row}:C${row}`).setFontWeight("bold").setBackground("#F3F3F3");
      
      row++;
      data.anomalies.incompletes.forEach(donnee => {
        sheet.getRange(`A${row}`).setValue(donnee.nom);
        sheet.getRange(`B${row}`).setValue(donnee.prenom);
        sheet.getRange(`C${row}`).setValue(donnee.champsManquants.join(", "));
        
        row++;
      });
    } else {
      row++;
      sheet.getRange(`A${row}`).setValue("Aucune donnée incomplète détectée");
      row++;
    }
    
    // Ajouter un espace après la section
    row += 2;
    
    return row;
  }
  
  /**
   * Génère la section des tendances
   * @private
   * @param {Sheet} sheet - Feuille de calcul
   * @param {number} startRow - Ligne de début
   * @param {object} data - Données d'analyse
   * @return {number} - Nouvelle ligne courante
   */
  function _generateTrendsSection(sheet, startRow, data) {
    let row = startRow;
    
    // Vérifier si des tendances sont disponibles
    if (!data.tendances || 
        (!data.tendances.effectifs.length && 
         !data.tendances.remuneration.length && 
         !data.tendances.absences.length)) {
      return row; // Pas de tendances disponibles
    }
    
    // Titre de section
    sheet.getRange(`A${row}:E${row}`).merge();
    sheet.getRange(`A${row}`).setValue("ÉVOLUTION ET TENDANCES")
      .setBackground("#F1F3F4")
      .setFontWeight("bold");
    
    row += 2;
    
    // Évolution des effectifs
    if (data.tendances.effectifs.length > 0) {
      sheet.getRange(`A${row}`).setValue("Évolution des effectifs");
      sheet.getRange(`A${row}`).setFontWeight("bold");
      
      row++;
      sheet.getRange(`A${row}:D${row}`).setValues([
        ["Période", "Total", "Hommes", "Femmes"]
      ]);
      sheet.getRange(`A${row}:D${row}`).setFontWeight("bold").setBackground("#F3F3F3");
      
      row++;
      data.tendances.effectifs.forEach(item => {
        sheet.getRange(`A${row}`).setValue(item.periode);
        sheet.getRange(`B${row}`).setValue(item.total);
        sheet.getRange(`C${row}`).setValue(item.hommes);
        sheet.getRange(`D${row}`).setValue(item.femmes);
        
        row++;
      });
      
      row++;
    }
    
    // Évolution des rémunérations
    if (data.tendances.remuneration.length > 0) {
      sheet.getRange(`A${row}`).setValue("Évolution des rémunérations");
      sheet.getRange(`A${row}`).setFontWeight("bold");
      
      row++;
      sheet.getRange(`A${row}:C${row}`).setValues([
        ["Période", "Rémunération moyenne", "Masse salariale"]
      ]);
      sheet.getRange(`A${row}:C${row}`).setFontWeight("bold").setBackground("#F3F3F3");
      
      row++;
      data.tendances.remuneration.forEach(item => {
        sheet.getRange(`A${row}`).setValue(item.periode);
        sheet.getRange(`B${row}`).setValue(item.moyenne).setNumberFormat("#,##0.00 €");
        sheet.getRange(`C${row}`).setValue(item.total).setNumberFormat("#,##0.00 €");
        
        row++;
      });
      
      row++;
    }
    
    // Évolution des absences
    if (data.tendances.absences.length > 0) {
      sheet.getRange(`A${row}`).setValue("Évolution des absences");
      sheet.getRange(`A${row}`).setFontWeight("bold");
      
      row++;
      sheet.getRange(`A${row}:C${row}`).setValues([
        ["Période", "Jours d'absence", "Taux d'absence"]
      ]);
      sheet.getRange(`A${row}:C${row}`).setFontWeight("bold").setBackground("#F3F3F3");
      
      row++;
      data.tendances.absences.forEach(item => {
        sheet.getRange(`A${row}`).setValue(item.periode);
        sheet.getRange(`B${row}`).setValue(item.totalJours);
        sheet.getRange(`C${row}`).setValue(item.tauxAbsence / 100).setNumberFormat("0.00%");
        
        row++;
      });
    }
    
    // Ajouter un espace après la section
    row += 2;
    
    return row;
  }
  
  /**
   * Ajoute des graphiques au rapport
   * @private
   * @param {Sheet} sheet - Feuille de calcul
   * @param {object} data - Données d'analyse
   */
  function _addCharts(sheet, data) {
    // Vérifier si suffisamment de données sont disponibles pour les graphiques
    if (!data.global || !data.global.effectifs) return;
    
    try {
      // 1. Graphique de répartition hommes/femmes
      const sexChartData = sheet.getRange("A14:C15"); // Supposons que ces cellules contiennent les données H/F
      const sexChart = sheet.newChart()
        .setChartType(Charts.ChartType.PIE)
        .addRange(sexChartData)
        .setPosition(5, 7, 0, 0)
        .setOption("title", "Répartition par sexe")
        .setOption("pieSliceText", "percentage")
        .setOption("legend", { position: "right" })
        .setOption("width", 400)
        .setOption("height", 300)
        .build();
      
      sheet.insertChart(sexChart);
      
      // 2. Graphique de répartition par tranche d'âge
      // Trouver les lignes contenant les données de tranches d'âge
      const lastRow = sheet.getLastRow();
      let ageDataStartRow = 0;
      for (let i = 10; i < 30; i++) {
        if (sheet.getRange(`A${i}`).getValue() === "Répartition par tranche d'âge") {
          ageDataStartRow = i + 2; // +2 pour sauter l'en-tête
          break;
        }
      }
      
      if (ageDataStartRow > 0) {
        // Calculer le nombre de tranches d'âge
        let ageCount = 0;
        for (const tranche in data.global.effectifs.repartitionAge) {
          ageCount++;
        }
        
        if (ageCount > 0) {
          const ageChartData = sheet.getRange(ageDataStartRow, 1, ageCount, 2);
          const ageChart = sheet.newChart()
            .setChartType(Charts.ChartType.COLUMN)
            .addRange(ageChartData)
            .setPosition(5, 7, 500, 0) // Position à droite du premier graphique
            .setOption("title", "Effectifs par tranche d'âge")
            .setOption("legend", { position: "none" })
            .setOption("hAxis", { title: "Tranche d'âge" })
            .setOption("vAxis", { title: "Nombre de salariés" })
            .setOption("width", 400)
            .setOption("height", 300)
            .build();
          
          sheet.insertChart(ageChart);
        }
      }
      
      // 3. Graphique d'évolution des rémunérations si disponible
      if (data.tendances && data.tendances.remuneration.length > 0) {
        // Trouver les lignes contenant les données d'évolution des rémunérations
        let remuDataStartRow = 0;
        for (let i = 50; i < 150; i++) {
          if (i >= lastRow) break;
          if (sheet.getRange(`A${i}`).getValue() === "Évolution des rémunérations") {
            remuDataStartRow = i + 2; // +2 pour sauter l'en-tête
            break;
          }
        }
        
        if (remuDataStartRow > 0) {
          const remuCount = data.tendances.remuneration.length;
          if (remuCount > 0) {
            const remuChartData = sheet.getRange(remuDataStartRow, 1, remuCount, 2);
            const remuChart = sheet.newChart()
              .setChartType(Charts.ChartType.LINE)
              .addRange(remuChartData)
              .setPosition(20, 7, 0, 0) // Position en bas du rapport
              .setOption("title", "Évolution de la rémunération moyenne")
              .setOption("legend", { position: "none" })
              .setOption("hAxis", { title: "Période" })
              .setOption("vAxis", { title: "Rémunération moyenne (€)" })
              .setOption("width", 500)
              .setOption("height", 300)
              .build();
            
            sheet.insertChart(remuChart);
          }
        }
      }
      
      // 4. Graphique des types de contrats si disponible
      // Trouver les lignes contenant les données de types de contrats
      let contratDataStartRow = 0;
      for (let i = 30; i < 100; i++) {
        if (i >= lastRow) break;
        if (sheet.getRange(`A${i}`).getValue() === "Répartition par type de contrat") {
          contratDataStartRow = i + 2; // +2 pour sauter l'en-tête
          break;
        }
      }
      
      if (contratDataStartRow > 0) {
        // Compter le nombre de types de contrats
        let typesCount = 0;
        for (const type in data.global.contrats.typesContrat) {
          typesCount++;
        }
        
        if (typesCount > 0) {
          const contratChartData = sheet.getRange(contratDataStartRow, 1, typesCount, 2);
          const contratChart = sheet.newChart()
            .setChartType(Charts.ChartType.PIE)
            .addRange(contratChartData)
            .setPosition(20, 7, 600, 0) // Position à droite du graphique de rémunération
            .setOption("title", "Répartition par type de contrat")
            .setOption("pieSliceText", "percentage")
            .setOption("legend", { position: "right" })
            .setOption("width", 400)
            .setOption("height", 300)
            .build();
          
          sheet.insertChart(contratChart);
        }
      }
    } catch (e) {
      Logger.log("Erreur lors de la création des graphiques: " + e.toString());
    }
  }
  
  /**
   * Formate une période de dates
   * @private
   * @param {object} periode - Objet contenant les dates debut et fin
   * @return {string} - Période formatée
   */
  function _formatPeriode(periode) {
    if (!periode || !periode.debut || !periode.fin) {
      return "Période non définie";
    }
    
    const debut = Utilities.formatDate(periode.debut, Session.getScriptTimeZone(), "MMMM yyyy");
    const fin = Utilities.formatDate(periode.fin, Session.getScriptTimeZone(), "MMMM yyyy");
    
    return debut + " à " + fin;
  }
  
  /**
   * Export manuel en CSV
   * @private
   * @param {array} data - Données à exporter
   * @return {string} - Contenu CSV
   */
  function _manualExportToCsv(data) {
    let csv = "section,libelle,valeur\n";
    
    data.forEach(row => {
      const section = _escapeCsvValue(row.section || "");
      const libelle = _escapeCsvValue(row.libelle || "");
      const valeur = _escapeCsvValue(row.valeur || "");
      
      csv += `${section},${libelle},${valeur}\n`;
    });
    
    return csv;
  }
  
  /**
   * Échappe les valeurs pour le CSV
   * @private
   * @param {string} value - Valeur à échapper
   * @return {string} - Valeur échappée
   */
  function _escapeCsvValue(value) {
    if (typeof value !== "string") {
      value = String(value);
    }
    
    // Si la valeur contient une virgule, des guillemets ou un saut de ligne, l'entourer de guillemets
    if (value.includes(",") || value.includes("\"") || value.includes("\n")) {
      // Doubler les guillemets existants
      value = value.replace(/"/g, "\"\"");
      return `"${value}"`;
    }
    
    return value;
  }
  
  return {
    generateSpreadsheetReport,
    generatePDFReport,
    generateCSVReport,
    generateEqualityIndexReport,
    generateEqualityIndexCertificate,
    generateGenderPayGapReport,
    exportDashboardToSheet
  };
})();
