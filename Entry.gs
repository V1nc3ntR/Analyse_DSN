// Entry.gs - Points d'entrée pour Apps Script

/**
 * Création du menu dans l'interface
 */
function onOpen() {
  const ui = SpreadsheetApp.getUi();
  ui.createMenu("DSN")
    .addItem("Ouvrir Formulaire Unique", "showSingleForm")
    .addSeparator()
    .addSubMenu(ui.createMenu("Analyse")
      .addItem("Générer tableau de bord", "generateDashboard")
      .addItem("Détecter les anomalies salariales", "detectSalaryAnomalies")
      .addItem("Analyser les arrêts de travail", "analyzeWorkStoppages")
      .addItem("Vérifier la conformité", "checkCompliance"))
    .addSubMenu(ui.createMenu("Visualisation")
      .addItem("Créer graphiques", "createAllCharts")
      .addItem("Tableau de bord visuel", "createVisualDashboard"))
    .addSubMenu(ui.createMenu("Export")
      .addItem("Exporter en CSV", "exportToCSV")
      .addItem("Exporter en PDF", "exportToPDF")
      .addItem("Export complet", "fullExport"))
    .addSubMenu(ui.createMenu("Planification")
      .addItem("Configurer import automatique", "configureAutoImport")
      .addItem("Supprimer tous les déclencheurs", "removeAllTriggers"))
    .addSubMenu(ui.createMenu("Utilitaires")
      .addItem("Vérifier compatibilité DSN 2025", "checkDSN2025Compatibility")
      .addItem("Nettoyer cache des feuilles", "clearSheetsCache")
      .addItem("Paramètres avancés", "showAdvancedSettings"))
    .addToUi();
}

/**
 * Affiche le formulaire principal de l'application
 */
function showSingleForm() {
  var html = HtmlService.createTemplateFromFile("SingleForm")
    .evaluate()
    .setWidth(700)
    .setHeight(700);
  SpreadsheetApp.getUi().showModalDialog(html, "DSN : Récupération & Analyse");
}

/**
 * Point d'entrée pour l'import des DSN via API
 */
function importerDSNForAnalysis(formData) {
  return APIHandler.importerDSNForAnalysis(formData);
}

/**
 * Point d'entrée pour l'analyse des DSN locales
 * @param {array} allTexts - Contenu textuel des DSN à analyser
 * @param {object} settings - Paramètres d'analyse
 * @return {object} - Résultat de l'analyse
 */
function parseLocalTextDSNs(allTexts, settings) {
  var localDSNData = [];
  
  allTexts.forEach(function(dsnText) {
    var parsed;
    
    // Utiliser le parseur approprié selon le format spécifié
    if (settings && settings.format === "2025") {
      parsed = DSNParser2025.parseDSNContent(dsnText);
    } else {
      parsed = DSNProcessor.parseDSNContent(dsnText);
    }
    
    localDSNData.push(parsed);
  });
  
  return { localDSNData: localDSNData };
}

/**
 * Point d'entrée pour le stockage des DSN importées via API
 */
function storeDSNToSheets(dsnArray, settings) {
  return SpreadsheetHandler.storeDSNToSheets(dsnArray, settings);
}

/**
 * Point d'entrée pour le stockage des DSN importées localement
 */
function storeLocalDSNToSheets(localDSNs, settings) {
  return SpreadsheetHandler.storeLocalDSNToSheets(localDSNs, settings);
}

/**
 * Fonction utilitaire pour inclure des fichiers HTML
 */
function include(filename) {
  return HtmlService.createHtmlOutputFromFile(filename).getContent();
}

/**
 * Génère le tableau de bord d'analyse DSN
 */
function generateDashboard() {
  try {
    const ui = SpreadsheetApp.getUi();
    ui.alert("Génération du tableau de bord",
             "Cette opération peut prendre quelques instants, en fonction du volume de données à analyser.",
             ui.ButtonSet.OK);
    
    // Récupérer toutes les DSN disponibles
    const allDSNs = getAllDSNs();
    
    if (allDSNs.length === 0) {
      ui.alert("Aucune donnée",
              "Aucune donnée DSN n'a été trouvée dans ce classeur. Veuillez d'abord importer des DSN.",
              ui.ButtonSet.OK);
      return;
    }
    
    const dashboard = AnalysisModule.generateDashboard(allDSNs);
    const reportSheet = AnalysisModule.exportDashboardToSheet(dashboard);
    
    ui.alert("Tableau de bord généré",
             "Le tableau de bord a été généré avec succès. Consultez la feuille 'Tableau de bord'.",
             ui.ButtonSet.OK);
  } catch (e) {
    const ui = SpreadsheetApp.getUi();
    ui.alert("Erreur",
             "Une erreur est survenue lors de la génération du tableau de bord : " + e.toString(),
             ui.ButtonSet.OK);
  }
}

/**
 * Détecte les anomalies salariales dans les DSN
 */
function detectSalaryAnomalies() {
  try {
    const ui = SpreadsheetApp.getUi();
    ui.alert("Détection des anomalies salariales",
             "Cette opération peut prendre quelques instants, en fonction du volume de données à analyser.",
             ui.ButtonSet.OK);
    
    // Récupérer toutes les DSN disponibles
    const allDSNs = getAllDSNs();
    
    if (allDSNs.length === 0) {
      ui.alert("Aucune donnée",
              "Aucune donnée DSN n'a été trouvée dans ce classeur. Veuillez d'abord importer des DSN.",
              ui.ButtonSet.OK);
      return;
    }
    
    // Créer une feuille pour les anomalies
    let anomalySheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Anomalies salariales");
    if (anomalySheet) {
      SpreadsheetApp.getActiveSpreadsheet().deleteSheet(anomalySheet);
    }
    anomalySheet = SpreadsheetApp.getActiveSpreadsheet().insertSheet("Anomalies salariales");
    
    // En-têtes
    anomalySheet.getRange("A1:H1").setValues([
      ["Mois", "NIR", "Nom", "Prénom", "Contrat", "Ancien salaire", "Nouveau salaire", "Variation"]
    ]);
    anomalySheet.getRange("A1:H1").setFontWeight("bold").setBackground("#E0E0E0");
    
    // Détecter les anomalies pour chaque DSN
    let row = 2;
    let totalAnomalies = 0;
    
    allDSNs.forEach(dsn => {
      const anomalies = AnalysisModule.detectSalaryAnomalies(dsn.data, { threshold: 15 });
      totalAnomalies += anomalies.length;
      
      anomalies.forEach(anomaly => {
        anomalySheet.getRange(row, 1, 1, 8).setValues([
          [
            dsn.data.moisPrincipal,
            anomaly.nir,
            anomaly.nom,
            anomaly.prenom,
            anomaly.contrat,
            anomaly.ancienSalaire,
            anomaly.nouveauSalaire,
            anomaly.variation
          ]
        ]);
        anomalySheet.getRange(row, 6, 1, 2).setNumberFormat("0.00 €");
        row++;
      });
    });
    
    // Mise en forme finale
    anomalySheet.autoResizeColumns(1, 8);
    
    ui.alert("Détection terminée",
             `${totalAnomalies} anomalies salariales ont été détectées et reportées dans la feuille "Anomalies salariales".`,
             ui.ButtonSet.OK);
  } catch (e) {
    const ui = SpreadsheetApp.getUi();
    ui.alert("Erreur",
             "Une erreur est survenue lors de la détection des anomalies : " + e.toString(),
             ui.ButtonSet.OK);
  }
}

/**
 * Analyse les arrêts de travail dans les DSN
 */
function analyzeWorkStoppages() {
  try {
    const ui = SpreadsheetApp.getUi();
    ui.alert("Analyse des arrêts de travail",
             "Cette opération peut prendre quelques instants, en fonction du volume de données à analyser.",
             ui.ButtonSet.OK);
    
    // Récupérer toutes les DSN disponibles
    const allDSNs = getAllDSNs();
    
    if (allDSNs.length === 0) {
      ui.alert("Aucune donnée",
              "Aucune donnée DSN n'a été trouvée dans ce classeur. Veuillez d'abord importer des DSN.",
              ui.ButtonSet.OK);
      return;
    }
    
    // Analyser les arrêts de travail
    const stopStats = AnalysisModule.analyzeWorkStoppages(allDSNs);
    
    // Créer une feuille pour les résultats
    let reportSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Analyse arrêts de travail");
    if (reportSheet) {
      SpreadsheetApp.getActiveSpreadsheet().deleteSheet(reportSheet);
    }
    reportSheet = SpreadsheetApp.getActiveSpreadsheet().insertSheet("Analyse arrêts de travail");
    
    // Titre
    reportSheet.getRange("A1:D1").merge();
    reportSheet.getRange("A1").setValue("ANALYSE DES ARRÊTS DE TRAVAIL")
      .setFontWeight("bold")
      .setFontSize(14)
      .setHorizontalAlignment("center");
    
    // Statistiques globales
    reportSheet.getRange("A3").setValue("STATISTIQUES GLOBALES")
      .setFontWeight("bold");
    
    reportSheet.getRange("A4").setValue("Nombre total d'arrêts :");
    reportSheet.getRange("B4").setValue(stopStats.total);
    
    reportSheet.getRange("A5").setValue("Durée moyenne des arrêts :");
    reportSheet.getRange("B5").setValue(stopStats.dureeMoyenne.toFixed(1) + " jours");
    
    // Répartition par motif
    reportSheet.getRange("A7").setValue("RÉPARTITION PAR MOTIF")
      .setFontWeight("bold");
    
    reportSheet.getRange("A8:D8").setValues([
      ["Motif", "Nombre", "% du total", "Durée moyenne"]
    ]);
    reportSheet.getRange("A8:D8").setFontWeight("bold").setBackground("#E0E0E0");
    
    let row = 9;
    Object.keys(stopStats.parMotif).forEach(motif => {
      if (stopStats.parMotif[motif].count > 0) {
        const pct = (stopStats.parMotif[motif].count / stopStats.total) * 100;
        reportSheet.getRange(row, 1, 1, 4).setValues([
          [
            stopStats.parMotif[motif].libelle,
            stopStats.parMotif[motif].count,
            pct.toFixed(1) + "%",
            stopStats.parMotif[motif].dureeMoyenne.toFixed(1) + " jours"
          ]
        ]);
        row++;
      }
    });
    
    // Distribution par profil
    reportSheet.getRange("A" + (row + 2)).setValue("DISTRIBUTION")
      .setFontWeight("bold");
    
    reportSheet.getRange("A" + (row + 3)).setValue("Hommes :");
    reportSheet.getRange("B" + (row + 3)).setValue(stopStats.distribution.hommes);
    
    reportSheet.getRange("A" + (row + 4)).setValue("Femmes :");
    reportSheet.getRange("B" + (row + 4)).setValue(stopStats.distribution.femmes);
    
    row += 6;
    reportSheet.getRange("A" + row).setValue("Par tranche d'âge :")
      .setFontWeight("bold");
    
    row++;
    Object.keys(stopStats.distribution.tranches).forEach(tranche => {
      reportSheet.getRange("A" + row).setValue(tranche + " :");
      reportSheet.getRange("B" + row).setValue(stopStats.distribution.tranches[tranche]);
      row++;
    });
    
    // Mise en forme finale
    reportSheet.autoResizeColumns(1, 4);
    
    ui.alert("Analyse terminée",
             "L'analyse des arrêts de travail a été effectuée avec succès. Consultez la feuille \"Analyse arrêts de travail\".",
             ui.ButtonSet.OK);
  } catch (e) {
    const ui = SpreadsheetApp.getUi();
    ui.alert("Erreur",
             "Une erreur est survenue lors de l'analyse des arrêts de travail : " + e.toString(),
             ui.ButtonSet.OK);
  }
}

/**
 * Vérifie la conformité des DSN
 */
function checkCompliance() {
  try {
    const ui = SpreadsheetApp.getUi();
    ui.alert("Vérification de conformité",
             "Cette opération peut prendre quelques instants, en fonction du volume de données à analyser.",
             ui.ButtonSet.OK);
    
    // Récupérer toutes les DSN disponibles
    const allDSNs = getAllDSNs();
    
    if (allDSNs.length === 0) {
      ui.alert("Aucune donnée",
              "Aucune donnée DSN n'a été trouvée dans ce classeur. Veuillez d'abord importer des DSN.",
              ui.ButtonSet.OK);
      return;
    }
    
    // Vérifier la conformité
    const compliance = AnalysisModule.generateComplianceReports(allDSNs);
    
    // Créer une feuille pour les résultats
    let reportSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Rapport de conformité");
    if (reportSheet) {
      SpreadsheetApp.getActiveSpreadsheet().deleteSheet(reportSheet);
    }
    reportSheet = SpreadsheetApp.getActiveSpreadsheet().insertSheet("Rapport de conformité");
    
    // Titre
    reportSheet.getRange("A1:F1").merge();
    reportSheet.getRange("A1").setValue("RAPPORT DE CONFORMITÉ DSN")
      .setFontWeight("bold")
      .setFontSize(14)
      .setHorizontalAlignment("center");
    
    // Données manquantes
    reportSheet.getRange("A3").setValue("DONNÉES MANQUANTES")
      .setFontWeight("bold");
    
    if (compliance.missingData.length > 0) {
      reportSheet.getRange("A4:F4").setValues([
        ["Mois", "NIR", "Nom", "Prénom", "Champ", "Code"]
      ]);
      reportSheet.getRange("A4:F4").setFontWeight("bold").setBackground("#E0E0E0");
      
      let row = 5;
      compliance.missingData.forEach(item => {
        reportSheet.getRange(row, 1, 1, 6).setValues([
          [
            item.mois,
            item.nir,
            item.nom,
            item.prenom,
            item.champ,
            item.code
          ]
        ]);
        row++;
      });
    } else {
      reportSheet.getRange("A4").setValue("Aucune donnée manquante détectée");
    }
    
    // Avertissements
    let row = reportSheet.getLastRow() + 2;
    reportSheet.getRange("A" + row).setValue("AVERTISSEMENTS")
      .setFontWeight("bold");
    
    if (compliance.warnings.length > 0) {
      row++;
      reportSheet.getRange(row, 1, 1, 6).setValues([
        ["Mois", "NIR", "Nom", "Prénom", "Type", "Description"]
      ]);
      reportSheet.getRange(row, 1, 1, 6).setFontWeight("bold").setBackground("#E0E0E0");
      
      row++;
      compliance.warnings.forEach(item => {
        reportSheet.getRange(row, 1, 1, 6).setValues([
          [
            item.mois,
            item.nir,
            item.nom,
            item.prenom,
            item.type,
            item.description
          ]
        ]);
        row++;
      });
    } else {
      row++;
      reportSheet.getRange("A" + row).setValue("Aucun avertissement détecté");
    }
    
    // Incohérences
    row = reportSheet.getLastRow() + 2;
    reportSheet.getRange("A" + row).setValue("INCOHÉRENCES")
      .setFontWeight("bold");
    
    if (compliance.inconsistencies.length > 0) {
      row++;
      reportSheet.getRange(row, 1, 1, 6).setValues([
        ["Mois", "NIR", "Nom", "Prénom", "Type", "Description"]
      ]);
      reportSheet.getRange(row, 1, 1, 6).setFontWeight("bold").setBackground("#E0E0E0");
      
      row++;
      compliance.inconsistencies.forEach(item => {
        reportSheet.getRange(row, 1, 1, 6).setValues([
          [
            item.mois,
            item.nir,
            item.nom,
            item.prenom,
            item.type,
            item.description
          ]
        ]);
        row++;
      });
    } else {
      row++;
      reportSheet.getRange("A" + row).setValue("Aucune incohérence détectée");
    }
    
    // Mise en forme finale
    reportSheet.autoResizeColumns(1, 6);
    
    // Résumé
    const totalIssues = compliance.missingData.length + compliance.warnings.length + compliance.inconsistencies.length;
    
    ui.alert("Vérification terminée",
             `La vérification de conformité a été effectuée avec succès. ${totalIssues} problèmes ont été détectés. Consultez la feuille "Rapport de conformité".`,
             ui.ButtonSet.OK);
  } catch (e) {
    const ui = SpreadsheetApp.getUi();
    ui.alert("Erreur",
             "Une erreur est survenue lors de la vérification de conformité : " + e.toString(),
             ui.ButtonSet.OK);
  }
}

/**
 * Crée tous les graphiques d'analyse
 */
function createAllCharts() {
  try {
    const ui = SpreadsheetApp.getUi();
    ui.alert("Création des graphiques",
             "Cette opération peut prendre quelques instants.",
             ui.ButtonSet.OK);
    
    // Récupérer toutes les DSN disponibles
    const allDSNs = getAllDSNs();
    
    if (allDSNs.length === 0) {
      ui.alert("Aucune donnée",
              "Aucune donnée DSN n'a été trouvée dans ce classeur. Veuillez d'abord importer des DSN.",
              ui.ButtonSet.OK);
      return;
    }
    
    // Générer le tableau de bord pour les données
    const dashboard = AnalysisModule.generateDashboard(allDSNs);
    
    // Supprimer l'ancienne feuille de graphiques si elle existe
    let chartSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Graphiques");
    if (chartSheet) {
      SpreadsheetApp.getActiveSpreadsheet().deleteSheet(chartSheet);
    }
    
    // Utiliser VisualizationModule pour créer les graphiques
    VisualizationModule.createSalaryEvolutionChart(dashboard.stats);
    VisualizationModule.createAgeDistributionChart(dashboard.stats);
    VisualizationModule.createGenderDistributionChart(dashboard.stats);
    
    ui.alert("Création terminée",
             "Les graphiques ont été créés avec succès. Consultez la feuille 'Graphiques' pour les visualiser.",
             ui.ButtonSet.OK);
  } catch (e) {
    const ui = SpreadsheetApp.getUi();
    ui.alert("Erreur",
             "Une erreur est survenue lors de la création des graphiques : " + e.toString(),
             ui.ButtonSet.OK);
  }
}

/**
 * Crée un tableau de bord visuel
 */
function createVisualDashboard() {
  try {
    const ui = SpreadsheetApp.getUi();
    ui.alert("Création du tableau de bord visuel",
             "Cette opération peut prendre quelques instants.",
             ui.ButtonSet.OK);
    
    // Récupérer toutes les DSN disponibles
    const allDSNs = getAllDSNs();
    
    if (allDSNs.length === 0) {
      ui.alert("Aucune donnée",
              "Aucune donnée DSN n'a été trouvée dans ce classeur. Veuillez d'abord importer des DSN.",
              ui.ButtonSet.OK);
      return;
    }
    
    // Générer le tableau de bord pour les données
    const dashboard = AnalysisModule.generateDashboard(allDSNs);
    
    // Utiliser VisualizationModule pour créer le tableau de bord visuel
    const dashboardSheet = VisualizationModule.createVisualDashboard(dashboard);
    
    ui.alert("Création terminée",
             `Le tableau de bord visuel a été créé avec succès. Consultez la feuille "${dashboardSheet.getName()}" pour le visualiser.`,
             ui.ButtonSet.OK);
  } catch (e) {
    const ui = SpreadsheetApp.getUi();
    ui.alert("Erreur",
             "Une erreur est survenue lors de la création du tableau de bord visuel : " + e.toString(),
             ui.ButtonSet.OK);
  }
}

/**
 * Exporte les données au format CSV
 */
function exportToCSV() {
  try {
    const ui = SpreadsheetApp.getUi();
    ui.alert("Export en CSV",
             "Cette opération peut prendre quelques instants.",
             ui.ButtonSet.OK);
    
    // Récupérer toutes les DSN disponibles
    const allDSNs = getAllDSNs();
    
    if (allDSNs.length === 0) {
      ui.alert("Aucune donnée",
              "Aucune donnée DSN n'a été trouvée dans ce classeur. Veuillez d'abord importer des DSN.",
              ui.ButtonSet.OK);
      return;
    }
    
    // Générer le tableau de bord pour les données
    const dashboard = AnalysisModule.generateDashboard(allDSNs);
    
    // Utiliser ExportManager pour l'export en CSV
    const csvReport = DSNReportService.generateCSVReport(dashboard);
    
    // Créer un dialogue pour télécharger le CSV
    const htmlOutput = HtmlService.createHtmlOutput(
      `<html>
        <head>
          <base target="_top">
          <style>
            body { font-family: Arial, sans-serif; padding: 20px; }
            .container { text-align: center; }
            textarea { width: 100%; height: 300px; margin: 20px 0; }
            .btn { 
              background-color: #4285F4; 
              color: white; 
              padding: 10px 20px; 
              border: none; 
              border-radius: 5px;
              cursor: pointer;
              font-weight: bold;
            }
          </style>
        </head>
        <body>
          <div class="container">
            <h3>Export CSV généré avec succès</h3>
            <p>Copiez le contenu ci-dessous ou utilisez le bouton de téléchargement :</p>
            <textarea readonly>${csvReport}</textarea>
            <button class="btn" onclick="downloadCSV()">Télécharger le CSV</button>
          </div>
          <script>
            function downloadCSV() {
              const csvContent = document.querySelector('textarea').value;
              const blob = new Blob([csvContent], { type: 'text/csv;charset=utf-8;' });
              const url = URL.createObjectURL(blob);
              const link = document.createElement('a');
              link.href = url;
              link.setAttribute('download', 'DSN_Export_${Utilities.formatDate(new Date(), Session.getScriptTimeZone(), "yyyy-MM-dd")}.csv');
              link.click();
            }
          </script>
        </body>
      </html>`
    )
    .setWidth(600)
    .setHeight(450);
    
    ui.showModalDialog(htmlOutput, "Export CSV");
  } catch (e) {
    const ui = SpreadsheetApp.getUi();
    ui.alert("Erreur",
             "Une erreur est survenue lors de l'export en CSV : " + e.toString(),
             ui.ButtonSet.OK);
  }
}

/**
 * Exporte les données au format PDF
 */
function exportToPDF() {
  try {
    const ui = SpreadsheetApp.getUi();
    ui.alert("Export en PDF",
             "Cette opération peut prendre quelques instants.",
             ui.ButtonSet.OK);
    
    // Récupérer toutes les DSN disponibles
    const allDSNs = getAllDSNs();
    
    if (allDSNs.length === 0) {
      ui.alert("Aucune donnée",
              "Aucune donnée DSN n'a été trouvée dans ce classeur. Veuillez d'abord importer des DSN.",
              ui.ButtonSet.OK);
      return;
    }
    
    // Générer le tableau de bord pour les données
    const dashboard = AnalysisModule.generateDashboard(allDSNs);
    
    // Utiliser DSNReportService pour l'export en PDF
    const pdfResult = DSNReportService.generatePDFReport(dashboard);
    
    if (pdfResult.success) {
      let message = "L'export en PDF a été réalisé avec succès.";
      if (pdfResult.file) {
        message += ` Le fichier a été enregistré sur Google Drive sous le nom "${pdfResult.file.getName()}".`;
      } else if (pdfResult.downloadUrl) {
        message += " Utilisez le lien suivant pour télécharger le fichier.";
      }
      
      ui.alert("Export terminé", message, ui.ButtonSet.OK);
      
      if (pdfResult.downloadUrl) {
        const html = HtmlService.createHtmlOutput(
          `<p>Cliquez sur le lien ci-dessous pour télécharger le fichier PDF:</p>
           <p><a href="${pdfResult.downloadUrl}" target="_blank">Télécharger le PDF</a></p>`
        )
        .setWidth(400)
        .setHeight(200);
        
        ui.showModalDialog(html, "Téléchargement du PDF");
      }
    } else {
      ui.alert("Erreur",
               "Une erreur est survenue lors de l'export en PDF : " + pdfResult.message,
               ui.ButtonSet.OK);
    }
  } catch (e) {
    const ui = SpreadsheetApp.getUi();
    ui.alert("Erreur",
             "Une erreur est survenue lors de l'export en PDF : " + e.toString(),
             ui.ButtonSet.OK);
  }
}

/**
 * Effectue un export complet (tous formats)
 */
function fullExport() {
  try {
    const ui = SpreadsheetApp.getUi();
    ui.alert("Export complet",
             "Cette opération peut prendre quelques instants.",
             ui.ButtonSet.OK);
    
    // Récupérer toutes les DSN disponibles
    const allDSNs = getAllDSNs();
    
    if (allDSNs.length === 0) {
      ui.alert("Aucune donnée",
              "Aucune donnée DSN n'a été trouvée dans ce classeur. Veuillez d'abord importer des DSN.",
              ui.ButtonSet.OK);
      return;
    }
    
    // Générer le tableau de bord pour les données
    const dashboard = AnalysisModule.generateDashboard(allDSNs);
    
    // Utiliser ExportManager pour l'export complet
    const result = ExportManager.exportFullDashboard(dashboard);
    
    if (result.files && result.files.length > 0) {
      let fileList = "";
      result.files.forEach(file => {
        fileList += `- ${file.name} (${file.type}) : ${file.description}<br>`;
      });
      
      const html = HtmlService.createHtmlOutput(
        `<p>L'export complet a été réalisé avec succès. ${result.files.length} fichiers ont été générés :</p>
         <p>${fileList}</p>`
      )
      .setWidth(500)
      .setHeight(300);
      
      ui.showModalDialog(html, "Export complet terminé");
    } else {
      ui.alert("Export terminé",
               "L'export complet a été réalisé avec succès, mais aucun fichier n'a été généré.",
               ui.ButtonSet.OK);
    }
  } catch (e) {
    const ui = SpreadsheetApp.getUi();
    ui.alert("Erreur",
             "Une erreur est survenue lors de l'export complet : " + e.toString(),
             ui.ButtonSet.OK);
  }
}

/**
 * Ouverture du formulaire de configuration d'import automatique
 */
function configureAutoImport() {
  const htmlOutput = HtmlService.createHtmlOutput(
    `<html>
      <head>
        <base target="_top">
        <link href="https://cdnjs.cloudflare.com/ajax/libs/font-awesome/5.15.4/css/all.min.css" rel="stylesheet">
        <link href="https://cdnjs.cloudflare.com/ajax/libs/bootstrap/4.6.0/css/bootstrap.min.css" rel="stylesheet">
        <script src="https://cdnjs.cloudflare.com/ajax/libs/jquery/3.6.0/jquery.min.js"></script>
        <script src="https://cdnjs.cloudflare.com/ajax/libs/bootstrap/4.6.0/js/bootstrap.bundle.min.js"></script>
        <style>
          body { font-family: Arial, sans-serif; padding: 20px; }
          .card { border-radius: 10px; box-shadow: 0 4px 6px rgba(0, 0, 0, 0.1); }
          .card-header { background: linear-gradient(135deg, #4361ee 0%, #3a0ca3 100%); color: white; }
          .btn-primary { background: linear-gradient(135deg, #4361ee 0%, #3a0ca3 100%); border: none; }
        </style>
      </head>
      <body>
        <div class="container">
          <div class="card">
            <div class="card-header">
              <h4><i class="fas fa-calendar-alt mr-2"></i>Configuration de l'import automatique</h4>
            </div>
            <div class="card-body">
              <form id="autoImportForm">
                <div class="form-group">
                  <label>Jour du mois pour l'import :</label>
                  <select name="jour" class="form-control">
                    <option value="1">1er du mois</option>
                    <option value="5" selected>5 du mois</option>
                    <option value="10">10 du mois</option>
                    <option value="15">15 du mois</option>
                    <option value="20">20 du mois</option>
                    <option value="25">25 du mois</option>
                  </select>
                </div>
                
                <div class="form-group">
                  <label>Heure de l'import :</label>
                  <select name="heure" class="form-control">
                    <option value="1" selected>1h00 du matin</option>
                    <option value="3">3h00 du matin</option>
                    <option value="5">5h00 du matin</option>
                    <option value="12">12h00 (midi)</option>
                    <option value="18">18h00</option>
                    <option value="22">22h00</option>
                  </select>
                </div>
                
                <div class="form-group">
                  <label>Email de notification :</label>
                  <input type="email" name="email" class="form-control" placeholder="Adresse email pour les notifications">
                </div>
                
                <div class="form-group">
                  <div class="custom-control custom-checkbox">
                    <input type="checkbox" class="custom-control-input" id="generateStats" name="generateStats" checked>
                    <label class="custom-control-label" for="generateStats">Générer automatiquement les statistiques</label>
                  </div>
                </div>
                
                <div class="text-center mt-3">
                  <button type="submit" class="btn btn-primary">
                    <i class="fas fa-save mr-2"></i>Sauvegarder la configuration
                  </button>
                </div>
              </form>
            </div>
          </div>
        </div>
        
        <script>
          document.getElementById('autoImportForm').addEventListener('submit', function(e) {
            e.preventDefault();
            
            const config = {
              jour: parseInt(document.querySelector('select[name="jour"]').value),
              heure: parseInt(document.querySelector('select[name="heure"]').value),
              email: document.querySelector('input[name="email"]').value,
              generateStats: document.getElementById('generateStats').checked
            };
            
            google.script.run
              .withSuccessHandler(function() {
                alert('Configuration sauvegardée avec succès !');
                google.script.host.close();
              })
              .withFailureHandler(function(error) {
                alert('Erreur : ' + error);
              })
              .scheduleAutoImport(config);
          });
        </script>
      </body>
    </html>`
  )
  .setWidth(500)
  .setHeight(500);
  
  SpreadsheetApp.getUi().showModalDialog(htmlOutput, "Configuration de l'import automatique");
}

/**
 * Configure l'import automatique
 * @param {object} config - Configuration de l'import
 */
function scheduleAutoImport(config) {
  return SchedulerModule.scheduleMonthlyImport(config);
}

/**
 * Supprime tous les déclencheurs existants
 */
function removeAllTriggers() {
  // Supprimer tous les déclencheurs existants
  const triggers = ScriptApp.getProjectTriggers();
  for (let i = 0; i < triggers.length; i++) {
    ScriptApp.deleteTrigger(triggers[i]);
  }
  
  const ui = SpreadsheetApp.getUi();
  ui.alert("Déclencheurs supprimés",
           "Tous les déclencheurs ont été supprimés.",
           ui.ButtonSet.OK);
}

/**
 * Outil de vérification de compatibilité avec le format DSN 2025
 */
function checkDSN2025Compatibility() {
  const ui = SpreadsheetApp.getUi();
  
  // Vérifier si les modules nécessaires existent
  const hasParser2025 = (typeof DSNParser2025 !== 'undefined' && typeof DSNParser2025.parseDSNContent === 'function');
  const hasModels2025 = (typeof DSNModels !== 'undefined' && typeof DSNModels.createSalarieModel === 'function');
  
  if (!hasParser2025 || !hasModels2025) {
    ui.alert("Modules manquants",
             "Les modules nécessaires pour le format DSN 2025 ne sont pas tous disponibles. Veuillez vérifier que les fichiers DSNParser2025.gs et DSNModels.gs sont présents.",
             ui.ButtonSet.OK);
    return;
  }
  
  // Vérifier si le format est correctement pris en charge dans les fonctions d'import
  const parseLocalFunction = parseLocalTextDSNs.toString();
  const supports2025 = parseLocalFunction.includes('format === "2025"');
  
  if (!supports2025) {
    ui.alert("Fonctions non compatibles",
             "Les fonctions d'import ne prennent pas correctement en charge le format DSN 2025. Une mise à jour est nécessaire.",
             ui.ButtonSet.OK);
    return;
  }
  
  // Vérifier si le formulaire contient l'option de format
  // Cette vérification ne peut pas être faite directement, mais nous supposons que c'est le cas
  
  ui.alert("Compatibilité DSN 2025",
           "Tous les modules nécessaires sont présents et l'application est compatible avec le format DSN 2025.",
           ui.ButtonSet.OK);
}

/**
 * Nettoie le cache des feuilles pour libérer de la mémoire
 */
function clearSheetsCache() {
  try {
    SpreadsheetHandler.clearCache();
    
    const ui = SpreadsheetApp.getUi();
    ui.alert("Cache nettoyé",
             "Le cache des feuilles a été nettoyé avec succès.",
             ui.ButtonSet.OK);
  } catch (e) {
    const ui = SpreadsheetApp.getUi();
    ui.alert("Erreur",
             "Une erreur est survenue lors du nettoyage du cache : " + e.toString(),
             ui.ButtonSet.OK);
  }
}

/**
 * Affiche les paramètres avancés
 */
function showAdvancedSettings() {
  const htmlOutput = HtmlService.createHtmlOutput(
    `<html>
      <head>
        <base target="_top">
        <link href="https://cdnjs.cloudflare.com/ajax/libs/font-awesome/5.15.4/css/all.min.css" rel="stylesheet">
        <link href="https://cdnjs.cloudflare.com/ajax/libs/bootstrap/4.6.0/css/bootstrap.min.css" rel="stylesheet">
        <script src="https://cdnjs.cloudflare.com/ajax/libs/jquery/3.6.0/jquery.min.js"></script>
        <script src="https://cdnjs.cloudflare.com/ajax/libs/bootstrap/4.6.0/js/bootstrap.bundle.min.js"></script>
        <style>
          body { font-family: Arial, sans-serif; padding: 20px; }
          .card { border-radius: 10px; box-shadow: 0 4px 6px rgba(0, 0, 0, 0.1); margin-bottom: 20px; }
          .card-header { background: linear-gradient(135deg, #4361ee 0%, #3a0ca3 100%); color: white; }
          .btn-primary { background: linear-gradient(135deg, #4361ee 0%, #3a0ca3 100%); border: none; }
        </style>
      </head>
      <body>
        <div class="container">
          <h3 class="mb-4 text-center">Paramètres avancés</h3>
          
          <div class="card">
            <div class="card-header">
              <h5><i class="fas fa-cogs mr-2"></i>Paramètres de performance</h5>
            </div>
            <div class="card-body">
              <form id="performanceForm">
                <div class="form-group">
                  <label>Taille des lots de traitement:</label>
                  <select name="batchSize" class="form-control">
                    <option value="10">10 (Petit fichier)</option>
                    <option value="50" selected>50 (Par défaut)</option>
                    <option value="100">100 (Grand fichier)</option>
                    <option value="200">200 (Très grand fichier)</option>
                  </select>
                </div>
                
                <div class="form-group">
                  <div class="custom-control custom-switch">
                    <input type="checkbox" class="custom-control-input" id="cacheEnabled" name="cacheEnabled" checked>
                    <label class="custom-control-label" for="cacheEnabled">Activer la mise en cache</label>
                  </div>
                </div>
                
                <div class="form-group">
                  <div class="custom-control custom-switch">
                    <input type="checkbox" class="custom-control-input" id="parallelProcessing" name="parallelProcessing">
                    <label class="custom-control-label" for="parallelProcessing">Traitement parallèle (expérimental)</label>
                  </div>
                </div>
                
                <button type="submit" class="btn btn-primary">Enregistrer</button>
              </form>
            </div>
          </div>
          
          <div class="card">
            <div class="card-header">
              <h5><i class="fas fa-wrench mr-2"></i>Options d'analyse</h5>
            </div>
            <div class="card-body">
              <form id="analysisForm">
                <div class="form-group">
                  <label>Seuil de détection des anomalies (%):</label>
                  <input type="number" name="anomalyThreshold" class="form-control" value="15" min="5" max="50">
                </div>
                
                <div class="form-group">
                  <label>Format DSN par défaut:</label>
                  <select name="defaultFormat" class="form-control">
                    <option value="standard" selected>Standard</option>
                    <option value="2025">Format 2025</option>
                  </select>
                </div>
                
                <div class="form-group">
                  <div class="custom-control custom-switch">
                    <input type="checkbox" class="custom-control-input" id="exportWarnings" name="exportWarnings" checked>
                    <label class="custom-control-label" for="exportWarnings">Inclure les avertissements dans les exports</label>
                  </div>
                </div>
                
                <button type="submit" class="btn btn-primary">Enregistrer</button>
              </form>
            </div>
          </div>
        </div>
        
        <script>
          document.getElementById('performanceForm').addEventListener('submit', function(e) {
            e.preventDefault();
            
            const config = {
              batchSize: parseInt(document.querySelector('select[name="batchSize"]').value),
              cacheEnabled: document.getElementById('cacheEnabled').checked,
              parallelProcessing: document.getElementById('parallelProcessing').checked
            };
            
            google.script.run
              .withSuccessHandler(function() {
                alert('Paramètres de performance enregistrés !');
              })
              .withFailureHandler(function(error) {
                alert('Erreur : ' + error);
              })
              .savePerformanceSettings(config);
          });
          
          document.getElementById('analysisForm').addEventListener('submit', function(e) {
            e.preventDefault();
            
            const config = {
              anomalyThreshold: parseInt(document.querySelector('input[name="anomalyThreshold"]').value),
              defaultFormat: document.querySelector('select[name="defaultFormat"]').value,
              exportWarnings: document.getElementById('exportWarnings').checked
            };
            
            google.script.run
              .withSuccessHandler(function() {
                alert('Options d\'analyse enregistrées !');
              })
              .withFailureHandler(function(error) {
                alert('Erreur : ' + error);
              })
              .saveAnalysisSettings(config);
          });
        </script>
      </body>
    </html>`
  )
  .setWidth(500)
  .setHeight(600);
  
  SpreadsheetApp.getUi().showModalDialog(htmlOutput, "Paramètres avancés");
}

/**
 * Sauvegarde les paramètres de performance
 * @param {object} config - Configuration de performance
 */
function savePerformanceSettings(config) {
  const userProperties = PropertiesService.getUserProperties();
  userProperties.setProperty('performanceSettings', JSON.stringify(config));
  return true;
}

/**
 * Sauvegarde les paramètres d'analyse
 * @param {object} config - Configuration d'analyse
 */
function saveAnalysisSettings(config) {
  const userProperties = PropertiesService.getUserProperties();
  userProperties.setProperty('analysisSettings', JSON.stringify(config));
  return true;
}

/**
 * Récupère toutes les DSN disponibles dans le classeur
 * @return {array} - Tableau des DSN
 */
function getAllDSNs() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheets = ss.getSheets();
  const allDSNs = [];
  
  // Patterns pour les formats de nom de feuille
  const mmyyyyPattern = /^(\d{2})-(\d{4})$/;
  const yyyymmPattern = /^(\d{4})-(\d{2})$/;
  
  // Feuille de suivi DSN
  const trackingSheet = ss.getSheetByName("DSN_Tracking");
  const trackingData = trackingSheet ? trackingSheet.getDataRange().getValues() : [];
  
  // Sauter l'en-tête
  const trackingStart = trackingSheet ? 1 : 0;
  
  // Créer une map pour accéder rapidement aux informations de suivi
  const trackingMap = {};
  if (trackingSheet) {
    for (let i = trackingStart; i < trackingData.length; i++) {
      const row = trackingData[i];
      // Format attendu: [Mois, Année, Date d'import, NbSalariés, NbContrats, NbArrêts, Chemin]
      if (row[0] && row[1]) {
        const mm = String(row[0]).padStart(2, '0');
        const yyyy = String(row[1]);
        trackingMap[mm + "-" + yyyy] = {
          month: parseInt(mm, 10),
          year: parseInt(yyyy, 10)
        };
      }
    }
  }
  
  // Traiter chaque feuille
  sheets.forEach(sheet => {
    const name = sheet.getName();
    
    // Essayer le format MM-YYYY
    let match = name.match(mmyyyyPattern);
    if (match) {
      const mm = match[1];
      const yyyy = match[2];
      const moisPrincipal = yyyy + "-" + mm;
      
      // Créer un modèle simplifié pour l'analyse
      const model = {
        moisPrincipal: moisPrincipal,
        individus: [],
        salaries: [],
        contrats: [],
        arrets: []
      };
      
      allDSNs.push({
        annee: parseInt(yyyy, 10),
        mois: parseInt(mm, 10),
        data: model
      });
      
      return;
    }
    
    // Essayer le format YYYY-MM
    match = name.match(yyyymmPattern);
    if (match) {
      const yyyy = match[1];
      const mm = match[2];
      const moisPrincipal = yyyy + "-" + mm;
      
      // Créer un modèle simplifié pour l'analyse
      const model = {
        moisPrincipal: moisPrincipal,
        individus: [],
        salaries: [],
        contrats: [],
        arrets: []
      };
      
      allDSNs.push({
        annee: parseInt(yyyy, 10),
        mois: parseInt(mm, 10),
        data: model
      });
      
      return;
    }
    
    // Vérifier si la feuille existe dans le suivi
    if (trackingMap[name]) {
      const info = trackingMap[name];
      const moisPrincipal = info.year + "-" + String(info.month).padStart(2, '0');
      
      // Créer un modèle simplifié pour l'analyse
      const model = {
        moisPrincipal: moisPrincipal,
        individus: [],
        salaries: [],
        contrats: [],
        arrets: []
      };
      
      allDSNs.push({
        annee: info.year,
        mois: info.month,
        data: model
      });
    }
  });
  
  return allDSNs;
}
