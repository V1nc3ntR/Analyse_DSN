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
 */
function parseLocalTextDSNs(allTexts, settings) {
  var localDSNData = [];
  allTexts.forEach(function(dsnText) {
    var parsed = DSNProcessor.parseDSNContent(dsnText);
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

// Fonctions UI pour les nouvelles fonctionnalités
function generateDashboard() {
  try {
    const ui = SpreadsheetApp.getUi();
    ui.alert("Génération du tableau de bord",
             "Cette opération peut prendre quelques instants, en fonction du volume de données à analyser.",
             ui.ButtonSet.OK);
    
    // Récupérer toutes les DSN disponibles
    const allDSNs = getAllDSNs();
    
    const dashboard = AnalysisModule.generateDashboard(allDSNs);
    AnalysisModule.exportDashboardToSheet(dashboard);
    
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
 * Récupère toutes les DSN disponibles dans le classeur
 */
function getAllDSNs() {
  // Cette fonction devrait extraire toutes les données DSN des feuilles
  // et les convertir au format attendu par AnalysisModule
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheets = ss.getSheets();
  const dsnSheets = sheets.filter(sheet => {
    const name = sheet.getName();
    return /^\d{2}-\d{4}$/.test(name); // Format MM-YYYY
  });
  
  // Pour l'instant, retournons un tableau vide
  // Dans une implémentation complète, cette fonction
  // extrairait les données de chaque feuille
  return [];
}

function detectSalaryAnomalies() {
  const ui = SpreadsheetApp.getUi();
  ui.alert("Fonctionnalité en cours de développement",
           "La détection des anomalies salariales sera bientôt disponible.",
           ui.ButtonSet.OK);
}

function analyzeWorkStoppages() {
  const ui = SpreadsheetApp.getUi();
  ui.alert("Fonctionnalité en cours de développement",
           "L'analyse des arrêts de travail sera bientôt disponible.",
           ui.ButtonSet.OK);
}

function checkCompliance() {
  const ui = SpreadsheetApp.getUi();
  ui.alert("Fonctionnalité en cours de développement",
           "La vérification de conformité sera bientôt disponible.",
           ui.ButtonSet.OK);
}

function createAllCharts() {
  const ui = SpreadsheetApp.getUi();
  ui.alert("Fonctionnalité en cours de développement",
           "La création de graphiques sera bientôt disponible.",
           ui.ButtonSet.OK);
}

function createVisualDashboard() {
  const ui = SpreadsheetApp.getUi();
  ui.alert("Fonctionnalité en cours de développement",
           "Le tableau de bord visuel sera bientôt disponible.",
           ui.ButtonSet.OK);
}

function exportToCSV() {
  const ui = SpreadsheetApp.getUi();
  ui.alert("Fonctionnalité en cours de développement",
           "L'export en CSV sera bientôt disponible.",
           ui.ButtonSet.OK);
}

function exportToPDF() {
  const ui = SpreadsheetApp.getUi();
  ui.alert("Fonctionnalité en cours de développement",
           "L'export en PDF sera bientôt disponible.",
           ui.ButtonSet.OK);
}

function fullExport() {
  const ui = SpreadsheetApp.getUi();
  ui.alert("Fonctionnalité en cours de développement",
           "L'export complet sera bientôt disponible.",
           ui.ButtonSet.OK);
}

function configureAutoImport() {
  const ui = SpreadsheetApp.getUi();
  ui.alert("Fonctionnalité en cours de développement",
           "La configuration d'import automatique sera bientôt disponible.",
           ui.ButtonSet.OK);
}

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
