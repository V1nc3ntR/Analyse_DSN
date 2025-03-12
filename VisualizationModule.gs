// VisualizationModule.gs - Module de visualisation des données

const VisualizationModule = (function() {
  return {
    /**
     * Génère un graphique pour l'évolution des salaires
     * @param {object} stats - Statistiques générées par AnalysisModule
     * @return {Chart} - Objet graphique
     */
    createSalaryEvolutionChart: function(stats) {
      const ss = SpreadsheetApp.getActiveSpreadsheet();
      let chartSheet = ss.getSheetByName("Graphiques");
      
      if (!chartSheet) {
        chartSheet = ss.insertSheet("Graphiques");
      }
      
      // Préparer les données
      const dataRange = chartSheet.getRange(1, 1, stats.evolution.salaireMoyen.length + 1, 2);
      
      // En-têtes
      const header = [["Mois", "Salaire moyen"]];
      
      // Données
      const data = stats.evolution.salaireMoyen.map(item => [item.mois, item.valeur]);
      
      // Écrire les données
      dataRange.setValues([...header, ...data]);
      
      // Créer le graphique
      const chart = chartSheet.newChart()
        .setChartType(Charts.ChartType.LINE)
        .addRange(dataRange)
        .setPosition(5, 1, 0, 0)
        .setOption('title', 'Évolution du salaire moyen')
        .setOption('legend', {position: 'bottom'})
        .setOption('hAxis', {title: 'Mois'})
        .setOption('vAxis', {title: 'Salaire (€)'})
        .build();
      
      chartSheet.insertChart(chart);
      
      return chart;
    },
    
    /**
     * Génère un graphique en secteurs pour la répartition par tranche d'âge
     * @param {object} stats - Statistiques générées par AnalysisModule
     * @return {Chart} - Objet graphique
     */
    createAgeDistributionChart: function(stats) {
      const ss = SpreadsheetApp.getActiveSpreadsheet();
      let chartSheet = ss.getSheetByName("Graphiques");
      
      if (!chartSheet) {
        chartSheet = ss.insertSheet("Graphiques");
      }
      
      // Préparer les données
      const tranches = Object.keys(stats.tranchesAge);
      const dataRange = chartSheet.getRange(1, 4, tranches.length + 1, 2);
      
      // En-têtes
      const header = [["Tranche d'âge", "Nombre"]];
      
      // Données
      const data = tranches.map(tranche => [tranche, stats.tranchesAge[tranche].total]);
      
      // Écrire les données
      dataRange.setValues([...header, ...data]);
      
      // Créer le graphique
      const chart = chartSheet.newChart()
        .setChartType(Charts.ChartType.PIE)
        .addRange(dataRange)
        .setPosition(5, 6, 0, 0)
        .setOption('title', 'Répartition par tranche d\'âge')
        .setOption('pieSliceText', 'percentage')
        .setOption('legend', {position: 'right'})
        .build();
      
      chartSheet.insertChart(chart);
      
      return chart;
    },
    
    /**
     * Génère un graphique en colonnes pour la répartition hommes/femmes par tranche d'âge
     * @param {object} stats - Statistiques générées par AnalysisModule
     * @return {Chart} - Objet graphique
     */
    createGenderDistributionChart: function(stats) {
      const ss = SpreadsheetApp.getActiveSpreadsheet();
      let chartSheet = ss.getSheetByName("Graphiques");
      
      if (!chartSheet) {
        chartSheet = ss.insertSheet("Graphiques");
      }
      
      // Préparer les données
      const tranches = Object.keys(stats.tranchesAge);
      const dataRange = chartSheet.getRange(1, 7, tranches.length + 1, 3);
      
      // En-têtes
      const header = [["Tranche d'âge", "Hommes", "Femmes"]];
      
      // Données
      const data = tranches.map(tranche => [
        tranche, 
        stats.tranchesAge[tranche].hommes,
        stats.tranchesAge[tranche].femmes
      ]);
      
      // Écrire les données
      dataRange.setValues([...header, ...data]);
      
      // Créer le graphique
      const chart = chartSheet.newChart()
        .setChartType(Charts.ChartType.COLUMN)
        .addRange(dataRange)
        .setPosition(25, 1, 0, 0)
        .setOption('title', 'Répartition hommes/femmes par tranche d\'âge')
        .setOption('isStacked', 'percent')
        .setOption('legend', {position: 'top'})
        .setOption('hAxis', {title: 'Tranche d\'âge'})
        .setOption('vAxis', {title: 'Pourcentage'})
        .build();
      
      chartSheet.insertChart(chart);
      
      return chart;
    },
    
    /**
     * Génère un tableau de bord visuel complet avec tous les graphiques
     * @param {object} dashboard - Tableau de bord généré par AnalysisModule
     * @return {Sheet} - Feuille de tableau de bord
     */
    createVisualDashboard: function(dashboard) {
      const ss = SpreadsheetApp.getActiveSpreadsheet();
      let dashboardSheet = ss.getSheetByName("Tableau de bord visuel");
      
      if (dashboardSheet) {
        ss.deleteSheet(dashboardSheet);
      }
      
      dashboardSheet = ss.insertSheet("Tableau de bord visuel");
      
      // Créer les graphiques
      this.createSalaryEvolutionChart(dashboard.stats);
      this.createAgeDistributionChart(dashboard.stats);
      this.createGenderDistributionChart(dashboard.stats);
      
      // Insérer les KPIs principaux
      dashboardSheet.getRange("A1:C1").merge();
      dashboardSheet.getRange("A1").setValue("TABLEAU DE BORD DSN - INDICATEURS CLÉS")
        .setFontWeight("bold")
        .setFontSize(16)
        .setHorizontalAlignment("center");
      
      // KPIs de la première ligne
      dashboardSheet.getRange("A3").setValue("Effectif total");
      dashboardSheet.getRange("A4").setValue(dashboard.stats.global.totalIndividus)
        .setFontWeight("bold")
        .setFontSize(24)
        .setHorizontalAlignment("center");
      
      dashboardSheet.getRange("B3").setValue("Salaire moyen");
      dashboardSheet.getRange("B4").setValue(dashboard.stats.global.salaireMoyen)
        .setFontWeight("bold")
        .setFontSize(24)
        .setHorizontalAlignment("center")
        .setNumberFormat("0.00 €");
      
      dashboardSheet.getRange("C3").setValue("Écart H/F");
      dashboardSheet.getRange("C4").setValue(dashboard.stats.global.ecartHommesFemmes / 100)
        .setFontWeight("bold")
        .setFontSize(24)
        .setHorizontalAlignment("center")
        .setNumberFormat("0.0%");
      
      // KPIs de la seconde ligne
      dashboardSheet.getRange("A6").setValue("Hommes");
      dashboardSheet.getRange("A7").setValue(dashboard.stats.global.hommes)
        .setFontWeight("bold")
        .setFontSize(18)
        .setHorizontalAlignment("center");
      
      dashboardSheet.getRange("B6").setValue("Femmes");
      dashboardSheet.getRange("B7").setValue(dashboard.stats.global.femmes)
        .setFontWeight("bold")
        .setFontSize(18)
        .setHorizontalAlignment("center");
      
      dashboardSheet.getRange("C6").setValue("Arrêts de travail");
      dashboardSheet.getRange("C7").setValue(dashboard.arrets.total)
        .setFontWeight("bold")
        .setFontSize(18)
        .setHorizontalAlignment("center");
      
      // Ajouter les graphiques depuis la feuille Graphiques
      const chartsSheet = ss.getSheetByName("Graphiques");
      const charts = chartsSheet.getCharts();
      
      let verticalPosition = 9;
      charts.forEach(chart => {
        const updatedChart = chart.modify()
          .setPosition(verticalPosition, 1, 0, 0)
          .build();
        
        dashboardSheet.insertChart(updatedChart);
        verticalPosition += 20;
      });
      
      // Formater la feuille
      dashboardSheet.getRange("A1:C7").setBorder(true, true, true, true, true, true);
      dashboardSheet.getRange("A3:C3").setBackground("#f0f0f0");
      dashboardSheet.getRange("A6:C6").setBackground("#f0f0f0");
      
      // Adapter la largeur des colonnes
      dashboardSheet.autoResizeColumns(1, 3);
      
      return dashboardSheet;
    }
  };
})();
