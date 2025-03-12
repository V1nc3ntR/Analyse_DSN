// ExportManager.gs - Module de gestion des exports

const ExportManager = (function() {
  return {
    /**
     * Exporte les données au format CSV
     * @param {array} data - Données à exporter
     * @param {object} options - Options d'export
     * @return {string} - Contenu CSV
     */
    exportToCsv: function(data, options = {}) {
      const delimiter = options.delimiter || ",";
      const includeHeader = options.includeHeader !== false;
      
      if (!data || !data.length) return "";
      
      // Extraire les en-têtes à partir des clés du premier objet
      const headers = Object.keys(data[0]);
      
      // Construire les lignes
      let csvContent = "";
      
      // Ajouter l'en-tête si demandé
      if (includeHeader) {
        csvContent += headers.join(delimiter) + "\n";
      }
      
      // Ajouter les données
      data.forEach(item => {
        const row = headers.map(header => {
          let value = item[header];
          
          // Gérer les valeurs null/undefined
          if (value === null || value === undefined) {
            return "";
          }
          
          // Formater les dates
          if (value instanceof Date) {
            value = Utilities.formatDate(value, Session.getScriptTimeZone(), "dd/MM/yyyy");
          }
          
          // Échapper les guillemets et entourer de guillemets si nécessaire
          value = String(value);
          if (value.includes(delimiter) || value.includes("\n") || value.includes("\"")) {
            value = "\"" + value.replace(/"/g, "\"\"") + "\"";
          }
          
          return value;
        });
        
        csvContent += row.join(delimiter) + "\n";
      });
      
      return csvContent;
    },
    
    /**
     * Enregistre un fichier sur Google Drive
     * @param {string|Blob} content - Contenu du fichier
     * @param {string} filename - Nom du fichier
     * @param {string} mimeType - Type MIME du fichier
     * @return {DriveFile} - Le fichier créé
     */
    saveFileToDrive: function(content, filename, mimeType) {
      // Créer un dossier s'il n'existe pas
      const folderName = "DSN Exports";
      let folder;
      
      try {
        const folderIterator = DriveApp.getFoldersByName(folderName);
        if (folderIterator.hasNext()) {
          folder = folderIterator.next();
        } else {
          folder = DriveApp.createFolder(folderName);
        }
        
        // Créer le fichier
        let file;
        if (content instanceof Blob) {
          file = folder.createFile(content);
        } else {
          file = folder.createFile(filename, content, mimeType);
        }
        
        return file;
      } catch (e) {
        Logger.log("Erreur lors de l'enregistrement du fichier : " + e.toString());
        throw e;
      }
    },
    
    /**
     * Crée un rapport PDF à partir d'une feuille de calcul
     * @param {Sheet} sheet - Feuille de calcul
     * @param {string} filename - Nom du fichier
     * @return {DriveFile} - Le fichier PDF créé
     */
    exportSheetToPdf: function(sheet, filename) {
      const spreadsheetId = sheet.getParent().getId();
      const sheetId = sheet.getSheetId();
      
      // URL pour l'export en PDF
      const url = `https://docs.google.com/spreadsheets/d/${spreadsheetId}/export?` +
                 `format=pdf&` +
                 `size=a4&` +
                 `portrait=true&` +
                 `fitw=true&` +
                 `gid=${sheetId}`;
      
      // Options d'authentification
      const token = ScriptApp.getOAuthToken();
      const options = {
        headers: {
          'Authorization': 'Bearer ' + token
        },
        muteHttpExceptions: true
      };
      
      // Récupérer le PDF
      const response = UrlFetchApp.fetch(url, options);
      
      if (response.getResponseCode() !== 200) {
        throw new Error("Erreur lors de l'export PDF: " + response.getContentText());
      }
      
      const pdfBlob = response.getBlob().setName(filename + ".pdf");
      
      // Enregistrer sur Drive
      return this.saveFileToDrive(pdfBlob, filename + ".pdf", "application/pdf");
    },
    
    /**
     * Exporte toutes les données d'analyse dans différents formats
     * @param {object} dashboard - Données du tableau de bord
     * @return {object} - Liens vers les fichiers exportés
     */
    exportFullDashboard: function(dashboard) {
      const timestamp = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), "yyyy-MM-dd_HH-mm");
      const results = {
        files: []
      };
      
      try {
        // 1. Exporter les statistiques globales en CSV
        const statsData = [];
        statsData.push({
          categorie: "Global",
          metrique: "Nombre total d'individus",
          valeur: dashboard.stats.global.totalIndividus
        });
        statsData.push({
          categorie: "Global",
          metrique: "Hommes",
          valeur: dashboard.stats.global.hommes
        });
        statsData.push({
          categorie: "Global",
          metrique: "Femmes",
          valeur: dashboard.stats.global.femmes
        });
        statsData.push({
          categorie: "Global",
          metrique: "Salaire moyen",
          valeur: dashboard.stats.global.salaireMoyen
        });
        statsData.push({
          categorie: "Global",
          metrique: "Salaire médian",
          valeur: dashboard.stats.global.salaireMedian
        });
        statsData.push({
          categorie: "Global",
          metrique: "Écart salarial H/F",
          valeur: dashboard.stats.global.ecartHommesFemmes
        });
        
        // Ajouter les tranches d'âge
        Object.keys(dashboard.stats.tranchesAge).forEach(tranche => {
          const t = dashboard.stats.tranchesAge[tranche];
          statsData.push({
            categorie: "Tranche d'âge",
            metrique: tranche + " - Total",
            valeur: t.total
          });
          statsData.push({
            categorie: "Tranche d'âge",
            metrique: tranche + " - Hommes",
            valeur: t.hommes
          });
          statsData.push({
            categorie: "Tranche d'âge",
            metrique: tranche + " - Femmes",
            valeur: t.femmes
          });
          statsData.push({
            categorie: "Tranche d'âge",
            metrique: tranche + " - Salaire moyen",
            valeur: t.salaireMoyen
          });
        });
        
        const statsCsv = this.exportToCsv(statsData);
        const statsFile = this.saveFileToDrive(
          statsCsv,
          `stats_globales_${timestamp}.csv`,
          "text/csv"
        );
        results.files.push({
          type: "CSV",
          name: statsFile.getName(),
          url: statsFile.getUrl(),
          description: "Statistiques globales"
        });
        
        // 2. Exporter les anomalies salariales en CSV
        if (dashboard.anomalies.length > 0) {
          const anomaliesCsv = this.exportToCsv(dashboard.anomalies);
          const anomaliesFile = this.saveFileToDrive(
            anomaliesCsv,
            `anomalies_salariales_${timestamp}.csv`,
            "text/csv"
          );
          results.files.push({
            type: "CSV",
            name: anomaliesFile.getName(),
            url: anomaliesFile.getUrl(),
            description: "Anomalies salariales"
          });
        }
        
        // 3. Exporter les arrêts de travail en CSV
        const arretsData = [];
        
        // Données globales
        arretsData.push({
          categorie: "Global",
          metrique: "Nombre total",
          valeur: dashboard.arrets.total
        });
        arretsData.push({
          categorie: "Global",
          metrique: "Durée moyenne",
          valeur: dashboard.arrets.dureeMoyenne.toFixed(1) + " jours"
        });
        
        // Par motif
        Object.keys(dashboard.arrets.parMotif).forEach(motif => {
          if (dashboard.arrets.parMotif[motif].count > 0) {
            arretsData.push({
              categorie: "Motif",
              metrique: dashboard.arrets.parMotif[motif].libelle,
              valeur: dashboard.arrets.parMotif[motif].count,
              pourcentage: ((dashboard.arrets.parMotif[motif].count / dashboard.arrets.total) * 100).toFixed(1) + "%",
              dureeMoyenne: dashboard.arrets.parMotif[motif].dureeMoyenne.toFixed(1) + " jours"
            });
          }
        });
        
        // Distribution
        arretsData.push({
          categorie: "Distribution",
          metrique: "Hommes",
          valeur: dashboard.arrets.distribution.hommes
        });
        arretsData.push({
          categorie: "Distribution",
          metrique: "Femmes",
          valeur: dashboard.arrets.distribution.femmes
        });
        
        Object.keys(dashboard.arrets.distribution.tranches).forEach(tranche => {
          arretsData.push({
            categorie: "Distribution",
            metrique: "Tranche " + tranche,
            valeur: dashboard.arrets.distribution.tranches[tranche]
          });
        });
        
        const arretsCsv = this.exportToCsv(arretsData);
        const arretsFile = this.saveFileToDrive(
          arretsCsv,
          `arrets_travail_${timestamp}.csv`,
          "text/csv"
        );
        results.files.push({
          type: "CSV",
          name: arretsFile.getName(),
          url: arretsFile.getUrl(),
          description: "Statistiques des arrêts de travail"
        });
        
        // 4. Exporter le rapport de conformité en CSV
        const complianceData = [
          ...dashboard.compliance.missingData.map(item => ({
            type: "Données manquantes",
            mois: item.mois,
            nir: item.nir,
            nom: item.nom,
            prenom: item.prenom,
            description: `Champ ${item.champ} (${item.code}) manquant`
          })),
          ...dashboard.compliance.warnings.map(item => ({
            type: "Avertissement",
            mois: item.mois,
            nir: item.nir,
            nom: item.nom,
            prenom: item.prenom,
            description: `${item.type}: ${item.description}`
          })),
          ...dashboard.compliance.inconsistencies.map(item => ({
            type: "Incohérence",
            mois: item.mois,
            nir: item.nir,
            nom: item.nom,
            prenom: item.prenom,
            description: `${item.type}: ${item.description}`
          }))
        ];
        
        if (complianceData.length > 0) {
          const complianceCsv = this.exportToCsv(complianceData);
          const complianceFile = this.saveFileToDrive(
            complianceCsv,
            `rapport_conformite_${timestamp}.csv`,
            "text/csv"
          );
          results.files.push({
            type: "CSV",
            name: complianceFile.getName(),
            url: complianceFile.getUrl(),
            description: "Rapport de conformité"
          });
        }
        
        // 5. Créer et exporter une feuille de calcul avec un tableau de bord
        const reportSheet = AnalysisModule.exportDashboardToSheet(dashboard);
        const pdfFile = this.exportSheetToPdf(reportSheet, `tableau_bord_dsn_${timestamp}`);
        
        results.files.push({
          type: "PDF",
          name: pdfFile.getName(),
          url: pdfFile.getUrl(),
          description: "Tableau de bord complet (PDF)"
        });
        
        return results;
      } catch (e) {
        Logger.log("Erreur lors de l'export du tableau de bord : " + e.toString());
        throw e;
      }
    }
  };
})();
