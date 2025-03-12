// SchedulerModule.gs - Module de planification pour automatiser les tâches

const SchedulerModule = (function() {
  return {
    /**
     * Planifie un import DSN automatique mensuel
     * @param {object} config - Configuration de la tâche
     * @return {Trigger} - Déclencheur créé
     */
    scheduleMonthlyImport: function(config) {
      // Enregistrer la configuration dans les propriétés du script
      const scriptProperties = PropertiesService.getScriptProperties();
      scriptProperties.setProperty('autoImportConfig', JSON.stringify(config));
      
      // Créer un déclencheur horaire pour vérifier si un import est nécessaire
      const trigger = ScriptApp.newTrigger('checkAndRunAutoImport')
        .timeBased()
        .everyDays(1) // Vérification quotidienne
        .atHour(config.heure || 1) // Par défaut à 1h du matin
        .create();
      
      return trigger;
    },
    
    /**
     * Vérifie si un import automatique doit être exécuté
     */
    checkAndRunAutoImport: function() {
      const scriptProperties = PropertiesService.getScriptProperties();
      const configJson = scriptProperties.getProperty('autoImportConfig');
      
      if (!configJson) return;
      
      const config = JSON.parse(configJson);
      const today = new Date();
      
      // Si nous sommes au jour spécifié du mois (par défaut le 5)
      if (today.getDate() === (config.jour || 5)) {
        // Calculer la période pour laquelle récupérer les DSN
        const lastMonth = new Date(today.getFullYear(), today.getMonth() - 1, 1);
        const lastMonthEnd = new Date(today.getFullYear(), today.getMonth(), 0);
        
        // Préparer les données pour l'import
        const formData = {
          apiIdentifiant: config.apiIdentifiant,
          apiMotDePasse: config.apiMotDePasse,
          codeDossier: config.codeDossier,
          codeEtablissement: config.codeEtablissement,
          dateDebut: Utilities.formatDate(lastMonth, Session.getScriptTimeZone(), "yyyy-MM-dd"),
          dateFin: Utilities.formatDate(lastMonthEnd, Session.getScriptTimeZone(), "yyyy-MM-dd"),
          settings: config.settings || {}
        };
        
        // Exécuter l'import
        try {
          const result = APIHandler.importerDSNForAnalysis(formData);
          if (result && result.apiDSNData) {
            SpreadsheetHandler.storeDSNToSheets(result.apiDSNData, formData.settings);
            
            // Envoyer une notification si configuré
            if (config.notification && config.notification.email) {
              this.sendNotificationEmail(config.notification.email, result);
            }
            
            // Mettre à jour les statistiques si nécessaire
            if (config.generateStats) {
              this.generateAndSaveStats();
            }
          }
        } catch (e) {
          Logger.log("Erreur lors de l'import automatique : " + e.toString());
          
          // Envoyer une notification d'erreur
          if (config.notification && config.notification.email) {
            this.sendErrorEmail(config.notification.email, e.toString());
          }
        }
      }
    },
    
    /**
     * Envoie un email de notification de succès
     * @param {string} email - Adresse email de destination
     * @param {object} result - Résultat de l'import
     */
    sendNotificationEmail: function(email, result) {
      const nbMois = result.apiDSNData ? result.apiDSNData.length : 0;
      const subject = "Import DSN automatique - Succès";
      const body = `L'import automatique des DSN a été effectué avec succès.
      
Détails:
- Nombre de mois traités: ${nbMois}
- Date et heure: ${Utilities.formatDate(new Date(), Session.getScriptTimeZone(), "dd/MM/yyyy HH:mm")}

Ce message est généré automatiquement, merci de ne pas y répondre.`;
      
      MailApp.sendEmail(email, subject, body);
    },
    
    /**
     * Envoie un email de notification d'erreur
     * @param {string} email - Adresse email de destination
     * @param {string} errorMessage - Message d'erreur
     */
    sendErrorEmail: function(email, errorMessage) {
      const subject = "Import DSN automatique - Erreur";
      const body = `Une erreur s'est produite lors de l'import automatique des DSN.
      
Détails de l'erreur:
${errorMessage}

Date et heure: ${Utilities.formatDate(new Date(), Session.getScriptTimeZone(), "dd/MM/yyyy HH:mm")}

Merci de vérifier la configuration et de réessayer manuellement.`;
      
      MailApp.sendEmail(email, subject, body);
    },
    
    /**
     * Génère et sauvegarde les statistiques basées sur toutes les données
     */
    generateAndSaveStats: function() {
      // Récupérer toutes les DSN
      const ss = SpreadsheetApp.getActiveSpreadsheet();
      const sheets = ss.getSheets();
      const dsnSheets = sheets.filter(sheet => {
        const name = sheet.getName();
        return /^\d{2}-\d{4}$/.test(name); // Format MM-YYYY
      });
      
      if (dsnSheets.length === 0) return;
      
      // Construire un tableau de DSN pour l'analyse
      const allDSNs = [];
      
      dsnSheets.forEach(sheet => {
        const [mm, yyyy] = sheet.getName().split("-");
        const moisPrincipal = yyyy + "-" + mm;
        
        // Construire un modèle simplifié pour l'analyse
        // Note: dans un cas réel, vous devriez récupérer les données complètes
        const data = {
          moisPrincipal: moisPrincipal,
          individus: []
        };
        
        // Ici, on pourrait lire les données de la feuille et reconstruire le modèle
        // Pour simplifier, nous passons juste l'objet data
        
        allDSNs.push({
          annee: Number(yyyy),
          mois: Number(mm),
          data: data
        });
      });
      
      // Générer le tableau de bord
      const dashboard = AnalysisModule.generateDashboard(allDSNs);
      
      // Exporter les résultats
      ExportManager.exportFullDashboard(dashboard);
    }
  };
})();
