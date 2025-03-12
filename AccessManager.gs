// AccessManager.gs - Module de gestion des accès et permissions

const AccessManager = (function() {
  return {
    /**
     * Définit les permissions d'accès à la feuille
     * @param {string} userEmail - Email de l'utilisateur
     * @param {string} role - Rôle (viewer, editor, owner)
     */
    setUserAccess: function(userEmail, role) {
      const ss = SpreadsheetApp.getActiveSpreadsheet();
      
      try {
        if (role === "viewer") {
          ss.addViewer(userEmail);
        } else if (role === "editor") {
          ss.addEditor(userEmail);
        } else if (role === "owner") {
          ss.addEditor(userEmail); // Notez que Drive API est nécessaire pour le transfert de propriété
        }
      } catch (e) {
        Logger.log("Erreur lors de la définition des accès : " + e.toString());
        throw e;
      }
    },
    
    /**
     * Protège certaines feuilles contre les modifications
     * @param {array} sheetNames - Noms des feuilles à protéger
     * @param {array} editorEmails - Emails des utilisateurs autorisés à modifier
     */
    protectSheets: function(sheetNames, editorEmails) {
      const ss = SpreadsheetApp.getActiveSpreadsheet();
      
      sheetNames.forEach(name => {
        const sheet = ss.getSheetByName(name);
        
        if (sheet) {
          const protection = sheet.protect().setDescription("Protection de la feuille " + name);
          
          // Supprimer les éditeurs par défaut
          protection.removeEditors(protection.getEditors());
          
          // Si une liste d'éditeurs est fournie, les ajouter
          if (editorEmails && editorEmails.length > 0) {
            protection.addEditors(editorEmails);
          }
        }
      });
    },
    
    /**
     * Protège certaines plages de cellules contre les modifications
     * @param {string} sheetName - Nom de la feuille
     * @param {string} range - Plage de cellules (ex: "A1:C10")
     * @param {array} editorEmails - Emails des utilisateurs autorisés à modifier
     */
    protectRange: function(sheetName, range, editorEmails) {
      const ss = SpreadsheetApp.getActiveSpreadsheet();
      const sheet = ss.getSheetByName(sheetName);
      
      if (sheet) {
        const rangeToProtect = sheet.getRange(range);
        const protection = rangeToProtect.protect().setDescription("Protection de la plage " + range);
        
        // Supprimer les éditeurs par défaut
        protection.removeEditors(protection.getEditors());
        
        // Si une liste d'éditeurs est fournie, les ajouter
        if (editorEmails && editorEmails.length > 0) {
          protection.addEditors(editorEmails);
        }
      }
    },
    
    /**
     * Crée un journal des accès à la feuille
     * @param {object} access - Informations sur l'accès
     */
    logAccess: function(access) {
      const ss = SpreadsheetApp.getActiveSpreadsheet();
      let logSheet = ss.getSheetByName("Journal d'accès");
      
      if (!logSheet) {
        logSheet = ss.insertSheet("Journal d'accès");
        
        // En-têtes
        logSheet.getRange("A1:D1").setValues([["Date", "Utilisateur", "Action", "Détails"]]);
        logSheet.getRange("A1:D1").setFontWeight("bold");
      }
      
      // Ajouter l'entrée de journal
      const row = logSheet.getLastRow() + 1;
      logSheet.getRange(row, 1, 1, 4).setValues([[
        new Date(),
        access.user || Session.getActiveUser().getEmail(),
        access.action,
        access.details || ""
      ]]);
      
      // Formater la date
      logSheet.getRange(row, 1).setNumberFormat("dd/MM/yyyy HH:mm:ss");
    }
  };
})();
