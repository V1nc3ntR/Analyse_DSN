// APIHandler.gs - Service de gestion des appels API pour récupérer les DSN

const APIHandler = (function() {
  // URL de base de l'API (à ajuster selon votre fournisseur DSN)
  const API_BASE_URL = "https://api.dsn-service.fr/v1";
  
  /**
   * Importe les données DSN via API pour analyse
   * @param {object} formData - Données du formulaire d'importation
   * @return {object} - Résultat de l'importation
   */
  function importerDSNForAnalysis(formData) {
    try {
      // Vérification des paramètres obligatoires
      if (!formData.apiIdentifiant || !formData.apiMotDePasse) {
        throw new Error("Identifiants API manquants");
      }
      
      if (!formData.codeDossier || !formData.codeEtablissement) {
        throw new Error("Paramètres de recherche incomplets");
      }
      
      if (!formData.dateDebut || !formData.dateFin) {
        throw new Error("Période de recherche non spécifiée");
      }
      
      // Ajuster les dates au format de l'API si nécessaire
      const dateDebut = Utilities.formatDate(new Date(formData.dateDebut), 
                                           Session.getScriptTimeZone(), 
                                           "yyyy-MM-dd");
      const dateFin = Utilities.formatDate(new Date(formData.dateFin), 
                                         Session.getScriptTimeZone(), 
                                         "yyyy-MM-dd");
      
      // Préparation des mois à rechercher
      const moisRecherches = _getMonthsBetweenDates(dateDebut, dateFin);
      
      // Récupération des données DSN pour chaque mois
      const apiDSNData = [];
      
      for (const periode of moisRecherches) {
        const dsnData = _fetchDSNForMonth(formData, periode.annee, periode.mois);
        if (dsnData && dsnData.success && dsnData.data) {
          apiDSNData.push({
            annee: periode.annee,
            mois: periode.mois,
            data: dsnData.data
          });
        }
      }
      
      return {
        success: true,
        apiDSNData: apiDSNData,
        message: `${apiDSNData.length} mois récupérés avec succès`
      };
    } catch (e) {
      Logger.log("Erreur lors de l'import API: " + e.toString());
      return {
        success: false,
        message: "Erreur lors de l'import: " + e.toString()
      };
    }
  }
  
  /**
   * Récupère les données DSN pour un mois spécifique
   * @private
   * @param {object} formData - Données du formulaire
   * @param {number} annee - Année
   * @param {number} mois - Mois
   * @return {object} - Résultat de la récupération
   */
  function _fetchDSNForMonth(formData, annee, mois) {
    try {
      // Dans une application réelle, vous feriez un appel API ici
      // Par exemple avec UrlFetchApp
      
      // Exemple de code pour une API réelle
      /*
      const url = `${API_BASE_URL}/dsn/${formData.codeDossier}/${formData.codeEtablissement}/${annee}/${mois}`;
      
      const options = {
        method: 'get',
        headers: {
          'Authorization': 'Basic ' + Utilities.base64Encode(`${formData.apiIdentifiant}:${formData.apiMotDePasse}`),
          'Content-Type': 'application/json'
        }
      };
      
      const response = UrlFetchApp.fetch(url, options);
      
      if (response.getResponseCode() === 200) {
        const jsonResponse = JSON.parse(response.getContentText());
        return {
          success: true,
          data: jsonResponse
        };
      } else {
        throw new Error(`Erreur API (${response.getResponseCode()}): ${response.getContentText()}`);
      }
      */
      
      // DÉMONSTRATION: Simuler une réponse DSN pour test
      // Dans une application réelle, vous supprimeriez ce code et utiliseriez celui ci-dessus
      
      // Simuler un délai (comme un vrai appel API)
      Utilities.sleep(500);
      
      // Créer des données DSN fictives pour démonstration
      return {
        success: true,
        data: _generateMockDSNData(annee, mois, formData)
      };
    } catch (e) {
      Logger.log(`Erreur lors de la récupération DSN pour ${mois}/${annee}: ${e.toString()}`);
      return {
        success: false,
        message: e.toString()
      };
    }
  }
  
  /**
   * Génère des données DSN fictives pour la démonstration
   * @private
   * @param {number} annee - Année
   * @param {number} mois - Mois
   * @param {object} formData - Données du formulaire
   * @return {string} - Texte DSN simulé
   */
  function _generateMockDSNData(annee, mois, formData) {
    const mm = mois < 10 ? "0" + mois : mois;
    const dsnText = `S21.G00.06.001,${annee}${mm}
S21.G00.06.002,${annee}-${mm}
S21.G00.06.003,12
S21.G00.06.004,DSN Mensuelle

S21.G00.11.001,${formData.codeDossier || "123456789"}
S21.G00.11.002,${formData.codeEtablissement || "00001"}
S21.G00.11.003,Entreprise Démo DSN
S21.G00.11.004,1

S21.G00.30.001,1234567890123
S21.G00.30.002,Dupont
S21.G00.30.004,Jean
S21.G00.30.005,01
S21.G00.30.006,01011980
S21.G00.30.022,01

S21.G00.40.001,01012022
S21.G00.40.002,01
S21.G00.40.004,3511
S21.G00.40.006,Développeur
S21.G00.40.007,01
S21.G00.40.009,CONT001
S21.G00.40.010,

S21.G00.30.001,2345678901234
S21.G00.30.002,Martin
S21.G00.30.004,Sophie
S21.G00.30.005,02
S21.G00.30.006,15021985
S21.G00.30.022,01

S21.G00.40.001,01022022
S21.G00.40.002,01
S21.G00.40.004,3821
S21.G00.40.006,Comptable
S21.G00.40.007,01
S21.G00.40.009,CONT002
S21.G00.40.010,

S21.G00.51.001,01${mm}${annee}
S21.G00.51.002,28${mm}${annee}
S21.G00.51.010,CONT001
S21.G00.51.011,001
S21.G00.51.013,3500.00

S21.G00.51.001,01${mm}${annee}
S21.G00.51.002,28${mm}${annee}
S21.G00.51.010,CONT002
S21.G00.51.011,001
S21.G00.51.013,3200.00

S21.G00.60.001,01
S21.G00.60.002,10${mm}${annee}
S21.G00.60.003,15${mm}${annee}
S21.G00.60.010,16${mm}${annee}
S21.G00.60.011,01`;

    return dsnText;
  }
  
  /**
   * Obtient la liste des mois entre deux dates
   * @private
   * @param {string} dateDebut - Date de début (YYYY-MM-DD)
   * @param {string} dateFin - Date de fin (YYYY-MM-DD)
   * @return {array} - Liste des périodes {annee, mois}
   */
  function _getMonthsBetweenDates(dateDebut, dateFin) {
    const debut = new Date(dateDebut);
    const fin = new Date(dateFin);
    
    const periodes = [];
    const currentDate = new Date(debut.getFullYear(), debut.getMonth(), 1);
    
    while (currentDate <= fin) {
      periodes.push({
        annee: currentDate.getFullYear(),
        mois: currentDate.getMonth() + 1
      });
      
      // Passer au mois suivant
      currentDate.setMonth(currentDate.getMonth() + 1);
    }
    
    return periodes;
  }
  
  /**
   * Récupère la DSN la plus récente
   * @param {object} options - Options de recherche
   * @return {object} - Résultat de la récupération
   */
  function getMostRecentDSN(options) {
    try {
      // Obtenir la date du mois dernier
      const today = new Date();
      const lastMonth = new Date(today.getFullYear(), today.getMonth() - 1, 1);
      const annee = lastMonth.getFullYear();
      const mois = lastMonth.getMonth() + 1;
      
      // Utiliser les options ou des valeurs par défaut
      const formData = {
        apiIdentifiant: options.apiIdentifiant || "",
        apiMotDePasse: options.apiMotDePasse || "",
        codeDossier: options.codeDossier || "",
        codeEtablissement: options.codeEtablissement || ""
      };
      
      // Récupérer la DSN du mois dernier
      const dsnData = _fetchDSNForMonth(formData, annee, mois);
      
      if (dsnData && dsnData.success && dsnData.data) {
        return {
          success: true,
          dsnData: {
            annee: annee,
            mois: mois,
            data: dsnData.data
          }
        };
      } else {
        throw new Error("Impossible de récupérer la DSN la plus récente");
      }
    } catch (e) {
      Logger.log("Erreur lors de la récupération de la DSN récente: " + e.toString());
      return {
        success: false,
        message: "Erreur: " + e.toString()
      };
    }
  }
  
  /**
   * Vérifie si les identifiants API sont valides
   * @param {string} apiIdentifiant - Identifiant API
   * @param {string} apiMotDePasse - Mot de passe API
   * @return {object} - Résultat de la vérification
   */
  function validateAPICredentials(apiIdentifiant, apiMotDePasse) {
    try {
      if (!apiIdentifiant || !apiMotDePasse) {
        throw new Error("Identifiants incomplets");
      }
      
      // Dans une version réelle, vous feriez un appel API pour vérifier
      // Ici, nous simulons une vérification
      if (apiIdentifiant === "demo" && apiMotDePasse === "demo") {
        return {
          success: true,
          message: "Identifiants valides"
        };
      } else {
        // Pour la démo, accepter n'importe quels identifiants fournis
        return {
          success: true,
          message: "Identifiants acceptés"
        };
        
        // Dans un cas réel, vous retourneriez plutôt:
        // throw new Error("Identifiants invalides");
      }
    } catch (e) {
      Logger.log("Erreur de validation API: " + e.toString());
      return {
        success: false,
        message: "Erreur: " + e.toString()
      };
    }
  }
  
  /**
   * Récupère la liste des établissements disponibles
   * @param {object} credentials - Identifiants API
   * @return {array} - Liste des établissements
   */
  function getAvailableEstablishments(credentials) {
    try {
      // Dans une version réelle, vous feriez un appel API
      // Ici, nous retournons des données factices
      return {
        success: true,
        establishments: [
          { code: "00001", name: "Siège Social" },
          { code: "00002", name: "Établissement Paris" },
          { code: "00003", name: "Établissement Lyon" },
          { code: "00004", name: "Établissement Marseille" }
        ]
      };
    } catch (e) {
      Logger.log("Erreur lors de la récupération des établissements: " + e.toString());
      return {
        success: false,
        message: "Erreur: " + e.toString()
      };
    }
  }
  
  // API publique
  return {
    importerDSNForAnalysis,
    getMostRecentDSN,
    validateAPICredentials,
    getAvailableEstablishments
  };
})();
