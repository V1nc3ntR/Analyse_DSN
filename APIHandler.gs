// APIHandler.gs - Gestion des appels API

const APIHandler = (function() {
  return {
    /**
     * Importe les DSN via l'API pour analyse
     * @param {object} formData - Données du formulaire d'import
     * @return {object} - Résultat de l'import
     */
    importerDSNForAnalysis: function(formData) {
      var apiIdentifiant    = formData.apiIdentifiant;
      var apiMotDePasse     = formData.apiMotDePasse;
      var codeDossier       = formData.codeDossier;
      var codeEtablissement = formData.codeEtablissement;
      var dateDebut         = new Date(formData.dateDebut);
      var dateFin           = new Date(formData.dateFin);
      var settings          = formData.settings || {}; // Paramètres additionnels

      var moisList = DateUtils.getMonthsBetween(dateDebut, dateFin);
      var results = [];

      moisList.forEach(function(item) {
        var annee = item.annee;
        var mois = item.mois;
        var url = this._buildApiUrl(formData, item);
        var options = this._buildApiOptions(formData);
        
        try {
          var response = UrlFetchApp.fetch(url, options);
          if(response.getResponseCode() === 200) {
            var dsnTxt = response.getContentText();
            var parsedModel = DSNProcessor.parseDSNContent(dsnTxt);
            results.push({
              annee: annee,
              mois: mois,
              data: parsedModel
            });
          } else {
            Logger.log("DSN introuvable pour %s-%s => HTTP %s", annee, mois, response.getResponseCode());
          }
        } catch(e) {
          Logger.log("Erreur DSN %s-%s => %s", annee, mois, e);
        }
      }, this);
      
      return { apiDSNData: results };
    },
    
    /**
     * Construit l'URL de l'API DSN
     * @private
     * @param {object} formData - Données du formulaire
     * @param {object} item - Item mois/année
     * @return {string} - URL formatée
     */
    _buildApiUrl: function(formData, item) {
      return "https://api.openpaye.co/DSNs?codeDossier=" + formData.codeDossier +
             "&codeEtablissement=" + formData.codeEtablissement +
             "&annee=" + item.annee +
             "&mois=" + item.mois +
             "&format=txt";
    },
    
    /**
     * Construit les options pour l'appel API
     * @private
     * @param {object} formData - Données du formulaire
     * @return {object} - Options de l'appel
     */
    _buildApiOptions: function(formData) {
      return {
        method: "GET",
        headers: {
          "Authorization": "Basic " + Utilities.base64Encode(formData.apiIdentifiant + ":" + formData.apiMotDePasse)
        },
        muteHttpExceptions: true
      };
    }
  };
})();
