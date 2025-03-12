// DSNProcessor.gs - Module principal de traitement des DSN

const DSNProcessor = (function() {
  // Variables privées et constantes
  const RUBRIC_PATTERNS = {
    BLOC: /^(S\d{2}\.G00\.\d{2,3})\.(\d{1,3})$/
  };
  
  // Méthodes publiques
  return {
    /**
     * Parse le contenu d'une DSN et construit le modèle de données
     * @param {string} dsnText - Contenu textuel de la DSN
     * @return {object} - Modèle de données structuré
     */
    parseDSNContent: function(dsnText) {
      var lines = dsnText.split(/\r?\n/);
      var rubrics = [];
      lines.forEach(function(line) {
        line = line.trim();
        if(!line) return;
        var parts = line.split(",");
        if(parts.length < 2) return;
        var left = parts[0].trim();
        var right = parts[1].trim();
        if(right.startsWith("'") && right.endsWith("'")) {
          right = right.substring(1, right.length - 1);
        }
        var match = left.match(RUBRIC_PATTERNS.BLOC);
        if(!match) return;
        rubrics.push({
          bloc: match[1],
          rubrique: match[2],
          valeur: right
        });
      });
      var blocMap = {};
      rubrics.forEach(function(r) {
        if(!blocMap[r.bloc]) blocMap[r.bloc] = [];
        blocMap[r.bloc].push({ rubrique: r.rubrique, valeur: r.valeur });
      });
      return this.buildDsnModel(blocMap);
    },
    
    /**
     * Version optimisée du parsing DSN
     * @param {string} dsnText - Contenu textuel de la DSN
     * @return {object} - Modèle de données structuré
     */
    parseDSNContentOptimized: function(dsnText) {
      // Utilisation d'expressions régulières pour la correspondance exacte
      const rubricRegex = /^(S\d{2}\.G00\.\d{2,3})\.(\d{1,3}),(.+)$/;
      const quotedValueRegex = /^'(.*)'$/;
      
      const blocMap = {};
      
      // Utiliser une seule expression régulière pour l'analyse complète
      const lines = dsnText.split(/\r?\n/);
      
      for (let i = 0; i < lines.length; i++) {
        const line = lines[i].trim();
        if (!line) continue;
        
        const match = line.match(rubricRegex);
        if (!match) continue;
        
        const [, bloc, rubrique, rawValue] = match;
        let valeur = rawValue.trim();
        
        // Traiter les valeurs entre guillemets
        const quotedMatch = valeur.match(quotedValueRegex);
        if (quotedMatch) {
          valeur = quotedMatch[1];
        }
        
        // Initialiser le tableau si nécessaire
        if (!blocMap[bloc]) {
          blocMap[bloc] = [];
        }
        
        // Ajouter la rubrique
        blocMap[bloc].push({ rubrique, valeur });
      }
      
      return this.buildDsnModel(blocMap);
    },
    
    /**
     * Construit le modèle de données à partir des blocs de rubriques
     * @param {object} blocs - Map des blocs de rubriques
     * @return {object} - Modèle de données structuré
     */
    buildDsnModel: function(blocs) {
      var model = { moisPrincipal: null, individus: [] };

      // S20.G00.05 => annee + mois
      var decl = blocs["S20.G00.05"];
      if(decl) {
        var annee = decl.find(r => r.rubrique === "005");
        var mois  = decl.find(r => r.rubrique === "006");
        if(annee && mois) {
          var mm = mois.valeur;
          if(mm.length === 1) mm = "0" + mm;
          model.moisPrincipal = annee.valeur + "-" + mm;
        }
      }
      if(!model.moisPrincipal) {
        model.moisPrincipal = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), "yyyy-MM");
      }

      // S21.G00.30 => individus
      var s2130 = blocs["S21.G00.30"] || [];
      var allInds = [];
      var currentInd = null;
      s2130.forEach(function(r) {
        if(r.rubrique === "001") {
          if(currentInd) allInds.push(currentInd);
          currentInd = {
            nir: r.valeur,
            rubriques: [],
            contrats: [],
            arrets: []
          };
        } else {
          if(!currentInd) {
            currentInd = { nir: null, rubriques: [], contrats: [], arrets: [] };
          }
          currentInd.rubriques.push(r);
        }
      });
      if(currentInd) allInds.push(currentInd);

      // S21.G00.40 => contrats
      var c40 = blocs["S21.G00.40"] || [];
      var contrats = [];
      var currentCtr = null;
      c40.forEach(function(r) {
        if(r.rubrique === "009") {
          if(currentCtr) contrats.push(currentCtr);
          currentCtr = { numContrat: r.valeur, rubriques: [], remunerations: [] };
          currentCtr.rubriques.push(r);
        } else {
          if(!currentCtr) {
            currentCtr = { numContrat: null, rubriques: [], remunerations: [] };
          }
          currentCtr.rubriques.push(r);
        }
      });
      if(currentCtr) contrats.push(currentCtr);

      // S21.G00.51 => rémunérations
      var c51 = blocs["S21.G00.51"] || [];
      var allRems = [];
      var currentRem = null;
      c51.forEach(function(r) {
        if(r.rubrique === "010") {
          if(currentRem) allRems.push(currentRem);
          currentRem = { numContrat: r.valeur, rubriques: [] };
        } else {
          if(!currentRem) {
            currentRem = { numContrat: null, rubriques: [] };
          }
          currentRem.rubriques.push(r);
        }
      });
      if(currentRem) allRems.push(currentRem);

      allRems.forEach(function(rem) {
        var found = contrats.find(c => c.numContrat === rem.numContrat);
        if(found) {
          found.remunerations.push(rem);
        }
      });

      // S21.G00.60 => arrêts
      var c60 = blocs["S21.G00.60"] || [];
      var allArr = [];
      var currentArr = null;
      c60.forEach(function(r) {
        if(r.rubrique === "001") {
          if(currentArr) allArr.push(currentArr);
          currentArr = { motif: r.valeur, rubriques: [] };
          currentArr.rubriques.push(r);
        } else {
          if(!currentArr) {
            currentArr = { motif: null, rubriques: [] };
          }
          currentArr.rubriques.push(r);
        }
      });
      if(currentArr) allArr.push(currentArr);

      // Répartition contrats/arrets => individus
      var indIndex = 0;
      contrats.forEach(function(ct) {
        if(!allInds[indIndex]) indIndex = 0;
        allInds[indIndex].contrats.push(ct);
        indIndex++;
        if(indIndex >= allInds.length) indIndex = 0;
      });
      indIndex = 0;
      allArr.forEach(function(ar) {
        if(!allInds[indIndex]) indIndex = 0;
        allInds[indIndex].arrets.push(ar);
        indIndex++;
        if(indIndex >= allInds.length) indIndex = 0;
      });

      model.individus = allInds;
      return model;
    },
    
    /**
     * Cherche la valeur d'une rubrique dans un tableau de rubriques
     * @param {array} arr - Tableau de rubriques
     * @param {string} rubId - Identifiant de la rubrique
     * @return {string|null} - Valeur de la rubrique ou null
     */
    findRubriqueValue: function(arr, rubId) {
      if(!arr) return null;
      var f = arr.find(function(x){ return x.rubrique === rubId; });
      return f ? f.valeur : null;
    },
    
    /**
     * Version optimisée de la recherche de rubrique
     * @param {array} arr - Tableau de rubriques
     * @param {string} rubId - Identifiant de la rubrique
     * @return {string|null} - Valeur de la rubrique ou null
     */
    findRubriqueValueOptimized: function(arr, rubId) {
      if (!arr) return null;
      
      // Convertir le tableau en objet pour une recherche O(1)
      const rubriquesMap = {};
      arr.forEach(item => {
        rubriquesMap[item.rubrique] = item.valeur;
      });
      
      return rubriquesMap[rubId] || null;
    }
  };
})();

// Exportation de la fonction findRubriqueValue pour compatibilité
function findRubriqueValue(arr, rubId) {
  return DSNProcessor.findRubriqueValue(arr, rubId);
}
