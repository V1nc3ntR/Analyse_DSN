// SpreadsheetHandler.gs - Gestion des opérations sur les feuilles

const SpreadsheetHandler = (function() {
  // Cache pour stocker les données des feuilles
  const sheetCache = {};
  
  return {
    /**
     * Stocke les DSN dans les feuilles de calcul
     * @param {array} dsnArray - Tableau de DSN à stocker
     * @param {object} settings - Paramètres optionnels
     */
    storeDSNToSheets: function(dsnArray, settings) {
      var ss = SpreadsheetApp.getActiveSpreadsheet();
      dsnArray.forEach(function(dsn) {
        var mm = dsn.mois < 10 ? "0" + dsn.mois : dsn.mois;
        var sheetName = mm + "-" + dsn.annee;
        var sheet = this._getOrCreateSheet(ss, sheetName);
        var individus = dsn.data.individus || [];
        
        // Traitement par lots pour améliorer les performances
        this._processIndividusBatch(sheet, dsn.data.moisPrincipal, individus, settings);
      }, this);
    },
    
    /**
     * Stocke les DSN locales dans les feuilles de calcul
     * @param {array} localDSNs - Tableau de DSN locales à stocker
     * @param {object} settings - Paramètres optionnels
     */
    storeLocalDSNToSheets: function(localDSNs, settings) {
      var ss = SpreadsheetApp.getActiveSpreadsheet();
      localDSNs.forEach(function(finalModel) {
        var splitted = finalModel.moisPrincipal.split("-");
        var sheetName = "SansMois";
        if(splitted.length === 2) {
          sheetName = splitted[1] + "-" + splitted[0];
        }
        var sheet = this._getOrCreateSheet(ss, sheetName);
        
        // Traitement par lots
        this._processIndividusBatch(sheet, finalModel.moisPrincipal, finalModel.individus || [], settings);
      }, this);
    },
    
    /**
     * Obtient le salaire précédent pour un contrat
     * @param {string} nir - NIR de l'individu
     * @param {string} numCtr - Numéro de contrat
     * @param {string} currentMoisPrincipal - Mois principal actuel
     * @return {number|null} - Salaire précédent ou null
     */
    getPreviousSalaryForContract: function(nir, numCtr, currentMoisPrincipal) {
      var parts = currentMoisPrincipal.split("-");
      if(parts.length < 2) return null;
      
      var yyyy = Number(parts[0]);
      var mm = Number(parts[1]);
      var prevDate = new Date(yyyy, mm-2, 1);
      var prevYear = prevDate.getFullYear();
      var prevMonth = prevDate.getMonth() + 1;
      var prevSheetName = (prevMonth < 10 ? "0" + prevMonth : prevMonth) + "-" + prevYear;
      
      // Utiliser le cache si disponible
      if (!sheetCache[prevSheetName]) {
        var ss = SpreadsheetApp.getActiveSpreadsheet();
        var sheet = ss.getSheetByName(prevSheetName);
        if(!sheet) return null;
        
        // Mettre en cache les données pour des recherches futures
        sheetCache[prevSheetName] = sheet.getDataRange().getValues();
      }
      
      var data = sheetCache[prevSheetName];
      
      // Recherche plus efficace
      for(var i = 1; i < data.length; i++) {
        if(data[i][0] === nir && data[i][19] === numCtr) {
          return Number(data[i][22]);
        }
      }
      return null;
    },
    
    /**
     * Vide le cache des feuilles
     */
    clearCache: function() {
      Object.keys(sheetCache).forEach(key => delete sheetCache[key]);
    },
    
    /**
     * Obtient ou crée une feuille
     * @private
     * @param {Spreadsheet} ss - Classeur
     * @param {string} sheetName - Nom de la feuille
     * @return {Sheet} - Feuille
     */
    _getOrCreateSheet: function(ss, sheetName) {
      var sheet = ss.getSheetByName(sheetName);
      if(!sheet) {
        sheet = ss.insertSheet(sheetName);
        this._initSheetHeader(sheet);
      }
      return sheet;
    },
    
    /**
     * Initialise les en-têtes de la feuille
     * @private
     * @param {Sheet} sheet - Feuille
     */
    _initSheetHeader: function(sheet) {
      var headers = [
        // Infos Individu
        "NIR (S21.G00.30.001)",
        "Nom (S21.G00.30.002)",
        "Nom d'usage (S21.G00.30.003)",
        "Prénom (S21.G00.30.004)",
        "Sexe (S21.G00.30.005)",
        "Date Naissance (S21.G00.30.006)",
        "Lieu Naissance (S21.G00.30.007)",
        "Dépt Naissance (S21.G00.30.014)",
        "Pays Naissance (S21.G00.30.015)",
        "Matricule (S21.G00.30.019)",
        "NTT (S21.G00.30.020)",
        "Statut Étranger (S21.G00.30.022)",
        "Cumul emploi retraite (S21.G00.30.023)",
        "Niveau Formation (S21.G00.30.024)",
        "Niveau Diplôme (S21.G00.30.025)",
        "Libellé Pays Naiss (S21.G00.30.029)",

        // Champs calculés
        "Âge",
        "Tranche d'âge",
        "Maternité/Adoption ?",

        // Contrat
        "Numéro Contrat (S21.G00.40.009)",
        "Plusieurs Contrats ?",
        "Nb Contrats",

        // Rémunérations
        "Rémunération Version 1",
        "Rémunération Version 2",

        // Evolutions
        "Evolution Salaire",
        "Evolution après maternité/adoption",

        // Date de Maj
        "Date de Maj"
      ];
      sheet.getRange(1,1,1,headers.length).setValues([headers]);

      sheet.getRange(1,6,sheet.getMaxRows(),1).setNumberFormat("dd/MM/yyyy");  // date de naissance
      sheet.getRange(1,27,sheet.getMaxRows(),1).setNumberFormat("dd/MM/yyyy"); // date de maj

      sheet.getRange(1,22,sheet.getMaxRows(),1).setNumberFormat("0");    // nb contrats
      sheet.getRange(1,23,sheet.getMaxRows(),1).setNumberFormat("0.00"); // remu1
      sheet.getRange(1,24,sheet.getMaxRows(),1).setNumberFormat("0.00"); // remu2
    },
    
    /**
     * Traite un lot d'individus pour insertion dans la feuille
     * @private
     * @param {Sheet} sheet - Feuille
     * @param {string} moisPrincipal - Mois principal
     * @param {array} individus - Tableau d'individus
     * @param {object} settings - Paramètres optionnels
     */
    _processIndividusBatch: function(sheet, moisPrincipal, individus, settings) {
      // Optimisation : traitement par lots de 50 individus
      const BATCH_SIZE = settings?.batchSize || 50;
      for (let i = 0; i < individus.length; i += BATCH_SIZE) {
        const batch = individus.slice(i, i + BATCH_SIZE);
        
        // Préparer les données pour insertion en masse
        let allRows = [];
        batch.forEach(function(individu) {
          const rows = this._prepareIndividuRows(sheet, moisPrincipal, individu, settings);
          allRows = allRows.concat(rows);
        }, this);
        
        // Insérer toutes les lignes d'un coup
        if (allRows.length > 0) {
          const startRow = sheet.getLastRow() + 1;
          sheet.getRange(startRow, 1, allRows.length, allRows[0].length).setValues(allRows);
          
          // Appliquer les formatages en une seule opération
          this._applyBatchFormatting(sheet, startRow, allRows.length);
        }
      }
    },
    
    /**
     * Prépare les lignes à insérer pour un individu
     * @private
     * @param {Sheet} sheet - Feuille
     * @param {string} moisPrincipal - Mois principal
     * @param {object} individu - Données de l'individu
     * @param {object} settings - Paramètres optionnels
     * @return {array} - Lignes préparées
     */
    _prepareIndividuRows: function(sheet, moisPrincipal, individu, settings) {
      const rows = [];
      const nir = individu.nir || "";
      const rubs = individu.rubriques || [];
      
      // Extraire les informations de base
      const nom = findRubriqueValue(rubs, "002") || "";
      const nomUsage = findRubriqueValue(rubs, "003") || "";
      const prenom = findRubriqueValue(rubs, "004") || "";
      const sexe = findRubriqueValue(rubs, "005") || "";
      const datnais = findRubriqueValue(rubs, "006") || "";
      const lieuNaiss = findRubriqueValue(rubs, "007") || "";
      const deptNaiss = findRubriqueValue(rubs, "014") || "";
      const paysNaiss = findRubriqueValue(rubs, "015") || "";
      const matricule = findRubriqueValue(rubs, "019") || "";
      const ntt = findRubriqueValue(rubs, "020") || "";
      const statutEtranger = findRubriqueValue(rubs, "022") || "";
      const cumulRetraite = findRubriqueValue(rubs, "023") || "";
      const nivFormation = findRubriqueValue(rubs, "024") || "";
      const nivDiplome = findRubriqueValue(rubs, "025") || "";
      const libPaysNaiss = findRubriqueValue(rubs, "029") || "";

      // Calculer l'âge et la tranche d'âge
      const dateObj = DateUtils.parseDateString(datnais);
      let age = "";
      let tranche = "";
      if(dateObj && !isNaN(dateObj.getTime())){
        const refDate = DateUtils.getLastDayOfMonth(moisPrincipal);
        age = DateUtils.calcAge(dateObj, refDate);
        tranche = DateUtils.getTrancheAge(age);
      }

      // Déterminer si c'est un arrêt maternité ou adoption
      let isMatAdopt = false;
      (individu.arrets || []).forEach(function(a){
        if(a.motif === "02" || a.motif === "09"){
          isMatAdopt = true;
        }
      });

      // Gérer les contrats
      const nbCtr = (individu.contrats || []).length;
      const isMulti = nbCtr > 1 ? "Oui" : "Non";

      // S'il n'y a pas de contrat, ajouter une ligne quand même
      let ctList = individu.contrats || [];
      if(ctList.length === 0) {
        ctList = [{ numContrat: "", rubriques: [], remunerations: [] }];
      }

      // Pour chaque contrat, créer une ligne
      ctList.forEach(function(contrat) {
        // Calculer les rémunérations
        const remu1 = calcRemunerationVersion1ForContract(contrat);
        const remu2 = calcRemunerationVersion2ForContract(contrat);
        
        // Calculer l'évolution du salaire
        const prevSalary = this.getPreviousSalaryForContract(nir, contrat.numContrat, moisPrincipal);
        let evolution = "N/A";
        if(prevSalary !== null && !isNaN(prevSalary) && remu1 !== null){
          if(remu1 > prevSalary) evolution = "Augmentation";
          else if(remu1 < prevSalary) evolution = "Baisse";
          else evolution = "Stable";
        }
        
        // Évolution après maternité/adoption
        let evolMatAdopt = "N/A";
        if(sexe === "02" && isMatAdopt && prevSalary !== null && remu1 !== null){
          evolMatAdopt = (remu1 !== prevSalary) ? "Oui" : "Non";
        }
        
        // Créer la ligne de données
        const rowData = [
          nir,
          nom,
          nomUsage,
          prenom,
          sexe,
          dateObj && !isNaN(dateObj.getTime()) ? dateObj : datnais,
          lieuNaiss,
          deptNaiss,
          paysNaiss,
          matricule,
          ntt,
          statutEtranger,
          cumulRetraite,
          nivFormation,
          nivDiplome,
          libPaysNaiss,
          age !== "" ? Number(age) : "",
          tranche,
          isMatAdopt ? "Oui" : "Non",
          contrat.numContrat || "",
          isMulti,
          nbCtr,
          remu1 !== null ? remu1 : "",
          remu2 !== null ? remu2 : "",
          evolution,
          evolMatAdopt,
          new Date()
        ];
        
        rows.push(rowData);
      }, this);
      
      return rows;
    },
    
    /**
     * Applique les formatages en masse pour les lignes insérées
     * @private
     * @param {Sheet} sheet - Feuille
     * @param {number} startRow - Ligne de début
     * @param {number} rowCount - Nombre de lignes
     */
    _applyBatchFormatting: function(sheet, startRow, rowCount) {
      if (rowCount > 0) {
        // Formatage des dates
        sheet.getRange(startRow, 6, rowCount, 1).setNumberFormat("dd/MM/yyyy");
        sheet.getRange(startRow, 27, rowCount, 1).setNumberFormat("dd/MM/yyyy");
        
        // Formatage des nombres
        sheet.getRange(startRow, 22, rowCount, 1).setNumberFormat("0");
        sheet.getRange(startRow, 23, rowCount, 1).setNumberFormat("0.00");
        sheet.getRange(startRow, 24, rowCount, 1).setNumberFormat("0.00");
      }
    }
  };
})();

// Fonctions de calcul des rémunérations
function calcRemunerationVersion1ForContract(contrat){
  var total = 0;
  (contrat.remunerations || []).forEach(function(rem){
    // On cherche la rubrique "011" => type, "013" => montant
    var typeRem = findRubriqueValue(rem.rubriques, "011") || "";
    var montant = parseFloat(findRubriqueValue(rem.rubriques, "013")) || 0;
    if(typeRem === "010" || typeRem === "999"){
      total += montant;
    }
  });
  return total;
}

function calcRemunerationVersion2ForContract(contrat){
  var total = 0;
  (contrat.remunerations || []).forEach(function(rem){
    var montant = parseFloat(findRubriqueValue(rem.rubriques, "013")) || 0;
    total += montant;
  });
  return total;
}
