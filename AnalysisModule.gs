// AnalysisModule.gs - Module d'analyse avancée des données

const AnalysisModule = (function() {
  return {
    /**
     * Détecte les anomalies de salaire en fonction des critères définis
     * @param {object} data - Données DSN avec individus et contrats
     * @param {object} options - Options de détection
     * @return {array} - Liste des anomalies détectées
     */
    detectSalaryAnomalies: function(data, options = {}) {
      const anomalies = [];
      const threshold = options.threshold || 10; // % de variation considérée comme anomalie
      
      (data.individus || []).forEach(individu => {
        (individu.contrats || []).forEach(contrat => {
          const remu1 = calcRemunerationVersion1ForContract(contrat);
          const prevSalary = SpreadsheetHandler.getPreviousSalaryForContract(
            individu.nir, contrat.numContrat, data.moisPrincipal
          );
          
          if (prevSalary !== null && remu1 !== null) {
            const variation = ((remu1 - prevSalary) / prevSalary) * 100;
            
            if (Math.abs(variation) > threshold) {
              anomalies.push({
                nir: individu.nir,
                nom: findRubriqueValue(individu.rubriques, "002") || "",
                prenom: findRubriqueValue(individu.rubriques, "004") || "",
                contrat: contrat.numContrat,
                ancienSalaire: prevSalary,
                nouveauSalaire: remu1,
                variation: variation.toFixed(2) + "%",
                type: variation > 0 ? "Augmentation" : "Diminution"
              });
            }
          }
        });
      });
      
      return anomalies;
    },
    
    /**
     * Génère des statistiques par tranche d'âge et par sexe
     * @param {array} allDSN - Ensemble des DSN à analyser
     * @return {object} - Statistiques calculées
     */
    generateStatistics: function(allDSN) {
      const stats = {
        tranchesAge: {
          "moins de 30 ans": { total: 0, hommes: 0, femmes: 0, salaireMoyen: 0, salaireTotalHommes: 0, salaireTotalFemmes: 0 },
          "30 à 39 ans": { total: 0, hommes: 0, femmes: 0, salaireMoyen: 0, salaireTotalHommes: 0, salaireTotalFemmes: 0 },
          "40 à 49 ans": { total: 0, hommes: 0, femmes: 0, salaireMoyen: 0, salaireTotalHommes: 0, salaireTotalFemmes: 0 },
          "50 ans et plus": { total: 0, hommes: 0, femmes: 0, salaireMoyen: 0, salaireTotalHommes: 0, salaireTotalFemmes: 0 }
        },
        global: {
          totalIndividus: 0,
          hommes: 0,
          femmes: 0,
          salaireMoyen: 0,
          salaireMedian: 0,
          ecartHommesFemmes: 0
        },
        evolution: {
          salaireMoyen: [],
          effectifs: []
        }
      };
      
      // Tous les salaires pour calculer la médiane
      const allSalaires = [];
      const salairesByTranche = {
        "moins de 30 ans": [],
        "30 à 39 ans": [],
        "40 à 49 ans": [],
        "50 ans et plus": []
      };
      
      // Traiter chaque DSN
      allDSN.forEach(dsn => {
        // Ajouter un point de données pour l'évolution
        const moisAnnee = dsn.data.moisPrincipal;
        let totalSalaire = 0;
        let nbIndividus = 0;
        
        (dsn.data.individus || []).forEach(individu => {
          const sexe = findRubriqueValue(individu.rubriques, "005");
          const datnais = findRubriqueValue(individu.rubriques, "006");
          const dateObj = DateUtils.parseDateString(datnais);
          
          if (dateObj && !isNaN(dateObj.getTime())) {
            const refDate = DateUtils.getLastDayOfMonth(dsn.data.moisPrincipal);
            const age = DateUtils.calcAge(dateObj, refDate);
            const tranche = DateUtils.getTrancheAge(age);
            
            // Incrémenter les compteurs
            stats.tranchesAge[tranche].total++;
            if (sexe === "01") { // Homme
              stats.tranchesAge[tranche].hommes++;
              stats.global.hommes++;
            } else if (sexe === "02") { // Femme
              stats.tranchesAge[tranche].femmes++;
              stats.global.femmes++;
            }
            stats.global.totalIndividus++;
            
            // Calculer le salaire total pour cet individu
            let salaireTotalIndividu = 0;
            (individu.contrats || []).forEach(contrat => {
              const salaire = calcRemunerationVersion1ForContract(contrat);
              if (salaire !== null) {
                salaireTotalIndividu += salaire;
                totalSalaire += salaire;
                
                // Ajouter à la liste des salaires pour la médiane
                allSalaires.push(salaire);
                salairesByTranche[tranche].push(salaire);
                
                // Ajouter aux totaux par sexe et tranche d'âge
                if (sexe === "01") { // Homme
                  stats.tranchesAge[tranche].salaireTotalHommes += salaire;
                } else if (sexe === "02") { // Femme
                  stats.tranchesAge[tranche].salaireTotalFemmes += salaire;
                }
              }
            });
            
            nbIndividus++;
          }
        });
        
        // Ajouter aux données d'évolution
        if (nbIndividus > 0) {
          stats.evolution.salaireMoyen.push({
            mois: moisAnnee,
            valeur: totalSalaire / nbIndividus
          });
          stats.evolution.effectifs.push({
            mois: moisAnnee,
            valeur: nbIndividus
          });
        }
      });
      
      // Calculer les moyennes finales
      Object.keys(stats.tranchesAge).forEach(tranche => {
        const t = stats.tranchesAge[tranche];
        if (t.total > 0) {
          const salairesTotal = t.salaireTotalHommes + t.salaireTotalFemmes;
          t.salaireMoyen = salairesTotal / t.total;
          
          // Calculer le salaire médian par tranche
          if (salairesByTranche[tranche].length > 0) {
            salairesByTranche[tranche].sort((a, b) => a - b);
            const midIndex = Math.floor(salairesByTranche[tranche].length / 2);
            t.salaireMedian = salairesByTranche[tranche][midIndex];
          }
          
          // Calculer l'écart salarial H/F
          if (t.hommes > 0 && t.femmes > 0) {
            const moyenneHommes = t.salaireTotalHommes / t.hommes;
            const moyenneFemmes = t.salaireTotalFemmes / t.femmes;
            t.ecartHommesFemmes = ((moyenneHommes - moyenneFemmes) / moyenneHommes) * 100;
          }
        }
      });
      
      // Calculer les statistiques globales
      if (stats.global.totalIndividus > 0) {
        stats.global.salaireMoyen = allSalaires.reduce((sum, val) => sum + val, 0) / allSalaires.length;
        
        // Calculer le salaire médian global
        if (allSalaires.length > 0) {
          allSalaires.sort((a, b) => a - b);
          const midIndex = Math.floor(allSalaires.length / 2);
          stats.global.salaireMedian = allSalaires[midIndex];
        }
        
        // Calculer l'écart salarial global H/F
        if (stats.global.hommes > 0 && stats.global.femmes > 0) {
          const salaireTotalHommes = Object.values(stats.tranchesAge).reduce((sum, t) => sum + t.salaireTotalHommes, 0);
          const salaireTotalFemmes = Object.values(stats.tranchesAge).reduce((sum, t) => sum + t.salaireTotalFemmes, 0);
          
          const moyenneHommes = salaireTotalHommes / stats.global.hommes;
          const moyenneFemmes = salaireTotalFemmes / stats.global.femmes;
          stats.global.ecartHommesFemmes = ((moyenneHommes - moyenneFemmes) / moyenneHommes) * 100;
        }
      }
      
      return stats;
    },
    
    /**
     * Analyse les périodes d'arrêt de travail
     * @param {array} allDSN - Ensemble des DSN à analyser
     * @return {object} - Statistiques des arrêts de travail
     */
    analyzeWorkStoppages: function(allDSN) {
      const stopStats = {
        total: 0,
        parMotif: {},
        dureeMoyenne: 0,
        distribution: {
          hommes: 0,
          femmes: 0,
          tranches: {
            "moins de 30 ans": 0,
            "30 à 39 ans": 0,
            "40 à 49 ans": 0,
            "50 ans et plus": 0
          }
        },
        evolution: []
      };
      
      // Codes motifs d'arrêt avec leurs libellés
      const motifsArret = {
        "01": "Maladie",
        "02": "Maternité",
        "03": "Paternité / accueil de l'enfant",
        "04": "Accident du travail",
        "05": "Accident de trajet",
        "06": "Maladie professionnelle",
        "07": "Temps partiel thérapeutique",
        "08": "Activité partielle",
        "09": "Adoption",
        "10": "Congé suite à accident ou maladie vie privée",
        "11": "Congé de formation professionnelle",
        "12": "Congé formation de cadres et animateurs pour la jeunesse",
        "13": "Autres"
      };
      
      // Initialiser les compteurs par motif
      Object.keys(motifsArret).forEach(code => {
        stopStats.parMotif[code] = {
          libelle: motifsArret[code],
          count: 0,
          dureeTotal: 0,
          dureeMoyenne: 0
        };
      });
      
      let totalDuree = 0;
      
      // Traiter chaque DSN
      allDSN.forEach(dsn => {
        const moisAnnee = dsn.data.moisPrincipal;
        let arretsduMois = 0;
        
        (dsn.data.individus || []).forEach(individu => {
          const sexe = findRubriqueValue(individu.rubriques, "005");
          const datnais = findRubriqueValue(individu.rubriques, "006");
          const dateObj = DateUtils.parseDateString(datnais);
          let tranche = "";
          
          if (dateObj && !isNaN(dateObj.getTime())) {
            const refDate = DateUtils.getLastDayOfMonth(dsn.data.moisPrincipal);
            const age = DateUtils.calcAge(dateObj, refDate);
            tranche = DateUtils.getTrancheAge(age);
          }
          
          (individu.arrets || []).forEach(arret => {
            const motif = arret.motif;
            if (motif && stopStats.parMotif[motif]) {
              stopStats.total++;
              stopStats.parMotif[motif].count++;
              arretsduMois++;
              
              // Calculer la durée si disponible
              const dateDebut = findRubriqueValue(arret.rubriques, "002");
              const dateFin = findRubriqueValue(arret.rubriques, "003");
              
              if (dateDebut && dateFin) {
                const debut = DateUtils.parseDateString(dateDebut);
                const fin = DateUtils.parseDateString(dateFin);
                
                if (debut && fin && !isNaN(debut.getTime()) && !isNaN(fin.getTime())) {
                  const duree = Math.round((fin - debut) / (1000 * 60 * 60 * 24)); // jours
                  stopStats.parMotif[motif].dureeTotal += duree;
                  totalDuree += duree;
                }
              }
              
              // Distribution par sexe et tranche d'âge
              if (sexe === "01") {
                stopStats.distribution.hommes++;
              } else if (sexe === "02") {
                stopStats.distribution.femmes++;
              }
              
              if (tranche) {
                stopStats.distribution.tranches[tranche]++;
              }
            }
          });
        });
        
        // Ajouter aux données d'évolution mensuelle
        stopStats.evolution.push({
          mois: moisAnnee,
          count: arretsduMois
        });
      });
      
      // Calculer les durées moyennes
      Object.keys(stopStats.parMotif).forEach(motif => {
        if (stopStats.parMotif[motif].count > 0) {
          stopStats.parMotif[motif].dureeMoyenne = stopStats.parMotif[motif].dureeTotal / stopStats.parMotif[motif].count;
        }
      });
      
      if (stopStats.total > 0) {
        stopStats.dureeMoyenne = totalDuree / stopStats.total;
      }
      
      return stopStats;
    },
    
    /**
     * Génère des rapports de conformité basés sur les données DSN
     * @param {array} allDSN - Ensemble des DSN à analyser
     * @return {object} - Rapports de conformité
     */
    generateComplianceReports: function(allDSN) {
      const compliance = {
        missingData: [],
        inconsistencies: [],
        warnings: []
      };
      
      // Traiter chaque DSN
      allDSN.forEach(dsn => {
        const moisAnnee = dsn.data.moisPrincipal;
        
        (dsn.data.individus || []).forEach(individu => {
          const nir = individu.nir || "";
          const nom = findRubriqueValue(individu.rubriques, "002") || "";
          const prenom = findRubriqueValue(individu.rubriques, "004") || "";
          
          // Vérifier les données obligatoires manquantes
          const requiredFields = [
            { code: "002", label: "Nom" },
            { code: "004", label: "Prénom" },
            { code: "005", label: "Sexe" },
            { code: "006", label: "Date de naissance" }
          ];
          
          requiredFields.forEach(field => {
            if (!findRubriqueValue(individu.rubriques, field.code)) {
              compliance.missingData.push({
                mois: moisAnnee,
                nir: nir,
                nom: nom,
                prenom: prenom,
                champ: field.label,
                code: field.code
              });
            }
          });
          
          // Vérifier les incohérences potentielles
          // Exemple: date de naissance incohérente
          const datnais = findRubriqueValue(individu.rubriques, "006");
          const dateObj = DateUtils.parseDateString(datnais);
          
          if (dateObj && !isNaN(dateObj.getTime())) {
            const now = new Date();
            const age = DateUtils.calcAge(dateObj, now);
            
            if (age < 16 || age > 80) {
              compliance.warnings.push({
                mois: moisAnnee,
                nir: nir,
                nom: nom,
                prenom: prenom,
                type: "Age inhabituel",
                description: `Age calculé: ${age} ans (Date de naissance: ${datnais})`
              });
            }
          }
          
          // Vérifier les contrats
          (individu.contrats || []).forEach(contrat => {
            const numContrat = contrat.numContrat || "";
            
            // Vérifier la présence de rémunération
            if ((contrat.remunerations || []).length === 0) {
              compliance.warnings.push({
                mois: moisAnnee,
                nir: nir,
                nom: nom,
                prenom: prenom,
                contrat: numContrat,
                type: "Contrat sans rémunération",
                description: "Aucune rémunération trouvée pour ce contrat"
              });
            }
            
            // Vérifier les valeurs de rémunération anormales
            (contrat.remunerations || []).forEach(rem => {
              const montant = parseFloat(findRubriqueValue(rem.rubriques, "013")) || 0;
              
              if (montant > 50000) { // Seuil arbitraire pour détecter les valeurs anormalement élevées
                compliance.inconsistencies.push({
                  mois: moisAnnee,
                  nir: nir,
                  nom: nom,
                  prenom: prenom,
                  contrat: numContrat,
                  type: "Rémunération potentiellement incorrecte",
                  description: `Montant anormalement élevé: ${montant.toFixed(2)} €`
                });
              }
            });
          });
        });
      });
      
      return compliance;
    },
    
    /**
     * Génère un tableau de bord complet pour l'ensemble des données
     * @param {array} allDSN - Ensemble des DSN à analyser
     * @return {object} - Données du tableau de bord
     */
    generateDashboard: function(allDSN) {
      // Statistiques générales
      const stats = this.generateStatistics(allDSN);
      
      // Analyse des arrêts de travail
      const arrets = this.analyzeWorkStoppages(allDSN);
      
      // Rapports de conformité
      const compliance = this.generateComplianceReports(allDSN);
      
      // Détection des anomalies salariales
      const anomalies = [];
      allDSN.forEach(dsn => {
        const detected = this.detectSalaryAnomalies(dsn.data, { threshold: 15 });
        detected.forEach(a => {
          a.mois = dsn.data.moisPrincipal;
          anomalies.push(a);
        });
      });
      
      // Agréger toutes les données dans un objet de tableau de bord
      return {
        stats: stats,
        arrets: arrets,
        compliance: compliance,
        anomalies: anomalies,
        lastUpdate: new Date()
      };
    },
    
    /**
     * Génère un rapport au format Excel pour exportation
     * @param {object} dashboard - Données du tableau de bord
     * @return {Spreadsheet} - Feuille Excel générée
     */
    exportDashboardToSheet: function(dashboard) {
      const ss = SpreadsheetApp.getActiveSpreadsheet();
      let reportSheet = ss.getSheetByName("Tableau de bord");
      
      if (reportSheet) {
        // Supprimer la feuille existante pour la recréer
        ss.deleteSheet(reportSheet);
      }
      
      reportSheet = ss.insertSheet("Tableau de bord");
      
      // Titre du rapport
      reportSheet.getRange("A1:F1").merge();
      reportSheet.getRange("A1").setValue("TABLEAU DE BORD DSN")
        .setFontWeight("bold")
        .setFontSize(16)
        .setHorizontalAlignment("center");
      
      // Date du rapport
      reportSheet.getRange("A2:F2").merge();
      reportSheet.getRange("A2").setValue("Généré le " + Utilities.formatDate(new Date(), Session.getScriptTimeZone(), "dd/MM/yyyy à HH:mm"))
        .setFontStyle("italic")
        .setHorizontalAlignment("center");
      
      // Section 1: Statistiques globales
      reportSheet.getRange("A4").setValue("STATISTIQUES GLOBALES")
        .setFontWeight("bold")
        .setFontSize(14);
      
      reportSheet.getRange("A5").setValue("Nombre total d'individus:");
      reportSheet.getRange("B5").setValue(dashboard.stats.global.totalIndividus);
      
      reportSheet.getRange("A6").setValue("Hommes:");
      reportSheet.getRange("B6").setValue(dashboard.stats.global.hommes);
      
      reportSheet.getRange("A7").setValue("Femmes:");
      reportSheet.getRange("B7").setValue(dashboard.stats.global.femmes);
      
      reportSheet.getRange("A8").setValue("Salaire moyen:");
      reportSheet.getRange("B8").setValue(dashboard.stats.global.salaireMoyen)
        .setNumberFormat("0.00 €");
      
      reportSheet.getRange("A9").setValue("Salaire médian:");
      reportSheet.getRange("B9").setValue(dashboard.stats.global.salaireMedian)
        .setNumberFormat("0.00 €");
      
      reportSheet.getRange("A10").setValue("Écart salarial H/F:");
      reportSheet.getRange("B10").setValue(dashboard.stats.global.ecartHommesFemmes + "%")
        .setNumberFormat("0.00%");
      
      // Section 2: Répartition par tranche d'âge
      reportSheet.getRange("A12").setValue("RÉPARTITION PAR TRANCHE D'ÂGE")
        .setFontWeight("bold")
        .setFontSize(14);
      
      reportSheet.getRange("A13:E13").setValues([["Tranche d'âge", "Total", "Hommes", "Femmes", "Salaire moyen"]]);
      
      let row = 14;
      Object.keys(dashboard.stats.tranchesAge).forEach(tranche => {
        const t = dashboard.stats.tranchesAge[tranche];
        reportSheet.getRange(`A${row}:E${row}`).setValues([
          [tranche, t.total, t.hommes, t.femmes, t.salaireMoyen]
        ]);
        reportSheet.getRange(`E${row}`).setNumberFormat("0.00 €");
        row++;
      });
      
      // Section 3: Anomalies salariales
      row += 2;
      reportSheet.getRange(`A${row}`).setValue("ANOMALIES SALARIALES")
        .setFontWeight("bold")
        .setFontSize(14);
      
      if (dashboard.anomalies.length > 0) {
        row++;
        reportSheet.getRange(`A${row}:H${row}`).setValues([
          ["Mois", "NIR", "Nom", "Prénom", "Contrat", "Ancien salaire", "Nouveau salaire", "Variation"]
        ]);
        
        dashboard.anomalies.forEach(a => {
          row++;
          reportSheet.getRange(`A${row}:H${row}`).setValues([
            [a.mois, a.nir, a.nom, a.prenom, a.contrat, a.ancienSalaire, a.nouveauSalaire, a.variation]
          ]);
          reportSheet.getRange(`F${row}:G${row}`).setNumberFormat("0.00 €");
        });
      } else {
        row++;
        reportSheet.getRange(`A${row}`).setValue("Aucune anomalie salariale détectée.")
          .setFontStyle("italic");
      }
      
      // Section 4: Arrêts de travail
      row += 2;
      reportSheet.getRange(`A${row}`).setValue("ARRÊTS DE TRAVAIL")
        .setFontWeight("bold")
        .setFontSize(14);
      
      row++;
      reportSheet.getRange(`A${row}`).setValue("Nombre total d'arrêts:");
      reportSheet.getRange(`B${row}`).setValue(dashboard.arrets.total);
      
      row++;
      reportSheet.getRange(`A${row}`).setValue("Durée moyenne des arrêts:");
      reportSheet.getRange(`B${row}`).setValue(dashboard.arrets.dureeMoyenne.toFixed(1) + " jours");
      
      row += 2;
      reportSheet.getRange(`A${row}`).setValue("Répartition par motif");
      
      row++;
      reportSheet.getRange(`A${row}:D${row}`).setValues([
        ["Motif", "Nombre", "% du total", "Durée moyenne"]
      ]);
      
      Object.keys(dashboard.arrets.parMotif).forEach(motif => {
        if (dashboard.arrets.parMotif[motif].count > 0) {
          row++;
          const pct = (dashboard.arrets.parMotif[motif].count / dashboard.arrets.total) * 100;
          reportSheet.getRange(`A${row}:D${row}`).setValues([
            [
              dashboard.arrets.parMotif[motif].libelle,
              dashboard.arrets.parMotif[motif].count,
              pct.toFixed(1) + "%",
              dashboard.arrets.parMotif[motif].dureeMoyenne.toFixed(1) + " jours"
            ]
          ]);
        }
      });
      
      // Appliquer un formatage global
      reportSheet.autoResizeColumns(1, 8);
      
      return reportSheet;
    }
  };
})();
