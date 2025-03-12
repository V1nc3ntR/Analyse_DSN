// DSNAnalysisService.gs - Service d'analyse des données DSN

const DSNAnalysisService = (function() {
  /**
   * Analyse complète d'un ensemble de DSN
   * @param {array} dsnDataArray - Tableau de données DSN au format 2025
   * @return {object} - Rapport d'analyse
   */
  function analyzeAllDSN(dsnDataArray) {
    // Initialiser le rapport d'analyse
    const report = {
      // Statistiques générales
      global: {
        totalDSN: dsnDataArray.length,
        periodeCouverte: {
          debut: null,
          fin: null
        },
        effectifs: {
          total: 0,
          hommes: 0,
          femmes: 0,
          repartitionAge: {
            "moins de 30 ans": 0,
            "30 à 39 ans": 0,
            "40 à 49 ans": 0,
            "50 ans et plus": 0
          }
        },
        remuneration: {
          masseSalarialeTotale: 0,
          remunMoyenne: 0,
          remunMediane: 0,
          remunMin: null,
          remunMax: null,
          ecartHommesFemmes: 0,
          remunParTrancheAge: {}
        },
        contrats: {
          total: 0,
          typesContrat: {},
          tempsPartiel: 0,
          tempPlein: 0
        },
        absences: {
          tauxGlobal: 0,
          parTypeAbsence: {},
          coutEstime: 0
        }
      },
      
      // Tendances
      tendances: {
        effectifs: [],
        remuneration: [],
        absences: []
      },
      
      // Écarts et anomalies
      anomalies: {
        remunerationAnormales: [],
        ecartsSalariaux: [],
        incompletes: []
      },
      
      // Indicateurs GEPP (Gestion des Emplois et des Parcours Professionnels)
      gepp: {
        pyramideAges: {},
        ancienneteParService: {},
        indicateursParite: {
          tauxFeminisation: 0,
          indexEgalite: null,
          ecartSalarial: 0
        }
      }
    };
    
    // Stocker tous les salariés et contrats pour analyse
    const allSalaries = {};
    const allContrats = {};
    const allRemunerations = [];
    
    // Tableau pour le calcul de la médiane
    const allSalaires = [];
    const salairesFemmes = [];
    const salairesHommes = [];
    
    // Déterminer la période couverte
    dsnDataArray.forEach(dsn => {
      if (dsn.moisPrincipal) {
        const moisDate = _parseYearMonth(dsn.moisPrincipal);
        
        if (!report.global.periodeCouverte.debut || 
            moisDate < report.global.periodeCouverte.debut) {
          report.global.periodeCouverte.debut = moisDate;
        }
        
        if (!report.global.periodeCouverte.fin || 
            moisDate > report.global.periodeCouverte.fin) {
          report.global.periodeCouverte.fin = moisDate;
        }
      }
    });
    
    // Premier passage pour collecter tous les salariés uniques par NIR
    dsnDataArray.forEach(dsn => {
      dsn.salaries.forEach(salarie => {
        // Ajouter le salarié s'il n'existe pas déjà
        if (!allSalaries[salarie.identite.nir]) {
          allSalaries[salarie.identite.nir] = salarie;
        }
      });
      
      dsn.contrats.forEach(contrat => {
        // Ajouter le contrat s'il n'existe pas déjà
        const contratKey = `${contrat.id.nirSalarie}_${contrat.id.numeroContrat}`;
        if (!allContrats[contratKey]) {
          allContrats[contratKey] = contrat;
        }
        
        // Ajouter les rémunérations à la liste globale
        if (contrat.remunerations && contrat.remunerations.length > 0) {
          allRemunerations.push(...contrat.remunerations);
        }
      });
    });
    
    // Calculer les effectifs
    report.global.effectifs.total = Object.keys(allSalaries).length;
    
    for (const nir in allSalaries) {
      const salarie = allSalaries[nir];
      
      // Répartition hommes/femmes
      if (salarie.identite.sexe === "01") { // Homme
        report.global.effectifs.hommes++;
      } else if (salarie.identite.sexe === "02") { // Femme
        report.global.effectifs.femmes++;
      }
      
      // Répartition par âge
      if (salarie.analytics.trancheAge) {
        report.global.effectifs.repartitionAge[salarie.analytics.trancheAge]++;
      }
    }
    
    // Calculer la parité
    if (report.global.effectifs.total > 0) {
      report.gepp.indicateursParite.tauxFeminisation = 
        (report.global.effectifs.femmes / report.global.effectifs.total) * 100;
    }
    
    // Analyser les contrats
    report.global.contrats.total = Object.keys(allContrats).length;
    
    for (const key in allContrats) {
      const contrat = allContrats[key];
      
      // Répartition par type de contrat
      const typeContrat = contrat.caracteristiques.nature || "Non spécifié";
      if (!report.global.contrats.typesContrat[typeContrat]) {
        report.global.contrats.typesContrat[typeContrat] = 0;
      }
      report.global.contrats.typesContrat[typeContrat]++;
      
      // Répartition temps plein / temps partiel
      if (contrat.tempsTravail.modalite === "20") { // Temps partiel
        report.global.contrats.tempsPartiel++;
      } else if (contrat.tempsTravail.modalite === "10") { // Temps plein
        report.global.contrats.tempPlein++;
      }
    }
    
    // Calculer les statistiques de rémunération
    let totalMasseSalariale = 0;
    let salairesMoyensProfil = {};
    
    allRemunerations.forEach(remu => {
      // On ne considère que les rémunérations principales
      if (remu.type === "001" || remu.type === "002" || remu.type === "003") {
        if (remu.montant) {
          const montant = remu.montant;
          const salarie = allSalaries[_findSalarieByContrat(allContrats, remu.numeroContrat)];
          
          totalMasseSalariale += montant;
          allSalaires.push(montant);
          
          // Mettre à jour min et max
          if (report.global.remuneration.remunMin === null || montant < report.global.remuneration.remunMin) {
            report.global.remuneration.remunMin = montant;
          }
          
          if (report.global.remuneration.remunMax === null || montant > report.global.remuneration.remunMax) {
            report.global.remuneration.remunMax = montant;
          }
          
          // Répartition par sexe pour calcul d'écart
          if (salarie) {
            if (salarie.identite.sexe === "01") { // Homme
              salairesHommes.push(montant);
            } else if (salarie.identite.sexe === "02") { // Femme
              salairesFemmes.push(montant);
            }
            
            // Répartition par tranche d'âge
            if (salarie.analytics.trancheAge) {
              const tranche = salarie.analytics.trancheAge;
              if (!report.global.remuneration.remunParTrancheAge[tranche]) {
                report.global.remuneration.remunParTrancheAge[tranche] = {
                  total: 0,
                  count: 0,
                  moyenne: 0
                };
              }
              
              report.global.remuneration.remunParTrancheAge[tranche].total += montant;
              report.global.remuneration.remunParTrancheAge[tranche].count++;
            }
            
            // Regrouper par profil (pour détecter les écarts)
            const profil = _buildSalarieProfil(salarie, allContrats);
            if (profil) {
              if (!salairesMoyensProfil[profil]) {
                salairesMoyensProfil[profil] = {
                  total: 0,
                  count: 0,
                  hommes: { total: 0, count: 0 },
                  femmes: { total: 0, count: 0 }
                };
              }
              
              salairesMoyensProfil[profil].total += montant;
              salairesMoyensProfil[profil].count++;
              
              if (salarie.identite.sexe === "01") { // Homme
                salairesMoyensProfil[profil].hommes.total += montant;
                salairesMoyensProfil[profil].hommes.count++;
              } else if (salarie.identite.sexe === "02") { // Femme
                salairesMoyensProfil[profil].femmes.total += montant;
                salairesMoyensProfil[profil].femmes.count++;
              }
            }
          }
        }
      }
    });
    
    // Calculer masse salariale et moyennes
    report.global.remuneration.masseSalarialeTotale = totalMasseSalariale;
    
    if (allSalaires.length > 0) {
      report.global.remuneration.remunMoyenne = totalMasseSalariale / allSalaires.length;
      
      // Calculer la médiane
      allSalaires.sort((a, b) => a - b);
      const midIndex = Math.floor(allSalaires.length / 2);
      report.global.remuneration.remunMediane = allSalaires[midIndex];
      
      // Calculer les moyennes par tranche d'âge
      for (const tranche in report.global.remuneration.remunParTrancheAge) {
        const data = report.global.remuneration.remunParTrancheAge[tranche];
        if (data.count > 0) {
          data.moyenne = data.total / data.count;
        }
      }
    }
    
    // Calculer l'écart salarial hommes/femmes
    if (salairesHommes.length > 0 && salairesFemmes.length > 0) {
      const moyenneHommes = salairesHommes.reduce((a, b) => a + b, 0) / salairesHommes.length;
      const moyenneFemmes = salairesFemmes.reduce((a, b) => a + b, 0) / salairesFemmes.length;
      
      if (moyenneHommes > 0) {
        report.global.remuneration.ecartHommesFemmes = 
          ((moyenneHommes - moyenneFemmes) / moyenneHommes) * 100;
        
        report.gepp.indicateursParite.ecartSalarial = report.global.remuneration.ecartHommesFemmes;
      }
    }
    
    // Détecter les écarts salariaux significatifs par profil
    for (const profil in salairesMoyensProfil) {
      const data = salairesMoyensProfil[profil];
      
      if (data.hommes.count > 0 && data.femmes.count > 0) {
        const moyenneHommes = data.hommes.total / data.hommes.count;
        const moyenneFemmes = data.femmes.total / data.femmes.count;
        
        if (moyenneHommes > 0) {
          const ecart = ((moyenneHommes - moyenneFemmes) / moyenneHommes) * 100;
          
          // Si l'écart est significatif (plus de 5%)
          if (Math.abs(ecart) > 5) {
            report.anomalies.ecartsSalariaux.push({
              profil: profil,
              moyenneHommes: moyenneHommes,
              moyenneFemmes: moyenneFemmes,
              ecartPourcentage: ecart,
              nbHommes: data.hommes.count,
              nbFemmes: data.femmes.count
            });
          }
        }
      }
    }
    
    // Calculer la pyramide des âges pour la GEPP
    for (const nir in allSalaries) {
      const salarie = allSalaries[nir];
      
      if (salarie.analytics.age) {
        // Grouper par tranches de 5 ans
        const tranche = Math.floor(salarie.analytics.age / 5) * 5;
        const trancheLabel = `${tranche}-${tranche+4} ans`;
        
        if (!report.gepp.pyramideAges[trancheLabel]) {
          report.gepp.pyramideAges[trancheLabel] = {
            total: 0,
            hommes: 0,
            femmes: 0
          };
        }
        
        report.gepp.pyramideAges[trancheLabel].total++;
        
        if (salarie.identite.sexe === "01") { // Homme
          report.gepp.pyramideAges[trancheLabel].hommes++;
        } else if (salarie.identite.sexe === "02") { // Femme
          report.gepp.pyramideAges[trancheLabel].femmes++;
        }
      }
    }
    
    // Détecter les anomalies de rémunération
    // Exemple: écart de plus de 30% par rapport à la moyenne du profil
    for (const profil in salairesMoyensProfil) {
      const data = salairesMoyensProfil[profil];
      
      if (data.count > 1) { // Au moins 2 personnes dans le profil
        const moyenneProfil = data.total / data.count;
        
        // Parcourir toutes les rémunérations pour ce profil
        allRemunerations.forEach(remu => {
          if ((remu.type === "001" || remu.type === "002" || remu.type === "003") && remu.montant) {
            const nirSalarie = _findSalarieByContrat(allContrats, remu.numeroContrat);
            const salarie = allSalaries[nirSalarie];
            
            if (salarie && _buildSalarieProfil(salarie, allContrats) === profil) {
              const ecartPct = ((remu.montant - moyenneProfil) / moyenneProfil) * 100;
              
              // Si l'écart est anormalement élevé
              if (Math.abs(ecartPct) > 30) {
                report.anomalies.remunerationAnormales.push({
                  profil: profil,
                  nir: nirSalarie,
                  nom: salarie.identite.nom,
                  prenom: salarie.identite.prenoms,
                  remuneration: remu.montant,
                  moyenneProfil: moyenneProfil,
                  ecartPourcentage: ecartPct,
                  periode: remu.periode.debut ? 
                    Utilities.formatDate(remu.periode.debut, Session.getScriptTimeZone(), "yyyy-MM") : null
                });
              }
            }
          }
        });
      }
    }
    
    // Rechercher les données incomplètes
    for (const nir in allSalaries) {
      const salarie = allSalaries[nir];
      const champsManquants = _verifierChampsSalarie(salarie);
      
      if (champsManquants.length > 0) {
        report.anomalies.incompletes.push({
          nir: nir,
          nom: salarie.identite.nom,
          prenom: salarie.identite.prenoms,
          champsManquants: champsManquants
        });
      }
    }
    
    // Analyser les tendances
    if (dsnDataArray.length > 1) {
      // Tri par mois
      dsnDataArray.sort((a, b) => {
        return _parseYearMonth(a.moisPrincipal) - _parseYearMonth(b.moisPrincipal);
      });
      
      // Analyser les évolutions mois par mois
      dsnDataArray.forEach(dsn => {
        const periode = dsn.moisPrincipal;
        
        // Tendance effectifs
        report.tendances.effectifs.push({
          periode: periode,
          total: dsn.salaries.length,
          hommes: dsn.salaries.filter(s => s.identite.sexe === "01").length,
          femmes: dsn.salaries.filter(s => s.identite.sexe === "02").length
        });
        
        // Tendance rémunération
        const remusMois = [];
        dsn.contrats.forEach(contrat => {
          if (contrat.remunerations) {
            contrat.remunerations.forEach(remu => {
              if ((remu.type === "001" || remu.type === "002" || remu.type === "003") && remu.montant) {
                remusMois.push(remu.montant);
              }
            });
          }
        });
        
        if (remusMois.length > 0) {
          const totalMois = remusMois.reduce((a, b) => a + b, 0);
          const moyenneMois = totalMois / remusMois.length;
          
          report.tendances.remuneration.push({
            periode: periode,
            moyenne: moyenneMois,
            total: totalMois
          });
        }
        
        // Tendance absences
        const totalJoursAbsence = dsn.arrets.reduce((total, arret) => {
          if (arret.dateDebut && arret.dateFin) {
            return total + _calculateDaysDifference(arret.dateDebut, arret.dateFin);
          }
          return total;
        }, 0);
        
        const totalSalaries = dsn.salaries.length;
        const tauxAbsence = totalSalaries > 0 ? 
          (totalJoursAbsence / (totalSalaries * 30)) * 100 : 0; // Approximation de 30 jours par mois
        
        report.tendances.absences.push({
          periode: periode,
          totalJours: totalJoursAbsence,
          tauxAbsence: tauxAbsence
        });
      });
    }
    
    return report;
  }
  
  /**
   * Analyse les données d'une seule DSN
   * @param {object} dsnData - Données d'une DSN au format 2025
   * @return {object} - Rapport d'analyse simplifié
   */
  function analyzeSingleDSN(dsnData) {
    return analyzeAllDSN([dsnData]);
  }
  
  /**
   * Génère un tableau de bord synthétique pour un ensemble de DSN
   * @param {array} dsnDataArray - Tableau de données DSN au format 2025
   * @return {object} - Tableau de bord
   */
  function generateDashboard(dsnDataArray) {
    // Effectuer l'analyse complète
    const analysis = analyzeAllDSN(dsnDataArray);
    
    // Extraire les KPIs et visualisations clés pour le tableau de bord
    const dashboard = {
      kpis: {
        effectifs: analysis.global.effectifs.total,
        masseSalariale: analysis.global.remuneration.masseSalarialeTotale,
        remunMoyenne: analysis.global.remuneration.remunMoyenne,
        tauxFeminisation: analysis.gepp.indicateursParite.tauxFeminisation,
        ecartSalarial: analysis.global.remuneration.ecartHommesFemmes
      },
      
      graphiques: {
        effectifs: _prepareDashboardData(analysis.tendances.effectifs, 'total', 'période'),
        repartitionSexe: _prepareGenderDistributionData(analysis.global.effectifs),
        pyramideAges: _preparePyramidAgeData(analysis.gepp.pyramideAges),
        evolutionRemuneration: _prepareDashboardData(analysis.tendances.remuneration, 'moyenne', 'période'),
        typesContrat: _prepareContractTypesData(analysis.global.contrats.typesContrat)
      },
      
      alertes: {
        ecartsSignificatifs: analysis.anomalies.ecartsSalariaux.length,
        remunerationAnormales: analysis.anomalies.remunerationAnormales.length,
        donneesMissingCount: analysis.anomalies.incompletes.length
      },
      
      periode: {
        debut: analysis.global.periodeCouverte.debut,
        fin: analysis.global.periodeCouverte.fin,
        formatted: _formatPeriode(analysis.global.periodeCouverte)
      }
    };
    
    return dashboard;
  }
  
  /**
   * Recherche des salariés selon des critères spécifiques
   * @param {array} dsnDataArray - Tableau de données DSN au format 2025
   * @param {object} criteria - Critères de recherche
   * @return {array} - Salariés correspondants aux critères
   */
  function searchSalaries(dsnDataArray, criteria) {
    // Collecter tous les salariés uniques
    const allSalaries = {};
    
    dsnDataArray.forEach(dsn => {
      dsn.salaries.forEach(salarie => {
        if (!allSalaries[salarie.identite.nir]) {
          allSalaries[salarie.identite.nir] = salarie;
        }
      });
    });
    
    // Filtrer selon les critères
    const results = [];
    
    for (const nir in allSalaries) {
      const salarie = allSalaries[nir];
      let match = true;
      
      if (criteria) {
        // Filtre par nom
        if (criteria.nom && !_matchesFilter(salarie.identite.nom, criteria.nom)) {
          match = false;
        }
        
        // Filtre par prénom
        if (criteria.prenom && !_matchesFilter(salarie.identite.prenoms, criteria.prenom)) {
          match = false;
        }
        
        // Filtre par sexe
        if (criteria.sexe && salarie.identite.sexe !== criteria.sexe) {
          match = false;
        }
        
        // Filtre par tranche d'âge
        if (criteria.trancheAge && salarie.analytics.trancheAge !== criteria.trancheAge) {
          match = false;
        }
        
        // Filtre par date de naissance
        if (criteria.dateNaissanceMin && (!salarie.identite.dateNaissance || 
            salarie.identite.dateNaissance < criteria.dateNaissanceMin)) {
          match = false;
        }
        
        if (criteria.dateNaissanceMax && (!salarie.identite.dateNaissance || 
            salarie.identite.dateNaissance > criteria.dateNaissanceMax)) {
          match = false;
        }
      }
      
      if (match) {
        results.push(salarie);
      }
    }
    
    return results;
  }
  
  /**
   * Recherche des contrats selon des critères spécifiques
   * @param {array} dsnDataArray - Tableau de données DSN au format 2025
   * @param {object} criteria - Critères de recherche
   * @return {array} - Contrats correspondants aux critères
   */
  function searchContrats(dsnDataArray, criteria) {
    // Collecter tous les contrats uniques
    const allContrats = {};
    
    dsnDataArray.forEach(dsn => {
      dsn.contrats.forEach(contrat => {
        const key = `${contrat.id.nirSalarie}_${contrat.id.numeroContrat}`;
        if (!allContrats[key]) {
          allContrats[key] = contrat;
        }
      });
    });
    
    // Filtrer selon les critères
    const results = [];
    
    for (const key in allContrats) {
      const contrat = allContrats[key];
      let match = true;
      
      if (criteria) {
        // Filtre par numéro de contrat
        if (criteria.numeroContrat && !_matchesFilter(contrat.id.numeroContrat, criteria.numeroContrat)) {
          match = false;
        }
        
        // Filtre par NIR de salarié
        if (criteria.nirSalarie && !_matchesFilter(contrat.id.nirSalarie, criteria.nirSalarie)) {
          match = false;
        }
        
        // Filtre par nature de contrat
        if (criteria.nature && contrat.caracteristiques.nature !== criteria.nature) {
          match = false;
        }
        
        // Filtre par dates de début
        if (criteria.dateDebutMin && (!contrat.id.dateDebut || 
            contrat.id.dateDebut < criteria.dateDebutMin)) {
          match = false;
        }
        
        if (criteria.dateDebutMax && (!contrat.id.dateDebut || 
            contrat.id.dateDebut > criteria.dateDebutMax)) {
          match = false;
        }
        
        // Filtre par type de temps de travail
        if (criteria.modaliteTempsTravail && contrat.tempsTravail.modalite !== criteria.modaliteTempsTravail) {
          match = false;
        }
      }
      
      if (match) {
        results.push(contrat);
      }
    }
    
    return results;
  }
  
  /**
   * Analyse l'évolution du salaire pour un salarié donné
   * @param {array} dsnDataArray - Tableau de données DSN au format 2025
   * @param {string} nir - NIR du salarié
   * @return {object} - Analyse de l'évolution salariale
   */
  function analyzeSalaryEvolution(dsnDataArray, nir) {
    const evolution = {
      nir: nir,
      identite: {
        nom: "",
        prenom: ""
      },
      historique: [],
      stats: {
        evolutionTotale: 0,
        evolutionMoyenne: 0,
        evolutionMediane: 0,
        volatilite: 0
      }
    };
    
    // Collecter toutes les rémunérations par mois
    const remuParMois = {};
    
    dsnDataArray.forEach(dsn => {
      const mois = dsn.moisPrincipal;
      
      // Récupérer le salarié
      const salarie = dsn.salaries.find(s => s.identite.nir === nir);
      if (salarie) {
        evolution.identite.nom = salarie.identite.nom;
        evolution.identite.prenom = salarie.identite.prenoms;
        
        // Récupérer les contrats du salarié
        const contrats = dsn.contrats.filter(c => c.id.nirSalarie === nir);
        
        // Calculer la rémunération totale pour ce mois
        let remuTotal = 0;
        
        contrats.forEach(contrat => {
          if (contrat.remunerations) {
            contrat.remunerations.forEach(remu => {
              if ((remu.type === "001" || remu.type === "002" || remu.type === "003") && remu.montant) {
                remuTotal += remu.montant;
              }
            });
          }
        });
        
        if (remuTotal > 0) {
          remuParMois[mois] = remuTotal;
        }
      }
    });
    
    // Trier les mois par ordre chronologique
    const moisTries = Object.keys(remuParMois).sort();
    
    // Construire l'historique d'évolution
    if (moisTries.length > 1) {
      for (let i = 1; i < moisTries.length; i++) {
        const moisPrec = moisTries[i-1];
        const moisCurr = moisTries[i];
        
        const remuPrec = remuParMois[moisPrec];
        const remuCurr = remuParMois[moisCurr];
        
        const evolutionPct = ((remuCurr - remuPrec) / remuPrec) * 100;
        
        evolution.historique.push({
          moisDebut: moisPrec,
          moisFin: moisCurr,
          remuDebut: remuPrec,
          remuFin: remuCurr,
          evolutionPct: evolutionPct
        });
      }
      
      // Calculer les statistiques
      if (evolution.historique.length > 0) {
        // Évolution totale
        const remuInitiale = remuParMois[moisTries[0]];
        const remuFinale = remuParMois[moisTries[moisTries.length - 1]];
        
        evolution.stats.evolutionTotale = ((remuFinale - remuInitiale) / remuInitiale) * 100;
        
        // Évolution moyenne
        const totalEvolution = evolution.historique.reduce((sum, item) => sum + item.evolutionPct, 0);
        evolution.stats.evolutionMoyenne = totalEvolution / evolution.historique.length;
        
        // Évolution médiane
        const evolutions = evolution.historique.map(item => item.evolutionPct).sort((a, b) => a - b);
        const midIndex = Math.floor(evolutions.length / 2);
        evolution.stats.evolutionMediane = evolutions[midIndex];
        
        // Volatilité (écart-type des évolutions)
        const moyenne = evolution.stats.evolutionMoyenne;
        const sumSquaredDiff = evolutions.reduce((sum, val) => sum + Math.pow(val - moyenne, 2), 0);
        evolution.stats.volatilite = Math.sqrt(sumSquaredDiff / evolutions.length);
      }
    }
    
    return evolution;
  }

  /**
   * Analyse détaillée des absences pour un salarié donné
   * @param {array} dsnDataArray - Tableau de données DSN au format 2025
   * @param {string} nir - NIR du salarié
   * @return {object} - Analyse des absences
   */
  function analyzeAbsences(dsnDataArray, nir) {
    const analyse = {
      nir: nir,
      identite: {
        nom: "",
        prenom: ""
      },
      absences: [],
      stats: {
        totalJours: 0,
        moyenneParArret: 0,
        repartitionParMotif: {},
        tauxAbsenteisme: 0,
        evolutionMensuelle: []
      }
    };
    
    // Récupérer toutes les informations du salarié
    let salarieTrouve = false;
    let joursActivite = 0;
    const absencesParMois = {};
    
    dsnDataArray.forEach(dsn => {
      const mois = dsn.moisPrincipal;
      const moisDate = _parseYearMonth(mois);
      const joursDansMois = new Date(moisDate.getFullYear(), moisDate.getMonth() + 1, 0).getDate();
      
      // Récupérer le salarié
      const salarie = dsn.salaries.find(s => s.identite.nir === nir);
      if (salarie) {
        if (!salarieTrouve) {
          analyse.identite.nom = salarie.identite.nom;
          analyse.identite.prenom = salarie.identite.prenoms;
          salarieTrouve = true;
        }
        
        // Ajouter les jours d'activité potentielle
        joursActivite += joursDansMois;
        
        // Récupérer les arrêts du salarié
        const arrets = dsn.arrets.filter(a => a.nirSalarie === nir);
        
        // Initialiser le compteur pour ce mois
        if (!absencesParMois[mois]) {
          absencesParMois[mois] = 0;
        }
        
        // Ajouter les arrêts
        arrets.forEach(arret => {
          if (arret.dateDebut && arret.dateFin) {
            const duree = _calculateDaysDifference(arret.dateDebut, arret.dateFin);
            
            // Ajouter à la liste des absences
            analyse.absences.push({
              motif: arret.motif,
              dateDebut: arret.dateDebut,
              dateFin: arret.dateFin,
              duree: duree,
              motifReprise: arret.motifReprise
            });
            
            // Mettre à jour les statistiques
            analyse.stats.totalJours += duree;
            absencesParMois[mois] += duree;
            
            // Mettre à jour la répartition par motif
            if (!analyse.stats.repartitionParMotif[arret.motif]) {
              analyse.stats.repartitionParMotif[arret.motif] = {
                count: 0,
                jours: 0
              };
            }
            
            analyse.stats.repartitionParMotif[arret.motif].count++;
            analyse.stats.repartitionParMotif[arret.motif].jours += duree;
          }
        });
      }
    });
    
    // Calculer les moyennes et pourcentages
    if (analyse.absences.length > 0) {
      analyse.stats.moyenneParArret = analyse.stats.totalJours / analyse.absences.length;
    }
    
    if (joursActivite > 0) {
      analyse.stats.tauxAbsenteisme = (analyse.stats.totalJours / joursActivite) * 100;
    }
    
    // Calculer l'évolution mensuelle
    for (const mois in absencesParMois) {
      const jours = absencesParMois[mois];
      const moisDate = _parseYearMonth(mois);
      const joursDansMois = new Date(moisDate.getFullYear(), moisDate.getMonth() + 1, 0).getDate();
      
      analyse.stats.evolutionMensuelle.push({
        mois: mois,
        jours: jours,
        tauxMensuel: (jours / joursDansMois) * 100
      });
    }
    
    // Trier les évolutions par mois
    analyse.stats.evolutionMensuelle.sort((a, b) => {
      return _parseYearMonth(a.mois) - _parseYearMonth(b.mois);
    });
    
    return analyse;
  }
  
  // ===== Fonctions utilitaires privées =====
  
  /**
   * Parse une date au format "YYYY-MM"
   * @private
   * @param {string} yearMonth - Date au format "YYYY-MM"
   * @return {Date} - Objet Date correspondant
   */
  function _parseYearMonth(yearMonth) {
    if (!yearMonth) return null;
    
    const parts = yearMonth.split("-");
    if (parts.length === 2) {
      const year = parseInt(parts[0], 10);
      const month = parseInt(parts[1], 10) - 1; // Les mois commencent à 0 en JavaScript
      
      return new Date(year, month, 1);
    }
    
    return null;
  }
  
  /**
   * Trouve le NIR du salarié associé à un contrat
   * @private
   * @param {object} allContrats - Map des contrats
   * @param {string} numeroContrat - Numéro de contrat
   * @return {string} - NIR du salarié associé
   */
  function _findSalarieByContrat(allContrats, numeroContrat) {
    for (const key in allContrats) {
      const contrat = allContrats[key];
      if (contrat.id.numeroContrat === numeroContrat) {
        return contrat.id.nirSalarie;
      }
    }
    
    return null;
  }
  
  /**
   * Construit un identifiant de profil pour un salarié
   * @private
   * @param {object} salarie - Objet salarié
   * @param {object} allContrats - Map des contrats
   * @return {string} - Identifiant de profil
   */
  function _buildSalarieProfil(salarie, allContrats) {
    if (!salarie) return null;
    
    let contratType = "";
    let csp = "";
    
    // Rechercher le contrat principal
    if (salarie.contrats && salarie.contrats.length > 0) {
      for (const key in allContrats) {
        const contrat = allContrats[key];
        if (contrat.id.nirSalarie === salarie.identite.nir) {
          contratType = contrat.caracteristiques.nature || "";
          csp = contrat.caracteristiques.codeCSP || "";
          break;
        }
      }
    }
    
    // Construire le profil sous forme de chaîne
    return `${csp}_${contratType}_${salarie.analytics.trancheAge || ""}`;
  }
  
  /**
   * Vérifie si une valeur correspond à un filtre (texte)
   * @private
   * @param {string} value - Valeur à vérifier
   * @param {string} filter - Filtre à appliquer
   * @return {boolean} - True si la valeur correspond au filtre
   */
  function _matchesFilter(value, filter) {
    if (!value) return false;
    if (!filter) return true;
    
    return value.toLowerCase().includes(filter.toLowerCase());
  }
  
  /**
   * Vérifie si un salarié a des champs manquants obligatoires
   * @private
   * @param {object} salarie - Objet salarié
   * @return {array} - Liste des champs manquants
   */
  function _verifierChampsSalarie(salarie) {
    const champsObligatoires = [
      { chemin: "identite.nir", nom: "NIR" },
      { chemin: "identite.nom", nom: "Nom" },
      { chemin: "identite.prenoms", nom: "Prénoms" },
      { chemin: "identite.sexe", nom: "Sexe" },
      { chemin: "identite.dateNaissance", nom: "Date de naissance" }
    ];
    
    const manquants = [];
    
    champsObligatoires.forEach(champ => {
      const valeur = _getObjectValueByPath(salarie, champ.chemin);
      if (!valeur) {
        manquants.push(champ.nom);
      }
    });
    
    return manquants;
  }
  
  /**
   * Récupère une valeur dans un objet à partir d'un chemin en notation pointée
   * @private
   * @param {object} obj - Objet à parcourir
   * @param {string} path - Chemin en notation pointée
   * @return {*} - Valeur correspondante ou undefined
   */
  function _getObjectValueByPath(obj, path) {
    if (!obj || !path) return undefined;
    
    const keys = path.split('.');
    let current = obj;
    
    for (let i = 0; i < keys.length; i++) {
      if (current === null || current === undefined) {
        return undefined;
      }
      
      current = current[keys[i]];
    }
    
    return current;
  }
  
  /**
   * Calcule le nombre de jours entre deux dates
   * @private
   * @param {Date} startDate - Date de début
   * @param {Date} endDate - Date de fin
   * @return {number} - Nombre de jours
   */
  function _calculateDaysDifference(startDate, endDate) {
    const diffTime = Math.abs(endDate - startDate);
    return Math.ceil(diffTime / (1000 * 60 * 60 * 24)) + 1; // +1 pour inclure le jour de début
  }
  
  /**
   * Prépare les données pour les visualisations du tableau de bord
   * @private
   * @param {array} dataArray - Tableau de données
   * @param {string} valueKey - Clé de la valeur à extraire
   * @param {string} labelKey - Clé du libellé
   * @return {object} - Données formatées pour visualisation
   */
  function _prepareDashboardData(dataArray, valueKey, labelKey) {
    if (!dataArray || !dataArray.length) {
      return { labels: [], values: [] };
    }
    
    return {
      labels: dataArray.map(item => item[labelKey]),
      values: dataArray.map(item => item[valueKey])
    };
  }
  
  /**
   * Prépare les données pour la visualisation de la répartition par sexe
   * @private
   * @param {object} effectifs - Données d'effectifs
   * @return {object} - Données formatées pour visualisation
   */
  function _prepareGenderDistributionData(effectifs) {
    return {
      labels: ["Hommes", "Femmes"],
      values: [effectifs.hommes, effectifs.femmes]
    };
  }
  
  /**
   * Prépare les données pour la visualisation de la pyramide des âges
   * @private
   * @param {object} pyramideAges - Données de pyramide des âges
   * @return {object} - Données formatées pour visualisation
   */
  function _preparePyramidAgeData(pyramideAges) {
    const tranches = Object.keys(pyramideAges).sort();
    
    return {
      labels: tranches,
      hommes: tranches.map(tranche => pyramideAges[tranche].hommes),
      femmes: tranches.map(tranche => pyramideAges[tranche].femmes)
    };
  }
  
  /**
   * Prépare les données pour la visualisation des types de contrat
   * @private
   * @param {object} typesContrat - Données des types de contrat
   * @return {object} - Données formatées pour visualisation
   */
  function _prepareContractTypesData(typesContrat) {
    const labels = [];
    const values = [];
    
    for (const type in typesContrat) {
      let label = type;
      // Transformer les codes en libellés plus lisibles
      if (type === "01") label = "CDI";
      else if (type === "02") label = "CDD";
      else if (type === "03") label = "Intérim";
      else if (type === "07") label = "Contrat à durée indéterminée intermittent";
      else if (type === "08") label = "Contrat à durée déterminée intermittent";
      else if (type === "09") label = "Contrat de travail temporaire";
      else if (type === "10") label = "Contrat de travail saisonnier";
      
      labels.push(label);
      values.push(typesContrat[type]);
    }
    
    return { labels, values };
  }
  
  /**
   * Formate la période de couverture
   * @private
   * @param {object} periode - Période avec debut et fin
   * @return {string} - Période formatée
   */
  function _formatPeriode(periode) {
    if (!periode.debut || !periode.fin) {
      return "Période inconnue";
    }
    
    return Utilities.formatDate(periode.debut, Session.getScriptTimeZone(), "MMMM yyyy") + 
           " à " + 
           Utilities.formatDate(periode.fin, Session.getScriptTimeZone(), "MMMM yyyy");
  }
  
  // API publique
  return {
    analyzeAllDSN,
    analyzeSingleDSN,
    generateDashboard,
    searchSalaries,
    searchContrats,
    analyzeSalaryEvolution,
    analyzeAbsences
  };
})();
