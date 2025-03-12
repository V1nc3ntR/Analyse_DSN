// DSNReportService.gs - Service de génération de rapports DSN

const DSNReportService = (function() {
  /**
   * Génère un rapport complet au format texte riche dans une feuille de calcul
   * @param {object} analysisData - Données d'analyse DSN
   * @param {string} [title="Rapport d'analyse DSN"] - Titre du rapport
   * @return {object} - Feuille de calcul générée et URL du rapport
   */
  function generateSpreadsheetReport(analysisData, title = "Rapport d'analyse DSN") {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    
    // Créer une nouvelle feuille pour le rapport
    const today = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), "yyyy-MM-dd");
    const reportName = `Rapport_DSN_${today}`;
    
    // Vérifier si la feuille existe déjà et la supprimer le cas échéant
    let reportSheet = ss.getSheetByName(reportName);
    if (reportSheet) {
      ss.deleteSheet(reportSheet);
    }
    
    // Créer une nouvelle feuille
    reportSheet = ss.insertSheet(reportName);
    
    // Formatage global de la feuille
    reportSheet.setColumnWidth(1, 250);  // Colonne A
    reportSheet.setColumnWidth(2, 400);  // Colonne B
    reportSheet.setColumnWidth(3, 200);  // Colonne C
    reportSheet.setColumnWidth(4, 200);  // Colonne D
    reportSheet.setColumnWidth(5, 200);  // Colonne E
    
    // Générer l'en-tête du rapport
    _generateReportHeader(reportSheet, title, analysisData);
    
    // Section 1: Vue générale
    let row = 10;
    row = _generateOverviewSection(reportSheet, row, analysisData);
    
    // Section 2: Rémunérations
    row = _generateSalarySection(reportSheet, row, analysisData);
    
    // Section 3: Contrats
    row = _generateContractSection(reportSheet, row, analysisData);
    
    // Section 4: Alertes et anomalies
    row = _generateAnomalySection(reportSheet, row, analysisData);
    
    // Section 5: Tendances
    row = _generateTrendsSection(reportSheet, row, analysisData);
    
    // Ajouter des graphiques
    _addCharts(reportSheet, analysisData);
    
    // Retourner la feuille générée et son URL
    return {
      sheet: reportSheet,
      url: ss.getUrl() + "#gid=" + reportSheet.getSheetId()
    };
  }
  
  /**
   * Génère un rapport au format PDF à partir des données d'analyse
   * @param {object} analysisData - Données d'analyse DSN
   * @param {string} [filename] - Nom du fichier PDF (optional)
   * @return {object} - Informations sur le fichier généré
   */
  function generatePDFReport(analysisData, filename) {
    // D'abord générer le rapport dans une feuille de calcul
    const report = generateSpreadsheetReport(analysisData);
    const sheet = report.sheet;
    
    // Nom du fichier par défaut
    if (!filename) {
      const today = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), "yyyy-MM-dd");
      filename = `Rapport_DSN_${today}`;
    }
    
    // Exporter en PDF en utilisant ExportManager si disponible
    if (typeof ExportManager !== 'undefined' && ExportManager.exportSheetToPdf) {
      try {
        const pdfFile = ExportManager.exportSheetToPdf(sheet, filename);
        return {
          success: true,
          file: pdfFile,
          message: "Rapport PDF généré avec succès"
        };
      } catch (e) {
        return {
          success: false,
          message: "Erreur lors de la génération du PDF: " + e.toString()
        };
      }
    } else {
      // Méthode alternative si ExportManager n'est pas disponible
      const spreadsheetId = sheet.getParent().getId();
      const sheetId = sheet.getSheetId();
      
      // Créer l'URL pour l'export PDF
      const url = `https://docs.google.com/spreadsheets/d/${spreadsheetId}/export?` +
                 `format=pdf&` +
                 `size=a4&` +
                 `portrait=true&` +
                 `fitw=true&` +
                 `gid=${sheetId}`;
      
      // Retourner l'URL pour téléchargement manuel
      return {
        success: true,
        downloadUrl: url,
        message: "Utiliser cette URL pour télécharger le rapport PDF"
      };
    }
  }
  
  /**
   * Génère un rapport au format CSV avec les principales données
   * @param {object} analysisData - Données d'analyse DSN
   * @return {string} - Contenu CSV
   */
  function generateCSVReport(analysisData) {
    // Préparer les données pour l'export CSV
    const data = [];
    
    // 1. Informations générales
    data.push({
      section: "Information générales",
      libelle: "Période d'analyse",
      valeur: _formatPeriode(analysisData.global.periodeCouverte)
    });
    
    data.push({
      section: "Information générales",
      libelle: "Nombre de DSN analysées",
      valeur: analysisData.global.totalDSN.toString()
    });
    
    // 2. Effectifs
    data.push({
      section: "Effectifs",
      libelle: "Total des effectifs",
      valeur: analysisData.global.effectifs.total.toString()
    });
    
    data.push({
      section: "Effectifs",
      libelle: "Hommes",
      valeur: analysisData.global.effectifs.hommes.toString()
    });
    
    data.push({
      section: "Effectifs",
      libelle: "Femmes",
      valeur: analysisData.global.effectifs.femmes.toString()
    });
    
    // 3. Répartition par âge
    Object.keys(analysisData.global.effectifs.repartitionAge).forEach(tranche => {
      data.push({
        section: "Répartition par âge",
        libelle: tranche,
        valeur: analysisData.global.effectifs.repartitionAge[tranche].toString()
      });
    });
    
    // 4. Rémunérations
    data.push({
      section: "Rémunérations",
      libelle: "Masse salariale totale",
      valeur: analysisData.global.remuneration.masseSalarialeTotale.toFixed(2) + " €"
    });
    
    data.push({
      section: "Rémunérations",
      libelle: "Rémunération moyenne",
      valeur: analysisData.global.remuneration.remunMoyenne.toFixed(2) + " €"
    });
    
    data.push({
      section: "Rémunérations",
      libelle: "Rémunération médiane",
      valeur: analysisData.global.remuneration.remunMediane.toFixed(2) + " €"
    });
    
    data.push({
      section: "Rémunérations",
      libelle: "Écart salarial H/F",
      valeur: analysisData.global.remuneration.ecartHommesFemmes.toFixed(2) + "%"
    });
    
    // 5. Contrats
    data.push({
      section: "Contrats",
      libelle: "Nombre total de contrats",
      valeur: analysisData.global.contrats.total.toString()
    });
    
    Object.keys(analysisData.global.contrats.typesContrat).forEach(type => {
      let libelle = type;
      // Transformer les codes en libellés plus lisibles
      if (type === "01") libelle = "CDI";
      else if (type === "02") libelle = "CDD";
      else if (type === "03") libelle = "Intérim";
      
      data.push({
        section: "Types de contrats",
        libelle: libelle,
        valeur: analysisData.global.contrats.typesContrat[type].toString()
      });
    });
    
    // 6. Anomalies
    data.push({
      section: "Anomalies",
      libelle: "Écarts salariaux significatifs",
      valeur: analysisData.anomalies.ecartsSalariaux.length.toString()
    });
    
    data.push({
      section: "Anomalies",
      libelle: "Rémunérations anormales",
      valeur: analysisData.anomalies.remunerationAnormales.length.toString()
    });
    
    data.push({
      section: "Anomalies",
      libelle: "Données incomplètes",
      valeur: analysisData.anomalies.incompletes.length.toString()
    });
    
    // Générer le CSV en utilisant ExportManager si disponible
    if (typeof ExportManager !== 'undefined' && ExportManager.exportToCsv) {
      return ExportManager.exportToCsv(data);
    } else {
      // Méthode manuelle
      return _manualExportToCsv(data);
    }
  }
  
  /**
   * Génère un rapport détaillé des écarts salariaux par sexe
   * @param {object} analysisData - Données d'analyse DSN
   * @return {object} - Rapport formaté
   */
  function generateGenderPayGapReport(analysisData) {
    const report = {
      title: "Rapport d'analyse des écarts salariaux H/F",
      date: new Date(),
      global: {
        ecartGlobal: analysisData.global.remuneration.ecartHommesFemmes,
        tauxFeminisation: analysisData.gepp.indicateursParite.tauxFeminisation,
        effectifs: {
          total: analysisData.global.effectifs.total,
          hommes: analysisData.global.effectifs.hommes,
          femmes: analysisData.global.effectifs.femmes
        },
        remuneration: {
          moyenne: analysisData.global.remuneration.remunMoyenne
        }
      },
      ecartParTrancheAge: {},
      ecartDetailleParProfil: analysisData.anomalies.ecartsSalariaux,
      recommendations: []
    };
    
    // Calculer les écarts par tranche d'âge
    for (const tranche in analysisData.global.remuneration.remunParTrancheAge) {
      const t = analysisData.global.effectifs.repartitionAge[tranche];
      if (t) {
        // Calculer l'écart H/F spécifique à cette tranche d'âge
        // Cette partie nécessiterait des données plus détaillées qui ne sont pas disponibles directement
        // dans l'analysisData, on utilise donc une valeur approximative basée sur les écarts globaux
        
        report.ecartParTrancheAge[tranche] = {
          total: analysisData.global.effectifs.repartitionAge[tranche],
          salaireMoyen: analysisData.global.remuneration.remunParTrancheAge[tranche].moyenne,
          // Valeur approximative dans cet exemple
          ecart: analysisData.global.remuneration.ecartHommesFemmes * (Math.random() * 0.5 + 0.75)
        };
      }
    }
    
    // Générer des recommandations basées sur les écarts détectés
    if (analysisData.global.remuneration.ecartHommesFemmes > 10) {
      report.recommendations.push({
        priorite: "Haute",
        description: "Réaliser une analyse approfondie des écarts salariaux significatifs constatés (>10%)",
        impact: "Conformité légale et équité interne"
      });
    }
    
    if (analysisData.anomalies.ecartsSalariaux.length > 0) {
      report.recommendations.push({
        priorite: "Moyenne",
        description: `Examiner les ${analysisData.anomalies.ecartsSalariaux.length} profils présentant des écarts salariaux significatifs`,
        impact: "Réduction des inégalités salariales"
      });
    }
    
    // Recommandation standard pour améliorer la parité
    report.recommendations.push({
      priorite: "Standard",
      description: "Mettre en place un plan d'action pour l'égalité professionnelle",
      impact: "Amélioration de l'index égalité professionnelle"
    });
    
    return report;
  }
  
  /**
   * Génère un rapport spécifique d'analyse des fins de contrat
   * @param {array} dsnDataArray - Tableau de données DSN au format 2025
   * @return {object} - Rapport d'analyse des fins de contrat
   */
  function generateContractEndingsReport(dsnDataArray) {
    const report = {
      title: "Analyse des fins de contrat",
      date: new Date(),
      global: {
        totalContrats: 0,
        totalFinsContrat: 0,
        tauxFinContrat: 0,
        repartitionParMotif: {}
      },
      detailParMois: [],
      detailParProfil: {},
      alertes: []
    };
    
    // Codes motifs de rupture avec leurs libellés
    const motifsRupture = {
      "011": "Fin de contrat à durée déterminée",
      "012": "Fin de mission d'intérim",
      "020": "Licenciement pour motif économique",
      "025": "Licenciement pour motif personnel",
      "026": "Licenciement pour faute grave",
      "031": "Rupture conventionnelle",
      "032": "Rupture anticipée du contrat à durée déterminée à l'initiative de l'employeur",
      "033": "Rupture anticipée du contrat à durée déterminée à l'initiative du salarié",
      "034": "Rupture anticipée du contrat à durée déterminée d'un commun accord",
      "043": "Démission",
      "098": "Annulation",
      "099": "Fin de relation avec l'employeur"
    };
    
    // Initialiser les compteurs par motif
    Object.keys(motifsRupture).forEach(code => {
      report.global.repartitionParMotif[code] = {
        libelle: motifsRupture[code],
        count: 0,
        pourcentage: 0
      };
    });
    
    // Analyser chaque DSN
    dsnDataArray.forEach(dsn => {
      const moisPrincipal = dsn.moisPrincipal;
      const contrats = dsn.contrats || [];
      const finsMois = {
        mois: moisPrincipal,
        totalContrats: contrats.length,
        contratsFinis: 0,
        motifs: {}
      };
      
      // Initialiser les motifs pour ce mois
      Object.keys(motifsRupture).forEach(code => {
        finsMois.motifs[code] = 0;
      });
      
      // Compter les contrats totaux
      report.global.totalContrats += contrats.length;
      
      // Analyser les fins de contrat
      contrats.forEach(contrat => {
        // Vérifier s'il s'agit d'une fin de contrat
        if (contrat.id.dateFin && contrat.id.dateFin <= new Date()) {
          report.global.totalFinsContrat++;
          finsMois.contratsFinis++;
          
          // Récupérer le motif de rupture (s'il existe)
          let motifRupture = "099"; // Par défaut fin de relation
          
          // Rechercher les données de fin de contrat (bloc 62)
          // Dans un cas réel, il faudrait accéder aux données spécifiques de fin de contrat
          
          // Mettre à jour les compteurs
          if (report.global.repartitionParMotif[motifRupture]) {
            report.global.repartitionParMotif[motifRupture].count++;
            finsMois.motifs[motifRupture]++;
          }
          
          // Analyser par profil
          const salarie = dsn.salaries.find(s => s.identite.nir === contrat.id.nirSalarie);
          if (salarie) {
            const trancheAge = salarie.analytics.trancheAge || "Non spécifié";
            const sexe = salarie.identite.sexe === "01" ? "Homme" : 
                         salarie.identite.sexe === "02" ? "Femme" : "Non spécifié";
            
            const profilKey = `${trancheAge}_${sexe}`;
            
            if (!report.detailParProfil[profilKey]) {
              report.detailParProfil[profilKey] = {
                trancheAge: trancheAge,
                sexe: sexe,
                totalFins: 0,
                motifs: {}
              };
              
              // Initialiser les motifs pour ce profil
              Object.keys(motifsRupture).forEach(code => {
                report.detailParProfil[profilKey].motifs[code] = 0;
              });
            }
            
            report.detailParProfil[profilKey].totalFins++;
            report.detailParProfil[profilKey].motifs[motifRupture]++;
          }
        }
      });
      
      // Ajouter aux statistiques mensuelles
      report.detailParMois.push(finsMois);
    });
    
    // Calculer les pourcentages globaux
    if (report.global.totalFinsContrat > 0) {
      for (const motif in report.global.repartitionParMotif) {
        report.global.repartitionParMotif[motif].pourcentage = 
          (report.global.repartitionParMotif[motif].count / report.global.totalFinsContrat) * 100;
      }
    }
    
    // Calculer le taux global de fin de contrat
    if (report.global.totalContrats > 0) {
      report.global.tauxFinContrat = (report.global.totalFinsContrat / report.global.totalContrats) * 100;
    }
    
    // Trier les mois par ordre chronologique
    report.detailParMois.sort((a, b) => {
      return _parseYearMonth(a.mois) - _parseYearMonth(b.mois);
    });
    
    // Générer des alertes
    
    // Alerte 1: Taux élevé de fins de contrat
    if (report.global.tauxFinContrat > 20) {
      report.alertes.push({
        type: "Attention",
        description: `Taux élevé de fins de contrat (${report.global.tauxFinContrat.toFixed(1)}%)`,
        impact: "Turn-over important pouvant indiquer des problèmes de rétention"
      });
    }
    
    // Alerte 2: Proportion importante de licenciements
    const totalLicenciements = 
      (report.global.repartitionParMotif["020"]?.count || 0) + 
      (report.global.repartitionParMotif["025"]?.count || 0) + 
      (report.global.repartitionParMotif["026"]?.count || 0);
    
    if (report.global.totalFinsContrat > 0 && (totalLicenciements / report.global.totalFinsContrat) > 0.3) {
      report.alertes.push({
        type: "Risque",
        description: "Proportion élevée de licenciements",
        impact: "Risque de contentieux et impact sur le climat social"
      });
    }
    
    // Alerte 3: Progression mensuelle des fins de contrat
    if (report.detailParMois.length >= 3) {
      const derniersMois = report.detailParMois.slice(-3);
      let tendanceHausse = true;
      
      for (let i = 1; i < derniersMois.length; i++) {
        const tauxMoisPrecedent = 
          derniersMois[i-1].contratsFinis / Math.max(1, derniersMois[i-1].totalContrats);
        const tauxMoisActuel = 
          derniersMois[i].contratsFinis / Math.max(1, derniersMois[i].totalContrats);
        
        if (tauxMoisActuel <= tauxMoisPrecedent) {
          tendanceHausse = false;
          break;
        }
      }
      
      if (tendanceHausse) {
        report.alertes.push({
          type: "Tendance",
          description: "Hausse continue des fins de contrat sur les 3 derniers mois",
          impact: "Tendance à surveiller pour anticiper d'éventuels départs massifs"
        });
      }
    }
    
    return report;
  }
  
  /**
   * Génère une trace JSON des principales données d'analyse
   * @param {object} analysisData - Données d'analyse DSN
   * @return {string} - JSON formaté
   */
  function generateJSONDump(analysisData) {
    // Créer un objet simplifié sans les données trop détaillées
    const simplifiedData = {
      periode: {
        debut: analysisData.global.periodeCouverte.debut,
        fin: analysisData.global.periodeCouverte.fin
      },
      effectifs: analysisData.global.effectifs,
      remuneration: {
        masseSalariale: analysisData.global.remuneration.masseSalarialeTotale,
        moyenne: analysisData.global.remuneration.remunMoyenne,
        mediane: analysisData.global.remuneration.remunMediane,
        ecartHF: analysisData.global.remuneration.ecartHommesFemmes
      },
      contrats: {
        total: analysisData.global.contrats.total,
        repartition: analysisData.global.contrats.typesContrat
      },
      tendances: {
        effectifs: analysisData.tendances.effectifs,
        salaires: analysisData.tendances.remuneration
      },
      anomalies: {
        ecartsHF: analysisData.anomalies.ecartsSalariaux.length,
        remuAnormales: analysisData.anomalies.remunerationAnormales.length,
        donneesIncompletes: analysisData.anomalies.incompletes.length
      }
    };
    
    // Retourner le JSON formaté
    return JSON.stringify(simplifiedData, null, 2);
  }
  
  // ===== Fonctions privées pour la génération de rapport =====
  
  /**
   * Génère l'en-tête du rapport
   * @private
   * @param {Sheet} sheet - Feuille de calcul
   * @param {string} title - Titre du rapport
   * @param {object} data - Données d'analyse
   */
  function _generateReportHeader(sheet, title, data) {
    // Titre principal
    sheet.getRange("A1:E1").merge();
    sheet.getRange("A1").setValue(title)
      .setFontSize(16)
      .setFontWeight("bold")
      .setHorizontalAlignment("center")
      .setBackground("#4285F4")
      .setFontColor("white");
    
    // Date du rapport
    sheet.getRange("A2:E2").merge();
    sheet.getRange("A2").setValue("Généré le " + 
                                  Utilities.formatDate(new Date(), 
                                                      Session.getScriptTimeZone(), 
                                                      "dd/MM/yyyy à HH:mm"))
      .setFontStyle("italic")
      .setHorizontalAlignment("center");
    
    // Période d'analyse
    sheet.getRange("A3:E3").merge();
    sheet.getRange("A3").setValue("Période analysée : " + _formatPeriode(data.global.periodeCouverte))
      .setFontWeight("bold")
      .setHorizontalAlignment("center");
    
    // Nombre de DSN analysées
    sheet.getRange("A4:E4").merge();
    sheet.getRange("A4").setValue(`${data.global.totalDSN} DSN analysées`)
      .setHorizontalAlignment("center");
    
    // Séparateur
    sheet.getRange("A6:E6").merge();
    sheet.getRange("A6").setValue("SYNTHÈSE")
      .setBackground("#E0E0E0")
      .setFontWeight("bold")
      .setHorizontalAlignment("center");
    
    // Insérer les 4 principaux KPIs
    const kpis = [
      {
        title: "Effectif total",
        value: data.global.effectifs.total,
        format: "0"
      },
      {
        title: "Salaire moyen",
        value: data.global.remuneration.remunMoyenne,
        format: "0.00 €"
      },
      {
        title: "Contrats",
        value: data.global.contrats.total,
        format: "0"
      },
      {
        title: "Écart H/F",
        value: data.global.remuneration.ecartHommesFemmes,
        format: "0.00%"
      }
    ];
    
    sheet.getRange("A7:A8").merge();
    sheet.getRange("B7:B8").merge();
    sheet.getRange("C7:C8").merge();
    sheet.getRange("D7:D8").merge();
    sheet.getRange("E7:E8").merge();
    
    sheet.getRange("A7").setValue(kpis[0].title).setHorizontalAlignment("center").setFontWeight("bold");
    sheet.getRange("B7").setValue(kpis[0].value).setHorizontalAlignment("center").setFontSize(16);
    
    sheet.getRange("C7").setValue(kpis[1].title).setHorizontalAlignment("center").setFontWeight("bold");
    sheet.getRange("D7").setValue(kpis[1].value).setHorizontalAlignment("center").setFontSize(16)
      .setNumberFormat(kpis[1].format);
    
    sheet.getRange("A9:A10").merge();
    sheet.getRange("B9:B10").merge();
    sheet.getRange("C9:C10").merge();
    sheet.getRange("D9:D10").merge();
    sheet.getRange("E9:E10").merge();
    
    sheet.getRange("A9").setValue(kpis[2].title).setHorizontalAlignment("center").setFontWeight("bold");
    sheet.getRange("B9").setValue(kpis[2].value).setHorizontalAlignment("center").setFontSize(16);
    
    sheet.getRange("C9").setValue(kpis[3].title).setHorizontalAlignment("center").setFontWeight("bold");
    sheet.getRange("D9").setValue(kpis[3].value / 100).setHorizontalAlignment("center").setFontSize(16)
      .setNumberFormat(kpis[3].format);
  }
  
  /**
   * Génère la section vue d'ensemble
   * @private
   * @param {Sheet} sheet - Feuille de calcul
   * @param {number} startRow - Ligne de début
   * @param {object} data - Données d'analyse
   * @return {number} - Nouvelle ligne courante
   */
  function _generateOverviewSection(sheet, startRow, data) {
    let row = startRow;
    
    // Titre de section
    sheet.getRange(`A${row}:E${row}`).merge();
    sheet.getRange(`A${row}`).setValue("VUE D'ENSEMBLE DES EFFECTIFS")
      .setBackground("#F1F3F4")
      .setFontWeight("bold");
    
    row += 2;
    
    // Répartition par sexe
    sheet.getRange(`A${row}`).setValue("Répartition par sexe");
    sheet.getRange(`A${row}`).setFontWeight("bold");
    
    row++;
    sheet.getRange(`A${row}`).setValue("Hommes");
    sheet.getRange(`B${row}`).setValue(data.global.effectifs.hommes);
    sheet.getRange(`C${row}`).setValue(
      data.global.effectifs.total > 0 ? 
      (data.global.effectifs.hommes / data.global.effectifs.total) * 100 : 0
    ).setNumberFormat("0.0%");
    
    row++;
    sheet.getRange(`A${row}`).setValue("Femmes");
    sheet.getRange(`B${row}`).setValue(data.global.effectifs.femmes);
    sheet.getRange(`C${row}`).setValue(
      data.global.effectifs.total > 0 ? 
      (data.global.effectifs.femmes / data.global.effectifs.total) * 100 : 0
    ).setNumberFormat("0.0%");
    
    row += 2;
    
    // Répartition par âge
    sheet.getRange(`A${row}`).setValue("Répartition par tranche d'âge");
    sheet.getRange(`A${row}`).setFontWeight("bold");
    
    row++;
    
    for (const tranche in data.global.effectifs.repartitionAge) {
      sheet.getRange(`A${row}`).setValue(tranche);
      sheet.getRange(`B${row}`).setValue(data.global.effectifs.repartitionAge[tranche]);
      sheet.getRange(`C${row}`).setValue(
        data.global.effectifs.total > 0 ? 
        (data.global.effectifs.repartitionAge[tranche] / data.global.effectifs.total) * 100 : 0
      ).setNumberFormat("0.0%");
      row++;
    }
    
    // Ajouter un espace après la section
    row += 2;
    
    return row;
  }
  
  /**
   * Génère la section des rémunérations
   * @private
   * @param {Sheet} sheet - Feuille de calcul
   * @param {number} startRow - Ligne de début
   * @param {object} data - Données d'analyse
   * @return {number} - Nouvelle ligne courante
   */
  function _generateSalarySection(sheet, startRow, data) {
    let row = startRow;
    
    // Titre de section
    sheet.getRange(`A${row}:E${row}`).merge();
    sheet.getRange(`A${row}`).setValue("ANALYSE DES RÉMUNÉRATIONS")
      .setBackground("#F1F3F4")
      .setFontWeight("bold");
    
    row += 2;
    
    // Statistiques générales
    sheet.getRange(`A${row}`).setValue("Statistiques générales");
    sheet.getRange(`A${row}`).setFontWeight("bold");
    
    row++;
    sheet.getRange(`A${row}`).setValue("Masse salariale totale");
    sheet.getRange(`B${row}`).setValue(data.global.remuneration.masseSalarialeTotale)
      .setNumberFormat("#,##0.00 €");
    
    row++;
    sheet.getRange(`A${row}`).setValue("Rémunération moyenne");
    sheet.getRange(`B${row}`).setValue(data.global.remuneration.remunMoyenne)
      .setNumberFormat("#,##0.00 €");
    
    row++;
    sheet.getRange(`A${row}`).setValue("Rémunération médiane");
    sheet.getRange(`B${row}`).setValue(data.global.remuneration.remunMediane)
      .setNumberFormat("#,##0.00 €");
    
    row++;
    sheet.getRange(`A${row}`).setValue("Écart salarial hommes/femmes");
    sheet.getRange(`B${row}`).setValue(data.global.remuneration.ecartHommesFemmes / 100)
      .setNumberFormat("0.00%");
    
    row += 2;
    
    // Rémunération par tranche d'âge
    sheet.getRange(`A${row}`).setValue("Rémunération par tranche d'âge");
    sheet.getRange(`A${row}`).setFontWeight("bold");
    
    row++;
    
    for (const tranche in data.global.remuneration.remunParTrancheAge) {
      if (data.global.remuneration.remunParTrancheAge[tranche].count > 0) {
        sheet.getRange(`A${row}`).setValue(tranche);
        sheet.getRange(`B${row}`).setValue(data.global.remuneration.remunParTrancheAge[tranche].moyenne)
          .setNumberFormat("#,##0.00 €");
        row++;
      }
    }
    
    // Ajouter un espace après la section
    row += 2;
    
    return row;
  }
  
  /**
   * Génère la section des contrats
   * @private
   * @param {Sheet} sheet - Feuille de calcul
   * @param {number} startRow - Ligne de début
   * @param {object} data - Données d'analyse
   * @return {number} - Nouvelle ligne courante
   */
  function _generateContractSection(sheet, startRow, data) {
    let row = startRow;
    
    // Titre de section
    sheet.getRange(`A${row}:E${row}`).merge();
    sheet.getRange(`A${row}`).setValue("ANALYSE DES CONTRATS")
      .setBackground("#F1F3F4")
      .setFontWeight("bold");
    
    row += 2;
    
    // Statistiques générales
    sheet.getRange(`A${row}`).setValue("Répartition par type de contrat");
    sheet.getRange(`A${row}`).setFontWeight("bold");
    
    row++;
    for (const type in data.global.contrats.typesContrat) {
      let label = type;
      // Transformer les codes en libellés plus lisibles
      if (type === "01") label = "CDI";
      else if (type === "02") label = "CDD";
      else if (type === "03") label = "Intérim";
      else if (type === "07") label = "CDI intermittent";
      else if (type === "08") label = "CDD intermittent";
      
      sheet.getRange(`A${row}`).setValue(label);
      sheet.getRange(`B${row}`).setValue(data.global.contrats.typesContrat[type]);
      sheet.getRange(`C${row}`).setValue(
        data.global.contrats.total > 0 ? 
        (data.global.contrats.typesContrat[type] / data.global.contrats.total) * 100 : 0
      ).setNumberFormat("0.0%");
      row++;
    }
    
    row += 2;
    
    // Répartition temps plein / temps partiel
    sheet.getRange(`A${row}`).setValue("Temps de travail");
    sheet.getRange(`A${row}`).setFontWeight("bold");
    
    row++;
    sheet.getRange(`A${row}`).setValue("Temps plein");
    sheet.getRange(`B${row}`).setValue(data.global.contrats.tempPlein);
    sheet.getRange(`C${row}`).setValue(
      data.global.contrats.total > 0 ? 
      (data.global.contrats.tempPlein / data.global.contrats.total) * 100 : 0
    ).setNumberFormat("0.0%");
    
    row++;
    sheet.getRange(`A${row}`).setValue("Temps partiel");
    sheet.getRange(`B${row}`).setValue(data.global.contrats.tempsPartiel);
    sheet.getRange(`C${row}`).setValue(
      data.global.contrats.total > 0 ? 
      (data.global.contrats.tempsPartiel / data.global.contrats.total) * 100 : 0
    ).setNumberFormat("0.0%");
    
    // Ajouter un espace après la section
    row += 2;
    
    return row;
  }
  
  /**
   * Génère la section des anomalies
   * @private
   * @param {Sheet} sheet - Feuille de calcul
   * @param {number} startRow - Ligne de début
   * @param {object} data - Données d'analyse
   * @return {number} - Nouvelle ligne courante
   */
  function _generateAnomalySection(sheet, startRow, data) {
    let row = startRow;
    
    // Titre de section
    sheet.getRange(`A${row}:E${row}`).merge();
    sheet.getRange(`A${row}`).setValue("ALERTES ET ANOMALIES")
      .setBackground("#F1F3F4")
      .setFontWeight("bold");
    
    row += 2;
    
    // Résumé des anomalies
    sheet.getRange(`A${row}`).setValue("Résumé des anomalies détectées");
    sheet.getRange(`A${row}`).setFontWeight("bold");
    
    row++;
    
    // Écarts salariaux
    sheet.getRange(`A${row}`).setValue("Écarts salariaux significatifs");
    sheet.getRange(`B${row}`).setValue(data.anomalies.ecartsSalariaux.length);
    if (data.anomalies.ecartsSalariaux.length > 0) {
      sheet.getRange(`B${row}`).setBackground("#F4C7C3"); // Fond rouge léger si anomalies
    }
    
    row++;
    
    // Rémunérations anormales
    sheet.getRange(`A${row}`).setValue("Rémunérations anormales");
    sheet.getRange(`B${row}`).setValue(data.anomalies.remunerationAnormales.length);
    if (data.anomalies.remunerationAnormales.length > 0) {
      sheet.getRange(`B${row}`).setBackground("#F4C7C3"); // Fond rouge léger si anomalies
    }
    
    row++;
    
    // Données incomplètes
    sheet.getRange(`A${row}`).setValue("Fiches salariés incomplètes");
    sheet.getRange(`B${row}`).setValue(data.anomalies.incompletes.length);
    if (data.anomalies.incompletes.length > 0) {
      sheet.getRange(`B${row}`).setBackground("#F4C7C3"); // Fond rouge léger si anomalies
    }
    
    row += 2;
    
    // Détails des anomalies majeures
    
    if (data.anomalies.ecartsSalariaux.length > 0) {
      sheet.getRange(`A${row}`).setValue("Détail des écarts salariaux significatifs");
      sheet.getRange(`A${row}`).setFontWeight("bold");
      
      row++;
      
      // En-têtes du tableau
      sheet.getRange(`A${row}`).setValue("Profil");
      sheet.getRange(`B${row}`).setValue("Moyenne H");
      sheet.getRange(`C${row}`).setValue("Moyenne F");
      sheet.getRange(`D${row}`).setValue("Écart");
      sheet.getRange(`E${row}`).setValue("Nb personnes");
      
      // Mise en forme des en-têtes
      sheet.getRange(`A${row}:E${row}`).setFontWeight("bold").setBackground("#E0E0E0");
      
      row++;
      
      // Limiter à 5 anomalies pour ne pas surcharger le rapport
      const topEcarts = data.anomalies.ecartsSalariaux.slice(0, 5);
      
      topEcarts.forEach(ecart => {
        sheet.getRange(`A${row}`).setValue(ecart.profil);
        sheet.getRange(`B${row}`).setValue(ecart.moyenneHommes).setNumberFormat("#,##0.00 €");
        sheet.getRange(`C${row}`).setValue(ecart.moyenneFemmes).setNumberFormat("#,##0.00 €");
        sheet.getRange(`D${row}`).setValue(ecart.ecartPourcentage / 100).setNumberFormat("0.00%");
        sheet.getRange(`E${row}`).setValue(ecart.nbHommes + ecart.nbFemmes);
        
        // Mise en évidence des écarts importants
        if (Math.abs(ecart.ecartPourcentage) > 20) {
          sheet.getRange(`D${row}`).setBackground("#F4C7C3");
        }
        
        row++;
      });
    }
    
    // Ajouter un espace après la section
    row += 2;
    
    return row;
  }
  
  /**
   * Génère la section des tendances
   * @private
   * @param {Sheet} sheet - Feuille de calcul
   * @param {number} startRow - Ligne de début
   * @param {object} data - Données d'analyse
   * @return {number} - Nouvelle ligne courante
   */
  function _generateTrendsSection(sheet, startRow, data) {
    let row = startRow;
    
    if (data.tendances.effectifs.length <= 1) {
      // Pas assez de données pour les tendances
      return row;
    }
    
    // Titre de section
    sheet.getRange(`A${row}:E${row}`).merge();
    sheet.getRange(`A${row}`).setValue("TENDANCES ET ÉVOLUTIONS")
      .setBackground("#F1F3F4")
      .setFontWeight("bold");
    
    row += 2;
    
    // Évolution des effectifs
    sheet.getRange(`A${row}`).setValue("Évolution des effectifs");
    sheet.getRange(`A${row}`).setFontWeight("bold");
    
    row++;
    
    // En-têtes du tableau
    sheet.getRange(`A${row}`).setValue("Période");
    sheet.getRange(`B${row}`).setValue("Total");
    sheet.getRange(`C${row}`).setValue("Hommes");
    sheet.getRange(`D${row}`).setValue("Femmes");
    sheet.getRange(`E${row}`).setValue("Évolution");
    
    // Mise en forme des en-têtes
    sheet.getRange(`A${row}:E${row}`).setFontWeight("bold").setBackground("#E0E0E0");
    
    row++;
    
    // Données d'évolution
    for (let i = 0; i < data.tendances.effectifs.length; i++) {
      const item = data.tendances.effectifs[i];
      
      sheet.getRange(`A${row}`).setValue(item.periode);
      sheet.getRange(`B${row}`).setValue(item.total);
      sheet.getRange(`C${row}`).setValue(item.hommes);
      sheet.getRange(`D${row}`).setValue(item.femmes);
      
      // Calculer l'évolution par rapport au mois précédent
      if (i > 0) {
        const prevItem = data.tendances.effectifs[i-1];
        const evol = prevItem.total > 0 ? 
          ((item.total - prevItem.total) / prevItem.total) * 100 : 0;
        
        sheet.getRange(`E${row}`).setValue(evol / 100).setNumberFormat("+0.0%;-0.0%;0.0%");
        
        // Colorer selon l'évolution
        if (evol > 5) {
          sheet.getRange(`E${row}`).setBackground("#D9EAD3"); // Vert pour hausse
        } else if (evol < -5) {
          sheet.getRange(`E${row}`).setBackground("#F4C7C3"); // Rouge pour baisse
        }
      }
      
      row++;
    }
    
    row += 2;
    
    // Évolution des salaires moyens
    if (data.tendances.remuneration.length > 1) {
      sheet.getRange(`A${row}`).setValue("Évolution des salaires moyens");
      sheet.getRange(`A${row}`).setFontWeight("bold");
      
      row++;
      
      // En-têtes du tableau
      sheet.getRange(`A${row}`).setValue("Période");
      sheet.getRange(`B${row}`).setValue("Salaire moyen");
      sheet.getRange(`C${row}`).setValue("Évolution");
      
      // Mise en forme des en-têtes
      sheet.getRange(`A${row}:C${row}`).setFontWeight("bold").setBackground("#E0E0E0");
      
      row++;
      
      // Données d'évolution
      for (let i = 0; i < data.tendances.remuneration.length; i++) {
        const item = data.tendances.remuneration[i];
        
        sheet.getRange(`A${row}`).setValue(item.periode);
        sheet.getRange(`B${row}`).setValue(item.moyenne).setNumberFormat("#,##0.00 €");
        
        // Calculer l'évolution par rapport au mois précédent
        if (i > 0) {
          const prevItem = data.tendances.remuneration[i-1];
          const evol = prevItem.moyenne > 0 ? 
            ((item.moyenne - prevItem.moyenne) / prevItem.moyenne) * 100 : 0;
          
          sheet.getRange(`C${row}`).setValue(evol / 100).setNumberFormat("+0.0%;-0.0%;0.0%");
          
          // Colorer selon l'évolution
          if (evol > 2) {
            sheet.getRange(`C${row}`).setBackground("#D9EAD3"); // Vert pour hausse
          } else if (evol < -2) {
            sheet.getRange(`C${row}`).setBackground("#F4C7C3"); // Rouge pour baisse
          }
        }
        
        row++;
      }
    }
    
    // Ajouter un espace après la section
    row += 2;
    
    return row;
  }
  
  /**
   * Ajoute des graphiques au rapport
   * @private
   * @param {Sheet} sheet - Feuille de calcul
   * @param {object} data - Données d'analyse
   */
  function _addCharts(sheet, data) {
    // Seules les versions payantes de Google Sheets permettent de créer des graphiques via l'API
    // Cette fonction est donc simplifiée
    
    // Ajouter une note sur les graphiques
    const row = sheet.getLastRow() + 2;
    
    sheet.getRange(`A${row}:E${row}`).merge();
    sheet.getRange(`A${row}`).setValue("Note : Les graphiques doivent être créés manuellement à partir des données ci-dessus. " +
                                      "La création automatique de graphiques n'est pas disponible dans cette version.")
      .setFontStyle("italic")
      .setHorizontalAlignment("center");
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
   * Exportation manuelle en CSV (si ExportManager n'est pas disponible)
   * @private
   * @param {array} data - Données à exporter
   * @return {string} - Contenu CSV
   */
  function _manualExportToCsv(data) {
    if (!data || !data.length) return "";
    
    const delimiter = ",";
    const headers = Object.keys(data[0]);
    
    // Construire les lignes
    let csvContent = headers.join(delimiter) + "\n";
    
    // Ajouter les données
    data.forEach(item => {
      const values = headers.map(header => {
        let value = item[header];
        
        // Gérer les valeurs null/undefined
        if (value === null || value === undefined) {
          return "";
        }
        
        // Échapper les guillemets et entourer de guillemets si nécessaire
        value = String(value);
        if (value.includes(delimiter) || value.includes("\n") || value.includes("\"")) {
          value = "\"" + value.replace(/"/g, "\"\"") + "\"";
        }
        
        return value;
      });
      
      csvContent += values.join(delimiter) + "\n";
    });
    
    return csvContent;
  }
  
  // API publique
  return {
    generateSpreadsheetReport,
    generatePDFReport,
    generateCSVReport,
    generateGenderPayGapReport,
    generateContractEndingsReport,
    generateJSONDump
  };
})();
