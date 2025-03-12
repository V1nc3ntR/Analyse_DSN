// DSNParser2025.gs - Fonctions de parsing des DSN selon le cahier technique 2025.1

const DSNParser2025 = (function() {
  // Constantes pour les regex et patterns
  const PATTERNS = {
    BLOC_RUBRIQUE: /^(S\d{2}\.G00\.\d{2,3})\.(\d{1,3})$/,
    DSN_DATE: /^(\d{2})(\d{2})(\d{4})$/,
    DSN_DECIMAL: /^-?\d+(\.\d+)?$/
  };
  
  // Constantes pour les codes de bloc DSN
  const BLOCS = {
    SALARIE: "S21.G00.30",
    CONTRAT: "S21.G00.40",
    REMUNERATION: "S21.G00.51",
    ARRET: "S21.G00.60",
    FIN_CONTRAT: "S21.G00.62",
    LIEU_TRAVAIL: "S21.G00.85",
    ANCIENNETE: "S21.G00.86",
    ACTIVITE_PARTIELLE: "S21.G00.95"
  };
  
  /**
   * Parse le contenu d'une DSN et construit le modèle de données complet
   * @param {string} dsnText - Contenu textuel de la DSN
   * @return {object} - Modèle de données DSN structuré
   */
  function parseDSNContent(dsnText) {
    // Extraire les blocs de rubriques
    const blocsMap = _extractBlocsFromText(dsnText);
    
    // Construire le modèle DSN
    return _buildDsnModel(blocsMap);
  }
  
  /**
   * Extrait les blocs de rubriques à partir du texte DSN
   * @private
   * @param {string} dsnText - Contenu textuel de la DSN
   * @return {object} - Map des blocs de rubriques
   */
  function _extractBlocsFromText(dsnText) {
    const lines = dsnText.split(/\r?\n/);
    const blocMap = {};
    
    lines.forEach(function(line) {
      line = line.trim();
      if(!line) return;
      
      const parts = line.split(",");
      if(parts.length < 2) return;
      
      const left = parts[0].trim();
      let right = parts.slice(1).join(",").trim();
      
      // Enlever les guillemets simples si présents
      if(right.startsWith("'") && right.endsWith("'")) {
        right = right.substring(1, right.length - 1);
      }
      
      const match = left.match(PATTERNS.BLOC_RUBRIQUE);
      if(!match) return;
      
      const bloc = match[1];
      const rubrique = match[2];
      
      // Initialiser le tableau du bloc si nécessaire
      if(!blocMap[bloc]) {
        blocMap[bloc] = [];
      }
      
      // Ajouter la rubrique avec sa valeur
      blocMap[bloc].push({
        rubrique: rubrique,
        valeur: right
      });
    });
    
    return blocMap;
  }
  
  /**
   * Construit le modèle DSN à partir des blocs extraits
   * @private
   * @param {object} blocsMap - Map des blocs de rubriques
   * @return {object} - Modèle de données DSN structuré
   */
  function _buildDsnModel(blocsMap) {
    const model = {
      moisPrincipal: _extractMoisPrincipal(blocsMap),
      etablissement: _extractEtablissement(blocsMap),
      salaries: _extractSalaries(blocsMap),
      contrats: _extractContrats(blocsMap),
      remunerations: _extractRemunerations(blocsMap),
      arrets: _extractArrets(blocsMap)
    };
    
    // Associer les contrats et arrêts aux salariés
    _linkEntities(model);
    
    // Enrichir le modèle avec des données analytiques
    _enrichModelWithAnalytics(model);
    
    // NOUVEAUTÉ: Enrichir avec les données pour l'Index d'Égalité Professionnelle
    _enrichModelWithEqualityData(model);
    
    return model;
  }
  
  /**
   * Extrait le mois principal de la DSN
   * @private
   * @param {object} blocsMap - Map des blocs de rubriques
   * @return {string} - Mois principal au format "YYYY-MM"
   */
  function _extractMoisPrincipal(blocsMap) {
    const decl = blocsMap["S20.G00.05"];
    
    if (decl) {
      const annee = _findRubriqueValue(decl, "005");
      const mois = _findRubriqueValue(decl, "006");
      
      if (annee && mois) {
        const mm = mois.length === 1 ? "0" + mois : mois;
        return annee + "-" + mm;
      }
    }
    
    // CORRECTION: Si non trouvé, utiliser la date de l'année courante au lieu de today
    // pour éviter les erreurs de nommage des feuilles
    const today = new Date();
    return Utilities.formatDate(today, Session.getScriptTimeZone(), "yyyy-MM");
  }
  
  /**
   * Extrait les informations de l'établissement
   * @private
   * @param {object} blocsMap - Map des blocs de rubriques
   * @return {object} - Informations sur l'établissement
   */
  function _extractEtablissement(blocsMap) {
    const etab = blocsMap["S21.G00.11"] || [];
    const adresse = blocsMap["S21.G00.13"] || [];
    
    return {
      siren: _findRubriqueValue(etab, "001") || "",
      nic: _findRubriqueValue(etab, "002") || "",
      siret: _findRubriqueValue(etab, "001") && _findRubriqueValue(etab, "002") 
             ? _findRubriqueValue(etab, "001") + _findRubriqueValue(etab, "002") 
             : "",
      raisonSociale: _findRubriqueValue(etab, "003") || "",
      adresse: {
        voie: _findRubriqueValue(adresse, "003") || "",
        codePostal: _findRubriqueValue(adresse, "004") || "",
        commune: _findRubriqueValue(adresse, "005") || "",
        pays: _findRubriqueValue(adresse, "007") || ""
      }
    };
  }
  
  /**
   * Extrait les données des salariés
   * @private
   * @param {object} blocsMap - Map des blocs de rubriques
   * @return {array} - Tableau des salariés
   */
  function _extractSalaries(blocsMap) {
    const s2130 = blocsMap[BLOCS.SALARIE] || [];
    const salaries = [];
    let currentSalarie = null;
    
    s2130.forEach(function(r) {
      // Nouveau salarié si on trouve le NIR (rubrique 001)
      if (r.rubrique === "001") {
        if (currentSalarie) {
          salaries.push(currentSalarie);
        }
        currentSalarie = DSNModels.createSalarieModel();
        currentSalarie.identite.nir = r.valeur;
      } else if (currentSalarie) {
        // Assigner les valeurs aux bons champs du modèle de salarié
        _assignSalarieValue(currentSalarie, r.rubrique, r.valeur);
      }
    });
    
    // Ajouter le dernier salarié
    if (currentSalarie) {
      salaries.push(currentSalarie);
    }
    
    return salaries;
  }
  
  /**
   * Assigne une valeur de rubrique au bon champ du modèle salarié
   * @private
   * @param {object} salarie - Modèle de salarié
   * @param {string} rubrique - Numéro de rubrique
   * @param {string} valeur - Valeur de la rubrique
   */
  function _assignSalarieValue(salarie, rubrique, valeur) {
    switch (rubrique) {
      // Identité
      case "001": salarie.identite.nir = valeur; break;
      case "002": salarie.identite.nom = valeur; break;
      case "003": salarie.identite.nomUsage = valeur; break;
      case "004": salarie.identite.prenoms = valeur; break;
      case "005": salarie.identite.sexe = valeur; break;
      case "006": salarie.identite.dateNaissance = _parseDSNDate(valeur); break;
      case "007": salarie.identite.lieuNaissance = valeur; break;
      case "014": salarie.identite.codeDeptNaissance = valeur; break;
      case "015": salarie.identite.codePaysNaissance = valeur; break;
      case "019": salarie.identite.matricule = valeur; break;
      case "020": salarie.identite.ntt = valeur; break;
      
      // Coordonnées - Adresse
      case "008": salarie.coordonnees.adresse.voie = valeur; break;
      case "009": salarie.coordonnees.adresse.codePostal = valeur; break;
      case "010": salarie.coordonnees.adresse.localite = valeur; break;
      case "011": salarie.coordonnees.adresse.codePays = valeur; break;
      case "016": salarie.coordonnees.adresse.complement1 = valeur; break;
      case "017": salarie.coordonnees.adresse.complement2 = valeur; break;
      case "018": salarie.coordonnees.email = valeur; break;
      
      // Situation professionnelle
      case "022": salarie.situation.statutEtranger = valeur; break;
      case "023": salarie.situation.cumulEmploiRetraite = valeur; break;
      case "024": salarie.situation.regimeMaladie = valeur; break;
      case "025": salarie.situation.regimeVieillesse = valeur; break;
      case "029": salarie.situation.paysResidence = valeur; break;
      
      // Formation
      case "027": salarie.formation.niveauFormation = valeur; break;
      case "028": salarie.formation.niveauDiplome = valeur; break;
    }
  }
  
  /**
   * Extrait les données des contrats
   * @private
   * @param {object} blocsMap - Map des blocs de rubriques
   * @return {array} - Tableau des contrats
   */
  function _extractContrats(blocsMap) {
    const s2140 = blocsMap[BLOCS.CONTRAT] || [];
    const contrats = [];
    let currentContrat = null;
    
    s2140.forEach(function(r) {
      // Nouveau contrat si on trouve le numéro de contrat (rubrique 009)
      if (r.rubrique === "009") {
        if (currentContrat) {
          contrats.push(currentContrat);
        }
        currentContrat = DSNModels.createContratModel();
        currentContrat.id.numeroContrat = r.valeur;
      } else if (currentContrat) {
        // Assigner les valeurs aux bons champs du modèle de contrat
        _assignContratValue(currentContrat, r.rubrique, r.valeur);
      }
    });
    
    // Ajouter le dernier contrat
    if (currentContrat) {
      contrats.push(currentContrat);
    }
    
    // Enrichir les contrats avec les informations des lieux de travail
    _enrichContratsWithLieuxTravail(contrats, blocsMap);
    
    // Enrichir les contrats avec les informations d'ancienneté
    _enrichContratsWithAnciennete(contrats, blocsMap);
    
    return contrats;
  }
  
  /**
   * Assigne une valeur de rubrique au bon champ du modèle contrat
   * @private
   * @param {object} contrat - Modèle de contrat
   * @param {string} rubrique - Numéro de rubrique
   * @param {string} valeur - Valeur de la rubrique
   */
  function _assignContratValue(contrat, rubrique, valeur) {
    switch (rubrique) {
      // Identification
      case "009": contrat.id.numeroContrat = valeur; break;
      case "001": contrat.id.dateDebut = _parseDSNDate(valeur); break;
      case "010": contrat.id.dateFin = _parseDSNDate(valeur); break;
      
      // Caractéristiques principales
      case "007": contrat.caracteristiques.nature = valeur; break;
      case "002": contrat.caracteristiques.statutSalarie = valeur; break;
      case "026": contrat.caracteristiques.statutEmploi = valeur; break;
      case "021": contrat.caracteristiques.motifRecours = valeur; break;
      case "006": contrat.caracteristiques.libelleEmploi = valeur; break;
      case "004": 
        contrat.caracteristiques.codeCSP = valeur; 
        // NOUVEAUTÉ: Remplir aussi la catégorie de poste pour l'égalité
        contrat.classificationEgalite.categoriePoste = _determinerCategorieCSP(valeur);
        break;
      case "005": contrat.caracteristiques.complementCSP = valeur; break;
      case "008": contrat.caracteristiques.dispositifPublic = valeur; break;
      case "073": contrat.caracteristiques.complementDispositif = valeur; break;
      
      // Temps de travail
      case "013": contrat.tempsTravail.modalite = valeur; break;
      case "011": contrat.tempsTravail.uniteMesure = valeur; break;
      case "012": contrat.tempsTravail.quotite = _parseNumber(valeur); break;
      case "043": contrat.tempsTravail.tauxTempsPartiel = _parseNumber(valeur); break;
      
      // Classification et rémunération
      case "017": contrat.classification.conventionCollective = valeur; break;
      case "018": contrat.classification.codeRegimeMaladie = valeur; break;
      case "020": contrat.classification.codeRegimeVieillesse = valeur; break;
      case "039": contrat.classification.codeRegimeAT = valeur; break;
      case "040": contrat.classification.codeRisqueAT = valeur; break;
      case "041": contrat.classification.positionConvention = valeur; break;
      case "048": 
        contrat.classification.niveauRemuneration = valeur; 
        // NOUVEAUTÉ: Remplir aussi le niveau pour l'égalité
        contrat.classificationEgalite.niveau = valeur;
        break;
      case "049": contrat.classification.echelon = valeur; break;
      case "050": 
        contrat.classification.coefficient = valeur; 
        // NOUVEAUTÉ: Remplir aussi le coefficient hiérarchique pour l'égalité
        contrat.classificationEgalite.coefficientHierarchique = valeur;
        break;
      
      // Lieu de travail
      case "019": contrat.lieuTravail.identifiant = valeur; break;
      
      // Données complémentaires
      case "047": contrat.complementaire.complementSante = valeur; break;
      case "042": contrat.complementaire.codeStatutApecita = valeur; break;
      case "023": contrat.complementaire.tauxDeduction = _parseNumber(valeur); break;
      case "024": contrat.complementaire.travailleurEtranger = valeur; break;
      case "025": contrat.complementaire.motifExclusion = valeur; break;
      case "078": contrat.complementaire.dispositifAide = valeur; break;
      case "079": contrat.complementaire.dateAdhesion = _parseDSNDate(valeur); break;
      case "080": contrat.complementaire.dateDenonciation = _parseDSNDate(valeur); break;
      
      // Statut BOETH pour le salarié associé
      case "072": 
        // À stocker temporairement, sera associé au salarié plus tard
        contrat.boeth = valeur;
        break;
    }
  }
  
  /**
   * Enrichit les contrats avec les informations des lieux de travail
   * @private
   * @param {array} contrats - Tableau des contrats
   * @param {object} blocsMap - Map des blocs de rubriques
   */
  function _enrichContratsWithLieuxTravail(contrats, blocsMap) {
    const lieuxTravail = blocsMap[BLOCS.LIEU_TRAVAIL] || [];
    
    // Regrouper les lieux de travail par identifiant
    const lieuxMap = {};
    let currentLieu = null;
    let currentId = null;
    
    lieuxTravail.forEach(function(r) {
      if (r.rubrique === "001") {
        currentId = r.valeur;
        currentLieu = {
          identifiant: currentId,
          libelle: "",
          adresse: {
            voie: "",
            codePostal: "",
            localite: "",
            codeDepartement: "",
            codePays: ""
          },
          codeAPET: "",
          codeINSEE: ""
        };
        lieuxMap[currentId] = currentLieu;
      } else if (currentLieu) {
        // Assigner les valeurs
        switch (r.rubrique) {
          case "002": currentLieu.libelle = r.valeur; break;
          case "003": currentLieu.adresse.voie = r.valeur; break;
          case "004": currentLieu.adresse.codePostal = r.valeur; break;
          case "005": currentLieu.adresse.localite = r.valeur; break;
          case "006": currentLieu.adresse.codeDepartement = r.valeur; break;
          case "007": currentLieu.adresse.codePays = r.valeur; break;
          case "009": currentLieu.codeAPET = r.valeur; break;
          case "010": currentLieu.codeINSEE = r.valeur; break;
        }
      }
    });
    
    // Associer les lieux de travail aux contrats
    contrats.forEach(function(contrat) {
      const lieuId = contrat.lieuTravail.identifiant;
      if (lieuId && lieuxMap[lieuId]) {
        contrat.lieuTravail = lieuxMap[lieuId];
      }
    });
  }
  
  /**
   * Enrichit les contrats avec les informations d'ancienneté
   * @private
   * @param {array} contrats - Tableau des contrats
   * @param {object} blocsMap - Map des blocs de rubriques
   */
  function _enrichContratsWithAnciennete(contrats, blocsMap) {
    const anciennetes = blocsMap[BLOCS.ANCIENNETE] || [];
    
    // Regrouper les anciennetés par contrat
    const ancMap = {};
    let currentAnciennete = null;
    let currentNumeroContrat = "";
    
    // Parcourir les rubriques pour trouver les informations d'ancienneté
    let index = 0;
    while (index < anciennetes.length) {
      // Rechercher le lien avec le contrat (décalage possible dans l'ordre)
      let trouve = false;
      for (let i = index; i < Math.min(index + 10, anciennetes.length); i++) {
        if (anciennetes[i].rubrique === "006") {
          currentNumeroContrat = anciennetes[i].valeur;
          currentAnciennete = {
            debutPeriodeReference: null,
            finPeriodeReference: null,
            uniteMesure: "",
            valeur: null
          };
          ancMap[currentNumeroContrat] = currentAnciennete;
          trouve = true;
          index = i + 1;
          break;
        }
      }
      
      if (!trouve) {
        // Si on n'a pas trouvé de référence contrat, on passe à la rubrique suivante
        index++;
        continue;
      }
      
      // Traiter les autres rubriques de l'ancienneté
      while (index < anciennetes.length && 
             !(anciennetes[index].rubrique === "006")) {
        const r = anciennetes[index];
        switch (r.rubrique) {
          case "001": currentAnciennete.debutPeriodeReference = _parseDSNDate(r.valeur); break;
          case "002": currentAnciennete.finPeriodeReference = _parseDSNDate(r.valeur); break;
          case "003": currentAnciennete.uniteMesure = r.valeur; break;
          case "004": currentAnciennete.valeur = _parseNumber(r.valeur); break;
        }
        index++;
      }
    }
    
    // Associer les anciennetés aux contrats
    contrats.forEach(function(contrat) {
      const numContrat = contrat.id.numeroContrat;
      if (numContrat && ancMap[numContrat]) {
        contrat.anciennete = ancMap[numContrat];
      }
    });
  }
  
  /**
   * Extrait les données des rémunérations
   * @private
   * @param {object} blocsMap - Map des blocs de rubriques
   * @return {array} - Tableau des rémunérations
   */
  function _extractRemunerations(blocsMap) {
    const s2151 = blocsMap[BLOCS.REMUNERATION] || [];
    const remunerations = [];
    let currentRemuneration = null;
    
    s2151.forEach(function(r) {
      // Nouvelle rémunération si on trouve le numéro de contrat (rubrique 010)
      if (r.rubrique === "010") {
        if (currentRemuneration) {
          remunerations.push(currentRemuneration);
        }
        currentRemuneration = DSNModels.createRemunerationModel();
        currentRemuneration.numeroContrat = r.valeur;
      } else if (currentRemuneration) {
        // Assigner les valeurs
        switch (r.rubrique) {
          case "001": currentRemuneration.periode.debut = _parseDSNDate(r.valeur); break;
          case "002": currentRemuneration.periode.fin = _parseDSNDate(r.valeur); break;
          case "011": currentRemuneration.type = r.valeur; break;
          case "013": currentRemuneration.montant = _parseNumber(r.valeur); break;
          case "014": currentRemuneration.tauxRemuneration = _parseNumber(r.valeur); break;
          case "012": currentRemuneration.nbHeures = _parseNumber(r.valeur); break;
          case "015": currentRemuneration.uniteMesure = r.valeur; break;
        }
      }
    });
    
    // Ajouter la dernière rémunération
    if (currentRemuneration) {
      remunerations.push(currentRemuneration);
    }
    
    return remunerations;
  }
  
  /**
   * Extrait les données des arrêts de travail
   * @private
   * @param {object} blocsMap - Map des blocs de rubriques
   * @return {array} - Tableau des arrêts de travail
   */
  function _extractArrets(blocsMap) {
    const s2160 = blocsMap[BLOCS.ARRET] || [];
    const arrets = [];
    let currentArret = null;
    
    s2160.forEach(function(r) {
      // Nouvel arrêt si on trouve le motif (rubrique 001)
      if (r.rubrique === "001") {
        if (currentArret) {
          arrets.push(currentArret);
        }
        currentArret = DSNModels.createArretTravailModel();
        currentArret.motif = r.valeur;
        
        // NOUVEAUTÉ: Détecter le type d'arrêt pour l'Index d'Égalité Professionnelle
        _detectArretType(currentArret);
      } else if (currentArret) {
        // Assigner les valeurs
        switch (r.rubrique) {
          case "002": currentArret.dateDebut = _parseDSNDate(r.valeur); break;
          case "003": currentArret.dateFin = _parseDSNDate(r.valeur); break;
          case "010": currentArret.dateReprise = _parseDSNDate(r.valeur); break;
          case "011": currentArret.motifReprise = r.valeur; break;
          case "004": currentArret.subrogation.debut = _parseDSNDate(r.valeur); break;
          case "005": currentArret.subrogation.fin = _parseDSNDate(r.valeur); break;
          case "007": currentArret.subrogation.montant = _parseNumber(r.valeur); break;
          // Lien avec le contrat et salarié - sera résolu plus tard
        }
      }
    });
    
    // Ajouter le dernier arrêt
    if (currentArret) {
      arrets.push(currentArret);
    }
    
    return arrets;
  }
  
  /**
   * Associe les entités entre elles (contrats aux salariés, etc.)
   * @private
   * @param {object} model - Modèle de données DSN
   */
  function _linkEntities(model) {
    const salariesByNir = {};
    
    // Créer un index des salariés par NIR
    model.salaries.forEach(function(salarie) {
      salariesByNir[salarie.identite.nir] = salarie;
    });
    
    // Associer les contrats aux salariés
    model.contrats.forEach(function(contrat) {
      // Recherche du salarié lié à ce contrat
      let salarieLie = null;
      
      // Stratégie 1: Utiliser les références directes si elles existent
      if (contrat.id.nirSalarie && salariesByNir[contrat.id.nirSalarie]) {
        salarieLie = salariesByNir[contrat.id.nirSalarie];
      } 
      // Stratégie 2: Utiliser la logique métier (premier salarié sans contrat par exemple)
      else {
        for (const nir in salariesByNir) {
          const salarie = salariesByNir[nir];
          // Pour cet exemple, on associe simplement au premier salarié trouvé
          // Dans une implémentation réelle, il faudrait utiliser une logique plus robuste
          salarieLie = salarie;
          break;
        }
      }
      
      if (salarieLie) {
        // Mettre à jour la référence dans le contrat
        contrat.id.nirSalarie = salarieLie.identite.nir;
        
        // Ajouter ce contrat à la liste des contrats du salarié
        salarieLie.contrats.push(contrat.id.numeroContrat);
        
        // Transférer le statut BOETH du contrat vers le salarié si présent
        if (contrat.hasOwnProperty('boeth') && contrat.boeth) {
          salarieLie.situation.boeth = contrat.boeth;
          delete contrat.boeth;
        }
      }
    });
    
    // Associer les rémunérations aux contrats
    model.contrats.forEach(function(contrat) {
      contrat.remunerations = model.remunerations.filter(function(remu) {
        return remu.numeroContrat === contrat.id.numeroContrat;
      });
    });
    
    // Associer les arrêts aux salariés et aux contrats
    model.arrets.forEach(function(arret) {
      // Pour cet exemple, on associe l'arrêt au premier contrat du premier salarié
      // Dans une implémentation réelle, il faudrait utiliser les liens explicites dans la DSN
      if (model.salaries.length > 0) {
        const salarie = model.salaries[0];
        arret.nirSalarie = salarie.identite.nir;
        salarie.arrets.push(arret);
        
        if (salarie.contrats.length > 0 && model.contrats.length > 0) {
          const contratId = salarie.contrats[0];
          const contrat = model.contrats.find(c => c.id.numeroContrat === contratId);
          if (contrat) {
            arret.numeroContrat = contrat.id.numeroContrat;
            contrat.arretsLies.push(arret);
          }
        }
      }
    });
  }
  
  /**
   * Enrichit le modèle avec des données analytiques
   * @private
   * @param {object} model - Modèle de données DSN
   */
  function _enrichModelWithAnalytics(model) {
    // Calculer les analytiques pour chaque salarié
    model.salaries.forEach(function(salarie) {
      // Calculer l'âge
      if (salarie.identite.dateNaissance) {
        const today = new Date();
        salarie.analytics.age = _calculateAge(salarie.identite.dateNaissance, today);
        salarie.analytics.trancheAge = _getTrancheAge(salarie.analytics.age);
      }
      
      // Calculer l'ancienneté de l'entreprise
      if (salarie.contrats.length > 0) {
        // Rechercher le contrat le plus ancien
        let dateDebutPlusAncienne = null;
        
        for (const numContrat of salarie.contrats) {
          const contrat = model.contrats.find(c => c.id.numeroContrat === numContrat);
          if (contrat && contrat.id.dateDebut) {
            if (!dateDebutPlusAncienne || contrat.id.dateDebut < dateDebutPlusAncienne) {
              dateDebutPlusAncienne = contrat.id.dateDebut;
            }
          }
        }
        
        if (dateDebutPlusAncienne) {
          const today = new Date();
          salarie.analytics.ancienneteEntreprise = _calculateMonthsDifference(dateDebutPlusAncienne, today);
          
          // NOUVEAUTÉ: Mettre à jour la date d'entrée dans l'entreprise
          salarie.carriere.dateEntree = dateDebutPlusAncienne;
        }
      }
      
      // Analyser les absences
      salarie.analytics.tempsAbsence.total = 0;
      
      for (const arret of model.arrets.filter(a => a.nirSalarie === salarie.identite.nir)) {
        if (arret.dateDebut && arret.dateFin) {
          const dureeArret = _calculateDaysDifference(arret.dateDebut, arret.dateFin);
          salarie.analytics.tempsAbsence.total += dureeArret;
          
          // Catégoriser par motif
          switch (arret.motif) {
            case "01": salarie.analytics.tempsAbsence.parMotif.maladie += dureeArret; break;
            case "02": 
            case "09": salarie.analytics.tempsAbsence.parMotif.maternite += dureeArret; break;
            case "04":
            case "05": 
            case "06": salarie.analytics.tempsAbsence.parMotif.accident += dureeArret; break;
            default: salarie.analytics.tempsAbsence.parMotif.autres += dureeArret; break;
          }
        }
      }
      
      // Calculer les évolutions salariales
      if (salarie.contrats.length > 0) {
        for (const numContrat of salarie.contrats) {
          const contrat = model.contrats.find(c => c.id.numeroContrat === numContrat);
          if (contrat && contrat.remunerations.length > 0) {
            // Trier les rémunérations par date
            const remus = contrat.remunerations.sort((a, b) => {
              return a.periode.debut - b.periode.debut;
            });
            
            // Utiliser seulement les principales (type 001, 002 ou 003)
            const remusPrincipales = remus.filter(r => 
              r.type === "001" || r.type === "002" || r.type === "003"
            );
            
            if (remusPrincipales.length > 0) {
              // Calculer l'évolution
              for (let i = 1; i < remusPrincipales.length; i++) {
                const remuPrev = remusPrincipales[i-1];
                const remuCurr = remusPrincipales[i];
                
                if (remuPrev.montant && remuCurr.montant) {
                  const evolution = {
                    periode: Utilities.formatDate(remuCurr.periode.debut, Session.getScriptTimeZone(), "yyyy-MM"),
                    tauxEvolution: ((remuCurr.montant - remuPrev.montant) / remuPrev.montant) * 100,
                    salaire: remuCurr.montant
                  };
                  
                  salarie.analytics.evolutionSalaire.push(evolution);
                }
              }
              
              // Calculer les moyennes
              const totalBrut = remusPrincipales.reduce((sum, r) => sum + (r.montant || 0), 0);
              salarie.analytics.salairesMoyens.brut = totalBrut / remusPrincipales.length;
              
              // Pour le net, il faudrait avoir accès aux versements individuels (S21.G00.50)
              // qui ne sont pas traités dans cet exemple
            }
          }
        }
      }
    });
    
    // Calculer les analytiques pour chaque contrat
    model.contrats.forEach(function(contrat) {
      // Calculer la durée du contrat
      if (contrat.id.dateDebut) {
        const dateFin = contrat.id.dateFin || new Date();
        contrat.analytics.dureeContrat = _calculateDaysDifference(contrat.id.dateDebut, dateFin);
      }
      
      // Calculer le coût total, salaire initial et final
      if (contrat.remunerations.length > 0) {
        // Trier les rémunérations par date
        const remus = contrat.remunerations.sort((a, b) => {
          return a.periode.debut - b.periode.debut;
        });
        
        // Utiliser seulement les principales (type 001, 002 ou 003)
        const remusPrincipales = remus.filter(r => 
          r.type === "001" || r.type === "002" || r.type === "003"
        );
        
        if (remusPrincipales.length > 0) {
          // Salaire initial et final
          contrat.analytics.salaireInitial = remusPrincipales[0].montant;
          contrat.analytics.salaireFinal = remusPrincipales[remusPrincipales.length - 1].montant;
          
          // NOUVEAUTÉ: Remplir les données d'évolution de rémunération
          contrat.evolutionRemuneration.remunerationInitiale = contrat.analytics.salaireInitial;
          contrat.evolutionRemuneration.remunerationActuelle = contrat.analytics.salaireFinal;
          
          // Évolution en pourcentage
          if (contrat.analytics.salaireInitial && contrat.analytics.salaireFinal) {
            contrat.analytics.evolutionPourcentage = 
              ((contrat.analytics.salaireFinal - contrat.analytics.salaireInitial) / 
               contrat.analytics.salaireInitial) * 100;
            
            // NOUVEAUTÉ: Remplir le pourcentage d'évolution
            contrat.evolutionRemuneration.pourcentageEvolution = contrat.analytics.evolutionPourcentage;
          }
          
          // Coût total
          contrat.analytics.coutTotal = remusPrincipales.reduce((sum, r) => sum + (r.montant || 0), 0);
        }
      }
      
      // Calculer le taux d'absentéisme
      if (contrat.arretsLies.length > 0 && contrat.analytics.dureeContrat) {
        const joursAbsence = contrat.arretsLies.reduce((sum, arret) => {
          if (arret.dateDebut && arret.dateFin) {
            return sum + _calculateDaysDifference(arret.dateDebut, arret.dateFin);
          }
          return sum;
        }, 0);
        
        if (contrat.analytics.dureeContrat > 0) {
          contrat.analytics.tauxAbsenteisme = (joursAbsence / contrat.analytics.dureeContrat) * 100;
        }
      }
      
      // NOUVEAUTÉ: Remplir les données initiales de la carrière
      contrat.evolutionCarriere.posteInitial = contrat.caracteristiques.libelleEmploi || "";
      contrat.evolutionCarriere.posteActuel = contrat.caracteristiques.libelleEmploi || "";
      contrat.evolutionCarriere.niveauInitial = contrat.classification.niveauRemuneration || "";
      contrat.evolutionCarriere.niveauActuel = contrat.classification.niveauRemuneration || "";
    });
  }
  
  /**
   * NOUVEAUTÉ: Enrichit le modèle avec des données spécifiques pour l'Index d'Égalité Professionnelle
   * @private
   * @param {object} model - Modèle de données DSN
   */
  function _enrichModelWithEqualityData(model) {
    // 1. Détecter les augmentations de salaire
    _detectSalaryIncreases(model);
    
    // 2. Détecter les promotions
    _detectPromotions(model);
    
    // 3. Analyser les retours de congé maternité
    _analyzeMaternityLeaveReturns(model);
    
    // 4. Identifier les 10 plus hautes rémunérations
    _identifyHighestRemunerations(model);
    
    // 5. Classer les salariés par niveau de rémunération
    _classifySalariesByRemunerationLevel(model);
  }
  
  /**
   * NOUVEAUTÉ: Détecte les augmentations de salaire
   * @private
   * @param {object} model - Modèle de données DSN
   */
  function _detectSalaryIncreases(model) {
    model.contrats.forEach(contrat => {
      if (contrat.remunerations.length < 2) return;
      
      // Trier les rémunérations par date
      const remusSorted = contrat.remunerations.sort((a, b) => {
        return a.periode.debut - b.periode.debut;
      });
      
      // Filtrer les rémunérations principales
      const remusPrincipales = remusSorted.filter(r => 
        r.type === "001" || r.type === "002" || r.type === "003"
      );
      
      if (remusPrincipales.length < 2) return;
      
      // Détecter les augmentations
      for (let i = 1; i < remusPrincipales.length; i++) {
        const prevRemu = remusPrincipales[i-1];
        const currRemu = remusPrincipales[i];
        
        if (prevRemu.montant && currRemu.montant && currRemu.montant > prevRemu.montant) {
          const augmentationPct = ((currRemu.montant - prevRemu.montant) / prevRemu.montant) * 100;
          const augmentationMontant = currRemu.montant - prevRemu.montant;
          
          // Marquer comme augmentation
          currRemu.indexEgalite.estAugmentation = true;
          currRemu.indexEgalite.tauxAugmentation = augmentationPct;
          currRemu.indexEgalite.montantAugmentation = augmentationMontant;
          
          // Ajouter à l'historique des augmentations
          contrat.evolutionRemuneration.historiqueAugmentations.push({
            dateDebut: prevRemu.periode.debut,
            dateFin: currRemu.periode.debut,
            remunerationAvant: prevRemu.montant,
            remunerationApres: currRemu.montant,
            augmentationPct: augmentationPct,
            augmentationMontant: augmentationMontant
          });
          
          // Si c'est la dernière période, mettre à jour l'augmentation de la dernière année
          if (i === remusPrincipales.length - 1) {
            contrat.evolutionRemuneration.augmentationDerniereAnnee = augmentationMontant;
            contrat.evolutionRemuneration.tauxAugmentation = augmentationPct;
          }
        }
      }
      
      // Mettre à jour les données d'augmentation du salarié
      const salarie = model.salaries.find(s => s.identite.nir === contrat.id.nirSalarie);
      if (salarie && contrat.evolutionRemuneration.historiqueAugmentations.length > 0) {
        // Dernière augmentation
        const derniereAugmentation = contrat.evolutionRemuneration.historiqueAugmentations[
          contrat.evolutionRemuneration.historiqueAugmentations.length - 1
        ];
        
        // Mettre à jour les données d'augmentation
        salarie.analytics.indexEgalite.augmentation.historique.push(...contrat.evolutionRemuneration.historiqueAugmentations);
        salarie.analytics.indexEgalite.augmentation.derniere = derniereAugmentation.dateFin;
        salarie.analytics.indexEgalite.augmentation.tauxAugmentation = derniereAugmentation.augmentationPct;
      }
    });
  }
  
  /**
   * NOUVEAUTÉ: Détecte les promotions
   * @private
   * @param {object} model - Modèle de données DSN
   */
  function _detectPromotions(model) {
    // Pour détecter les promotions, on examine les changements de niveau, d'échelon ou de coefficient
    model.contrats.forEach(contrat => {
      if (contrat.evolutionCarriere.historiquePromotions.length > 0) return; // Déjà détecté
      
      // Rechercher les événements de contrat avec changements de classification
      // Dans ce modèle simplifié, on suppose que toute augmentation de salaire > 5% 
      // non liée à une augmentation générale est une promotion
      if (contrat.evolutionRemuneration.historiqueAugmentations.length > 0) {
        const augmentationsSignificatives = contrat.evolutionRemuneration.historiqueAugmentations.filter(
          aug => aug.augmentationPct > 5.0
        );
        
        if (augmentationsSignificatives.length > 0) {
          // Considérer ces augmentations significatives comme des promotions
          augmentationsSignificatives.forEach(aug => {
            // Créer un enregistrement de promotion
            const promotion = {
              date: aug.dateFin,
              ancien: {
                poste: contrat.evolutionCarriere.posteInitial,
                niveau: contrat.evolutionCarriere.niveauInitial,
                remuneration: aug.remunerationAvant
              },
              nouveau: {
                poste: contrat.evolutionCarriere.posteActuel,
                niveau: contrat.evolutionCarriere.niveauActuel,
                remuneration: aug.remunerationApres
              },
              augmentationPct: aug.augmentationPct
            };
            
            // Ajouter à l'historique des promotions
            contrat.evolutionCarriere.historiquePromotions.push(promotion);
            
            // Si c'est la promotion la plus récente dans les 12 derniers mois, la marquer
            const today = new Date();
            const datePromotion = new Date(aug.dateFin);
            const diffMois = _calculateMonthsDifference(datePromotion, today);
            
            if (diffMois <= 12) {
              contrat.evolutionCarriere.estPromuDerniereAnnee = true;
            }
          });
        }
      }
      
      // Mettre à jour les données de promotion du salarié
      const salarie = model.salaries.find(s => s.identite.nir === contrat.id.nirSalarie);
      if (salarie && contrat.evolutionCarriere.historiquePromotions.length > 0) {
        // Dernière promotion
        const dernierePromotion = contrat.evolutionCarriere.historiquePromotions[
          contrat.evolutionCarriere.historiquePromotions.length - 1
        ];
        
        // Mettre à jour les données de promotion
        salarie.analytics.indexEgalite.promotion.historique.push(...contrat.evolutionCarriere.historiquePromotions);
        salarie.analytics.indexEgalite.promotion.derniere = dernierePromotion.date;
        salarie.analytics.indexEgalite.promotion.estPromu = contrat.evolutionCarriere.estPromuDerniereAnnee;
      }
    });
  }
  
  /**
   * NOUVEAUTÉ: Analyse les retours de congé maternité
   * @private
   * @param {object} model - Modèle de données DSN
   */
  function _analyzeMaternityLeaveReturns(model) {
    model.arrets.forEach(arret => {
      // Vérifier s'il s'agit d'un congé maternité
      if (arret.suiviEgalite.estCongeMaternite && arret.dateReprise) {
        const salarie = model.salaries.find(s => s.identite.nir === arret.nirSalarie);
        if (!salarie) return;
        
        // Vérifier s'il y a eu une augmentation au retour
        let augmentationAuRetour = false;
        
        // Rechercher des augmentations dans les 3 mois suivant le retour
        const contrats = model.contrats.filter(c => c.id.nirSalarie === salarie.identite.nir);
        
        for (const contrat of contrats) {
          for (const aug of contrat.evolutionRemuneration.historiqueAugmentations) {
            const dateAugmentation = new Date(aug.dateFin);
            const dateReprise = new Date(arret.dateReprise);
            
            // Vérifier si l'augmentation a eu lieu dans les 3 mois suivant le retour
            const diffJours = _calculateDaysDifference(dateReprise, dateAugmentation);
            
            if (diffJours >= 0 && diffJours <= 90) {
              augmentationAuRetour = true;
              
              // Mettre à jour les données d'arrêt
              arret.suiviEgalite.augmentationAuRetour = true;
              arret.suiviEgalite.dateAugmentation = dateAugmentation;
              arret.suiviEgalite.montantAugmentation = aug.augmentationMontant;
              
              // Mettre à jour les données d'augmentation du salarié
              salarie.analytics.indexEgalite.augmentation.apresCongeMaternite = true;
              
              break;
            }
          }
          
          if (augmentationAuRetour) break;
        }
      }
    });
  }
  
  /**
   * NOUVEAUTÉ: Identifie les 10 plus hautes rémunérations
   * @private
   * @param {object} model - Modèle de données DSN
   */
  function _identifyHighestRemunerations(model) {
    // Récupérer la dernière rémunération de chaque salarié
    const salariesRemunerations = [];
    
    model.salaries.forEach(salarie => {
      let remunerationTotale = 0;
      
      // Calculer la rémunération totale du salarié (tous contrats)
      salarie.contrats.forEach(numContrat => {
        const contrat = model.contrats.find(c => c.id.numeroContrat === numContrat);
        if (contrat && contrat.remunerations.length > 0) {
          // Utiliser la dernière rémunération connue
          const remusSorted = contrat.remunerations
            .filter(r => r.type === "001" || r.type === "002" || r.type === "003")
            .sort((a, b) => b.periode.debut - a.periode.debut);
          
          if (remusSorted.length > 0) {
            remunerationTotale += remusSorted[0].montant || 0;
          }
        }
      });
      
      if (remunerationTotale > 0) {
        salariesRemunerations.push({
          nir: salarie.identite.nir,
          sexe: salarie.identite.sexe,
          remuneration: remunerationTotale
        });
      }
    });
    
    // Trier par rémunération décroissante
    salariesRemunerations.sort((a, b) => b.remuneration - a.remuneration);
    
    // Prendre les 10 premiers (ou moins si moins de 10 salariés)
    const topRemunerations = salariesRemunerations.slice(0, Math.min(10, salariesRemunerations.length));
    
    // Mettre à jour les données des salariés concernés
    topRemunerations.forEach((top, index) => {
      const salarie = model.salaries.find(s => s.identite.nir === top.nir);
      if (salarie) {
        salarie.analytics.indexEgalite.niveauRemuneration.estParmiLes10Plus = true;
        salarie.analytics.indexEgalite.niveauRemuneration.classement = index + 1;
      }
    });
  }
  
  /**
   * NOUVEAUTÉ: Classe les salariés par quartile de rémunération
   * @private
   * @param {object} model - Modèle de données DSN
   */
  function _classifySalariesByRemunerationLevel(model) {
    // Regrouper les salariés par catégorie professionnelle et tranche d'âge
    const groupes = {};
    
    model.salaries.forEach(salarie => {
      // Récupérer les informations de classification
      let categoriePoste = "";
      let trancheAge = salarie.analytics.trancheAge || "";
      
      if (salarie.contrats.length > 0) {
        const contrat = model.contrats.find(c => c.id.numeroContrat === salarie.contrats[0]);
        if (contrat) {
          categoriePoste = contrat.classificationEgalite.categoriePoste || _determinerCategorieCSP(contrat.caracteristiques.codeCSP);
        }
      }
      
      if (!categoriePoste || !trancheAge) return;
      
      // Créer la clé du groupe
      const groupKey = `${categoriePoste}_${trancheAge}`;
      
      if (!groupes[groupKey]) {
        groupes[groupKey] = [];
      }
      
      // Calculer la rémunération totale du salarié
      let remunerationTotale = 0;
      
      salarie.contrats.forEach(numContrat => {
        const contrat = model.contrats.find(c => c.id.numeroContrat === numContrat);
        if (contrat && contrat.remunerations.length > 0) {
          // Utiliser la dernière rémunération connue
          const remusSorted = contrat.remunerations
            .filter(r => r.type === "001" || r.type === "002" || r.type === "003")
            .sort((a, b) => b.periode.debut - a.periode.debut);
          
          if (remusSorted.length > 0) {
            remunerationTotale += remusSorted[0].montant || 0;
          }
        }
      });
      
      // Ajouter le salarié au groupe
      groupes[groupKey].push({
        nir: salarie.identite.nir,
        sexe: salarie.identite.sexe,
        remuneration: remunerationTotale
      });
    });
    
    // Pour chaque groupe, diviser en quartiles et assigner aux salariés
    for (const groupKey in groupes) {
      const groupe = groupes[groupKey];
      
      if (groupe.length < 4) continue; // Pas assez de salariés pour calculer des quartiles
      
      // Trier par rémunération
      groupe.sort((a, b) => a.remuneration - b.remuneration);
      
      // Calculer les indices des quartiles
      const q1Index = Math.floor(groupe.length * 0.25);
      const q2Index = Math.floor(groupe.length * 0.5);
      const q3Index = Math.floor(groupe.length * 0.75);
      
      // Assigner les quartiles
      groupe.forEach((s, index) => {
        const salarie = model.salaries.find(sal => sal.identite.nir === s.nir);
        if (!salarie) return;
        
        if (index < q1Index) {
          salarie.analytics.indexEgalite.niveauRemuneration.quartile = 1;
        } else if (index < q2Index) {
          salarie.analytics.indexEgalite.niveauRemuneration.quartile = 2;
        } else if (index < q3Index) {
          salarie.analytics.indexEgalite.niveauRemuneration.quartile = 3;
        } else {
          salarie.analytics.indexEgalite.niveauRemuneration.quartile = 4;
        }
      });
    }
  }
  
  /**
   * NOUVEAUTÉ: Détermine la catégorie professionnelle à partir du code CSP
   * @private
   * @param {string} codeCSP - Code CSP du contrat
   * @return {string} - Catégorie professionnelle
   */
  function _determinerCategorieCSP(codeCSP) {
    if (!codeCSP) return "";
    
    // Le premier caractère du code CSP détermine la catégorie
    const premiereLettre = codeCSP.charAt(0);
    
    switch (premiereLettre) {
      case "2": return "Cadres";
      case "3": 
      case "4": return "TAM"; // Techniciens et Agents de Maîtrise
      case "5": return "Employés";
      case "6": 
      case "7": return "Ouvriers";
      default: return "";
    }
  }
  
  /**
   * NOUVEAUTÉ: Détecte le type d'arrêt pour l'Index d'Égalité Professionnelle
   * @private
   * @param {object} arret - Modèle d'arrêt de travail
   */
  function _detectArretType(arret) {
    if (arret.motif === "02") {
      arret.suiviEgalite.estCongeMaternite = true;
    } else if (arret.motif === "03") {
      arret.suiviEgalite.estCongePaternite = true;
    } else if (arret.motif === "09") {
      arret.suiviEgalite.estCongeAdoption = true;
    }
  }
  
  /**
   * Recherche la valeur d'une rubrique dans un tableau de rubriques
   * @private
   * @param {array} arr - Tableau de rubriques
   * @param {string} rubId - Identifiant de la rubrique
   * @return {string|null} - Valeur de la rubrique ou null
   */
  function _findRubriqueValue(arr, rubId) {
    if (!arr) return null;
    const rubrique = arr.find(function(r) { 
      return r.rubrique === rubId; 
    });
    return rubrique ? rubrique.valeur : null;
  }
  
  /**
   * Parse une date au format DSN (JJMMAAAA)
   * @private
   * @param {string} dateStr - Chaîne de date DSN
   * @return {Date|null} - Objet Date ou null si invalide
   */
  function _parseDSNDate(dateStr) {
    if (!dateStr) return null;
    
    // Format JJMMAAAA (DSN standard)
    const match = dateStr.match(PATTERNS.DSN_DATE);
    if (match) {
      const jour = parseInt(match[1], 10);
      const mois = parseInt(match[2], 10) - 1; // Les mois commencent à 0 en JavaScript
      const annee = parseInt(match[3], 10);
      
      const date = new Date(annee, mois, jour);
      if (!isNaN(date.getTime())) {
        return date;
      }
    }
    
    // Tenter avec le parseur de DateUtils si disponible
    if (typeof DateUtils !== 'undefined' && DateUtils.parseDateString) {
      return DateUtils.parseDateString(dateStr);
    }
    
    return null;
  }
  
  /**
   * Parse une valeur numérique
   * @private
   * @param {string} valueStr - Chaîne numérique
   * @return {number|null} - Nombre ou null si invalide
   */
  function _parseNumber(valueStr) {
    if (!valueStr) return null;
    
    // Vérifier si c'est un nombre décimal valide
    if (PATTERNS.DSN_DECIMAL.test(valueStr)) {
      return parseFloat(valueStr);
    }
    
    return null;
  }
  
  /**
   * Calcule l'âge à partir d'une date de naissance
   * @private
   * @param {Date} birthDate - Date de naissance
   * @param {Date} refDate - Date de référence
   * @return {number} - Âge calculé
   */
  function _calculateAge(birthDate, refDate) {
    let age = refDate.getFullYear() - birthDate.getFullYear();
    const m = refDate.getMonth() - birthDate.getMonth();
    
    if (m < 0 || (m === 0 && refDate.getDate() < birthDate.getDate())) {
      age--;
    }
    
    return age;
  }
  
  /**
   * Détermine la tranche d'âge
   * @private
   * @param {number} age - Âge
   * @return {string} - Tranche d'âge
   */
  function _getTrancheAge(age) {
    if (age < 30) return "moins de 30 ans";
    if (age < 40) return "30 à 39 ans";
    if (age < 50) return "40 à 49 ans";
    return "50 ans et plus";
  }
  
  /**
   * Calcule le nombre de mois entre deux dates
   * @private
   * @param {Date} startDate - Date de début
   * @param {Date} endDate - Date de fin
   * @return {number} - Nombre de mois
   */
  function _calculateMonthsDifference(startDate, endDate) {
    return (endDate.getFullYear() - startDate.getFullYear()) * 12 + 
           (endDate.getMonth() - startDate.getMonth());
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
    return Math.ceil(diffTime / (1000 * 60 * 60 * 24));
  }
  
  // API publique
  return {
    parseDSNContent: parseDSNContent
  };
})();
