// DSNModels.gs - Définitions des modèles de données conformes au cahier technique DSN 2025.1

const DSNModels = (function() {
  /**
   * Modèle de données pour un salarié (S21.G00.30)
   * @returns {Object} Modèle vide d'un salarié
   */
  function createSalarieModel() {
    return {
      // Identité
      identite: {
        nir: "",                      // S21.G00.30.001 - NIR
        nom: "",                      // S21.G00.30.002 - Nom de famille
        nomUsage: "",                 // S21.G00.30.003 - Nom d'usage
        prenoms: "",                  // S21.G00.30.004 - Prénoms
        sexe: "",                     // S21.G00.30.005 - Sexe
        dateNaissance: null,          // S21.G00.30.006 - Date de naissance
        lieuNaissance: "",            // S21.G00.30.007 - Lieu de naissance
        codeDeptNaissance: "",        // S21.G00.30.014 - Code département de naissance
        codePaysNaissance: "",        // S21.G00.30.015 - Code pays de naissance
        matricule: "",                // S21.G00.30.019 - Matricule
        ntt: ""                       // S21.G00.30.020 - Numéro technique temporaire
      },
      
      // Coordonnées
      coordonnees: {
        adresse: {
          voie: "",                   // S21.G00.30.008 - Numéro et libellé de voie
          complement1: "",            // S21.G00.30.016 - Complément d'adresse
          complement2: "",            // S21.G00.30.017 - Service de distribution
          codePostal: "",             // S21.G00.30.009 - Code postal
          localite: "",               // S21.G00.30.010 - Localité
          codePays: ""                // S21.G00.30.011 - Code pays
        },
        email: "",                    // S21.G00.30.018 - Adresse mél
        telephone: ""                 // Nouveau - À collecter séparément si nécessaire
      },
      
      // Situation professionnelle
      situation: {
        statutEtranger: "",           // S21.G00.30.022 - Statut à l'étranger au sens fiscal
        paysResidence: "",            // S21.G00.30.029 - Pays de résidence
        cumulEmploiRetraite: "",      // S21.G00.30.023 - Cumul emploi retraite
        regimeMaladie: "",            // S21.G00.30.024 - Code régime base risque maladie
        regimeVieillesse: "",         // S21.G00.30.025 - Code régime base risque vieillesse  
        boeth: ""                     // S21.G00.40.072 - Statut BOETH
      },
      
      // Formation et compétences
      formation: {
        niveauFormation: "",          // S21.G00.30.027 - Niveau de formation le plus élevé
        niveauDiplome: ""             // S21.G00.30.028 - Niveau de diplôme
      },
      
      // NOUVEAUTÉ: Carrière et évolution professionnelle
      carriere: {
        dateEntree: null,             // Date d'entrée dans l'entreprise
        fonctionActuelle: "",         // Fonction ou poste actuel
        categorieCSP: "",             // Catégorie socio-professionnelle
        niveauHierarchique: "",       // Niveau dans la hiérarchie
        estCadre: false,              // Statut cadre
        promotions: [],               // Historique des promotions
        congesMaternite: [],          // Historique des congés maternité
        autrementRetourConge: [],     // Historique des retours de congés maternité/paternité/adoption
        evolutionsSalariales: []      // Historique des évolutions salariales
      },
      
      // Agrégats et données calculées
      analytics: {
        age: null,                    // Calculé à partir de la date de naissance
        trancheAge: "",               // Calculé en catégories d'âge
        ancienneteEntreprise: null,   // En mois (calculé)
        salairesMoyens: {             // Moyennes des 12 derniers mois
          brut: null,
          net: null
        },
        evolutionSalaire: [],         // Historique d'évolution salariale
        tempsAbsence: {               // Statistiques d'absence
          total: 0,                   // Jours d'absence cumulés
          parMotif: {                 // Répartition par motif
            maladie: 0,
            maternite: 0,
            accident: 0,
            autres: 0
          }
        },
        // NOUVEAUTÉ: Indicateurs Index Égalité Professionnelle
        indexEgalite: {
          augmentation: {
            historique: [],           // Historique des augmentations
            derniere: null,           // Date de la dernière augmentation
            tauxAugmentation: null,   // Taux de la dernière augmentation
            apresCongeMaternite: null // Augmentation après congé maternité (oui/non)
          },
          promotion: {
            historique: [],           // Historique des promotions
            derniere: null,           // Date de la dernière promotion
            estPromu: false           // Indicateur si promu sur la période
          },
          niveauRemuneration: {
            classement: null,         // Classement salarial dans l'entreprise
            estParmiLes10Plus: false, // Fait partie des 10 plus hautes rémunérations
            quartile: null            // Quartile de rémunération (1, 2, 3 ou 4)
          }
        }
      },
      
      // Relations avec d'autres entités
      contrats: [],                   // Liste des ID de contrats
      arrets: [],                     // Liste des ID des arrêts de travail
      versements: []                  // Liste des ID des versements
    };
  }
  
  /**
   * Modèle de données pour un contrat (S21.G00.40)
   * @returns {Object} Modèle vide d'un contrat
   */
  function createContratModel() {
    return {
      // Identification
      id: {
        numeroContrat: "",            // S21.G00.40.009 - Numéro du contrat
        dateDebut: null,              // S21.G00.40.001 - Date de début du contrat
        dateFin: null,                // S21.G00.40.010 - Date de fin prévisionnelle
        nirSalarie: ""                // S21.G00.30.001 - Référence au salarié
      },
      
      // Caractéristiques principales
      caracteristiques: {
        nature: "",                   // S21.G00.40.007 - Nature du contrat (CDI, CDD, etc.)
        statutSalarie: "",            // S21.G00.40.002 - Statut du salarié (cadre, non-cadre)
        statutEmploi: "",             // S21.G00.40.026 - Statut d'emploi du salarié
        motifRecours: "",             // S21.G00.40.021 - Motif de recours (pour CDD)
        libelleEmploi: "",            // S21.G00.40.006 - Libellé de l'emploi
        codeCSP: "",                  // S21.G00.40.004 - Code PCS-ESE
        complementCSP: "",            // S21.G00.40.005 - Complément PCS-ESE
        dispositifPublic: "",         // S21.G00.40.008 - Dispositif de politique publique
        complementDispositif: ""      // S21.G00.40.073 - Complément dispositif
      },
      
      // Temps de travail
      tempsTravail: {
        modalite: "",                 // S21.G00.40.013 - Modalité d'exercice du temps de travail
        uniteMesure: "",              // S21.G00.40.011 - Unité de mesure de la quotité
        quotite: null,                // S21.G00.40.012 - Quotité de travail
        tauxTempsPartiel: null,       // S21.G00.40.043 - Taux de travail à temps partiel
        horairesContractuels: null    // Nombre d'heures contractuelles
      },
      
      // Classification et rémunération
      classification: {
        conventionCollective: "",     // S21.G00.40.017 - Code convention collective
        codeRegimeMaladie: "",        // S21.G00.40.018 - Code régime maladie
        codeRegimeVieillesse: "",     // S21.G00.40.020 - Code régime vieillesse
        codeRegimeAT: "",             // S21.G00.40.039 - Code régime accident du travail
        codeRisqueAT: "",             // S21.G00.40.040 - Code risque accident du travail
        positionConvention: "",       // S21.G00.40.041 - Positionnement dans la convention 
        niveauRemuneration: "",       // S21.G00.40.048 - Niveau de rémunération
        echelon: "",                  // S21.G00.40.049 - Échelon
        coefficient: ""               // S21.G00.40.050 - Coefficient
      },
      
      // NOUVEAUTÉ: Classification détaillée pour l'index d'égalité
      classificationEgalite: {
        categoriePoste: "",           // Catégorie de poste (Ouvriers, Employés, TAM, Cadres)
        niveau: "",                   // Niveau dans la catégorie
        coefficientHierarchique: "",  // Coefficient hiérarchique
        niveauResponsabilite: "",     // Niveau de responsabilité
        fonctionEncadrement: false,   // Fonction d'encadrement
        nombreSubordonnes: 0          // Nombre de personnes encadrées
      },
      
      // Lieu de travail
      lieuTravail: {
        identifiant: "",              // S21.G00.40.019 - Identifiant du lieu de travail
        libelle: "",                  // S21.G00.85.002 - Libellé du lieu de travail
        adresse: {
          voie: "",                   // S21.G00.85.003 - Numéro et libellé de voie
          codePostal: "",             // S21.G00.85.004 - Code postal
          localite: "",               // S21.G00.85.005 - Localité
          codeDepartement: "",        // S21.G00.85.006 - Code département
          codePays: ""                // S21.G00.85.007 - Code pays
        },
        codeAPET: "",                 // S21.G00.85.009 - Code APET
        codeINSEE: ""                 // S21.G00.85.010 - Code INSEE commune
      },
      
      // Données complémentaires
      complementaire: {
        complementSante: "",          // S21.G00.40.047 - Complémentaire santé
        codeStatutApecita: "",        // S21.G00.40.042 - Code statut catégoriel APECITA
        tauxDeduction: null,          // S21.G00.40.023 - Taux déduction forfaitaire
        travailleurEtranger: "",      // S21.G00.40.024 - Travailleur à l'étranger
        motifExclusion: "",           // S21.G00.40.025 - Motif d'exclusion DSN
        dispositifAide: "",           // S21.G00.40.078 - Dispositif d'aide à l'emploi
        dateAdhesion: null,           // S21.G00.40.079 - Date d'adhésion
        dateDenonciation: null        // S21.G00.40.080 - Date de dénonciation
      },
      
      // Informations d'ancienneté
      anciennete: {
        debutPeriodeReference: null,  // S21.G00.86.001 - Début période de référence
        finPeriodeReference: null,    // S21.G00.86.002 - Fin période de référence
        uniteMesure: "",              // S21.G00.86.003 - Unité de mesure
        valeur: null                  // S21.G00.86.004 - Valeur
      },
      
      // Historique des rémunérations
      remunerations: [],
      
      // NOUVEAUTÉ: Évolution de la rémunération pour l'index d'égalité
      evolutionRemuneration: {
        remunerationInitiale: null,      // Rémunération à la signature du contrat
        remunerationActuelle: null,      // Rémunération actuelle
        historiqueAugmentations: [],     // Historique des augmentations
        pourcentageEvolution: null,      // Évolution en % depuis la signature
        augmentationDerniereAnnee: null, // Montant de l'augmentation sur la dernière année
        tauxAugmentation: null           // Taux d'augmentation sur la dernière année
      },
      
      // NOUVEAUTÉ: Suivi de carrière pour l'index d'égalité
      evolutionCarriere: {
        posteInitial: "",               // Poste à la signature du contrat
        posteActuel: "",                // Poste actuel
        historiquePromotions: [],       // Historique des promotions
        estPromuDerniereAnnee: false,   // Promu dans les 12 derniers mois
        niveauInitial: "",              // Niveau hiérarchique initial
        niveauActuel: ""                // Niveau hiérarchique actuel
      },
      
      // Historique des arrêts de travail liés à ce contrat
      arretsLies: [],
      
      // Données agrégées et calculées pour analyse
      analytics: {
        dureeContrat: null,           // En jours (calculé)
        coutTotal: null,              // Coût total sur la durée du contrat
        salaireInitial: null,         // Salaire au début du contrat
        salaireFinal: null,           // Dernier salaire connu
        evolutionPourcentage: null,   // Évolution en % sur la période
        heuresTravaillees: null,      // Total des heures travaillées
        joursTravailles: null,        // Total des jours travaillés
        tauxAbsenteisme: null,        // Taux d'absentéisme
        periodeEssai: {               // Informations sur période d'essai
          debut: null,
          fin: null,
          prolongation: false,
          reussie: null
        }
      }
    };
  }
  
  /**
   * Modèle de données pour une rémunération (S21.G00.51)
   * @returns {Object} Modèle vide d'une rémunération
   */
  function createRemunerationModel() {
    return {
      periode: {
        debut: null,                  // S21.G00.51.001 - Date début période de paie
        fin: null                     // S21.G00.51.002 - Date fin période de paie
      },
      numeroContrat: "",              // S21.G00.51.010 - Numéro du contrat
      type: "",                       // S21.G00.51.011 - Type de rémunération
      montant: null,                  // S21.G00.51.013 - Montant
      tauxRemuneration: null,         // S21.G00.51.014 - Taux de rémunération
      nbHeures: null,                 // S21.G00.51.012 - Nombre d'heures
      uniteMesure: "",                // S21.G00.51.015 - Unité de mesure
      
      // NOUVEAUTÉ: Indicateurs pour l'analyse de l'index d'égalité
      indexEgalite: {
        estAugmentation: false,        // Indique s'il s'agit d'une augmentation
        tauxAugmentation: null,        // Taux d'augmentation en %
        montantAugmentation: null,     // Montant de l'augmentation
        lieePromotion: false,          // Liée à une promotion
        apresRetourCongeMaternite: false, // Après un retour de congé maternité
        retroactive: false             // Augmentation rétroactive
      }
    };
  }
  
  /**
   * Modèle de données pour un arrêt de travail (S21.G00.60)
   * @returns {Object} Modèle vide d'un arrêt de travail
   */
  function createArretTravailModel() {
    return {
      motif: "",                      // S21.G00.60.001 - Motif de l'arrêt
      dateDebut: null,                // S21.G00.60.002 - Date du dernier jour travaillé
      dateFin: null,                  // S21.G00.60.003 - Date de fin prévisionnelle
      dateReprise: null,              // S21.G00.60.010 - Date de la reprise
      motifReprise: "",               // S21.G00.60.011 - Motif de la reprise
      nirSalarie: "",                 // NIR du salarié concerné
      numeroContrat: "",              // Numéro du contrat concerné
      subrogation: {
        debut: null,                  // S21.G00.60.004 - Date début subrogation
        fin: null,                    // S21.G00.60.005 - Date fin subrogation
        montant: null                 // S21.G00.60.007 - Montant de la subrogation
      },
      
      // NOUVEAUTÉ: Suivi pour l'index égalité
      suiviEgalite: {
        estCongeMaternite: false,     // S'agit-il d'un congé maternité
        estCongePaternite: false,     // S'agit-il d'un congé paternité
        estCongeAdoption: false,      // S'agit-il d'un congé d'adoption
        augmentationAuRetour: null,   // Y a-t-il eu une augmentation au retour
        dateAugmentation: null,       // Date de l'augmentation après retour
        montantAugmentation: null     // Montant de l'augmentation après retour
      }
    };
  }
  
  /**
   * Modèle de données pour un événement de contrat
   * @returns {Object} Modèle vide d'un événement de contrat
   */
  function createEvenementContratModel() {
    return {
      // Identification
      id: {
        idEvenement: "",              // Identifiant unique
        numeroContrat: "",            // S21.G00.40.009 - Numéro du contrat lié
        nirSalarie: ""                // S21.G00.30.001 - NIR du salarié
      },
      
      // Caractéristiques de l'événement
      type: "",                       // Type d'événement (début, modification, fin, suspension, etc.)
      dateEvenement: null,            // Date effective de l'événement
      motif: "",                      // Motif lié à l'événement
      
      // Contenu détaillé selon le type
      contenu: {
        // Pour modifications contractuelles
        modifications: {
          avenant: "",                // Numéro d'avenant au contrat
          champModifie: "",           // Champ modifié (ex: "salaire", "temps de travail")
          ancienneValeur: "",         // Valeur avant modification
          nouvelleValeur: "",         // Valeur après modification
          pourcentageEvolution: null  // Pourcentage d'évolution (si applicable)
        },
        
        // Pour les fins de contrat
        finContrat: {
          motifRupture: "",           // S21.G00.62.001 - Motif de la rupture du contrat
          dateFinContrat: null,       // S21.G00.62.002 - Date de fin du contrat
          dateNotification: null,     // S21.G00.62.003 - Date de notification
          typePreavis: "",            // S21.G00.62.012 - Type de préavis
          dateDebuttPreavis: null,    // S21.G00.62.013 - Date de début de préavis
          dateFinPreavis: null,       // S21.G00.62.014 - Date de fin de préavis
          indemnites: {
            indemniteConventionnelle: {
              montant: null,          // Montant de l'indemnité conventionnelle
              assiette: null          // Assiette de l'indemnité
            },
            indemniteLegale: {
              montant: null,          // Montant de l'indemnité légale
              assiette: null          // Assiette de l'indemnité
            },
            indemniteSupralegale: {
              montant: null           // Montant de l'indemnité supralégale
            }
          }
        },
        
        // Pour les activités partielles
        activitePartielle: {
          mois: "",                   // S21.G00.95.001 - Mois de la suspension
          typeHeures: "",             // S21.G00.95.003 - Type d'heure
          nombreHeures: null,         // S21.G00.95.004 - Nombre d'heures
          montantVerse: null          // S21.G00.95.005 - Montant versé
        },
        
        // Pour les arrêts de travail
        arretTravail: {
          motif: "",                  // S21.G00.60.001 - Motif de l'arrêt
          dateDebut: null,            // S21.G00.60.002 - Date du dernier jour travaillé
          dateFin: null,              // S21.G00.60.003 - Date de fin prévisionnelle
          dateReprise: null,          // S21.G00.60.010 - Date de la reprise
          motifReprise: "",           // S21.G00.60.011 - Motif de la reprise
          subrogation: false          // Indicateur de subrogation
        }
      },
      
      // Données liées à la rémunération
      impactRemuneration: {
        indicateur: false,            // Indique si l'événement impacte la rémunération
        montantAvant: null,           // Montant avant l'événement
        montantApres: null,           // Montant après l'événement
        differenceAbsolue: null,      // Différence en valeur absolue
        differencePourcentage: null   // Différence en pourcentage
      },
      
      // NOUVEAUTÉ: Indicateurs pour l'index d'égalité
      indexEgalite: {
        estPromotion: false,          // S'agit-il d'une promotion
        estAugmentation: false,       // S'agit-il d'une augmentation
        tauxAugmentation: null,       // Taux d'augmentation si applicable
        apresCongeMaternite: false,   // Événement après un congé maternité
        nouveauNiveau: "",            // Nouveau niveau hiérarchique si promotion
        ancienNiveau: ""              // Ancien niveau hiérarchique si promotion
      },
      
      // Métadonnées pour l'analyse
      metadata: {
        source: "",                   // Source de l'événement (DSN, SIRH, etc.)
        dateCreation: null,           // Date de création de l'enregistrement
        dateModification: null,       // Date de dernière modification
        utilisateur: "",              // Identifiant de l'utilisateur responsable
        commentaire: ""               // Commentaire libre
      }
    };
  }
  
  /**
   * NOUVEAUTÉ: Modèle de données pour l'index d'égalité professionnelle
   * @returns {Object} Modèle d'index d'égalité
   */
  function createIndexEgaliteModel() {
    return {
      // Informations générales
      metadata: {
        dateCalcul: new Date(),           // Date du calcul
        periodeReference: {               // Période de référence
          debut: null,
          fin: null
        },
        effectifTotal: 0,                 // Effectif total pris en compte
        effectifHommes: 0,                // Nombre d'hommes
        effectifFemmes: 0,                // Nombre de femmes
        repartitionCategories: {}         // Répartition par catégorie professionnelle
      },
      
      // Indicateur 1: Écart de rémunération (40 points)
      indicateur1: {
        score: null,                      // Score sur 40 points
        ecartGlobal: null,                // Écart global pondéré
        ecartRemuneration: null,          // Écart de rémunération entre hommes et femmes (%)
        parCategorie: [],                 // Écarts par catégorie d'âge et CSP
        tauxSexesCompares: 0,             // Taux de groupes de salaires comparables
        populationComparable: 0,          // Nombre de salariés dans les groupes comparables
        nonCalculable: false,             // Indicateur non calculable
        motifNonCalculable: ""            // Motif si non calculable
      },
      
      // Indicateur 2: Écart de taux d'augmentations (20 points)
      indicateur2: {
        score: null,                      // Score sur 20 points
        tauxAugmentationHommes: null,     // Taux d'augmentation chez les hommes (%)
        tauxAugmentationFemmes: null,     // Taux d'augmentation chez les femmes (%)
        ecartTauxAugmentation: null,      // Écart entre les taux d'augmentation (points de %)
        nombreHommes: 0,                  // Nombre d'hommes considérés
        nombreFemmes: 0,                  // Nombre de femmes considérées
        nombreAugmentesHommes: 0,         // Nombre d'hommes augmentés
        nombreAugmenteesFemmes: 0,        // Nombre de femmes augmentées
        nonCalculable: false,             // Indicateur non calculable
        motifNonCalculable: ""            // Motif si non calculable
      },
      
      // Indicateur 3: Écart de taux de promotions (15 points)
      indicateur3: {
        score: null,                      // Score sur 15 points
        tauxPromotionHommes: null,        // Taux de promotion chez les hommes (%)
        tauxPromotionFemmes: null,        // Taux de promotion chez les femmes (%)
        ecartTauxPromotion: null,         // Écart entre les taux de promotion (points de %)
        nombreHommes: 0,                  // Nombre d'hommes considérés
        nombreFemmes: 0,                  // Nombre de femmes considérées
        nombrePromusHommes: 0,            // Nombre d'hommes promus
        nombrePromuesFemmes: 0,           // Nombre de femmes promues
        nonCalculable: false,             // Indicateur non calculable
        motifNonCalculable: ""            // Motif si non calculable
      },
      
      // Indicateur 4: Retour de congé maternité (15 points)
      indicateur4: {
        score: null,                      // Score sur 15 points
        nombreRetourCongeMaternite: 0,    // Nombre de retours de congé maternité
        nombreAugmentees: 0,              // Nombre de femmes augmentées au retour
        tauxRespect: null,                // Taux de respect de l'obligation légale (%)
        nonCalculable: false,             // Indicateur non calculable
        motifNonCalculable: ""            // Motif si non calculable
      },
      
      // Indicateur 5: Nombre de femmes dans les 10 plus hautes rémunérations (10 points)
      indicateur5: {
        score: null,                      // Score sur 10 points
        nombreFemmesPlusHautesRemunerations: 0, // Nombre de femmes dans les 10 plus hautes rémunérations
        nombreHommesPlusHautesRemunerations: 0,  // Nombre d'hommes dans les 10 plus hautes rémunérations
        nonCalculable: false,             // Indicateur non calculable
        motifNonCalculable: ""            // Motif si non calculable
      },
      
      // Indice global
      resultat: {
        total: null,                      // Score total sur 100
        totalCalculable: null,            // Score maximal calculable
        indicateurCalculables: 0,         // Nombre d'indicateurs calculables
        publication: {                   // Informations pour la publication
          date: null,                    // Date de publication
          methodologie: "",              // Commentaire sur la méthodologie
          mesuresCorrection: "",         // Mesures de correction prévues
          objectifsProgression: ""       // Objectifs de progression
        }
      }
    };
  }
  
  // Méthodes à exposer publiquement
  return {
    createSalarieModel,
    createContratModel,
    createRemunerationModel,
    createArretTravailModel,
    createEvenementContratModel,
    createIndexEgaliteModel // NOUVEAUTÉ: Fonction pour le modèle d'index d'égalité
  };
})();
