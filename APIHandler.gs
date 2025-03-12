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
          // Si nous utilisons le format 2025, nous créons directement un modèle compatible
          if (formData.settings && formData.settings.format === "2025") {
            // Créer un modèle DSN 2025 directement
            const parsedModel = _createDSN2025Model(periode.annee, periode.mois, dsnData.data, formData);
            
            apiDSNData.push({
              annee: periode.annee,
              mois: periode.mois,
              data: parsedModel
            });
          } else {
            // Format standard - envoyer les données brutes pour être parsées
            apiDSNData.push({
              annee: periode.annee,
              mois: periode.mois,
              data: dsnData.data
            });
          }
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
   * Crée directement un modèle DSN 2025 avec des données de démonstration
   * @private
   * @param {number} annee - Année
   * @param {number} mois - Mois  
   * @param {string} rawData - Données brutes (non utilisées dans cette démo)
   * @param {object} formData - Données du formulaire
   * @return {object} - Modèle DSN 2025 complet
   */
  function _createDSN2025Model(annee, mois, rawData, formData) {
    const mm = mois < 10 ? "0" + mois : mois;
    const moisPrincipal = `${annee}-${mm}`;
    
    // Créer un modèle complet pour la démo avec des données détaillées
    const model = {
      moisPrincipal: moisPrincipal,
      etablissement: {
        siren: formData.codeDossier || "123456789",
        nic: formData.codeEtablissement || "00001",
        siret: (formData.codeDossier || "123456789") + (formData.codeEtablissement || "00001"),
        raisonSociale: "Entreprise Démo DSN",
        adresse: {
          voie: "123 Avenue de la Démonstration",
          codePostal: "75000",
          commune: "Paris",
          pays: "FR"
        }
      },
      salaries: [
        // Salarié 1 - Homme
        {
          identite: {
            nir: "1840363420123",
            nom: "DUPONT",
            nomUsage: "",
            prenoms: "Jean",
            sexe: "01",
            dateNaissance: new Date(1980, 0, 15), // 15/01/1980
            lieuNaissance: "Paris",
            codeDeptNaissance: "75",
            codePaysNaissance: "FR",
            matricule: "EMP001",
            ntt: ""
          },
          coordonnees: {
            adresse: {
              voie: "45 rue de Paris",
              codePostal: "75010",
              localite: "Paris",
              codePays: "FR"
            },
            email: "jean.dupont@exemple.fr"
          },
          situation: {
            statutEtranger: "",
            paysResidence: "FR",
            cumulEmploiRetraite: "01",
            regimeMaladie: "200",
            regimeVieillesse: "200",
            boeth: ""
          },
          formation: {
            niveauFormation: "02",
            niveauDiplome: "44"
          },
          carriere: {
            dateEntree: new Date(2018, 2, 1), // 01/03/2018
            fonctionActuelle: "Développeur informatique",
            categorieCSP: "Cadres",
            niveauHierarchique: "2",
            estCadre: true
          },
          analytics: {
            age: 44,
            trancheAge: "40 à 49 ans",
            ancienneteEntreprise: 72,
            salairesMoyens: {
              brut: 3950,
              net: 3080
            }
          },
          contrats: ["CONT001"],
          arrets: []
        },
        // Salarié 2 - Femme
        {
          identite: {
            nir: "2860363420456",
            nom: "MARTIN",
            nomUsage: "DUBOIS",
            prenoms: "Sophie",
            sexe: "02",
            dateNaissance: new Date(1986, 2, 28), // 28/03/1986
            lieuNaissance: "Lyon",
            codeDeptNaissance: "69",
            codePaysNaissance: "FR",
            matricule: "EMP002",
            ntt: ""
          },
          coordonnees: {
            adresse: {
              voie: "12 rue des Fleurs",
              codePostal: "69000",
              localite: "Lyon",
              codePays: "FR"
            },
            email: "sophie.martin@exemple.fr"
          },
          situation: {
            statutEtranger: "",
            paysResidence: "FR",
            cumulEmploiRetraite: "01",
            regimeMaladie: "200",
            regimeVieillesse: "200",
            boeth: ""
          },
          formation: {
            niveauFormation: "02",
            niveauDiplome: "44"
          },
          carriere: {
            dateEntree: new Date(2019, 5, 1), // 01/06/2019
            fonctionActuelle: "Comptable",
            categorieCSP: "Employés",
            niveauHierarchique: "3",
            estCadre: false
          },
          analytics: {
            age: 38,
            trancheAge: "30 à 39 ans",
            ancienneteEntreprise: 58,
            salairesMoyens: {
              brut: 3200,
              net: 2496
            }
          },
          contrats: ["CONT002"],
          arrets: []
        },
        // Salarié 3 - Homme
        {
          identite: {
            nir: "1740363420789",
            nom: "BERNARD",
            nomUsage: "",
            prenoms: "Pierre",
            sexe: "01",
            dateNaissance: new Date(1974, 8, 10), // 10/09/1974
            lieuNaissance: "Bordeaux",
            codeDeptNaissance: "33",
            codePaysNaissance: "FR",
            matricule: "EMP003",
            ntt: ""
          },
          coordonnees: {
            adresse: {
              voie: "7 avenue des Champs",
              codePostal: "33000",
              localite: "Bordeaux",
              codePays: "FR"
            },
            email: "pierre.bernard@exemple.fr"
          },
          situation: {
            statutEtranger: "",
            paysResidence: "FR",
            cumulEmploiRetraite: "01",
            regimeMaladie: "200",
            regimeVieillesse: "200",
            boeth: ""
          },
          formation: {
            niveauFormation: "03",
            niveauDiplome: "41"
          },
          carriere: {
            dateEntree: new Date(2015, 0, 15), // 15/01/2015
            fonctionActuelle: "Directeur Commercial",
            categorieCSP: "Cadres",
            niveauHierarchique: "1",
            estCadre: true
          },
          analytics: {
            age: 50,
            trancheAge: "50 ans et plus",
            ancienneteEntreprise: 110,
            salairesMoyens: {
              brut: 5200,
              net: 4056
            }
          },
          contrats: ["CONT003"],
          arrets: []
        },
        // Salarié 4 - Femme
        {
          identite: {
            nir: "2910363420142",
            nom: "PETIT",
            nomUsage: "",
            prenoms: "Marie",
            sexe: "02",
            dateNaissance: new Date(1991, 6, 22), // 22/07/1991
            lieuNaissance: "Marseille",
            codeDeptNaissance: "13",
            codePaysNaissance: "FR",
            matricule: "EMP004",
            ntt: ""
          },
          coordonnees: {
            adresse: {
              voie: "56 boulevard de la Mer",
              codePostal: "13000",
              localite: "Marseille",
              codePays: "FR"
            },
            email: "marie.petit@exemple.fr"
          },
          situation: {
            statutEtranger: "",
            paysResidence: "FR",
            cumulEmploiRetraite: "01",
            regimeMaladie: "200",
            regimeVieillesse: "200",
            boeth: ""
          },
          formation: {
            niveauFormation: "01",
            niveauDiplome: "46"
          },
          carriere: {
            dateEntree: new Date(2020, 8, 1), // 01/09/2020
            fonctionActuelle: "Assistante Administrative",
            categorieCSP: "Employés",
            niveauHierarchique: "4",
            estCadre: false
          },
          analytics: {
            age: 33,
            trancheAge: "30 à 39 ans",
            ancienneteEntreprise: 43,
            salairesMoyens: {
              brut: 2800,
              net: 2184
            }
          },
          contrats: ["CONT004"],
          arrets: []
        }
      ],
      contrats: [
        // Contrat 1
        {
          id: {
            numeroContrat: "CONT001",
            dateDebut: new Date(2018, 2, 1), // 01/03/2018
            dateFin: null,
            nirSalarie: "1840363420123"
          },
          caracteristiques: {
            nature: "01", // CDI
            statutSalarie: "03", // Cadre
            statutEmploi: "01",
            motifRecours: "",
            libelleEmploi: "Développeur informatique",
            codeCSP: "353a", // Ingénieurs et cadres
            complementCSP: "",
            dispositifPublic: "",
            complementDispositif: ""
          },
          tempsTravail: {
            modalite: "10", // Temps plein
            uniteMesure: "10", // Heure
            quotite: 35,
            tauxTempsPartiel: null,
            horairesContractuels: 151.67
          },
          classification: {
            conventionCollective: "1486",
            codeRegimeMaladie: "200",
            codeRegimeVieillesse: "200",
            positionConvention: "5",
            niveauRemuneration: "3",
            echelon: "2",
            coefficient: "100"
          },
          classificationEgalite: {
            categoriePoste: "Cadres",
            niveau: "3",
            coefficientHierarchique: "100",
            niveauResponsabilite: "2",
            fonctionEncadrement: false,
            nombreSubordonnes: 0
          },
          remunerations: [
            {
              periode: {
                debut: new Date(annee, mois-1, 1),
                fin: new Date(annee, mois-1, 
                              new Date(annee, mois, 0).getDate())
              },
              numeroContrat: "CONT001",
              type: "001", // Rémunération brute
              montant: 3950,
              tauxRemuneration: null,
              nbHeures: 151.67,
              uniteMesure: "10"
            }
          ],
          evolutionRemuneration: {
            remunerationInitiale: 3800,
            remunerationActuelle: 3950,
            pourcentageEvolution: 3.95,
            augmentationDerniereAnnee: 150
          },
          evolutionCarriere: {
            posteInitial: "Développeur informatique",
            posteActuel: "Développeur informatique",
            niveauInitial: "3",
            niveauActuel: "3"
          },
          analytics: {
            dureeContrat: 1095,
            coutTotal: 47400,
            salaireInitial: 3800,
            salaireFinal: 3950
          }
        },
        // Contrat 2
        {
          id: {
            numeroContrat: "CONT002",
            dateDebut: new Date(2019, 5, 1), // 01/06/2019
            dateFin: null,
            nirSalarie: "2860363420456"
          },
          caracteristiques: {
            nature: "01", // CDI
            statutSalarie: "04", // Non cadre
            statutEmploi: "01",
            motifRecours: "",
            libelleEmploi: "Comptable",
            codeCSP: "543a", // Employés qualifiés
            complementCSP: "",
            dispositifPublic: "",
            complementDispositif: ""
          },
          tempsTravail: {
            modalite: "10", // Temps plein
            uniteMesure: "10", // Heure
            quotite: 35,
            tauxTempsPartiel: null,
            horairesContractuels: 151.67
          },
          classification: {
            conventionCollective: "1486",
            codeRegimeMaladie: "200",
            codeRegimeVieillesse: "200",
            positionConvention: "3",
            niveauRemuneration: "2",
            echelon: "1",
            coefficient: "90"
          },
          classificationEgalite: {
            categoriePoste: "Employés",
            niveau: "2",
            coefficientHierarchique: "90",
            niveauResponsabilite: "3",
            fonctionEncadrement: false,
            nombreSubordonnes: 0
          },
          remunerations: [
            {
              periode: {
                debut: new Date(annee, mois-1, 1),
                fin: new Date(annee, mois-1, 
                              new Date(annee, mois, 0).getDate())
              },
              numeroContrat: "CONT002",
              type: "001", // Rémunération brute
              montant: 3200,
              tauxRemuneration: null,
              nbHeures: 151.67,
              uniteMesure: "10"
            }
          ],
          evolutionRemuneration: {
            remunerationInitiale: 3000,
            remunerationActuelle: 3200,
            pourcentageEvolution: 6.67,
            augmentationDerniereAnnee: 200
          },
          evolutionCarriere: {
            posteInitial: "Comptable",
            posteActuel: "Comptable",
            niveauInitial: "2",
            niveauActuel: "2"
          },
          analytics: {
            dureeContrat: 730,
            coutTotal: 38400,
            salaireInitial: 3000,
            salaireFinal: 3200
          }
        },
        // Contrat 3
        {
          id: {
            numeroContrat: "CONT003",
            dateDebut: new Date(2015, 0, 15), // 15/01/2015
            dateFin: null,
            nirSalarie: "1740363420789"
          },
          caracteristiques: {
            nature: "01", // CDI
            statutSalarie: "03", // Cadre
            statutEmploi: "01",
            motifRecours: "",
            libelleEmploi: "Directeur Commercial",
            codeCSP: "374a", // Cadres d'entreprise
            complementCSP: "",
            dispositifPublic: "",
            complementDispositif: ""
          },
          tempsTravail: {
            modalite: "10", // Temps plein
            uniteMesure: "10", // Heure
            quotite: 35,
            tauxTempsPartiel: null,
            horairesContractuels: 151.67
          },
          classification: {
            conventionCollective: "1486",
            codeRegimeMaladie: "200",
            codeRegimeVieillesse: "200",
            positionConvention: "7",
            niveauRemuneration: "5",
            echelon: "3",
            coefficient: "120"
          },
          classificationEgalite: {
            categoriePoste: "Cadres",
            niveau: "5",
            coefficientHierarchique: "120",
            niveauResponsabilite: "1",
            fonctionEncadrement: true,
            nombreSubordonnes: 5
          },
          remunerations: [
            {
              periode: {
                debut: new Date(annee, mois-1, 1),
                fin: new Date(annee, mois-1, 
                              new Date(annee, mois, 0).getDate())
              },
              numeroContrat: "CONT003",
              type: "001", // Rémunération brute
              montant: 5200,
              tauxRemuneration: null,
              nbHeures: 151.67,
              uniteMesure: "10"
            }
          ],
          evolutionRemuneration: {
            remunerationInitiale: 4500,
            remunerationActuelle: 5200,
            pourcentageEvolution: 15.56,
            augmentationDerniereAnnee: 300
          },
          evolutionCarriere: {
            posteInitial: "Responsable Commercial",
            posteActuel: "Directeur Commercial",
            niveauInitial: "4",
            niveauActuel: "5"
          },
          analytics: {
            dureeContrat: 1825,
            coutTotal: 62400,
            salaireInitial: 4500,
            salaireFinal: 5200
          }
        },
        // Contrat 4
        {
          id: {
            numeroContrat: "CONT004",
            dateDebut: new Date(2020, 8, 1), // 01/09/2020
            dateFin: null,
            nirSalarie: "2910363420142"
          },
          caracteristiques: {
            nature: "01", // CDI
            statutSalarie: "04", // Non cadre
            statutEmploi: "01",
            motifRecours: "",
            libelleEmploi: "Assistante Administrative",
            codeCSP: "543b", // Employés qualifiés
            complementCSP: "",
            dispositifPublic: "",
            complementDispositif: ""
          },
          tempsTravail: {
            modalite: "20", // Temps partiel
            uniteMesure: "10", // Heure
            quotite: 28,
            tauxTempsPartiel: 80,
            horairesContractuels: 121.33
          },
          classification: {
            conventionCollective: "1486",
            codeRegimeMaladie: "200",
            codeRegimeVieillesse: "200",
            positionConvention: "2",
            niveauRemuneration: "1",
            echelon: "2",
            coefficient: "85"
          },
          classificationEgalite: {
            categoriePoste: "Employés",
            niveau: "1",
            coefficientHierarchique: "85",
            niveauResponsabilite: "4",
            fonctionEncadrement: false,
            nombreSubordonnes: 0
          },
          remunerations: [
            {
              periode: {
                debut: new Date(annee, mois-1, 1),
                fin: new Date(annee, mois-1, 
                              new Date(annee, mois, 0).getDate())
              },
              numeroContrat: "CONT004",
              type: "001", // Rémunération brute
              montant: 2800,
              tauxRemuneration: null,
              nbHeures: 121.33,
              uniteMesure: "10"
            }
          ],
          evolutionRemuneration: {
            remunerationInitiale: 2600,
            remunerationActuelle: 2800,
            pourcentageEvolution: 7.69,
            augmentationDerniereAnnee: 200
          },
          evolutionCarriere: {
            posteInitial: "Assistante Administrative",
            posteActuel: "Assistante Administrative",
            niveauInitial: "1",
            niveauActuel: "1"
          },
          analytics: {
            dureeContrat: 365,
            coutTotal: 33600,
            salaireInitial: 2600,
            salaireFinal: 2800
          }
        }
      ],
      arrets: [
        {
          motif: "01", // Maladie
          dateDebut: new Date(annee, mois-1, 10), // 10 du mois courant
          dateFin: new Date(annee, mois-1, 15),   // 15 du mois courant
          dateReprise: new Date(annee, mois-1, 16), // 16 du mois courant
          motifReprise: "01",
          nirSalarie: "1840363420123",
          numeroContrat: "CONT001",
          subrogation: {
            debut: null,
            fin: null,
            montant: null
          },
          suiviEgalite: {
            estCongeMaternite: false,
            estCongePaternite: false,
            estCongeAdoption: false,
            augmentationAuRetour: null
          }
        },
        {
          motif: "02", // Maternité
          dateDebut: new Date(annee, mois-3, 15), // 15 du mois-3
          dateFin: new Date(annee, mois-1, 20),   // 20 du mois courant
          dateReprise: new Date(annee, mois-1, 21), // 21 du mois courant
          motifReprise: "01",
          nirSalarie: "2910363420142",
          numeroContrat: "CONT004",
          subrogation: {
            debut: null,
            fin: null,
            montant: null
          },
          suiviEgalite: {
            estCongeMaternite: true,
            estCongePaternite: false,
            estCongeAdoption: false,
            augmentationAuRetour: true,
            dateAugmentation: new Date(annee, mois-1, 25),
            montantAugmentation: 200
          }
        }
      ]
    };
    
    return model;
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
    const dsnText = `S20.G00.05.005,${annee}
S20.G00.05.006,${mois}
S21.G00.06.001,${annee}${mm}
S21.G00.06.002,${annee}-${mm}
S21.G00.06.003,12
S21.G00.06.004,DSN Mensuelle

S21.G00.11.001,${formData.codeDossier || "123456789"}
S21.G00.11.002,${formData.codeEtablissement || "00001"}
S21.G00.11.003,Entreprise Démo DSN
S21.G00.11.004,1

S21.G00.30.001,1840363420123
S21.G00.30.002,DUPONT
S21.G00.30.004,Jean
S21.G00.30.005,01
S21.G00.30.006,15011980
S21.G00.30.019,EMP001
S21.G00.30.022,01
S21.G00.30.024,200

S21.G00.40.001,01032018
S21.G00.40.002,03
S21.G00.40.004,353a
S21.G00.40.006,Développeur informatique
S21.G00.40.007,01
S21.G00.40.009,CONT001
S21.G00.40.010,
S21.G00.40.011,10
S21.G00.40.012,35
S21.G00.40.013,10

S21.G00.51.001,01${mm}${annee}
S21.G00.51.002,${new Date(annee, mois, 0).getDate()}${mm}${annee}
S21.G00.51.010,CONT001
S21.G00.51.011,001
S21.G00.51.013,3950.00

S21.G00.30.001,2860363420456
S21.G00.30.002,MARTIN
S21.G00.30.003,DUBOIS
S21.G00.30.004,Sophie
S21.G00.30.005,02
S21.G00.30.006,28031986
S21.G00.30.019,EMP002
S21.G00.30.022,01
S21.G00.30.024,200

S21.G00.40.001,01062019
S21.G00.40.002,04
S21.G00.40.004,543a
S21.G00.40.006,Comptable
S21.G00.40.007,01
S21.G00.40.009,CONT002
S21.G00.40.010,
S21.G00.40.011,10
S21.G00.40.012,35
S21.G00.40.013,10

S21.G00.51.001,01${mm}${annee}
S21.G00.51.002,${new Date(annee, mois, 0).getDate()}${mm}${annee}
S21.G00.51.010,CONT002
S21.G00.51.011,001
S21.G00.51.013,3200.00

S21.G00.30.001,1740363420789
S21.G00.30.002,BERNARD
S21.G00.30.004,Pierre
S21.G00.30.005,01
S21.G00.30.006,10091974
S21.G00.30.019,EMP003
S21.G00.30.022,01
S21.G00.30.024,200

S21.G00.40.001,15012015
S21.G00.40.002,03
S21.G00.40.004,374a
S21.G00.40.006,Directeur Commercial
S21.G00.40.007,01
S21.G00.40.009,CONT003
S21.G00.40.010,
S21.G00.40.011,10
S21.G00.40.012,35
S21.G00.40.013,10

S21.G00.51.001,01${mm}${annee}
S21.G00.51.002,${new Date(annee, mois, 0).getDate()}${mm}${annee}
S21.G00.51.010,CONT003
S21.G00.51.011,001
S21.G00.51.013,5200.00

S21.G00.30.001,2910363420142
S21.G00.30.002,PETIT
S21.G00.30.004,Marie
S21.G00.30.005,02
S21.G00.30.006,22071991
S21.G00.30.019,EMP004
S21.G00.30.022,01
S21.G00.30.024,200

S21.G00.40.001,01092020
S21.G00.40.002,04
S21.G00.40.004,543b
S21.G00.40.006,Assistante Administrative
S21.G00.40.007,01
S21.G00.40.009,CONT004
S21.G00.40.010,
S21.G00.40.011,10
S21.G00.40.012,28
S21.G00.40.013,20

S21.G00.51.001,01${mm}${annee}
S21.G00.51.002,${new Date(annee, mois, 0).getDate()}${mm}${annee}
S21.G00.51.010,CONT004
S21.G00.51.011,001
S21.G00.51.013,2800.00

S21.G00.60.001,01
S21.G00.60.002,10${mm}${annee}
S21.G00.60.003,15${mm}${annee}
S21.G00.60.010,16${mm}${annee}
S21.G00.60.011,01

S21.G00.60.001,02
S21.G00.60.002,15${annee}${mm-3 < 10 ? '0' + (mm-3) : mm-3}
S21.G00.60.003,20${mm}${annee}
S21.G00.60.010,21${mm}${annee}
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
      }
    } catch (e) {
      Logger.log("Erreur de validation API: " + e.toString());
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
    validateAPICredentials
  };
})();
