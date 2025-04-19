/**
 * Yoozak - Module d'initialisation des commandes
 * 
 * Ce script contient les fonctions pour récupérer et traiter les données depuis Google Sheets
 * Version: 1.0
 * Date: 2023
 */

// La configuration globale est déjà définie dans Main.js
// Utiliser celle qui existe déjà

/**
 * Fonction pour journaliser les erreurs
 * Remplace la fonction globale si elle n'est pas disponible
 * 
 * @param {Error} error L'erreur à journaliser
 * @param {string} source La source de l'erreur
 */
function logError(error, source) {
  try {
    // Si la fonction globale existe, l'utiliser
    if (typeof this.logError === 'function' && this !== window) {
      this.logError(error, source);
      return;
    }
    
    // Sinon, journaliser dans la console
    console.error(`[${source}] ${error.toString()}`);
    
    // Essayer d'écrire dans les logs
    try {
      Logger.log(`ERREUR [${source}]: ${error.toString()}`);
    } catch (e) {
      // Ignorer si Logger n'est pas disponible
    }
  } catch (e) {
    // En cas d'erreur, simplement ignorer
    console.error("Erreur lors de la journalisation:", e);
  }
}

/**
 * Récupère les données depuis le Google Sheet externe spécifié
 * 
 * @return {Object} Les données récupérées sous forme d'objet
 */
function recupererDonneesExterne() {
  try {
    // URL du fichier Google Sheet externe
    const url = CONFIG.URL_EXTERNE || "https://docs.google.com/spreadsheets/d/1OK2Ndvc9dyUV99sJDLJ3tYSvtECijG-CM0QlGClYduU/edit";
    
    // Ouvrir le fichier externe en utilisant son URL
    const fichierExterne = SpreadsheetApp.openByUrl(url);
    
    // Récupérer les noms de toutes les feuilles
    const feuilles = fichierExterne.getSheets();
    const nomsFeuilles = feuilles.map(feuille => feuille.getName());
    
    // Récupérer les données de chaque feuille
    const donnees = {};
    
    feuilles.forEach(feuille => {
      const nomFeuille = feuille.getName();
      donnees[nomFeuille] = feuille.getDataRange().getValues();
    });
    
    return {
      success: true,
      nomsFeuilles: nomsFeuilles,
      donnees: donnees
    };
  } catch (error) {
    logError(error, 'recupererDonneesExterne');
    return {
      success: false,
      message: error.toString()
    };
  }
}

/**
 * Récupère les données d'une feuille spécifique du Google Sheet externe
 * 
 * @param {string} nomFeuille - Le nom de la feuille à récupérer
 * @return {Object} Les données récupérées sous forme d'objet
 */
function recupererFeuilleSpecifique(nomFeuille) {
  try {
    // URL du fichier Google Sheet externe
    const url = CONFIG.URL_EXTERNE || "https://docs.google.com/spreadsheets/d/1OK2Ndvc9dyUV99sJDLJ3tYSvtECijG-CM0QlGClYduU/edit";
    
    // Ouvrir le fichier externe en utilisant son URL
    const fichierExterne = SpreadsheetApp.openByUrl(url);
    
    // Récupérer la feuille spécifiée
    const feuille = fichierExterne.getSheetByName(nomFeuille);
    
    if (!feuille) {
      return {
        success: false,
        message: `La feuille "${nomFeuille}" n'existe pas dans le fichier.`
      };
    }
    
    // Récupérer toutes les données de la feuille
    const donnees = feuille.getDataRange().getValues();
    
    return {
      success: true,
      nomFeuille: nomFeuille,
      donnees: donnees
    };
  } catch (error) {
    logError(error, `recupererFeuilleSpecifique(${nomFeuille})`);
    return {
      success: false,
      message: error.toString()
    };
  }
}

/**
 * Importe les données d'une feuille spécifique du Google Sheet externe vers la feuille active
 * 
 * @param {string} nomFeuille - Le nom de la feuille à importer
 * @param {string} feuilleDestination - Le nom de la feuille de destination dans le fichier actif
 * @return {Object} Résultat de l'importation
 */
function importerDonnees(nomFeuille, feuilleDestination) {
  try {
    // Récupérer les données de la feuille spécifiée
    const resultat = recupererFeuilleSpecifique(nomFeuille);
    
    if (!resultat.success) {
      return resultat;
    }
    
    // Récupérer le fichier actif
    const fichierActif = SpreadsheetApp.getActiveSpreadsheet();
    
    // Vérifier si la feuille de destination existe, sinon la créer
    let feuilleActive = fichierActif.getSheetByName(feuilleDestination);
    if (!feuilleActive) {
      feuilleActive = fichierActif.insertSheet(feuilleDestination);
    } else {
      // Effacer les données existantes
      feuilleActive.clear();
    }
    
    // Vérifier si des données ont été récupérées
    if (resultat.donnees.length === 0) {
      return {
        success: false,
        message: "Aucune donnée n'a été trouvée dans la feuille spécifiée."
      };
    }
    
    // Importer les données
    const range = feuilleActive.getRange(1, 1, resultat.donnees.length, resultat.donnees[0].length);
    range.setValues(resultat.donnees);
    
    return {
      success: true,
      message: `Les données ont été importées avec succès vers la feuille "${feuilleDestination}".`,
      nbLignes: resultat.donnees.length,
      nbColonnes: resultat.donnees[0].length
    };
  } catch (error) {
    logError(error, `importerDonnees(${nomFeuille}, ${feuilleDestination})`);
    return {
      success: false,
      message: error.toString()
    };
  }
}

/**
 * Fonction pour récupérer la liste des feuilles du Google Sheet externe
 * 
 * @return {Object} La liste des noms de feuilles
 */
function getListeFeuilles() {
  try {
    console.log("Récupération de la liste des feuilles disponibles...");
    
    // Récupérer le fichier externe
    const fichierExterne = SpreadsheetApp.openByUrl(CONFIG.URL_EXTERNE);
    if (!fichierExterne) {
      return { 
        success: false, 
        message: "Impossible d'accéder au fichier externe" 
      };
    }
    
    // Récupérer toutes les feuilles
    const feuilles = fichierExterne.getSheets();
    const nomsFeuilles = feuilles.map(feuille => feuille.getName());
    
    console.log(`Feuilles trouvées : ${nomsFeuilles.join(', ')}`);
    
    return {
      success: true,
      message: `${nomsFeuilles.length} feuilles trouvées`,
      nomsFeuilles: nomsFeuilles
    };
  } catch (error) {
    console.error("Erreur lors de la récupération des feuilles :", error);
    return {
      success: false,
      message: "Erreur lors de la récupération des feuilles : " + error.toString()
    };
  }
}

/**
 * Fonction simplifiée pour récupérer un aperçu des données d'une feuille
 * Cette fonction est optimisée pour l'aperçu et retourne directement un HTML formaté
 * 
 * @param {string} nomFeuille - Le nom de la feuille à récupérer
 * @return {string} Contenu HTML formaté pour l'aperçu
 */
function getApercuHTML(nomFeuille) {
  try {
    console.log(`Génération de l'aperçu HTML pour la feuille ${nomFeuille}...`);
    
    // Récupérer le fichier externe
    const fichierExterne = SpreadsheetApp.openByUrl(CONFIG.URL_EXTERNE);
    if (!fichierExterne) {
      return "Erreur : Impossible d'accéder au fichier externe";
    }
    
    // Récupérer la feuille spécifiée
    const feuille = fichierExterne.getSheetByName(nomFeuille);
    if (!feuille) {
      return `Erreur : Feuille ${nomFeuille} introuvable`;
    }
    
    // Récupérer les données de la feuille
    const donneesRange = feuille.getDataRange();
    const donnees = donneesRange.getValues();
    
    if (donnees.length <= 1) {
      return "Erreur : La feuille ne contient pas assez de données";
    }
    
    // Générer le tableau HTML
    let html = `
      <div class="card mb-3">
        <div class="card-header bg-info text-white">Aperçu des données (${donnees.length - 1} lignes)</div>
        <div class="card-body">
          <div class="table-responsive">
            <table class="table table-striped table-bordered">
              <thead class="thead-dark">
                <tr>`;
    
    // En-têtes
    const headers = donnees[0];
    headers.forEach(header => {
      html += `<th>${header}</th>`;
    });
    
    html += `
                </tr>
              </thead>
              <tbody>`;
    
    // Lignes de données (limiter à 10 pour l'aperçu)
    const maxRows = Math.min(donnees.length, 11); // En-tête + 10 lignes de données
    for (let i = 1; i < maxRows; i++) {
      html += '<tr>';
      donnees[i].forEach(cell => {
        html += `<td>${cell}</td>`;
      });
      html += '</tr>';
    }
    
    html += `
              </tbody>
            </table>
          </div>`;
    
    // Ajouter un message si toutes les lignes ne sont pas affichées
    if (donnees.length > 11) {
      html += `<p class="text-info">Note : Seules les 10 premières lignes sont affichées sur un total de ${donnees.length - 1}.</p>`;
    }
    
    html += `
        </div>
      </div>`;
    
    return html;
  } catch (error) {
    console.error("Erreur lors de la génération de l'aperçu HTML :", error);
    return "Erreur : " + error.toString();
  }
}

/**
 * Fonction pour traiter les données du fichier importé et les préparer pour CMD initiale
 * Selon le cahier des charges, cette fonction doit:
 * 1. Insérer les lignes depuis Shopify/Youcan dans CMD initiale
 * 2. Formater les données téléphone et date
 * 3. Identifier les doublons mais les garder dans la feuille
 * 4. Exporter les cas d'erreur dans CMD problème
 * 
 * @param {string} sourceSheet - Le nom de la feuille source (shopify-orders ou Youcan-Orders)
 * @return {Object} Résultat du traitement
 */
function traiterDonneesCMDInitiale(sourceSheet) {
  try {
    // Récupérer le classeur actif
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    
    // 1. Vérifier que les feuilles nécessaires existent
    let feuilleSource = ss.getSheetByName(sourceSheet);
    let feuilleInitiale = ss.getSheetByName(CONFIG.SHEETS.INITIALE);
    let feuilleProbleme = ss.getSheetByName(CONFIG.SHEETS.PROBLEM);
    
    if (!feuilleSource) {
      return {
        success: false,
        message: `La feuille source "${sourceSheet}" n'existe pas.`
      };
    }
    
    if (!feuilleInitiale) {
      feuilleInitiale = ss.insertSheet(CONFIG.SHEETS.INITIALE);
      // Ajouter les en-têtes pour CMD initiale
      const enTetesInitiale = [
        "N° Commande", "Source", "Date", "Nom", "Prénom", "Téléphone", "Adresse", 
        "Ville", "Région", "Produits", "Opérateur", "Statut", "Date Affectation", "Remarque"
      ];
      feuilleInitiale.getRange(1, 1, 1, enTetesInitiale.length).setValues([enTetesInitiale]);
    }
    
    if (!feuilleProbleme) {
      feuilleProbleme = ss.insertSheet(CONFIG.SHEETS.PROBLEM);
      // Ajouter les en-têtes pour CMD Problem
      const enTetesProbleme = [
        "N° Commande", "Source", "Date", "Nom", "Prénom", "Téléphone", "Adresse", 
        "Ville", "Région", "Produits", "Type Problème", "Description"
      ];
      feuilleProbleme.getRange(1, 1, 1, enTetesProbleme.length).setValues([enTetesProbleme]);
    }
    
    // 2. Récupérer les données de la source
    const donneesSource = feuilleSource.getDataRange().getValues();
    if (donneesSource.length <= 1) {
      return {
        success: false,
        message: "Aucune donnée n'a été trouvée dans la feuille source."
      };
    }
    
    // Récupérer les en-têtes
    const enTetes = donneesSource[0];
    
    // Déterminer la source (Shopify ou Youcan)
    const source = sourceSheet.toLowerCase().includes("shopify") ? CONFIG.SOURCES.SHOPIFY : CONFIG.SOURCES.YOUCAN;
    
    // 3. Préparer les données à ajouter dans CMD initiale
    const donneesTraitees = [];
    const problemes = [];
    const commandesExistantes = new Set();
    
    // Récupérer les commandes existantes pour identifier les doublons
    const donneesInitiale = feuilleInitiale.getDataRange().getValues();
    for (let i = 1; i < donneesInitiale.length; i++) {
      commandesExistantes.add(donneesInitiale[i][0]); // Ajouter le numéro de commande
    }
    
    // Traiter chaque ligne de données
    for (let i = 1; i < donneesSource.length; i++) {
      const ligne = donneesSource[i];
      
      // Extraction des données selon la source
      let numeroCommande, date, nom, prenom, telephone, adresse, ville, region, produits;
      
      if (source === CONFIG.SOURCES.SHOPIFY) {
        // Extraction des données Shopify
        numeroCommande = ligne[getColIndex(enTetes, "Order Number")] || "";
        date = ligne[getColIndex(enTetes, "Created at")] || new Date();
        const nomComplet = (ligne[getColIndex(enTetes, "Customer Name")] || "").split(" ");
        nom = nomComplet.length > 0 ? nomComplet[0] : "";
        prenom = nomComplet.length > 1 ? nomComplet.slice(1).join(" ") : "";
        telephone = ligne[getColIndex(enTetes, "Phone")] || "";
        adresse = ligne[getColIndex(enTetes, "Shipping Address")] || "";
        ville = ligne[getColIndex(enTetes, "Shipping City")] || "";
        region = determinerRegion(ville); // Fonction à implémenter pour déterminer la région en fonction de la ville
        produits = ligne[getColIndex(enTetes, "Lineitem name")] || "";
      } else {
        // Extraction des données Youcan
        numeroCommande = ligne[getColIndex(enTetes, "Order ID")] || "";
        date = ligne[getColIndex(enTetes, "Date Order")] || new Date();
        nom = ligne[getColIndex(enTetes, "First Name")] || "";
        prenom = ligne[getColIndex(enTetes, "Last Name")] || "";
        telephone = ligne[getColIndex(enTetes, "Phone")] || "";
        adresse = ligne[getColIndex(enTetes, "Address")] || "";
        ville = ligne[getColIndex(enTetes, "City")] || "";
        region = determinerRegion(ville);
        produits = ligne[getColIndex(enTetes, "Product Name")] || "";
      }
      
      // Formatage des données
      telephone = formaterTelephone(telephone);
      date = formaterDate(date);
      
      // Vérifier si c'est un doublon
      let estDoublon = commandesExistantes.has(numeroCommande);
      let remarque = "";
      
      // Vérification des données obligatoires
      if (!numeroCommande || !telephone || !ville) {
        // Données incomplètes, l'ajouter aux problèmes
        problemes.push([
          numeroCommande, source, date, nom, prenom, telephone, adresse, ville, region, produits,
          "Données Incomplètes", "Manque d'informations essentielles"
        ]);
        continue;
      }
      
      // Préparer la ligne pour CMD initiale
      // Si c'est un doublon, on l'ajoute quand même mais avec une remarque et on l'ajoute aussi aux problèmes
      if (estDoublon) {
        remarque = "Doublon";
        
        // Ajouter aux problèmes pour référence
        problemes.push([
          numeroCommande, source, date, nom, prenom, telephone, adresse, ville, region, produits,
          "Doublon", "Commande déjà existante mais conservée"
        ]);
      }
      
      // Tout est bon, ajouter à CMD initiale avec statut "Non affectée"
      donneesTraitees.push([
        numeroCommande, source, date, nom, prenom, telephone, adresse, ville, region, produits,
        "", CONFIG.STATUTS.NON_AFFECTEE, "", remarque // Ajout de la remarque de doublon si applicable
      ]);
      
      // Ajouter à l'ensemble des commandes existantes pour éviter les doublons futurs
      commandesExistantes.add(numeroCommande);
    }
    
    // 4. Ajouter les données traitées à CMD initiale
    if (donneesTraitees.length > 0) {
      const lastRow = feuilleInitiale.getLastRow();
      feuilleInitiale.getRange(lastRow + 1, 1, donneesTraitees.length, donneesTraitees[0].length)
                     .setValues(donneesTraitees);
    }
    
    // 5. Ajouter les problèmes à CMD Problem
    if (problemes.length > 0) {
      const lastRow = feuilleProbleme.getLastRow();
      feuilleProbleme.getRange(lastRow + 1, 1, problemes.length, problemes[0].length)
                    .setValues(problemes);
    }
    
    return {
      success: true,
      message: "Les données ont été traitées avec succès.",
      commandesTraitees: donneesTraitees.length,
      problemes: problemes.length
    };
    
  } catch (error) {
    logError(error, 'traiterDonneesCMDInitiale');
    return {
      success: false,
      message: "Erreur lors du traitement des données: " + error.toString()
    };
  }
}

/**
 * Utilitaire pour obtenir l'index d'une colonne à partir de son nom
 * 
 * @param {Array} enTetes - Tableau contenant les en-têtes des colonnes
 * @param {string} nomColonne - Nom de la colonne recherchée
 * @return {number} Index de la colonne (-1 si non trouvée)
 */
function getColIndex(enTetes, nomColonne) {
  return enTetes.findIndex(col => col === nomColonne);
}

/**
 * Fonction pour déterminer la région à partir de la ville
 * Cette fonction devrait utiliser une référence des villes/régions dans CONFIG
 * 
 * @param {string} ville - Nom de la ville
 * @return {string} Nom de la région
 */
function determinerRegion(ville) {
  try {
    // Récupérer le classeur actif
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    
    // Récupérer la feuille de configuration
    const feuilleConfig = ss.getSheetByName(CONFIG.SHEETS.CONFIG);
    
    if (!feuilleConfig) {
      return "Région inconnue";
    }
    
    // Trouver l'onglet CMD Region dans la feuille de configuration
    const donneesConfig = feuilleConfig.getDataRange().getValues();
    let indexRegion = -1;
    
    for (let i = 0; i < donneesConfig.length; i++) {
      if (donneesConfig[i][0] === "CMD Region") {
        indexRegion = i;
        break;
      }
    }
    
    if (indexRegion === -1) {
      return "Région inconnue";
    }
    
    // Parcourir les lignes suivantes pour trouver la ville
    for (let i = indexRegion + 1; i < donneesConfig.length; i++) {
      // Si on trouve une nouvelle section, arrêter la recherche
      if (donneesConfig[i][0] && donneesConfig[i][0].startsWith("CMD")) {
        break;
      }
      
      // Vérifier si la ville correspond
      if (donneesConfig[i][0] && donneesConfig[i][0].toLowerCase() === ville.toLowerCase()) {
        return donneesConfig[i][1] || "Région inconnue";
      }
    }
    
    // Ville non trouvée
    return "Région inconnue";
    
  } catch (error) {
    logError(error, 'determinerRegion');
    return "Région inconnue";
  }
}

/**
 * Fonction pour formater un numéro de téléphone marocain
 * 
 * @param {string} telephone - Numéro de téléphone à formater
 * @return {string} Numéro de téléphone formaté
 */
function formaterTelephone(telephone) {
  if (!telephone) return "";
  
  // Supprimer tous les caractères non numériques
  let numero = telephone.toString().replace(/\D/g, '');
  
  // Si commence par 0, remplacer par +212
  if (numero.startsWith('0')) {
    numero = "212" + numero.substring(1);
  }
  
  // Si ne commence pas par 212, ajouter 212
  if (!numero.startsWith('212')) {
    numero = "212" + numero;
  }
  
  // Formater avec +
  if (!numero.startsWith('+')) {
    numero = "+" + numero;
  }
  
  return numero;
}

/**
 * Fonction pour formater une date
 * 
 * @param {Date|string} date - Date à formater
 * @return {string} Date formatée (JJ/MM/AAAA)
 */
function formaterDate(date) {
  if (!date) return "";
  
  try {
    // Si la date est déjà au format String, la convertir en objet Date
    if (typeof date === 'string') {
      date = new Date(date);
    }
    
    const jour = date.getDate().toString().padStart(2, '0');
    const mois = (date.getMonth() + 1).toString().padStart(2, '0');
    const annee = date.getFullYear();
    
    return `${jour}/${mois}/${annee}`;
  } catch (error) {
    logError(error, 'formaterDate');
    return date.toString();
  }
}

/**
 * Fonction pour fusionner automatiquement deux feuilles (Shopify et Youcan)
 * 
 * @return {Object} Résultat de la fusion
 */
function fusionnerFeuillesExterne() {
  try {
    // URL du fichier Google Sheet externe
    const url = CONFIG.URL_EXTERNE || "https://docs.google.com/spreadsheets/d/1OK2Ndvc9dyUV99sJDLJ3tYSvtECijG-CM0QlGClYduU/edit";
    
    // Ouvrir le fichier externe en utilisant son URL
    const fichierExterne = SpreadsheetApp.openByUrl(url);
    
    // Récupérer les noms de toutes les feuilles
    const feuilles = fichierExterne.getSheets();
    const nomsFeuilles = feuilles.map(feuille => feuille.getName());
    
    // Filtrer pour obtenir uniquement les feuilles Shopify et Youcan
    const feuillesShopify = nomsFeuilles.filter(nom => nom.toLowerCase().includes('shopify'));
    const feuillesYoucan = nomsFeuilles.filter(nom => nom.toLowerCase().includes('youcan'));
    
    // Journaliser les feuilles trouvées pour le débogage
    Logger.log("Feuilles Shopify trouvées: " + JSON.stringify(feuillesShopify));
    Logger.log("Feuilles Youcan trouvées: " + JSON.stringify(feuillesYoucan));
    
    // Vérifier si nous avons les deux types de feuilles
    if (feuillesShopify.length > 0 && feuillesYoucan.length > 0) {
      // Sélectionner les premières feuilles de chaque type
      const nomFeuilleShopify = feuillesShopify[0];
      const nomFeuilleYoucan = feuillesYoucan[0];
      
      Logger.log("Feuille Shopify sélectionnée: " + nomFeuilleShopify);
      Logger.log("Feuille Youcan sélectionnée: " + nomFeuilleYoucan);
      
      // Vérifier que les feuilles existent bien avant de les utiliser
      let feuilleShopify = null;
      let feuilleYoucan = null;
      
      try {
        feuilleShopify = fichierExterne.getSheetByName(nomFeuilleShopify);
        if (!feuilleShopify) {
          return {
            success: false,
            message: `La feuille Shopify "${nomFeuilleShopify}" n'a pas pu être trouvée.`
          };
        }
      } catch (e) {
        return {
          success: false,
          message: `Erreur lors de l'accès à la feuille Shopify "${nomFeuilleShopify}": ${e.toString()}`
        };
      }
      
      try {
        feuilleYoucan = fichierExterne.getSheetByName(nomFeuilleYoucan);
        if (!feuilleYoucan) {
          return {
            success: false,
            message: `La feuille Youcan "${nomFeuilleYoucan}" n'a pas pu être trouvée.`
          };
        }
      } catch (e) {
        return {
          success: false,
          message: `Erreur lors de l'accès à la feuille Youcan "${nomFeuilleYoucan}": ${e.toString()}`
        };
      }
      
      // Nom de la feuille fusionnée - doit être 'CMDinit'
      const nomFeuilleFusionnee = "CMDinit";
      
      Logger.log("Tentative de création de la feuille fusionnée: " + nomFeuilleFusionnee);
      
      // Vérifier si la feuille fusionnée existe déjà
      let feuilleFusionnee = null;
      try {
        feuilleFusionnee = fichierExterne.getSheetByName(nomFeuilleFusionnee);
        if (feuilleFusionnee) {
          // Si elle existe, la supprimer pour la recréer
          Logger.log("Suppression de la feuille fusionnée existante");
          fichierExterne.deleteSheet(feuilleFusionnee);
        }
      } catch (e) {
        // La feuille n'existe pas, on continue
        Logger.log("La feuille fusionnée n'existe pas encore: " + e.toString());
      }
      
      // Créer la nouvelle feuille fusionnée
      try {
        feuilleFusionnee = fichierExterne.insertSheet(nomFeuilleFusionnee);
        Logger.log("Nouvelle feuille fusionnée créée avec succès");
      } catch (e) {
        Logger.log("Erreur lors de la création de la feuille fusionnée: " + e.toString());
        return {
          success: false,
          message: `Erreur lors de la création de la feuille fusionnée: ${e.toString()}`
        };
      }
      
      // Récupérer les données des deux feuilles
      let donneesShopify = [];
      let donneesYoucan = [];
      
      try {
        donneesShopify = feuilleShopify.getDataRange().getValues();
        Logger.log("Données Shopify récupérées: " + donneesShopify.length + " lignes");
      } catch (e) {
        Logger.log("Erreur lors de la récupération des données Shopify: " + e.toString());
        return {
          success: false,
          message: `Erreur lors de la récupération des données Shopify: ${e.toString()}`
        };
      }
      
      try {
        donneesYoucan = feuilleYoucan.getDataRange().getValues();
        Logger.log("Données Youcan récupérées: " + donneesYoucan.length + " lignes");
      } catch (e) {
        Logger.log("Erreur lors de la récupération des données Youcan: " + e.toString());
        return {
          success: false,
          message: `Erreur lors de la récupération des données Youcan: ${e.toString()}`
        };
      }
      
      // Vérifier si des données ont été récupérées
      if (donneesShopify.length === 0 && donneesYoucan.length === 0) {
        Logger.log("Aucune donnée trouvée dans les feuilles");
        return {
          success: false,
          message: "Aucune donnée n'a été trouvée dans les feuilles Shopify et Youcan."
        };
      }
      
      try {
        // Déterminer le nombre maximum de colonnes entre les deux ensembles de données
        const maxColonnesShopify = donneesShopify.length > 0 ? donneesShopify[0].length : 0;
        const maxColonnesYoucan = donneesYoucan.length > 0 ? donneesYoucan[0].length : 0;
        const nombreColonnes = Math.max(maxColonnesShopify, maxColonnesYoucan);
        
        Logger.log("Nombre de colonnes: Shopify=" + maxColonnesShopify + ", Youcan=" + maxColonnesYoucan + ", Max=" + nombreColonnes);
        
        // Créer les en-têtes fusionnés si les deux feuilles ont des en-têtes
        let enTetes = [];
        if (donneesShopify.length > 0) {
          enTetes = donneesShopify[0].slice(0);
        } else if (donneesYoucan.length > 0) {
          enTetes = donneesYoucan[0].slice(0);
        }
        
        // S'assurer que l'en-tête a le bon nombre de colonnes
        while (enTetes.length < nombreColonnes) {
          enTetes.push(""); // Ajouter des colonnes vides si nécessaire
        }
        
        Logger.log("Écriture de l'en-tête avec " + nombreColonnes + " colonnes");
        
        // Écrire l'en-tête
        feuilleFusionnee.getRange(1, 1, 1, nombreColonnes).setValues([enTetes]);
        
        // Fonction pour s'assurer que chaque ligne a le bon nombre de colonnes
        function normaliserLigne(ligne, nombreColonnes) {
          const nouvelleLigne = ligne.slice(0); // Copier la ligne
          while (nouvelleLigne.length < nombreColonnes) {
            nouvelleLigne.push(""); // Ajouter des cellules vides si nécessaire
          }
          return nouvelleLigne;
        }
        
        // Compteur pour la position d'insertion
        let ligneActuelle = 2;
        
        // Copier les données de Shopify (sans l'en-tête)
        if (donneesShopify.length > 1) {
          Logger.log("Préparation de " + (donneesShopify.length - 1) + " lignes de données Shopify");
          
          // Normaliser chaque ligne pour qu'elle ait le bon nombre de colonnes
          const dataShopify = [];
          for (let i = 1; i < donneesShopify.length; i++) {
            const ligneTrns = normaliserLigne(donneesShopify[i], nombreColonnes);
            // Ajouter une source pour identifier l'origine
            if (enTetes.includes("Source")) {
              const indexSource = enTetes.indexOf("Source");
              ligneTrns[indexSource] = "Shopify";
            }
            dataShopify.push(ligneTrns);
          }
          
          Logger.log("Copie de " + dataShopify.length + " lignes de données Shopify");
          feuilleFusionnee.getRange(ligneActuelle, 1, dataShopify.length, nombreColonnes).setValues(dataShopify);
          ligneActuelle += dataShopify.length;
        }
        
        // Copier les données de Youcan (sans l'en-tête)
        if (donneesYoucan.length > 1) {
          Logger.log("Préparation de " + (donneesYoucan.length - 1) + " lignes de données Youcan");
          
          // Normaliser chaque ligne pour qu'elle ait le bon nombre de colonnes
          const dataYoucan = [];
          for (let i = 1; i < donneesYoucan.length; i++) {
            const ligneTrns = normaliserLigne(donneesYoucan[i], nombreColonnes);
            // Ajouter une source pour identifier l'origine
            if (enTetes.includes("Source")) {
              const indexSource = enTetes.indexOf("Source");
              ligneTrns[indexSource] = "Youcan";
            }
            dataYoucan.push(ligneTrns);
          }
          
          Logger.log("Copie de " + dataYoucan.length + " lignes de données Youcan");
          feuilleFusionnee.getRange(ligneActuelle, 1, dataYoucan.length, nombreColonnes).setValues(dataYoucan);
        }
        
        // Appliquer une mise en forme conditionnelle pour les doublons
        try {
          // Déterminer l'index de la colonne des numéros de commande
          const indexNumCommande = enTetes.indexOf("Order Number") >= 0 ? 
                                 enTetes.indexOf("Order Number") : 
                                 enTetes.indexOf("Order ID") >= 0 ?
                                 enTetes.indexOf("Order ID") : 0;

          if (indexNumCommande >= 0) {
            // Créer une règle pour colorer les doublons
            const range = feuilleFusionnee.getRange(2, 1, feuilleFusionnee.getLastRow() - 1, nombreColonnes);
            const rule = SpreadsheetApp.newConditionalFormatRule()
                .whenFormulaSatisfied(`=COUNTIF($${String.fromCharCode(65 + indexNumCommande)}$2:$${String.fromCharCode(65 + indexNumCommande)}$${feuilleFusionnee.getLastRow()}, $${String.fromCharCode(65 + indexNumCommande)}2)>1`)
                .setBackground("#F4CCCC") // Rouge clair
                .setRanges([range])
                .build();
            
            const rules = feuilleFusionnee.getConditionalFormatRules();
            rules.push(rule);
            feuilleFusionnee.setConditionalFormatRules(rules);
            Logger.log("Mise en forme conditionnelle appliquée pour les doublons");
          }
        } catch (e) {
          Logger.log("Erreur lors de l'application de la mise en forme conditionnelle: " + e.toString());
          // On continue malgré l'erreur, ce n'est pas bloquant
        }
        
        Logger.log("Fusion des données terminée avec succès");
      } catch (e) {
        Logger.log("Erreur lors de la fusion des données: " + e.toString());
        return {
          success: false,
          message: `Erreur lors de la fusion des données: ${e.toString()}`
        };
      }
      
      return {
        success: true,
        message: "Les feuilles ont été fusionnées avec succès.",
        feuilleFusionnee: nomFeuilleFusionnee,
        totalLignes: (donneesShopify.length > 1 ? donneesShopify.length - 1 : 0) + 
                     (donneesYoucan.length > 1 ? donneesYoucan.length - 1 : 0)
      };
    } else if (feuillesShopify.length > 0) {
      // Seulement Shopify est présent
      return {
        success: true,
        message: "Seule la feuille Shopify est présente, aucune fusion nécessaire.",
        feuilleUnique: feuillesShopify[0]
      };
    } else if (feuillesYoucan.length > 0) {
      // Seulement Youcan est présent
      return {
        success: true,
        message: "Seule la feuille Youcan est présente, aucune fusion nécessaire.",
        feuilleUnique: feuillesYoucan[0]
      };
    } else {
      // Aucune feuille n'a été trouvée
      return {
        success: false,
        message: "Aucune feuille Shopify ou Youcan n'a été trouvée dans le document."
      };
    }
  } catch (error) {
    logError(error, 'fusionnerFeuillesExterne');
    Logger.log("Erreur globale dans fusionnerFeuillesExterne: " + error.toString());
    return {
      success: false,
      message: "Erreur lors de la fusion des feuilles: " + error.toString()
    };
  }
}

/**
 * Fonction pour obtenir les données de CMD Initiale pour l'affichage après traitement
 * 
 * @return {Object} Données formatées de CMD Initiale
 */
function getCMDInitialeData() {
  try {
    console.log("Récupération des données de CMD Initiale...");
    
    // Récupérer le fichier actuel
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const cmdInitialeSheet = ss.getSheetByName("CMD initiale");
    
    if (!cmdInitialeSheet) {
      return {
        success: false,
        message: "Feuille CMD initiale introuvable"
      };
    }
    
    // Récupérer les données
    const donnees = cmdInitialeSheet.getDataRange().getValues();
    
    if (donnees.length <= 1) {
      return {
        success: false,
        message: "Aucune donnée dans CMD initiale"
      };
    }
    
    // Générer l'HTML pour afficher les données
    const htmlApercu = generateCMDInitialeHTML(donnees);
    
    // Compter les différents statuts
    let nonAffectees = 0;
    let affectees = 0;
    let problemes = 0;
    
    // L'index de la colonne Statut (supposons qu'il est à l'index 8)
    const statutIndex = 8;
    
    for (let i = 1; i < donnees.length; i++) {
      const statut = donnees[i][statutIndex];
      
      if (!statut || statut === "Non affectée") {
        nonAffectees++;
      } else if (statut === "Problème") {
        problemes++;
      } else {
        affectees++;
      }
    }
    
    return {
      success: true,
      htmlApercu: htmlApercu,
      statuts: {
        nonAffectees: nonAffectees,
        affectees: affectees,
        problemes: problemes
      }
    };
  } catch (error) {
    console.error("Erreur lors de la récupération des données de CMD Initiale :", error);
    return {
      success: false,
      message: "Erreur : " + error.toString()
    };
  }
}

/**
 * Génère un tableau HTML pour afficher les données CMD Initiale
 * @param {Array} donnees - Données de la feuille CMD Initiale
 * @return {string} - Code HTML pour afficher le tableau
 */
function generateCMDInitialeHTML(donnees) {
  try {
    // Extraire les en-têtes (première ligne)
    const enTetes = donnees[0];
    
    // Trouver l'index des colonnes importantes pour la coloration
    const refCmdIndex = enTetes.indexOf("Référence CMD");
    const numCommandeIndex = enTetes.indexOf("Numéro de commande");
    const statutIndex = enTetes.indexOf("Statut");
    
    // Générer le tableau HTML
    let html = `
      <div class="table-responsive">
        <table id="cmdInitialeTable" class="table table-striped table-bordered">
          <thead class="thead-dark">
            <tr>`;
    
    // Ajouter les en-têtes
    enTetes.forEach(header => {
      html += `<th>${header}</th>`;
    });
    
    html += `
            </tr>
          </thead>
          <tbody>`;
    
    // Map pour suivre les références de commande en double
    const refCmdMap = new Map();
    const numCommandeMap = new Map();
    
    // D'abord, compter les occurrences
    for (let i = 1; i < donnees.length; i++) {
      const refCmd = donnees[i][refCmdIndex];
      if (refCmd) {
        refCmdMap.set(refCmd, (refCmdMap.get(refCmd) || 0) + 1);
      }
      
      const numCommande = donnees[i][numCommandeIndex];
      if (numCommande) {
        numCommandeMap.set(numCommande, (numCommandeMap.get(numCommande) || 0) + 1);
      }
    }
    
    // Ajouter les lignes de données
    for (let i = 1; i < donnees.length; i++) {
      const ligne = donnees[i];
      const statut = ligne[statutIndex];
      
      // Déterminer la classe CSS pour la ligne en fonction du statut
      let trClass = "";
      if (statut === "Problème") {
        trClass = "table-danger";
      } else if (statut === "Affectée") {
        trClass = "table-success";
      } else {
        trClass = "table-warning";
      }
      
      // Vérifier si la référence CMD est en double
      const refCmd = ligne[refCmdIndex];
      const numCommande = ligne[numCommandeIndex];
      
      // Ajouter un attribut data-duplicates pour le filtrage
      const hasDuplicates = (refCmd && refCmdMap.get(refCmd) > 1) || 
                          (numCommande && numCommandeMap.get(numCommande) > 1);
      
      html += `<tr class="${trClass}" data-duplicates="${hasDuplicates}">`;
      
      // Ajouter chaque cellule
      for (let j = 0; j < ligne.length; j++) {
        const cellValue = ligne[j] !== null ? ligne[j].toString() : "";
        
        // Colorer en rouge les références CMD et numéros de commande en double
        if ((j === refCmdIndex && refCmd && refCmdMap.get(refCmd) > 1) ||
            (j === numCommandeIndex && numCommande && numCommandeMap.get(numCommande) > 1)) {
          html += `<td class="text-danger font-weight-bold">${cellValue}</td>`;
        } else {
          html += `<td>${cellValue}</td>`;
        }
      }
      
      html += `</tr>`;
    }
    
    html += `
          </tbody>
        </table>
      </div>`;
    
    return html;
  } catch (error) {
    console.error("Erreur lors de la génération du HTML pour CMD Initiale:", error);
    return "Erreur : " + error.toString();
  }
}

/**
 * Fonction pour traiter automatiquement toutes les étapes:
 * 1. Récupérer les feuilles
 * 2. Fusionner si nécessaire
 * 3. Traiter les données
 * 4. Afficher CMD initiale
 * 
 * @return {Object} Résultat du traitement complet
 */
function processusAutomatique() {
  try {
    // Journal de débogage
    Logger.log("Démarrage du processus automatique");
    
    // Étape 1: Récupérer les feuilles
    Logger.log("Étape 1: Récupération des feuilles");
    const resultFeuilles = getListeFeuilles();
    
    if (!resultFeuilles.success) {
      Logger.log("Échec de l'étape 1: " + resultFeuilles.message);
      return {
        success: false,
        etape: 1,
        message: "Erreur lors de la récupération des feuilles: " + resultFeuilles.message
      };
    }
    
    // Filtrer pour obtenir uniquement les feuilles Shopify et Youcan
    const feuillesShopify = resultFeuilles.nomsFeuilles.filter(nom => nom.toLowerCase().includes('shopify'));
    const feuillesYoucan = resultFeuilles.nomsFeuilles.filter(nom => nom.toLowerCase().includes('youcan'));
    
    Logger.log("Feuilles trouvées - Shopify: " + JSON.stringify(feuillesShopify) + ", Youcan: " + JSON.stringify(feuillesYoucan));
    
    // Vérifier qu'au moins un type de feuille est disponible
    if (feuillesShopify.length === 0 && feuillesYoucan.length === 0) {
      Logger.log("Aucune feuille Shopify ou Youcan trouvée");
      return {
        success: false,
        etape: 1,
        message: "Aucune feuille Shopify ou Youcan n'a été trouvée. Veuillez vérifier le document source."
      };
    }
    
    let feuilleATraiter = null;
    let rapportFusion = null;
    
    // Étape 2: Fusionner si nécessaire
    Logger.log("Étape 2: Fusion des feuilles si nécessaire");
    
    if (feuillesShopify.length > 0 && feuillesYoucan.length > 0) {
      // Les deux types de feuilles sont présents, on les fusionne
      Logger.log("Tentative de fusion des feuilles Shopify et Youcan");
      rapportFusion = fusionnerFeuillesExterne();
      
      if (!rapportFusion.success) {
        Logger.log("Échec de l'étape 2: " + rapportFusion.message);
        return {
          success: false,
          etape: 2,
          message: "Erreur lors de la fusion des feuilles: " + rapportFusion.message
        };
      }
      
      feuilleATraiter = rapportFusion.feuilleFusionnee;
      Logger.log("Fusion réussie, feuille à traiter: " + feuilleATraiter);
    } else if (feuillesShopify.length > 0) {
      // Seulement Shopify est présent
      feuilleATraiter = feuillesShopify[0];
      rapportFusion = {
        message: "Seule la feuille Shopify est présente, aucune fusion nécessaire."
      };
      Logger.log("Utilisation de la feuille Shopify: " + feuilleATraiter);
    } else if (feuillesYoucan.length > 0) {
      // Seulement Youcan est présent
      feuilleATraiter = feuillesYoucan[0];
      rapportFusion = {
        message: "Seule la feuille Youcan est présente, aucune fusion nécessaire."
      };
      Logger.log("Utilisation de la feuille Youcan: " + feuilleATraiter);
    }
    
    // Étape 3: Génération de l'aperçu
    Logger.log("Étape 3: Génération de l'aperçu pour la feuille " + feuilleATraiter);
    const apercu = getApercuHTML(feuilleATraiter);
    
    if (apercu.includes("Erreur") || apercu.includes("n'existe pas")) {
      Logger.log("Échec de l'étape 3: Problème lors de la génération de l'aperçu");
      return {
        success: false,
        etape: 3,
        message: "Erreur lors de la génération de l'aperçu. La feuille source pourrait être inaccessible.",
        detailErreur: apercu
      };
    }
    
    // Étape 4: Importer la feuille si nécessaire avant traitement
    Logger.log("Étape intermédiaire: Importation de la feuille " + feuilleATraiter);
    const feuilleDestination = feuilleATraiter.includes("Fusion") ? "Fusion_Temp" : feuilleATraiter;
    
    const resultImport = importerDonnees(feuilleATraiter, feuilleDestination);
    if (!resultImport.success) {
      Logger.log("Échec de l'importation: " + resultImport.message);
      return {
        success: false,
        etape: 3,
        message: "Erreur lors de l'importation de la feuille: " + resultImport.message
      };
    }
    
    // Étape 5: Traiter les données
    Logger.log("Étape 4: Traitement des données de la feuille " + feuilleDestination);
    const resultTraitement = traiterDonneesCMDInitiale(feuilleDestination);
    
    if (!resultTraitement.success) {
      Logger.log("Échec de l'étape 4: " + resultTraitement.message);
      return {
        success: false,
        etape: 4,
        message: "Erreur lors du traitement des données: " + resultTraitement.message
      };
    }
    
    // Étape 6: Récupérer les données de CMD Initiale après traitement
    Logger.log("Étape 5: Récupération des données CMD Initiale");
    const resultCMDInitiale = getCMDInitialeData();
    
    if (!resultCMDInitiale.success) {
      Logger.log("Échec de l'étape 5: " + resultCMDInitiale.message);
      // Même si cette étape échoue, on considère le processus comme réussi car les données ont été traitées
      Logger.log("Le processus automatique est considéré comme réussi malgré l'échec de l'affichage");
    }
    
    Logger.log("Processus automatique terminé avec succès");
    
    // Tout s'est bien passé, on retourne un rapport complet
    return {
      success: true,
      etape: 5,
      rapportFusion: rapportFusion.message,
      feuilleTraitee: feuilleATraiter,
      commandesTraitees: resultTraitement.commandesTraitees,
      problemes: resultTraitement.problemes,
      apercu: apercu,
      cmdInitiale: resultCMDInitiale.success ? resultCMDInitiale : null
    };
  } catch (error) {
    Logger.log("Erreur globale dans processusAutomatique: " + error.toString());
    logError(error, 'processusAutomatique');
    return {
      success: false,
      etape: 0,
      message: "Erreur lors du processus automatique: " + error.toString()
    };
  }
}
