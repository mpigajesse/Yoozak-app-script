/**
 * Yoozak - Module d'initialisation des commandes
 * 
 * Ce script contient les fonctions pour récupérer les données depuis Google Sheets
 * Version: 1.0
 * Date: 2023
 */

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
    const url = "https://docs.google.com/spreadsheets/d/1OK2Ndvc9dyUV99sJDLJ3tYSvtECijG-CM0QlGClYduU/edit";
    
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
    const url = "https://docs.google.com/spreadsheets/d/1OK2Ndvc9dyUV99sJDLJ3tYSvtECijG-CM0QlGClYduU/edit";
    
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
    // URL du fichier Google Sheet externe
    const url = "https://docs.google.com/spreadsheets/d/1OK2Ndvc9dyUV99sJDLJ3tYSvtECijG-CM0QlGClYduU/edit";
    
    // Ouvrir le fichier externe en utilisant son URL
    const fichierExterne = SpreadsheetApp.openByUrl(url);
    
    // Récupérer les noms de toutes les feuilles
    const feuilles = fichierExterne.getSheets();
    const nomsFeuilles = feuilles.map(feuille => feuille.getName());
    
    return {
      success: true,
      nomsFeuilles: nomsFeuilles
    };
  } catch (error) {
    logError(error, 'getListeFeuilles');
    return {
      success: false,
      message: error.toString()
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
    // URL du fichier Google Sheet externe
    const url = "https://docs.google.com/spreadsheets/d/1OK2Ndvc9dyUV99sJDLJ3tYSvtECijG-CM0QlGClYduU/edit";
    
    // Ouvrir le fichier externe en utilisant son URL
    const fichierExterne = SpreadsheetApp.openByUrl(url);
    
    // Récupérer la feuille spécifiée
    const feuille = fichierExterne.getSheetByName(nomFeuille);
    
    if (!feuille) {
      return "<p style='color: red;'>La feuille \"" + nomFeuille + "\" n'existe pas dans le fichier.</p>";
    }
    
    // Récupérer les données (limité à 10 lignes pour l'aperçu)
    const plage = feuille.getRange(1, 1, Math.min(11, feuille.getLastRow()), feuille.getLastColumn());
    const donnees = plage.getValues();
    
    if (donnees.length === 0) {
      return "<p style='color: red;'>Aucune donnée n'a été trouvée dans la feuille \"" + nomFeuille + "\".</p>";
    }
    
    // Créer le tableau HTML
    let html = "<table style='width:100%; border-collapse:collapse; margin-top:10px;'>";
    
    // En-tête (première ligne)
    html += "<thead><tr>";
    for (let i = 0; i < donnees[0].length; i++) {
      html += "<th style='border:1px solid #ddd; padding:8px; text-align:left; background-color:#f2f2f2;'>" + 
              (donnees[0][i] !== null ? donnees[0][i].toString() : "") + "</th>";
    }
    html += "</tr></thead>";
    
    // Corps du tableau (autres lignes)
    html += "<tbody>";
    for (let i = 1; i < donnees.length; i++) {
      html += "<tr>";
      for (let j = 0; j < donnees[i].length; j++) {
        html += "<td style='border:1px solid #ddd; padding:8px; text-align:left;'>" + 
                (donnees[i][j] !== null ? donnees[i][j].toString() : "") + "</td>";
      }
      html += "</tr>";
    }
    html += "</tbody></table>";
    
    // Ajouter un message de succès
    const infoMessage = "<p style='color: green;'>Aperçu de la feuille \"" + nomFeuille + 
                       "\" (" + (donnees.length-1) + " lignes sur " + (feuille.getLastRow()-1) + " au total)</p>";
    
    return infoMessage + html;
  } catch (error) {
    console.error("Erreur dans getApercuHTML:", error);
    return "<p style='color: red;'>Erreur lors de la récupération des données: " + error.toString() + "</p>";
  }
}
