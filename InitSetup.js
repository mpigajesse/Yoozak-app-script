/**
 * Yoozak - Module d'initialisation du système
 * 
 * Ce script contient les fonctions pour initialiser le système
 * Version: 1.0
 * Date: 17/04/2025
 */

/**
 * Initialise le système en créant toutes les feuilles nécessaires
 * Cette fonction est appelée depuis le menu ou au premier démarrage de l'application web
 * @param {boolean} fromWebApp Indique si la fonction est appelée depuis l'application web
 * @return {boolean} True si l'initialisation a réussi, sinon False
 */
function initialiserSysteme(fromWebApp) {
  const ss = getSpreadsheet();
  let confirmation = true;
  
  // Ne demander confirmation que si on n'est pas dans l'application web
  if (!fromWebApp) {
    try {
  const ui = SpreadsheetApp.getUi();
      const response = ui.alert(
    'Initialisation du système',
        'Attention ! Cette opération va réinitialiser toutes les feuilles du système. Voulez-vous continuer ?',
    ui.ButtonSet.YES_NO
  );
      confirmation = (response === ui.Button.YES);
    } catch (e) {
      // Si getUi() échoue, on continue sans demander de confirmation
      Logger.log("Impossible d'afficher l'interface utilisateur: " + e.toString());
    }
  }
  
  if (!confirmation) {
    return false;
  }
  
  try {
    // Créer ou réinitialiser chaque feuille
    creerFeuille(ss, CONFIG.SHEETS.CONFIG, [
      ['Email', 'Nom', 'Autorisé', 'Rôle'],
      [Session.getActiveUser().getEmail(), 'Admin', 'Oui', 'Admin']
    ]);
    
    creerFeuille(ss, CONFIG.SHEETS.INITIALE, [
      ['N° Commande', 'ID Source', 'Statut', 'Opérateur', 'Client', 'Téléphone', 'Adresse', 'Ville', 'Produit', 'Quantité', 'Prix', 'Date Création', 'Date Formatée']
    ]);
    
    creerFeuille(ss, CONFIG.SHEETS.TMP, [
      ['N° Commande', 'Client', 'Action', 'Statut', 'Téléphone', 'Adresse', 'Ville', 'Produit', 'Quantité', 'Prix', 'Date Création']
    ]);
    
    creerFeuille(ss, CONFIG.SHEETS.PRODUITS, [
      ['Référence', 'Nom', 'Prix', 'Stock', 'Actif'],
      ['PROD001', 'Produit Test 1', '100', '10', 'Oui'],
      ['PROD002', 'Produit Test 2', '200', '5', 'Oui']
    ]);
  
    creerFeuille(ss, CONFIG.SHEETS.TMP_LOG, [
      ['Opérateur', 'N° Commande', 'Action', 'Date', 'Date Formatée']
    ]);
    
    creerFeuille(ss, CONFIG.SHEETS.CONFIRME, [
      ['N° Commande', 'ID Source', 'Statut', 'Opérateur', 'Client', 'Téléphone', 'Adresse', 'Ville', 'Produit', 'Quantité', 'Prix', 'Date Confirmation', 'Date Formatée', 'Date Commande']
    ]);
    
    creerFeuille(ss, CONFIG.SHEETS.CONFIRME_LOG, [
      ['Opérateur', 'N° Commande', 'Action', 'Date', 'Date Formatée']
    ]);
    
    creerFeuille(ss, CONFIG.SHEETS.ANNULEE, [
      ['N° Commande', 'ID Source', 'Statut', 'Opérateur', 'Client', 'Téléphone', 'Adresse', 'Ville', 'Produit', 'Quantité', 'Prix', 'Date Annulation', 'Date Formatée', 'Date Commande', 'Motif']
    ]);
    
    creerFeuille(ss, CONFIG.SHEETS.RETOURNEE, [
      ['N° Commande', 'ID Source', 'Statut', 'Opérateur', 'Client', 'Téléphone', 'Adresse', 'Ville', 'Produit', 'Quantité', 'Prix', 'Date Retour', 'Date Formatée', 'Date Commande', 'Motif']
    ]);
    
    creerFeuille(ss, CONFIG.SHEETS.PROBLEM, [
      ['ID Source', 'Description', 'Client', 'Téléphone', 'Adresse', 'Ville', 'Produit', 'Quantité', 'Prix', 'Date Commande', 'Date Détection', 'Date Formatée']
    ]);
    
    // Créer les feuilles d'importation
    creerFeuille(ss, 'Import Shopify', [
      ['ID Commande', 'Client', 'Téléphone', 'Adresse', 'Ville', 'Produit', 'Quantité', 'Prix', 'Date', 'Source']
    ]);
    
    creerFeuille(ss, 'Import Youcan', [
      ['ID Commande', 'Client', 'Téléphone', 'Adresse', 'Ville', 'Produit', 'Quantité', 'Prix', 'Date', 'Source']
    ]);
  
    // Créer les plages nommées
    creerPlageNommee(ss, CONFIG.SHEETS.CONFIG, 1, 1, 1, 4, 'CMD Headers');
    creerPlageNommee(ss, CONFIG.SHEETS.CONFIG, 2, 1, -1, 4, 'CMD AllowedUsers');
    creerPlageNommee(ss, CONFIG.SHEETS.PRODUITS, 2, 1, -1, 5, 'CMD Products');
    
    // Ajouter des régions de test dans la feuille CONFIG
    ajouterRegions(ss);
  
    // Message de succès si on n'est pas dans l'application web
    if (!fromWebApp) {
      try {
        const ui = SpreadsheetApp.getUi();
        ui.alert('Initialisation réussie', 'Le système a été initialisé avec succès.', ui.ButtonSet.OK);
      } catch (e) {
        // Si getUi() échoue, on continue sans afficher de message
        Logger.log("Impossible d'afficher le message de succès: " + e.toString());
      }
    }
    
    // Enregistrer l'action dans le journal
    enregistrerLog(
      Session.getActiveUser().getEmail(),
      'SYSTEM',
      'Initialisation du système'
    );
    
    return true;
  } catch (e) {
    // Ajouter cette ligne pour journaliser l'erreur complète
    console.error("Erreur détaillée: " + e.toString() + "\nStack: " + e.stack);
    
    // Message d'erreur si on n'est pas dans l'application web
    if (!fromWebApp) {
      try {
        const ui = SpreadsheetApp.getUi();
        ui.alert('Erreur', 'Une erreur est survenue lors de l\'initialisation : ' + e.toString(), ui.ButtonSet.OK);
      } catch (e2) {
        // Si getUi() échoue, on continue sans afficher de message
        Logger.log("Impossible d'afficher le message d'erreur: " + e2.toString());
      }
    }
    Logger.log('Erreur lors de l\'initialisation : ' + e.toString());
    return false;
  }
}

/**
 * Crée ou réinitialise une feuille avec les données spécifiées
 * 
 * @param {Spreadsheet} ss Le classeur
 * @param {string} nom Le nom de la feuille
 * @param {Array} donnees Les données à écrire dans la feuille
 */
function creerFeuille(ss, nom, donnees) {
  let feuille = ss.getSheetByName(nom);
  
  // Si la feuille existe, la supprimer
  if (feuille) {
    ss.deleteSheet(feuille);
  }
  
  // Créer une nouvelle feuille
  feuille = ss.insertSheet(nom);
  
  // Écrire les données
  if (donnees && donnees.length > 0) {
    feuille.getRange(1, 1, donnees.length, donnees[0].length).setValues(donnees);
    // Mettre en forme la première ligne comme en-tête
    feuille.getRange(1, 1, 1, donnees[0].length).setFontWeight('bold').setBackground('#f3f3f3');
  }
  
  return feuille;
}

/**
 * Crée une plage nommée dans une feuille
 * 
 * @param {Spreadsheet} ss Le classeur
 * @param {string} nomFeuille Le nom de la feuille
 * @param {number} ligne La ligne de début
 * @param {number} colonne La colonne de début
 * @param {number} lignes Le nombre de lignes (ou -1 pour toutes les lignes jusqu'à la fin)
 * @param {number} colonnes Le nombre de colonnes
 * @param {string} nomPlage Le nom de la plage
 */
function creerPlageNommee(ss, nomFeuille, ligne, colonne, lignes, colonnes, nomPlage) {
  const feuille = ss.getSheetByName(nomFeuille);
  if (!feuille) return;
  
  // Si lignes = -1, prendre toutes les lignes jusqu'à la fin
  if (lignes === -1) {
    lignes = Math.max(1, feuille.getLastRow() - ligne + 1);
  }
  
  // Créer la plage
  const plage = feuille.getRange(ligne, colonne, lignes, colonnes);
  
  // Donner un nom à la plage
  ss.setNamedRange(nomPlage, plage);
}

/**
 * Ajoute des régions de test dans la feuille CONFIG
 * 
 * @param {Spreadsheet} ss Le classeur
 */
function ajouterRegions(ss) {
  const feuille = ss.getSheetByName(CONFIG.SHEETS.CONFIG);
  if (!feuille) return;
  
  // Ajouter une section pour les régions
  feuille.getRange(10, 1).setValue('RÉGIONS');
  feuille.getRange(11, 1, 1, 2).setValues([['Région', 'Frais de livraison']]);
  feuille.getRange(11, 1, 1, 2).setFontWeight('bold').setBackground('#f3f3f3');
  
  // Ajouter quelques régions de test
  const regions = [
    ['Casablanca', '30'],
    ['Rabat', '40'],
    ['Marrakech', '50'],
    ['Fès', '60'],
    ['Tanger', '70']
  ];
  
  feuille.getRange(12, 1, regions.length, 2).setValues(regions);
  
  // Créer une plage nommée pour les régions
  creerPlageNommee(ss, CONFIG.SHEETS.CONFIG, 12, 1, regions.length, 2, 'CMD Region');
}

/**
 * Vérifie si le système a été initialisé
 * 
 * @return {boolean} True si le système est initialisé, sinon False
 */
function estInitialise() {
  const ss = getSpreadsheet();
  
  // Vérifier l'existence des feuilles principales
  const sheetConfig = ss.getSheetByName(CONFIG.SHEETS.CONFIG);
  const sheetInitiale = ss.getSheetByName(CONFIG.SHEETS.INITIALE);
  
  return (sheetConfig && sheetInitiale);
}