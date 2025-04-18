/**
 * Yoozak - Module d'importation des données
 * 
 * Ce script contient les fonctions pour importer les données depuis Shopify et Youcan
 * Version: 1.0
 * Date: 17/04/2025
 */

/**
 * Configure les triggers pour l'importation automatique des données
 * Cette fonction est appelée depuis le menu Administrateur
 */
function configurerTriggers() {
  const ui = SpreadsheetApp.getUi();
  const response = ui.alert(
    'Configuration des triggers',
    'Souhaitez-vous configurer les triggers pour l\'importation automatique des données ? Cela supprimera tous les triggers existants.',
    ui.ButtonSet.YES_NO
  );
  
  if (response !== ui.Button.YES) {
    return;
  }
  
  try {
    // Supprimer tous les triggers existants
    const triggers = ScriptApp.getProjectTriggers();
    for (let i = 0; i < triggers.length; i++) {
      ScriptApp.deleteTrigger(triggers[i]);
    }
    
    // Créer un trigger pour l'importation depuis Shopify (toutes les 3 heures)
    ScriptApp.newTrigger('importerShopify')
      .timeBased()
      .everyHours(3)
      .create();
    
    // Créer un trigger pour l'importation depuis Youcan (toutes les 3 heures)
    ScriptApp.newTrigger('importerYoucan')
      .timeBased()
      .everyHours(3)
      .create();
    
    ui.alert('Triggers configurés', 'Les triggers ont été configurés avec succès.', ui.ButtonSet.OK);
  } catch (e) {
    ui.alert('Erreur', 'Une erreur est survenue lors de la configuration des triggers: ' + e.toString(), ui.ButtonSet.OK);
    Logger.log('Erreur lors de la configuration des triggers: ' + e.toString());
  }
}

/**
 * Importe les données depuis Shopify
 * Cette fonction est appelée par un trigger
 */
function importerShopify() {
  const ss = SpreadsheetApp.getActive();
  const sheetImport = ss.getSheetByName('Import Shopify');
  const sheetProblem = ss.getSheetByName(CONFIG.SHEETS.PROBLEM);
  const sheetInitiale = ss.getSheetByName(CONFIG.SHEETS.INITIALE);
  
  // Vérifier que les feuilles requises existent
  if (!sheetImport || !sheetProblem || !sheetInitiale) {
    Logger.log('Une ou plusieurs feuilles requises n\'existent pas.');
    return;
  }
  
  // Récupérer les données à importer
  const commandesData = sheetImport.getDataRange().getValues();
  if (commandesData.length <= 1) {
    Logger.log('Aucune commande à importer depuis Shopify.');
    return;
  }
  
  let importees = 0;
  let problemes = 0;
  
  // Parcourir les données à partir de la deuxième ligne (ignorer les en-têtes)
  for (let i = 1; i < commandesData.length; i++) {
    const idCommande = commandesData[i][0];
    
    // Vérifier si la commande existe déjà dans la feuille initiale
    const commandeExistante = verifierCommandeExistante(idCommande);
    
    if (commandeExistante) {
      // Ajouter la commande à la feuille des problèmes
      ajouterCommande_Problem(
        idCommande,
        'Doublon: Cette commande existe déjà dans le système',
        commandesData[i][1], // Nom client
        commandesData[i][2], // Téléphone
        commandesData[i][3], // Adresse
        commandesData[i][4], // Ville
        commandesData[i][5], // Produit
        commandesData[i][6], // Quantité
        commandesData[i][7], // Prix
        commandesData[i][8]  // Date commande
      );
      
      problemes++;
      continue;
    }
    
    // Générer un numéro de commande unique
    const numeroCommande = genererNumeroCommande(CONFIG.SOURCES.SHOPIFY);
    
    // Formater le numéro de téléphone
    const telephone = formaterTelephone(commandesData[i][2]);
    
    // Ajouter la commande à la feuille initiale
    sheetInitiale.appendRow([
      numeroCommande,
      idCommande,
      CONFIG.STATUTS.NON_AFFECTEE,
      '',
      commandesData[i][1], // Nom client
      telephone,           // Téléphone formaté
      commandesData[i][3], // Adresse
      commandesData[i][4], // Ville
      commandesData[i][5], // Produit
      commandesData[i][6], // Quantité
      commandesData[i][7], // Prix
      new Date(),          // Date de création
      formaterDate(commandesData[i][8] || new Date()) // Date commande
    ]);
    
    importees++;
  }
  
  // Effacer les commandes importées du fichier d'importation
  if (importees > 0 || problemes > 0) {
    // Conserver uniquement la ligne d'en-tête
    sheetImport.deleteRows(2, commandesData.length - 1);
  }
  
  Logger.log('Import Shopify terminé. ' + importees + ' commandes importées, ' + problemes + ' problèmes.');
}

/**
 * Importe les données depuis Youcan
 * Cette fonction est appelée par un trigger
 */
function importerYoucan() {
  const ss = SpreadsheetApp.getActive();
  const sheetImport = ss.getSheetByName('Import Youcan');
  const sheetProblem = ss.getSheetByName(CONFIG.SHEETS.PROBLEM);
  const sheetInitiale = ss.getSheetByName(CONFIG.SHEETS.INITIALE);
  
  // Vérifier que les feuilles requises existent
  if (!sheetImport || !sheetProblem || !sheetInitiale) {
    Logger.log('Une ou plusieurs feuilles requises n\'existent pas.');
    return;
  }
  
  // Récupérer les données à importer
  const commandesData = sheetImport.getDataRange().getValues();
  if (commandesData.length <= 1) {
    Logger.log('Aucune commande à importer depuis Youcan.');
    return;
  }
  
  let importees = 0;
  let problemes = 0;
  
  // Parcourir les données à partir de la deuxième ligne (ignorer les en-têtes)
  for (let i = 1; i < commandesData.length; i++) {
    const idCommande = commandesData[i][0];
    
    // Vérifier si la commande existe déjà dans la feuille initiale
    const commandeExistante = verifierCommandeExistante(idCommande);
    
    if (commandeExistante) {
      // Ajouter la commande à la feuille des problèmes
      ajouterCommande_Problem(
        idCommande,
        'Doublon: Cette commande existe déjà dans le système',
        commandesData[i][1], // Nom client
        commandesData[i][2], // Téléphone
        commandesData[i][3], // Adresse
        commandesData[i][4], // Ville
        commandesData[i][5], // Produit
        commandesData[i][6], // Quantité
        commandesData[i][7], // Prix
        commandesData[i][8]  // Date commande
      );
      
      problemes++;
      continue;
    }
    
    // Générer un numéro de commande unique
    const numeroCommande = genererNumeroCommande(CONFIG.SOURCES.YOUCAN);
    
    // Formater le numéro de téléphone
    const telephone = formaterTelephone(commandesData[i][2]);
    
    // Ajouter la commande à la feuille initiale
    sheetInitiale.appendRow([
      numeroCommande,
      idCommande,
      CONFIG.STATUTS.NON_AFFECTEE,
      '',
      commandesData[i][1], // Nom client
      telephone,           // Téléphone formaté
      commandesData[i][3], // Adresse
      commandesData[i][4], // Ville
      commandesData[i][5], // Produit
      commandesData[i][6], // Quantité
      commandesData[i][7], // Prix
      new Date(),          // Date de création
      formaterDate(commandesData[i][8] || new Date()) // Date commande
    ]);
    
    importees++;
  }
  
  // Effacer les commandes importées du fichier d'importation
  if (importees > 0 || problemes > 0) {
    // Conserver uniquement la ligne d'en-tête
    sheetImport.deleteRows(2, commandesData.length - 1);
  }
  
  Logger.log('Import Youcan terminé. ' + importees + ' commandes importées, ' + problemes + ' problèmes.');
}

/**
 * Vérifie si une commande existe déjà dans la feuille initiale
 * 
 * @param {string} idCommande L'ID de la commande à vérifier
 * @return {boolean} True si la commande existe déjà, sinon False
 */
function verifierCommandeExistante(idCommande) {
  const ss = SpreadsheetApp.getActive();
  const sheetInitiale = ss.getSheetByName(CONFIG.SHEETS.INITIALE);
  
  if (!sheetInitiale) {
    return false;
  }
  
  const commandesData = sheetInitiale.getDataRange().getValues();
  
  for (let i = 1; i < commandesData.length; i++) {
    if (commandesData[i][1] === idCommande) {
      return true;
    }
  }
  
  // Vérifier également dans la feuille confirmée
  const sheetConfirme = ss.getSheetByName(CONFIG.SHEETS.CONFIRME);
  
  if (!sheetConfirme) {
    return false;
  }
  
  const commandesConfirmees = sheetConfirme.getDataRange().getValues();
  
  for (let i = 1; i < commandesConfirmees.length; i++) {
    if (commandesConfirmees[i][1] === idCommande) {
      return true;
    }
  }
  
  // Vérifier également dans la feuille des commandes annulées
  const sheetAnnulee = ss.getSheetByName(CONFIG.SHEETS.ANNULEE);
  
  if (!sheetAnnulee) {
    return false;
  }
  
  const commandesAnnulees = sheetAnnulee.getDataRange().getValues();
  
  for (let i = 1; i < commandesAnnulees.length; i++) {
    if (commandesAnnulees[i][1] === idCommande) {
      return true;
    }
  }
  
  return false;
}

/**
 * Ajoute une commande avec problème à la feuille de problèmes
 * 
 * @param {string} idCommande ID de la commande source
 * @param {string} description Description du problème
 * @param {string} nomClient Nom du client
 * @param {string} telephone Numéro de téléphone
 * @param {string} adresse Adresse de livraison
 * @param {string} ville Ville
 * @param {string} produit Nom du produit
 * @param {number} quantite Quantité commandée
 * @param {number} prix Prix de la commande
 * @param {Date|string} dateCommande Date de la commande originale
 */
function ajouterCommande_Problem(idCommande, description, nomClient, telephone, adresse, ville, produit, quantite, prix, dateCommande) {
  const ss = SpreadsheetApp.getActive();
  const sheetProblem = ss.getSheetByName(CONFIG.SHEETS.PROBLEM);
  
  if (!sheetProblem) {
    return;
  }
  
  const now = new Date();
  
  sheetProblem.appendRow([
    now,                       // Date du problème
    formaterDate(now),         // Date formatée
    idCommande,                // ID commande source
    description,               // Description du problème
    nomClient,                 // Nom client
    formaterTelephone(telephone), // Téléphone formaté
    adresse,                   // Adresse
    ville,                     // Ville
    produit,                   // Produit
    quantite,                  // Quantité
    prix,                      // Prix
    dateCommande || now        // Date commande originale
  ]);
}