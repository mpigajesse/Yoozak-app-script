/**
 * Yoozak - Module Opérateur
 * 
 * Ce script contient les fonctions pour les opérateurs du système
 * Version: 1.0
 * Date: 17/04/2025
 */

/**
 * Implémentation de la fonction de création d'une commande unitaire
 * Cette fonction remplace le placeholder dans Main.js
 */
function creerCommande() {
  // Vérifier si l'utilisateur est autorisé
  if (!verifierUtilisateur()) {
    SpreadsheetApp.getUi().alert('Vous n\'êtes pas autorisé à utiliser cette fonction.');
    return;
  }
  
  const ui = SpreadsheetApp.getUi();
  const ss = SpreadsheetApp.getActive();
  
  // Déterminer l'opérateur
  const email = Session.getActiveUser().getEmail();
  const sheetConfig = ss.getSheetByName(CONFIG.SHEETS.CONFIG);
  const operateursData = sheetConfig.getRange("CMD SheetUsers").getValues();
  
  let operateur = null;
  for (let i = 0; i < operateursData.length; i++) {
    if (operateursData[i][1] === email) {
      operateur = operateursData[i][0];
      break;
    }
  }
  
  if (!operateur) {
    ui.alert('Vous n\'êtes pas autorisé à créer une commande. Contactez un administrateur.');
    return;
  }
  
  // Récupérer la liste des produits
  const produitsData = sheetConfig.getRange("CMD Products").getValues();
  if (produitsData.length <= 1) {
    ui.alert('Aucun produit disponible. Contactez un administrateur.');
    return;
  }
  
  // Récupérer la liste des régions
  const regionsData = sheetConfig.getRange("CMD Region").getValues();
  if (regionsData.length <= 1) {
    ui.alert('Aucune région disponible. Contactez un administrateur.');
    return;
  }
  
  // Créer le formulaire de création de commande
  const htmlFormulaire = HtmlService.createTemplateFromFile('CreerCommandeForm')
    .evaluate()
    .setWidth(600)
    .setHeight(550);
  
  ui.showModalDialog(htmlFormulaire, 'Créer une nouvelle commande');
}

/**
 * Fonction appelée depuis le formulaire HTML pour créer effectivement la commande
 * 
 * @param {Object} formData Les données du formulaire
 * @return {Object} Le résultat de l'opération
 */
function creerCommandeSubmit(formData) {
  const ss = SpreadsheetApp.getActive();
  const sheetConfig = ss.getSheetByName(CONFIG.SHEETS.CONFIG);
  const sheetInitiale = ss.getSheetByName(CONFIG.SHEETS.INITIALE);
  
  // Vérifier si les feuilles requises existent
  if (!sheetConfig || !sheetInitiale) {
    return { success: false, message: 'Les feuilles requises n\'existent pas.' };
  }
  
  try {
    // Déterminer l'opérateur
    const email = Session.getActiveUser().getEmail();
    const operateursData = sheetConfig.getRange("CMD SheetUsers").getValues();
    
    let operateur = null;
    for (let i = 0; i < operateursData.length; i++) {
      if (operateursData[i][1] === email) {
        operateur = operateursData[i][0];
        break;
      }
    }
    
    if (!operateur) {
      return { success: false, message: 'Vous n\'êtes pas autorisé à créer une commande.' };
    }
    
    // Générer un numéro de commande
    const numeroCommande = genererNumeroCommande('O', operateur.substring(0, 2).toUpperCase());
    
    // Préparer les données pour l'insertion
    const dateNow = new Date();
    
    // Formater le numéro de téléphone
    const telephone = formaterTelephone(formData.telephone);
    
    // Ajouter la commande à la feuille initiale
    sheetInitiale.appendRow([
      numeroCommande,
      'O-' + numeroCommande,  // Identifiant source (O pour Operateur)
      CONFIG.STATUTS.AFFECTEE, // Statut (affectée directement)
      operateur,              // Nom de l'opérateur
      formData.nomClient,     // Nom du client
      telephone,              // Téléphone formaté
      formData.adresse,       // Adresse
      formData.ville,         // Ville
      formData.produit,       // Produit
      formData.quantite,      // Quantité
      formData.prix,          // Prix
      dateNow,                // Date de création
      formaterDate(dateNow)   // Date formatée
    ]);
    
    // Ajouter la commande à la feuille de l'opérateur
    ajouterCommandeAOperateur(operateur, [
      numeroCommande,
      'O-' + numeroCommande,
      CONFIG.STATUTS.AFFECTEE,
      operateur,
      formData.nomClient,
      telephone,
      formData.adresse,
      formData.ville,
      formData.produit,
      formData.quantite,
      formData.prix,
      dateNow,
      formaterDate(dateNow)
    ]);
    
    // Enregistrer l'action dans le journal
    enregistrerLog(
      operateur,
      numeroCommande,
      'Création d\'une nouvelle commande'
    );
    
    return { 
      success: true, 
      message: 'Commande créée avec succès.',
      numeroCommande: numeroCommande
    };
  } catch (e) {
    return { 
      success: false, 
      message: 'Erreur lors de la création de la commande: ' + e.toString() 
    };
  }
}

/**
 * Implémentation de la fonction de modification d'une commande
 * Cette fonction remplace le placeholder dans Main.js
 */
function modifierCommande() {
  // Vérifier si l'utilisateur est autorisé
  if (!verifierUtilisateur()) {
    SpreadsheetApp.getUi().alert('Vous n\'êtes pas autorisé à utiliser cette fonction.');
    return;
  }
  
  const ui = SpreadsheetApp.getUi();
  const ss = SpreadsheetApp.getActive();
  
  // Déterminer l'opérateur
  const email = Session.getActiveUser().getEmail();
  const sheetConfig = ss.getSheetByName(CONFIG.SHEETS.CONFIG);
  const operateursData = sheetConfig.getRange("CMD SheetUsers").getValues();
  
  let operateur = null;
  for (let i = 0; i < operateursData.length; i++) {
    if (operateursData[i][1] === email) {
      operateur = operateursData[i][0];
      break;
    }
  }
  
  if (!operateur) {
    ui.alert('Vous n\'êtes pas un opérateur. Contactez un administrateur.');
    return;
  }
  
  // Demander le numéro de commande à modifier
  const response = ui.prompt(
    'Modifier une commande',
    'Entrez le numéro de la commande à modifier:',
    ui.ButtonSet.OK_CANCEL
  );
  
  if (response.getSelectedButton() !== ui.Button.OK) {
    return;
  }
  
  const numeroCommande = response.getResponseText().trim();
  if (!numeroCommande) {
    ui.alert('Numéro de commande invalide.');
    return;
  }
  
  // Vérifier l'existence de la commande et si elle est affectée à cet opérateur
  const sheetOperateur = ss.getSheetByName('TMP ' + operateur);
  if (!sheetOperateur) {
    ui.alert('Vous n\'avez pas de commandes affectées.');
    return;
  }
  
  // Rechercher la commande
  const commandesData = sheetOperateur.getDataRange().getValues();
  let commandeTrouvee = false;
  let indexLigne = 0;
  let donnees = null;
  
  for (let i = 1; i < commandesData.length; i++) {
    if (commandesData[i][0] === numeroCommande) {
      commandeTrouvee = true;
      indexLigne = i + 1;
      donnees = commandesData[i];
      break;
    }
  }
  
  if (!commandeTrouvee) {
    ui.alert('Commande non trouvée ou non affectée à vous.');
    return;
  }
  
  // Stocker les données de la commande dans CacheService pour le formulaire
  const cache = CacheService.getUserCache();
  cache.put('commande_' + numeroCommande, JSON.stringify({
    numeroCommande: donnees[0],
    idSource: donnees[1],
    nomClient: donnees[3],
    telephone: donnees[4],
    adresse: donnees[5],
    ville: donnees[6],
    produit: donnees[7],
    quantite: donnees[8],
    prix: donnees[9],
    dateCommande: donnees[10],
    indexLigne: indexLigne,
    operateur: operateur
  }), 600); // Cache pour 10 minutes
  
  // Afficher le formulaire de modification
  const htmlFormulaire = HtmlService.createTemplateFromFile('ModifierCommandeForm')
    .evaluate()
    .setWidth(600)
    .setHeight(550);
  
  ui.showModalDialog(htmlFormulaire, 'Modifier la commande');
}

/**
 * Récupère les données d'une commande pour le formulaire de modification
 * 
 * @param {string} numeroCommande Le numéro de la commande à récupérer
 * @return {Object} Les données de la commande
 */
function getCommandeData(numeroCommande) {
  const cache = CacheService.getUserCache();
  const data = cache.get('commande_' + numeroCommande);
  
  if (!data) {
    return { success: false, message: 'Données de commande non trouvées' };
  }
  
  try {
    return { 
      success: true, 
      data: JSON.parse(data) 
    };
  } catch (e) {
    return { 
      success: false, 
      message: 'Erreur lors de la récupération des données: ' + e.toString() 
    };
  }
}

/**
 * Enregistre les modifications d'une commande
 * 
 * @param {Object} formData Les données du formulaire
 * @return {Object} Le résultat de l'opération
 */
function modifierCommandeSubmit(formData) {
  const ss = SpreadsheetApp.getActive();
  const sheetInitiale = ss.getSheetByName(CONFIG.SHEETS.INITIALE);
  
  try {
    // Récupérer les données de la commande depuis le cache
    const cache = CacheService.getUserCache();
    const dataCache = cache.get('commande_' + formData.numeroCommande);
    
    if (!dataCache) {
      return { 
        success: false, 
        message: 'Données de commande expirées. Veuillez réessayer.' 
      };
    }
    
    const commandeData = JSON.parse(dataCache);
    
    // Formater le téléphone
    const telephone = formaterTelephone(formData.telephone);
    
    // Mettre à jour la feuille de l'opérateur
    const sheetOperateur = ss.getSheetByName('TMP ' + commandeData.operateur);
    const rowOperateur = commandeData.indexLigne;
    
    sheetOperateur.getRange(rowOperateur, 4).setValue(formData.nomClient);
    sheetOperateur.getRange(rowOperateur, 5).setValue(telephone);
    sheetOperateur.getRange(rowOperateur, 6).setValue(formData.adresse);
    sheetOperateur.getRange(rowOperateur, 7).setValue(formData.ville);
    sheetOperateur.getRange(rowOperateur, 8).setValue(formData.produit);
    sheetOperateur.getRange(rowOperateur, 9).setValue(formData.quantite);
    sheetOperateur.getRange(rowOperateur, 10).setValue(formData.prix);
    
    // Mettre à jour la feuille principale
    // Trouver la commande dans la feuille initiale
    const commandesData = sheetInitiale.getDataRange().getValues();
    let rowInitiale = 0;
    
    for (let i = 1; i < commandesData.length; i++) {
      if (commandesData[i][0] === formData.numeroCommande) {
        rowInitiale = i + 1;
        break;
      }
    }
    
    if (rowInitiale > 0) {
      sheetInitiale.getRange(rowInitiale, 5).setValue(formData.nomClient);
      sheetInitiale.getRange(rowInitiale, 6).setValue(telephone);
      sheetInitiale.getRange(rowInitiale, 7).setValue(formData.adresse);
      sheetInitiale.getRange(rowInitiale, 8).setValue(formData.ville);
      sheetInitiale.getRange(rowInitiale, 9).setValue(formData.produit);
      sheetInitiale.getRange(rowInitiale, 10).setValue(formData.quantite);
      sheetInitiale.getRange(rowInitiale, 11).setValue(formData.prix);
    }
    
    // Enregistrer l'action dans le journal
    enregistrerLog(
      commandeData.operateur,
      formData.numeroCommande,
      'Modification des informations de la commande'
    );
    
    return { 
      success: true, 
      message: 'Commande modifiée avec succès.' 
    };
  } catch (e) {
    return { 
      success: false, 
      message: 'Erreur lors de la modification de la commande: ' + e.toString() 
    };
  }
}

/**
 * Implémentation de la fonction de confirmation d'une commande
 * Cette fonction remplace le placeholder dans Main.js
 */
function confirmerCommande() {
  // Vérifier si l'utilisateur est autorisé
  if (!verifierUtilisateur()) {
    SpreadsheetApp.getUi().alert('Vous n\'êtes pas autorisé à utiliser cette fonction.');
    return;
  }
  
  const ui = SpreadsheetApp.getUi();
  const ss = SpreadsheetApp.getActive();
  
  // Déterminer l'opérateur
  const email = Session.getActiveUser().getEmail();
  const sheetConfig = ss.getSheetByName(CONFIG.SHEETS.CONFIG);
  const operateursData = sheetConfig.getRange("CMD SheetUsers").getValues();
  
  let operateur = null;
  for (let i = 0; i < operateursData.length; i++) {
    if (operateursData[i][1] === email) {
      operateur = operateursData[i][0];
      break;
    }
  }
  
  if (!operateur) {
    ui.alert('Vous n\'êtes pas un opérateur. Contactez un administrateur.');
    return;
  }
  
  // Demander le numéro de commande à confirmer
  const response = ui.prompt(
    'Confirmer une commande',
    'Entrez le numéro de la commande à confirmer:',
    ui.ButtonSet.OK_CANCEL
  );
  
  if (response.getSelectedButton() !== ui.Button.OK) {
    return;
  }
  
  const numeroCommande = response.getResponseText().trim();
  if (!numeroCommande) {
    ui.alert('Numéro de commande invalide.');
    return;
  }
  
  // Vérifier l'existence de la commande et si elle est affectée à cet opérateur
  const sheetOperateur = ss.getSheetByName('TMP ' + operateur);
  if (!sheetOperateur) {
    ui.alert('Vous n\'avez pas de commandes affectées.');
    return;
  }
  
  // Rechercher la commande
  const commandesData = sheetOperateur.getDataRange().getValues();
  let commandeTrouvee = false;
  let indexLigne = 0;
  let donnees = null;
  
  for (let i = 1; i < commandesData.length; i++) {
    if (commandesData[i][0] === numeroCommande) {
      commandeTrouvee = true;
      indexLigne = i + 1;
      donnees = commandesData[i];
      break;
    }
  }
  
  if (!commandeTrouvee) {
    ui.alert('Commande non trouvée ou non affectée à vous.');
    return;
  }
  
  // Confirmer l'action
  const confirmation = ui.alert(
    'Confirmer la commande',
    'Êtes-vous sûr de vouloir confirmer la commande ' + numeroCommande + ' ?',
    ui.ButtonSet.YES_NO
  );
  
  if (confirmation !== ui.Button.YES) {
    return;
  }
  
  // Passer la commande à l'état confirmé
  try {
    const sheetConfirme = ss.getSheetByName(CONFIG.SHEETS.CONFIRME);
    const sheetInitiale = ss.getSheetByName(CONFIG.SHEETS.INITIALE);
    const now = new Date();
    
    // Ajouter la commande à la feuille des commandes confirmées
    sheetConfirme.appendRow([
      numeroCommande,             // Numéro commande
      donnees[1],                 // ID commande source
      CONFIG.STATUTS.CONFIRMEE,   // Statut (Confirmé)
      operateur,                  // Opérateur
      donnees[3],                 // Nom client
      donnees[4],                 // Téléphone
      donnees[5],                 // Adresse
      donnees[6],                 // Ville
      donnees[7],                 // Produit
      donnees[8],                 // Quantité
      donnees[9],                 // Prix
      now,                        // Date de confirmation
      formaterDate(now),          // Date formatée
      donnees[10]                 // Date commande originale
    ]);
    
    // Supprimer la commande de la feuille de l'opérateur
    sheetOperateur.deleteRow(indexLigne);
    
    // Supprimer la commande de la feuille initiale
    // Trouver la commande dans la feuille initiale
    const commandesInitiales = sheetInitiale.getDataRange().getValues();
    for (let i = 1; i < commandesInitiales.length; i++) {
      if (commandesInitiales[i][0] === numeroCommande) {
        sheetInitiale.deleteRow(i + 1);
        break;
      }
    }
    
    // Enregistrer l'action dans le journal
    enregistrerLog(
      operateur,
      numeroCommande,
      'Confirmation de la commande'
    );
    
    ui.alert('La commande a été confirmée avec succès.');
  } catch (e) {
    ui.alert('Erreur lors de la confirmation de la commande: ' + e.toString());
  }
}

/**
 * Implémentation de la fonction d'annulation d'une commande
 * Cette fonction remplace le placeholder dans Main.js
 */
function annulerCommande() {
  // Vérifier si l'utilisateur est autorisé
  if (!verifierUtilisateur()) {
    SpreadsheetApp.getUi().alert('Vous n\'êtes pas autorisé à utiliser cette fonction.');
    return;
  }
  
  const ui = SpreadsheetApp.getUi();
  const ss = SpreadsheetApp.getActive();
  
  // Déterminer l'opérateur
  const email = Session.getActiveUser().getEmail();
  const sheetConfig = ss.getSheetByName(CONFIG.SHEETS.CONFIG);
  const operateursData = sheetConfig.getRange("CMD SheetUsers").getValues();
  
  let operateur = null;
  for (let i = 0; i < operateursData.length; i++) {
    if (operateursData[i][1] === email) {
      operateur = operateursData[i][0];
      break;
    }
  }
  
  if (!operateur) {
    ui.alert('Vous n\'êtes pas un opérateur. Contactez un administrateur.');
    return;
  }
  
  // Demander le numéro de commande à annuler
  const response = ui.prompt(
    'Annuler une commande',
    'Entrez le numéro de la commande à annuler:',
    ui.ButtonSet.OK_CANCEL
  );
  
  if (response.getSelectedButton() !== ui.Button.OK) {
    return;
  }
  
  const numeroCommande = response.getResponseText().trim();
  if (!numeroCommande) {
    ui.alert('Numéro de commande invalide.');
    return;
  }
  
  // Vérifier l'existence de la commande et si elle est affectée à cet opérateur
  const sheetOperateur = ss.getSheetByName('TMP ' + operateur);
  if (!sheetOperateur) {
    ui.alert('Vous n\'avez pas de commandes affectées.');
    return;
  }
  
  // Rechercher la commande
  const commandesData = sheetOperateur.getDataRange().getValues();
  let commandeTrouvee = false;
  let indexLigne = 0;
  let donnees = null;
  
  for (let i = 1; i < commandesData.length; i++) {
    if (commandesData[i][0] === numeroCommande) {
      commandeTrouvee = true;
      indexLigne = i + 1;
      donnees = commandesData[i];
      break;
    }
  }
  
  if (!commandeTrouvee) {
    ui.alert('Commande non trouvée ou non affectée à vous.');
    return;
  }
  
  // Demander le motif d'annulation
  const motifResponse = ui.prompt(
    'Annuler la commande',
    'Veuillez indiquer le motif de l\'annulation:',
    ui.ButtonSet.OK_CANCEL
  );
  
  if (motifResponse.getSelectedButton() !== ui.Button.OK) {
    return;
  }
  
  const motif = motifResponse.getResponseText().trim();
  if (!motif) {
    ui.alert('Vous devez indiquer un motif d\'annulation.');
    return;
  }
  
  // Confirmer l'action
  const confirmation = ui.alert(
    'Annuler la commande',
    'Êtes-vous sûr de vouloir annuler la commande ' + numeroCommande + ' ?\n\nMotif: ' + motif,
    ui.ButtonSet.YES_NO
  );
  
  if (confirmation !== ui.Button.YES) {
    return;
  }
  
  // Annuler la commande
  try {
    const sheetAnnulee = ss.getSheetByName(CONFIG.SHEETS.ANNULEE);
    const sheetInitiale = ss.getSheetByName(CONFIG.SHEETS.INITIALE);
    const now = new Date();
    
    // Ajouter la commande à la feuille des commandes annulées
    sheetAnnulee.appendRow([
      numeroCommande,             // Numéro commande
      donnees[1],                 // ID commande source
      CONFIG.STATUTS.ANNULE,      // Statut (Annulé)
      operateur,                  // Opérateur
      donnees[3],                 // Nom client
      donnees[4],                 // Téléphone
      donnees[5],                 // Adresse
      donnees[6],                 // Ville
      donnees[7],                 // Produit
      donnees[8],                 // Quantité
      donnees[9],                 // Prix
      now,                        // Date d'annulation
      formaterDate(now),          // Date formatée
      donnees[10],                // Date commande originale
      motif                       // Motif d'annulation
    ]);
    
    // Supprimer la commande de la feuille de l'opérateur
    sheetOperateur.deleteRow(indexLigne);
    
    // Supprimer la commande de la feuille initiale
    const commandesInitiales = sheetInitiale.getDataRange().getValues();
    for (let i = 1; i < commandesInitiales.length; i++) {
      if (commandesInitiales[i][0] === numeroCommande) {
        sheetInitiale.deleteRow(i + 1);
        break;
      }
    }
    
    // Enregistrer l'action dans le journal
    enregistrerLog(
      operateur,
      numeroCommande,
      'Annulation de la commande. Motif: ' + motif
    );
    
    ui.alert('La commande a été annulée avec succès.');
  } catch (e) {
    ui.alert('Erreur lors de l\'annulation de la commande: ' + e.toString());
  }
}