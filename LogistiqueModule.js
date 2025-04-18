/**
 * Yoozak - Module Logistique
 * 
 * Ce script contient les fonctions pour le service logistique
 * Version: 1.0
 * Date: 17/04/2025
 */

/**
 * Implémentation de la fonction de changement de statut d'une commande
 * Cette fonction remplace le placeholder dans Main.js
 */
function changerStatutCommande() {
  // Vérifier si l'utilisateur est autorisé
  if (!verifierUtilisateur()) {
    SpreadsheetApp.getUi().alert('Vous n\'êtes pas autorisé à utiliser cette fonction.');
    return;
  }
  
  const ui = SpreadsheetApp.getUi();
  const ss = SpreadsheetApp.getActive();
  const sheetConfirme = ss.getSheetByName(CONFIG.SHEETS.CONFIRME);
  
  if (!sheetConfirme) {
    ui.alert('La feuille requise n\'existe pas.');
    return;
  }
  
  // Demander le numéro de commande dont le statut doit être changé
  const response = ui.prompt(
    'Changer le statut d\'une commande',
    'Entrez le numéro de la commande à traiter:',
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
  
  // Rechercher la commande dans la feuille des commandes confirmées
  const commandesData = sheetConfirme.getDataRange().getValues();
  let commandeTrouvee = false;
  let indexLigne = 0;
  let statutActuel = '';
  let donnees = null;
  
  for (let i = 1; i < commandesData.length; i++) {
    if (commandesData[i][0] === numeroCommande) {
      commandeTrouvee = true;
      indexLigne = i + 1;
      statutActuel = commandesData[i][2];
      donnees = commandesData[i];
      break;
    }
  }
  
  if (!commandeTrouvee) {
    ui.alert('Commande non trouvée dans les commandes confirmées.');
    return;
  }
  
  // Déterminer les statuts possibles selon le statut actuel
  let statutsPossibles = [];
  
  switch (statutActuel) {
    case CONFIG.STATUTS.CONFIRMEE:
      statutsPossibles = [CONFIG.STATUTS.EN_PREPARATION];
      break;
    case CONFIG.STATUTS.EN_PREPARATION:
      statutsPossibles = [CONFIG.STATUTS.EXPEDIE];
      break;
    case CONFIG.STATUTS.EXPEDIE:
      statutsPossibles = [CONFIG.STATUTS.LIVRE, CONFIG.STATUTS.RETOURNE];
      break;
    case CONFIG.STATUTS.LIVRE:
      statutsPossibles = [CONFIG.STATUTS.RETOURNE];
      break;
    default:
      ui.alert('Le statut actuel ne permet pas de changement.');
      return;
  }
  
  // Demander le nouveau statut
  let htmlSelectStatut = '<select id="statut">';
  for (let i = 0; i < statutsPossibles.length; i++) {
    htmlSelectStatut += '<option value="' + statutsPossibles[i] + '">' + statutsPossibles[i] + '</option>';
  }
  htmlSelectStatut += '</select>';
  
  const responseStatut = ui.prompt(
    'Changer le statut',
    'Statut actuel: ' + statutActuel + '\n\nSélectionnez le nouveau statut:\n' + htmlSelectStatut,
    ui.ButtonSet.OK_CANCEL
  );
  
  if (responseStatut.getSelectedButton() !== ui.Button.OK) {
    return;
  }
  
  const nouveauStatut = responseStatut.getResponseText().trim();
  if (!statutsPossibles.includes(nouveauStatut)) {
    ui.alert('Statut non valide.');
    return;
  }
  
  // Confirmer le changement de statut
  const confirmation = ui.alert(
    'Confirmer le changement de statut',
    'Êtes-vous sûr de vouloir changer le statut de la commande ' + numeroCommande + ' de "' + statutActuel + '" à "' + nouveauStatut + '" ?',
    ui.ButtonSet.YES_NO
  );
  
  if (confirmation !== ui.Button.YES) {
    return;
  }
  
  // Mettre à jour le statut de la commande
  sheetConfirme.getRange(indexLigne, 3).setValue(nouveauStatut);
  
  // Si la commande est retournée, ajouter une entrée dans la feuille des retours
  if (nouveauStatut === CONFIG.STATUTS.RETOURNE) {
    const sheetRetournee = ss.getSheetByName(CONFIG.SHEETS.RETOURNEE);
    const now = new Date();
    
    // Demander la raison du retour
    const raisonResponse = ui.prompt(
      'Raison du retour',
      'Veuillez indiquer la raison du retour:',
      ui.ButtonSet.OK_CANCEL
    );
    
    if (raisonResponse.getSelectedButton() !== ui.Button.OK) {
      return;
    }
    
    const raison = raisonResponse.getResponseText().trim() || 'Non spécifiée';
    
    // Ajouter la commande à la feuille des retours
    sheetRetournee.appendRow([
      donnees[0],                 // Numéro commande
      donnees[1],                 // ID commande source
      nouveauStatut,              // Statut (Retourné)
      donnees[3],                 // Opérateur
      donnees[4],                 // Nom client
      donnees[5],                 // Téléphone
      donnees[6],                 // Adresse
      donnees[7],                 // Ville
      donnees[8],                 // Produit
      donnees[9],                 // Quantité
      donnees[10],                // Prix
      now,                        // Date du retour
      formaterDate(now),          // Date formatée
      donnees[13],                // Date commande originale
      raison                      // Raison du retour
    ]);
    
    // Supprimer la commande de la feuille des commandes confirmées
    sheetConfirme.deleteRow(indexLigne);
    
    // Mettre à jour le stock (décrémenter la quantité du produit retourné)
    mettreAJourStock(donnees[8], -parseInt(donnees[9]));
  }
  
  // Enregistrer l'action dans le journal
  enregistrerLog(
    Session.getActiveUser().getEmail(),
    numeroCommande,
    'Changement de statut: ' + statutActuel + ' -> ' + nouveauStatut
  );
  
  ui.alert('Le statut de la commande a été mis à jour avec succès.');
}

/**
 * Met à jour le stock d'un produit
 * 
 * @param {string} nomProduit Le nom du produit
 * @param {number} quantite La quantité à ajouter (positif) ou à soustraire (négatif)
 * @return {boolean} True si la mise à jour a réussi, sinon False
 */
function mettreAJourStock(nomProduit, quantite) {
  const ss = SpreadsheetApp.getActive();
  const sheetConfig = ss.getSheetByName(CONFIG.SHEETS.CONFIG);
  
  if (!sheetConfig) {
    return false;
  }
  
  // Récupérer la liste des produits
  const produitsRange = sheetConfig.getRange("CMD Products");
  const produitsData = produitsRange.getValues();
  
  // Rechercher le produit
  for (let i = 1; i < produitsData.length; i++) {
    if (produitsData[i][0] === nomProduit) {
      // Mettre à jour la quantité en stock
      const stockActuel = parseInt(produitsData[i][2]) || 0;
      const nouveauStock = stockActuel + quantite;
      
      // Vérifier que le stock ne devient pas négatif
      if (nouveauStock >= 0) {
        produitsRange.getCell(i + 1, 3).setValue(nouveauStock);
        return true;
      }
      break;
    }
  }
  
  return false;
}

/**
 * Implémentation de la fonction d'impression des tickets de commande
 * Cette fonction remplace le placeholder dans Main.js
 */
function imprimerTickets() {
  // Vérifier si l'utilisateur est autorisé
  if (!verifierUtilisateur()) {
    SpreadsheetApp.getUi().alert('Vous n\'êtes pas autorisé à utiliser cette fonction.');
    return;
  }
  
  const ui = SpreadsheetApp.getUi();
  const ss = SpreadsheetApp.getActive();
  const sheetConfirme = ss.getSheetByName(CONFIG.SHEETS.CONFIRME);
  
  if (!sheetConfirme) {
    ui.alert('La feuille requise n\'existe pas.');
    return;
  }
  
  // Demander le mode d'impression (unitaire ou en masse)
  const modeResponse = ui.alert(
    'Impression des tickets',
    'Comment souhaitez-vous imprimer les tickets ?\n\n' +
    'Cliquez sur OUI pour l\'impression unitaire (une commande spécifique).\n' +
    'Cliquez sur NON pour l\'impression en masse (toutes les commandes en préparation).',
    ui.ButtonSet.YES_NO_CANCEL
  );
  
  if (modeResponse === ui.Button.CANCEL) {
    return;
  }
  
  if (modeResponse === ui.Button.YES) {
    // Impression unitaire
    imprimerTicketUnitaire();
  } else {
    // Impression en masse
    imprimerTicketsEnMasse();
  }
}

/**
 * Gère l'impression d'un ticket pour une commande spécifique
 */
function imprimerTicketUnitaire() {
  const ui = SpreadsheetApp.getUi();
  const ss = SpreadsheetApp.getActive();
  const sheetConfirme = ss.getSheetByName(CONFIG.SHEETS.CONFIRME);
  
  // Demander le numéro de commande
  const response = ui.prompt(
    'Impression d\'un ticket',
    'Entrez le numéro de la commande à imprimer:',
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
  
  // Rechercher la commande
  const commandesData = sheetConfirme.getDataRange().getValues();
  let commandeTrouvee = false;
  let donnees = null;
  
  for (let i = 1; i < commandesData.length; i++) {
    if (commandesData[i][0] === numeroCommande) {
      commandeTrouvee = true;
      donnees = commandesData[i];
      break;
    }
  }
  
  if (!commandeTrouvee) {
    ui.alert('Commande non trouvée.');
    return;
  }
  
  // Vérifier que la commande est en préparation ou confirmée
  if (donnees[2] !== CONFIG.STATUTS.EN_PREPARATION && donnees[2] !== CONFIG.STATUTS.CONFIRMEE) {
    ui.alert('Seules les commandes confirmées ou en préparation peuvent être imprimées.');
    return;
  }
  
  // Générer le ticket
  genererTicket(donnees);
}

/**
 * Gère l'impression de tickets pour toutes les commandes en préparation
 */
function imprimerTicketsEnMasse() {
  const ui = SpreadsheetApp.getUi();
  const ss = SpreadsheetApp.getActive();
  const sheetConfirme = ss.getSheetByName(CONFIG.SHEETS.CONFIRME);
  
  // Récupérer toutes les commandes en préparation
  const commandesData = sheetConfirme.getDataRange().getValues();
  let commandesPreparation = [];
  
  for (let i = 1; i < commandesData.length; i++) {
    if (commandesData[i][2] === CONFIG.STATUTS.EN_PREPARATION) {
      commandesPreparation.push(commandesData[i]);
    }
  }
  
  if (commandesPreparation.length === 0) {
    ui.alert('Aucune commande en préparation.');
    return;
  }
  
  // Confirmer l'impression
  const confirmation = ui.alert(
    'Impression des tickets en masse',
    'Vous êtes sur le point d\'imprimer ' + commandesPreparation.length + ' ticket(s) de commandes. Continuer ?',
    ui.ButtonSet.YES_NO
  );
  
  if (confirmation !== ui.Button.YES) {
    return;
  }
  
  // Générer les tickets
  let ticketsGeneres = 0;
  for (let i = 0; i < commandesPreparation.length; i++) {
    if (genererTicket(commandesPreparation[i])) {
      ticketsGeneres++;
    }
  }
  
  ui.alert(ticketsGeneres + ' ticket(s) ont été générés avec succès.');
}

/**
 * Génère un ticket pour une commande
 * 
 * @param {Array} donnees Les données de la commande
 * @return {boolean} True si le ticket a été généré avec succès, sinon False
 */
function genererTicket(donnees) {
  try {
    const ss = SpreadsheetApp.getActive();
    
    // Créer ou récupérer la feuille de tickets
    let sheetTickets = ss.getSheetByName('Tickets');
    if (!sheetTickets) {
      sheetTickets = ss.insertSheet('Tickets');
      sheetTickets.appendRow([
        'Numéro commande',
        'Client',
        'Téléphone',
        'Adresse',
        'Ville',
        'Produit',
        'Quantité',
        'Prix',
        'Date génération',
        'Numéro de ticket'
      ]);
      
      // Formater l'en-tête
      sheetTickets.getRange('A1:J1').setFontWeight('bold');
      sheetTickets.setFrozenRows(1);
    }
    
    // Générer un numéro de ticket unique
    const numeroTicket = 'TK-' + new Date().getTime().toString().substring(7) + Math.floor(Math.random() * 1000).toString().padStart(3, '0');
    
    // Ajouter le ticket à la feuille
    const now = new Date();
    sheetTickets.appendRow([
      donnees[0],                 // Numéro commande
      donnees[4],                 // Client
      donnees[5],                 // Téléphone
      donnees[6],                 // Adresse
      donnees[7],                 // Ville
      donnees[8],                 // Produit
      donnees[9],                 // Quantité
      donnees[10],                // Prix
      formaterDate(now),          // Date génération
      numeroTicket                // Numéro de ticket
    ]);
    
    // Créer une nouvelle feuille pour le ticket imprimable
    const templateTicket = ss.insertSheet('Ticket - ' + numeroTicket);
    
    // Formater le ticket
    templateTicket.getRange('A1:G1').merge();
    templateTicket.getRange('A1').setValue('TICKET DE COMMANDE YOOZAK').setFontSize(16).setFontWeight('bold').setHorizontalAlignment('center');
    
    templateTicket.getRange('A3').setValue('Numéro de ticket:');
    templateTicket.getRange('B3').setValue(numeroTicket).setFontWeight('bold');
    
    templateTicket.getRange('A4').setValue('Numéro de commande:');
    templateTicket.getRange('B4').setValue(donnees[0]).setFontWeight('bold');
    
    templateTicket.getRange('A6').setValue('Client:');
    templateTicket.getRange('B6').setValue(donnees[4]);
    
    templateTicket.getRange('A7').setValue('Téléphone:');
    templateTicket.getRange('B7').setValue(donnees[5]);
    
    templateTicket.getRange('A8').setValue('Adresse:');
    templateTicket.getRange('B8').setValue(donnees[6]);
    
    templateTicket.getRange('A9').setValue('Ville:');
    templateTicket.getRange('B9').setValue(donnees[7]);
    
    templateTicket.getRange('A11').setValue('Produit:');
    templateTicket.getRange('B11').setValue(donnees[8]);
    
    templateTicket.getRange('A12').setValue('Quantité:');
    templateTicket.getRange('B12').setValue(donnees[9]);
    
    templateTicket.getRange('A13').setValue('Prix:');
    templateTicket.getRange('B13').setValue(donnees[10] + ' MAD');
    
    templateTicket.getRange('A15:G15').merge();
    templateTicket.getRange('A15').setValue('Date d\'impression: ' + formaterDate(now)).setHorizontalAlignment('center');
    
    templateTicket.getRange('A17:G17').merge();
    templateTicket.getRange('A17').setValue('MERCI POUR VOTRE COMMANDE!').setFontWeight('bold').setHorizontalAlignment('center');
    
    // Ajuster les largeurs de colonnes
    templateTicket.setColumnWidth(1, 150);
    templateTicket.setColumnWidth(2, 250);
    
    // Enregistrer l'action dans le journal
    enregistrerLog(
      Session.getActiveUser().getEmail(),
      donnees[0],
      'Génération du ticket ' + numeroTicket
    );
    
    return true;
  } catch (e) {
    Logger.log('Erreur lors de la génération du ticket: ' + e.toString());
    return false;
  }
}