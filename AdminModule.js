/**
 * Yoozak - Module Administrateur
 * 
 * Ce script contient les fonctions pour l'administrateur du système
 * Version: 1.0
 * Date: 17/04/2025
 */

/**
 * Implémentation de la fonction d'affectation des commandes aux opérateurs
 * Cette fonction remplace le placeholder dans Main.js
 */
function affecterCommandes() {
  // Vérifier si l'utilisateur est autorisé
  if (!verifierUtilisateur()) {
    SpreadsheetApp.getUi().alert('Vous n\'êtes pas autorisé à utiliser cette fonction.');
    return;
  }
  
  const ui = SpreadsheetApp.getUi();
  const ss = SpreadsheetApp.getActive();
  const sheetInitiale = ss.getSheetByName(CONFIG.SHEETS.INITIALE);
  const sheetConfig = ss.getSheetByName(CONFIG.SHEETS.CONFIG);
  
  // Vérifier que les feuilles requises existent
  if (!sheetInitiale || !sheetConfig) {
    ui.alert('Les feuilles requises n\'existent pas.');
    return;
  }
  
  try {
    // Récupérer la liste des commandes non affectées
    const commandesData = sheetInitiale.getDataRange().getValues();
    let commandesNonAffectees = [];
    
    for (let i = 1; i < commandesData.length; i++) {
      if (commandesData[i][2] === CONFIG.STATUTS.NON_AFFECTEE) {
        commandesNonAffectees.push({
          row: i + 1,
          numeroCommande: commandesData[i][0],
          idSource: commandesData[i][1],
          client: commandesData[i][4],
          produit: commandesData[i][8]
        });
      }
    }
    
    if (commandesNonAffectees.length === 0) {
      ui.alert('Aucune commande non affectée disponible.');
      return;
    }
    
    // Récupérer la liste des opérateurs
    const operateursRange = sheetConfig.getRange("CMD SheetUsers");
    const operateursData = operateursRange.getValues();
    let operateursActifs = [];
    
    for (let i = 1; i < operateursData.length; i++) {
      if (operateursData[i][2] === 'Oui') {
        operateursActifs.push(operateursData[i][0]);
      }
    }
    
    if (operateursActifs.length === 0) {
      ui.alert('Aucun opérateur actif disponible. Veuillez configurer au moins un opérateur.');
      return;
    }
    
    // Préparer la boîte de dialogue de sélection
    let htmlCommandes = '<div style="max-height: 200px; overflow-y: auto;">';
    for (let i = 0; i < commandesNonAffectees.length; i++) {
      htmlCommandes += '<input type="checkbox" id="cmd' + i + '" name="commande" value="' + commandesNonAffectees[i].numeroCommande + '"> ';
      htmlCommandes += '<label for="cmd' + i + '">' + commandesNonAffectees[i].numeroCommande + ' - ' + commandesNonAffectees[i].client + ' - ' + commandesNonAffectees[i].produit + '</label><br>';
    }
    htmlCommandes += '</div>';
    
    let htmlOperateurs = '<select id="operateur">';
    for (let i = 0; i < operateursActifs.length; i++) {
      htmlOperateurs += '<option value="' + operateursActifs[i] + '">' + operateursActifs[i] + '</option>';
    }
    htmlOperateurs += '</select>';
    
    // Demander le mode d'affectation (unitaire/masse)
    const modeResponse = ui.alert(
      'Mode d\'affectation',
      'Comment souhaitez-vous affecter les commandes ?\n\n' +
      'Cliquez sur OUI pour une affectation unitaire (vous sélectionnez les commandes).\n' +
      'Cliquez sur NON pour une affectation en masse (toutes les commandes non affectées).',
      ui.ButtonSet.YES_NO_CANCEL
    );
    
    if (modeResponse === ui.Button.CANCEL) {
      return;
    }
    
    let numCommandesSelectionnees = 0;
    let commandesSelectionnees = [];
    
    if (modeResponse === ui.Button.YES) {
      // Affectation unitaire
      const htmlContent = 
        '<p>Sélectionnez les commandes à affecter :</p>' +
        htmlCommandes +
        '<p>Sélectionnez l\'opérateur :</p>' +
        htmlOperateurs;
      
      const htmlOutput = HtmlService.createHtmlOutput(htmlContent)
        .setWidth(500)
        .setHeight(400)
        .setTitle('Affecter des commandes');
      
      ui.showModalDialog(htmlOutput, 'Affecter des commandes');
      
      // Cette partie est gérée par le code HTML et la fonction d'affectation callback
      return;
    } else {
      // Affectation en masse
      // Demander l'opérateur
      const operateurResponse = ui.prompt(
        'Affectation en masse',
        'Sélectionnez l\'opérateur auquel affecter toutes les commandes non affectées (' + commandesNonAffectees.length + ' commandes) :\n\n' + 
        operateursActifs.join('\n'),
        ui.ButtonSet.OK_CANCEL
      );
      
      if (operateurResponse.getSelectedButton() !== ui.Button.OK) {
        return;
      }
      
      const operateurSelectionne = operateurResponse.getResponseText().trim();
      if (!operateursActifs.includes(operateurSelectionne)) {
        ui.alert('Opérateur non valide.');
        return;
      }
      
      // Confirmer l'affectation
      const confirmResponse = ui.alert(
        'Confirmation d\'affectation',
        'Vous êtes sur le point d\'affecter ' + commandesNonAffectees.length + ' commande(s) à l\'opérateur ' + operateurSelectionne + '. Continuer ?',
        ui.ButtonSet.YES_NO
      );
      
      if (confirmResponse !== ui.Button.YES) {
        return;
      }
      
      // Affecter les commandes
      commandesSelectionnees = commandesNonAffectees;
      numCommandesSelectionnees = affecterCommandesAOperateur(operateurSelectionne, commandesSelectionnees, sheetInitiale);
    }
    
    ui.alert('Affectation terminée', numCommandesSelectionnees + ' commande(s) ont été affectées avec succès.', ui.ButtonSet.OK);
  } catch (e) {
    ui.alert('Erreur', 'Une erreur est survenue lors de l\'affectation des commandes: ' + e.toString(), ui.ButtonSet.OK);
    Logger.log('Erreur lors de l\'affectation des commandes: ' + e.toString());
  }
}

/**
 * Fonction d'affectation de commandes à un opérateur
 * 
 * @param {string} operateur Nom de l'opérateur
 * @param {Array} commandes Liste des commandes à affecter
 * @param {Sheet} sheetInitiale Feuille initiale des commandes
 * @return {number} Nombre de commandes affectées
 */
function affecterCommandesAOperateur(operateur, commandes, sheetInitiale) {
  const ss = SpreadsheetApp.getActive();
  
  // Vérifier si la feuille de l'opérateur existe, sinon la créer
  let sheetOperateur = ss.getSheetByName('TMP ' + operateur);
  if (!sheetOperateur) {
    sheetOperateur = ss.insertSheet('TMP ' + operateur);
    sheetOperateur.appendRow([
      'Numéro commande',
      'ID commande source',
      'Action',
      'Nom client',
      'Téléphone',
      'Adresse',
      'Ville',
      'Produit',
      'Quantité',
      'Prix',
      'Date commande'
    ]);
    
    // Formater l'en-tête
    const headerRange = sheetOperateur.getRange('A1:K1');
    headerRange.setFontWeight('bold');
    headerRange.setBackground('#e6e6e6');
    
    // Figer la première ligne
    sheetOperateur.setFrozenRows(1);
    
    // Ajuster les largeurs de colonnes
    sheetOperateur.setColumnWidth(1, 150); // A
    sheetOperateur.setColumnWidth(2, 150); // B
    sheetOperateur.setColumnWidth(3, 150); // C
    sheetOperateur.setColumnWidth(4, 150); // D
    sheetOperateur.setColumnWidth(6, 200); // F
    sheetOperateur.setColumnWidth(8, 150); // H
    sheetOperateur.setColumnWidth(11, 150); // K
  }
  
  // Compter les commandes affectées avec succès
  let commandesAffectees = 0;
  
  // Parcourir les commandes à affecter
  for (let i = 0; i < commandes.length; i++) {
    const commande = commandes[i];
    const row = commande.row;
    
    // Récupérer les données de la commande
    const commandeData = sheetInitiale.getRange(row, 1, 1, 13).getValues()[0];
    
    // Mettre à jour le statut et l'opérateur dans la feuille initiale
    sheetInitiale.getRange(row, 3).setValue(CONFIG.STATUTS.AFFECTEE);
    sheetInitiale.getRange(row, 4).setValue(operateur);
    
    // Ajouter la commande à la feuille de l'opérateur
    sheetOperateur.appendRow([
      commandeData[0],  // Numéro commande
      commandeData[1],  // ID commande source
      '',               // Action (vide initialement)
      commandeData[4],  // Nom client
      commandeData[5],  // Téléphone
      commandeData[6],  // Adresse
      commandeData[7],  // Ville
      commandeData[8],  // Produit
      commandeData[9],  // Quantité
      commandeData[10], // Prix
      commandeData[12]  // Date commande
    ]);
    
    commandesAffectees++;
    
    // Enregistrer l'action dans le journal
    enregistrerLog(
      Session.getActiveUser().getEmail(),
      commandeData[0],
      'Commande affectée à l\'opérateur ' + operateur
    );
  }
  
  return commandesAffectees;
}

/**
 * Implémentation de la fonction de désaffectation des commandes
 * Cette fonction remplace le placeholder dans Main.js
 */
function desaffecterCommandes() {
  // Vérifier si l'utilisateur est autorisé
  if (!verifierUtilisateur()) {
    SpreadsheetApp.getUi().alert('Vous n\'êtes pas autorisé à utiliser cette fonction.');
    return;
  }
  
  const ui = SpreadsheetApp.getUi();
  const ss = SpreadsheetApp.getActive();
  const sheetInitiale = ss.getSheetByName(CONFIG.SHEETS.INITIALE);
  
  // Vérifier que la feuille requise existe
  if (!sheetInitiale) {
    ui.alert('La feuille requise n\'existe pas.');
    return;
  }
  
  try {
    // Demander le numéro de commande à désaffecter
    const response = ui.prompt(
      'Désaffecter une commande',
      'Entrez le numéro de la commande à désaffecter:',
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
    
    // Rechercher la commande dans la feuille initiale
    const commandesData = sheetInitiale.getDataRange().getValues();
    let commandeTrouvee = false;
    let rowInitiale = 0;
    let operateur = '';
    
    for (let i = 1; i < commandesData.length; i++) {
      if (commandesData[i][0] === numeroCommande) {
        commandeTrouvee = true;
        rowInitiale = i + 1;
        operateur = commandesData[i][3];
        break;
      }
    }
    
    if (!commandeTrouvee) {
      ui.alert('Commande non trouvée.');
      return;
    }
    
    // Vérifier que la commande est affectée
    if (operateur === '' || commandesData[rowInitiale - 1][2] !== CONFIG.STATUTS.AFFECTEE) {
      ui.alert('Cette commande n\'est pas affectée à un opérateur.');
      return;
    }
    
    // Confirmer la désaffectation
    const confirmation = ui.alert(
      'Confirmer la désaffectation',
      'Vous êtes sur le point de désaffecter la commande ' + numeroCommande + ' de l\'opérateur ' + operateur + '. Continuer ?',
      ui.ButtonSet.YES_NO
    );
    
    if (confirmation !== ui.Button.YES) {
      return;
    }
    
    // Désaffecter la commande
    // 1. Mettre à jour la feuille initiale
    sheetInitiale.getRange(rowInitiale, 3).setValue(CONFIG.STATUTS.NON_AFFECTEE);
    sheetInitiale.getRange(rowInitiale, 4).setValue('');
    
    // 2. Supprimer la commande de la feuille de l'opérateur
    const sheetOperateur = ss.getSheetByName('TMP ' + operateur);
    if (sheetOperateur) {
      const commandesOperateur = sheetOperateur.getDataRange().getValues();
      for (let i = 1; i < commandesOperateur.length; i++) {
        if (commandesOperateur[i][0] === numeroCommande) {
          sheetOperateur.deleteRow(i + 1);
          break;
        }
      }
    }
    
    // Enregistrer l'action dans le journal
    enregistrerLog(
      Session.getActiveUser().getEmail(),
      numeroCommande,
      'Commande désaffectée de l\'opérateur ' + operateur
    );
    
    ui.alert('La commande a été désaffectée avec succès.');
  } catch (e) {
    ui.alert('Erreur', 'Une erreur est survenue lors de la désaffectation de la commande: ' + e.toString(), ui.ButtonSet.OK);
    Logger.log('Erreur lors de la désaffectation de la commande: ' + e.toString());
  }
}

/**
 * Implémentation de la fonction de résolution des problèmes
 * Cette fonction remplace le placeholder dans Main.js
 */
function resoudreProblemes() {
  // Vérifier si l'utilisateur est autorisé
  if (!verifierUtilisateur()) {
    SpreadsheetApp.getUi().alert('Vous n\'êtes pas autorisé à utiliser cette fonction.');
    return;
  }
  
  const ui = SpreadsheetApp.getUi();
  const ss = SpreadsheetApp.getActive();
  const sheetProblem = ss.getSheetByName(CONFIG.SHEETS.PROBLEM);
  
  // Vérifier que la feuille requise existe
  if (!sheetProblem) {
    ui.alert('La feuille des problèmes n\'existe pas.');
    return;
  }
  
  // Récupérer la liste des problèmes
  const problemesData = sheetProblem.getDataRange().getValues();
  if (problemesData.length <= 1) {
    ui.alert('Aucun problème à résoudre.');
    return;
  }
  
  // Préparer la liste des problèmes
  let htmlProblemes = '<div style="max-height: 300px; overflow-y: auto;"><table style="width:100%;border-collapse:collapse;">';
  htmlProblemes += '<tr style="background-color:#e6e6e6;font-weight:bold;"><th>ID</th><th>Date</th><th>Source</th><th>Description</th><th>Client</th><th>Produit</th></tr>';
  
  for (let i = 1; i < problemesData.length; i++) {
    htmlProblemes += '<tr style="border-bottom:1px solid #ddd;">';
    htmlProblemes += '<td>' + i + '</td>';
    htmlProblemes += '<td>' + problemesData[i][1] + '</td>'; // Date formatée
    htmlProblemes += '<td>' + problemesData[i][2] + '</td>'; // ID commande source
    htmlProblemes += '<td>' + problemesData[i][3] + '</td>'; // Description problème
    htmlProblemes += '<td>' + problemesData[i][4] + '</td>'; // Nom client
    htmlProblemes += '<td>' + problemesData[i][8] + '</td>'; // Produit
    htmlProblemes += '</tr>';
  }
  
  htmlProblemes += '</table></div>';
  
  // Demander l'ID du problème à résoudre
  const response = ui.prompt(
    'Résoudre un problème',
    'Liste des problèmes :\n\n' + htmlProblemes.replace(/<[^>]*>/g, '') + '\n\n' +
    'Entrez l\'ID du problème à résoudre (numéro de ligne):',
    ui.ButtonSet.OK_CANCEL
  );
  
  if (response.getSelectedButton() !== ui.Button.OK) {
    return;
  }
  
  const idProbleme = parseInt(response.getResponseText().trim());
  if (isNaN(idProbleme) || idProbleme < 1 || idProbleme >= problemesData.length) {
    ui.alert('ID de problème invalide.');
    return;
  }
  
  // Récupérer les données du problème
  const probleme = problemesData[idProbleme];
  
  // Demander l'action à effectuer
  const actionResponse = ui.alert(
    'Action',
    'Que souhaitez-vous faire avec ce problème ?\n\n' +
    'Cliquez sur OUI pour créer une nouvelle commande à partir de ce problème.\n' +
    'Cliquez sur NON pour simplement marquer le problème comme résolu et le supprimer.',
    ui.ButtonSet.YES_NO_CANCEL
  );
  
  if (actionResponse === ui.Button.CANCEL) {
    return;
  }
  
  if (actionResponse === ui.Button.YES) {
    // Créer une nouvelle commande
    const sheetInitiale = ss.getSheetByName(CONFIG.SHEETS.INITIALE);
    if (!sheetInitiale) {
      ui.alert('La feuille initiale n\'existe pas.');
      return;
    }
    
    // Générer un numéro de commande unique
    const numeroCommande = genererNumeroCommande('P'); // P pour Problème résolu
    
    // Ajouter la commande à la feuille initiale
    const now = new Date();
    sheetInitiale.appendRow([
      numeroCommande,
      probleme[2],             // ID commande source
      CONFIG.STATUTS.NON_AFFECTEE,
      '',                      // Pas d'opérateur initialement
      probleme[4],             // Nom client
      probleme[5],             // Téléphone
      probleme[6],             // Adresse
      probleme[7],             // Ville
      probleme[8],             // Produit
      probleme[9],             // Quantité
      probleme[10],            // Prix
      now,                     // Date de création
      formaterDate(now)        // Date formatée
    ]);
    
    // Enregistrer l'action dans le journal
    enregistrerLog(
      Session.getActiveUser().getEmail(),
      numeroCommande,
      'Création d\'une nouvelle commande à partir d\'un problème résolu'
    );
  }
  
  // Supprimer le problème de la feuille
  sheetProblem.deleteRow(idProbleme + 1);
  
  ui.alert('Le problème a été traité avec succès.');
}

/**
 * Ajoute une commande à la feuille d'un opérateur
 * 
 * @param {string} operateur Nom de l'opérateur
 * @param {Array} commandeData Les données de la commande
 * @return {boolean} True si l'opération a réussi, sinon False
 */
function ajouterCommandeAOperateur(operateur, commandeData) {
  const ss = SpreadsheetApp.getActive();
  
  // Vérifier si la feuille de l'opérateur existe, sinon la créer
  let sheetOperateur = ss.getSheetByName('TMP ' + operateur);
  if (!sheetOperateur) {
    sheetOperateur = ss.insertSheet('TMP ' + operateur);
    sheetOperateur.appendRow([
      'Numéro commande',
      'ID commande source',
      'Action',
      'Nom client',
      'Téléphone',
      'Adresse',
      'Ville',
      'Produit',
      'Quantité',
      'Prix',
      'Date commande'
    ]);
    
    // Formater l'en-tête
    const headerRange = sheetOperateur.getRange('A1:K1');
    headerRange.setFontWeight('bold');
    headerRange.setBackground('#e6e6e6');
    
    // Figer la première ligne
    sheetOperateur.setFrozenRows(1);
    
    // Ajuster les largeurs de colonnes
    sheetOperateur.setColumnWidth(1, 150); // A
    sheetOperateur.setColumnWidth(2, 150); // B
    sheetOperateur.setColumnWidth(3, 150); // C
    sheetOperateur.setColumnWidth(4, 150); // D
    sheetOperateur.setColumnWidth(6, 200); // F
    sheetOperateur.setColumnWidth(8, 150); // H
    sheetOperateur.setColumnWidth(11, 150); // K
  }
  
  // Ajouter la commande à la feuille de l'opérateur
  sheetOperateur.appendRow([
    commandeData[0],  // Numéro commande
    commandeData[1],  // ID commande source
    '',               // Action (vide initialement)
    commandeData[3],  // Nom client
    commandeData[4],  // Téléphone
    commandeData[5],  // Adresse
    commandeData[6],  // Ville
    commandeData[7],  // Produit
    commandeData[8],  // Quantité
    commandeData[9],  // Prix
    commandeData[11]  // Date formatée
  ]);
  
  return true;
}