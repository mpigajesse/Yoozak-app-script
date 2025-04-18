/**
 * Yoozak - Système de gestion des commandes
 * 
 * Ce script contient les fonctions principales pour la gestion des commandes Yoozak
 * Version: 1.0
 * Date: 17/04/2025
 */

// Configuration globale
const CONFIG = {
  SHEETS: {
    CONFIG: 'CMD config',
    INITIALE: 'CMD initiale',
    TMP: 'CMD TMP',
    PRODUITS: 'CMD produits',
    TMP_LOG: 'CMD TMP LOG',
    CONFIRME: 'CMD confirme',
    CONFIRME_LOG: 'CMD confirme LOG',
    ANNULEE: 'CMD Annulée',
    RETOURNEE: 'CMD Retournée',
    PROBLEM: 'CMD Problem'
  },
  STATUTS: {
    AFFECTEE: 'Affectée',
    NON_AFFECTEE: 'Non affectée',
    PROBLEME: 'Problème',
    CONFIRMEE: 'Confirmé',
    EN_PREPARATION: 'En cours de préparation',
    EXPEDIE: 'Expédié',
    LIVRE: 'Livré',
    RETOURNE: 'Retourné',
    ANNULE: 'Annulé'
  },
  SOURCES: {
    YOUCAN: 'Y',
    SHOPIFY: 'S'
  }
};

/**
 * Point d'entrée de l'application web
 * Cette fonction remplace onOpen() pour une application web
 */
function doGet(e) {
  // Vérifier si l'utilisateur est autorisé
  if (!verifierUtilisateur()) {
    return HtmlService.createHtmlOutput('<h1>Accès non autorisé</h1>')
      .setTitle('Yoozak - Accès refusé');
  }
  
  // Déterminer la page à afficher en fonction du paramètre
  let page = e.parameter.page || 'dashboard';
  let htmlOutput;
  
  switch (page) {
    case 'admin':
      htmlOutput = HtmlService.createTemplateFromFile('AdminPanel').evaluate();
      break;
    case 'operateur':
      htmlOutput = HtmlService.createTemplateFromFile('OperateurPanel').evaluate();
      break;
    case 'logistique':
      htmlOutput = HtmlService.createTemplateFromFile('LogistiquePanel').evaluate();
      break;
    case 'cmdinit':
      htmlOutput = HtmlService.createTemplateFromFile('CMDInitPanel').evaluate();
      break;
    default:
      // Page d'accueil/dashboard par défaut
      htmlOutput = HtmlService.createTemplateFromFile('DashboardPanel').evaluate();
  }
  
  return htmlOutput
    .setTitle('Yoozak - Gestion des Commandes')
    .setFaviconUrl('https://www.google.com/images/icons/product/sheets-32.png');
}

/**
 * Fonction onOpen qui s'exécute à l'ouverture du document
 * Crée le menu personnalisé pour l'application
 * Maintenue pour la compatibilité avec Google Sheets
 */
function onOpen() {
  const ui = SpreadsheetApp.getUi();
  ui.createMenu('Yoozak')
    .addItem('Initialiser le système', 'initialiserSysteme')
    .addSeparator()
    .addSubMenu(ui.createMenu('Administrateur')
      .addItem('Affecter commandes', 'affecterCommandes')
      .addItem('Désaffecter commandes', 'desaffecterCommandes')
      .addItem('Résoudre problèmes', 'resoudreProblemes'))
    .addSubMenu(ui.createMenu('Opérateur')
      .addItem('Créer commande', 'creerCommande')
      .addItem('Modifier commande', 'modifierCommande')
      .addItem('Confirmer commande', 'confirmerCommande')
      .addItem('Annuler commande', 'annulerCommande'))
    .addSubMenu(ui.createMenu('Logistique')
      .addItem('Changer statut commande', 'changerStatutCommande')
      .addItem('Imprimer tickets', 'imprimerTickets'))
    .addToUi();
}

/**
 * Vérifie si l'utilisateur est autorisé à accéder à l'application
 * 
 * @return {boolean} True si l'utilisateur est autorisé, sinon False
 */
function verifierUtilisateur() {
  const email = Session.getActiveUser().getEmail();
  const sheet = SpreadsheetApp.getActive().getSheetByName(CONFIG.SHEETS.CONFIG);
  
  // Si la feuille de configuration n'existe pas encore, autoriser l'accès
  if (!sheet) {
    return true;
  }
  
  try {
    // Récupérer la liste des utilisateurs autorisés
    const usersRange = sheet.getRange("CMD AllowedUsers");
    const usersValues = usersRange.getValues();
    
    for (let i = 0; i < usersValues.length; i++) {
      if (usersValues[i][0] === email && usersValues[i][2] === 'Oui') {
        return true;
      }
    }
  } catch (e) {
    // Si la plage nommée n'existe pas encore, autoriser l'accès
    return true;
  }
  
  return false;
}

/**
 * Fonction pour formater un numéro de téléphone marocain
 * 
 * @param {string} telephone Le numéro de téléphone à formater
 * @return {string} Le numéro de téléphone formaté
 */
function formaterTelephone(telephone) {
  if (!telephone) return '';
  
  // Supprimer tous les caractères non numériques
  let numero = telephone.toString().replace(/\D/g, '');
  
  // Vérifier si c'est un numéro marocain
  if (numero.startsWith('212')) {
    if (numero.length === 12) {
      return '0' + numero.substring(3);
    }
  } else if (numero.startsWith('0')) {
    if (numero.length === 10) {
      return numero;
    }
  }
  
  // Si le format ne correspond pas, retourner le numéro tel quel
  return telephone;
}

/**
 * Fonction pour formater une date selon le format marocain
 * 
 * @param {Date|string} date La date à formater
 * @return {string} La date formatée
 */
function formaterDate(date) {
  if (!date) return '';
  
  let dateObj;
  if (typeof date === 'string') {
    dateObj = new Date(date);
  } else {
    dateObj = date;
  }
  
  return Utilities.formatDate(dateObj, Session.getScriptTimeZone(), 'dd/MM/yyyy HH:mm:ss');
}

/**
 * Génère un numéro de commande unique
 * 
 * @param {string} source La source de la commande (Y: Youcan, S: Shopify, etc.)
 * @param {string} codeOperateur Code de l'opérateur (facultatif)
 * @return {string} Le numéro de commande généré
 */
function genererNumeroCommande(source, codeOperateur) {
  const prefix = source + (codeOperateur || '');
  const date = new Date();
  const timestamp = date.getTime().toString().substring(6);
  const random = Math.floor(Math.random() * 1000).toString().padStart(3, '0');
  
  return prefix + timestamp + random;
}

/**
 * Fonction pour enregistrer une action dans le journal
 * 
 * @param {string} operateur Nom de l'opérateur
 * @param {string} numeroCommande Numéro de la commande
 * @param {string} action Description de l'action
 */
function enregistrerLog(operateur, numeroCommande, action) {
  const sheet = SpreadsheetApp.getActive().getSheetByName(CONFIG.SHEETS.TMP_LOG);
  const now = new Date();
  
  sheet.appendRow([
    operateur,
    numeroCommande,
    action,
    now,
    formaterDate(now)
  ]);
}

/**
 * Fonction exécutée lors de la modification d'une cellule
 * 
 * @param {Object} e L'événement de modification
 */
function onEdit(e) {
  const range = e.range;
  const sheet = range.getSheet();
  const sheetName = sheet.getName();
  
  // Vérifier si la modification concerne une feuille d'opérateur
  if (sheetName.startsWith('TMP ')) {
    const row = range.getRow();
    const col = range.getColumn();
    
    // Si la modification concerne la colonne "Action" (colonne C = 3)
    if (col === 3 && row > 1) {
      const operateur = sheetName.substring(4); // Extraire le nom de l'opérateur
      const numeroCommande = sheet.getRange(row, 1).getValue();
      const nouvelleAction = e.value;
      
      // Enregistrer l'action dans le journal
      enregistrerLog(
        operateur,
        numeroCommande,
        'Action modifiée: ' + nouvelleAction
      );
    }
  }
}