/**
 * Yoozak - Utilitaires pour les formulaires
 * 
 * Ce script contient les fonctions utilitaires utilisées par les formulaires HTML
 * Version: 1.0
 * Date: 17/04/2025
 */

/**
 * Récupère la liste des produits pour les formulaires
 * 
 * @return {Array} Liste des produits [nom, prix, stock]
 */
function getProduits() {
  const ss = SpreadsheetApp.getActive();
  const sheetConfig = ss.getSheetByName(CONFIG.SHEETS.CONFIG);
  
  if (!sheetConfig) {
    return [];
  }
  
  try {
    const produitsRange = sheetConfig.getRange("CMD Products");
    const produitsData = produitsRange.getValues();
    
    // Supprimer la première ligne (en-têtes)
    return produitsData.slice(1);
  } catch (e) {
    Logger.log('Erreur lors de la récupération des produits: ' + e.toString());
    return [];
  }
}

/**
 * Récupère la liste des villes pour les formulaires
 * 
 * @return {Array} Liste des villes [nom, région]
 */
function getVilles() {
  const ss = SpreadsheetApp.getActive();
  const sheetConfig = ss.getSheetByName(CONFIG.SHEETS.CONFIG);
  
  if (!sheetConfig) {
    return [];
  }
  
  try {
    const villesRange = sheetConfig.getRange("CMD Region");
    const villesData = villesRange.getValues();
    
    // Supprimer la première ligne (en-têtes)
    return villesData.slice(1);
  } catch (e) {
    Logger.log('Erreur lors de la récupération des villes: ' + e.toString());
    return [];
  }
}

/**
 * Inclut un fichier HTML dans un autre
 * 
 * @param {string} filename Nom du fichier à inclure
 * @return {string} Contenu du fichier
 */
function include(filename) {
  return HtmlService.createHtmlOutputFromFile(filename).getContent();
}

/**
 * Retourne un timestamp unique pour éviter le cache des fichiers CSS/JS
 * 
 * @return {string} Timestamp
 */
function getTimestamp() {
  return new Date().getTime().toString();
}