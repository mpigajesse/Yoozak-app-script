/**
 * Yoozak - Module CMDInit
 * 
 * Ce script contient les fonctions pour l'importation et le traitement initial des commandes
 * Version: 1.0
 * Date: 17/04/2025
 */

/**
 * Récupère et retourne l'email de l'utilisateur actif
 * Utilisé par l'interface web
 */
function getActiveUserEmail() {
    return Session.getActiveUser().getEmail();
  }
  
  /**
   * Récupère les données du tableau de bord pour l'interface web
   */
  function getDashboardData() {
    const ss = getSpreadsheet();
    const sheetInitiale = ss.getSheetByName(CONFIG.SHEETS.INITIALE);
    const sheetConfirme = ss.getSheetByName(CONFIG.SHEETS.CONFIRME);
    const sheetProblem = ss.getSheetByName(CONFIG.SHEETS.PROBLEM);
    const sheetLog = ss.getSheetByName(CONFIG.SHEETS.TMP_LOG);
    
    // Nombre total de commandes
    const totalCommandes = sheetInitiale ? sheetInitiale.getLastRow() - 1 : 0;
    
    // Nombre de commandes confirmées
    const commandesConfirmees = sheetConfirme ? sheetConfirme.getLastRow() - 1 : 0;
    
    // Nombre de commandes en attente
    const commandesAttente = totalCommandes;
    
    // Nombre de commandes avec problèmes
    const commandesProblemes = sheetProblem ? sheetProblem.getLastRow() - 1 : 0;
    
    // Activité récente
    let activiteRecente = [];
    if (sheetLog && sheetLog.getLastRow() > 1) {
      const logData = sheetLog.getRange(2, 1, Math.min(10, sheetLog.getLastRow() - 1), 5).getValues();
      logData.forEach(function(row) {
        activiteRecente.push({
          operateur: row[0],
          numeroCommande: row[1],
          action: row[2],
          date: row[4]
        });
      });
    }
    
    return {
      totalCommandes: totalCommandes,
      commandesConfirmees: commandesConfirmees,
      commandesAttente: commandesAttente,
      commandesProblemes: commandesProblemes,
      activiteRecente: activiteRecente
    };
  }
  
  /**
   * Importe les données CSV de Shopify
   * 
   * @param {string} csvContent Le contenu du fichier CSV
   * @return {Object} Résultat de l'importation
   */
  function importShopifyCSV(csvContent) {
    // Créer une feuille temporaire pour l'importation
    const ss = getSpreadsheet();
    let tempSheet = ss.getSheetByName('TempImport');
    
    if (tempSheet) {
      ss.deleteSheet(tempSheet);
    }
    
    tempSheet = ss.insertSheet('TempImport');
    
    // Convertir le CSV en données tabulaires
    const csvData = Utilities.parseCsv(csvContent);
    
    // Si le CSV est vide, retourner une erreur
    if (csvData.length <= 1) {
      ss.deleteSheet(tempSheet);
      return { success: false, message: 'Le fichier CSV est vide ou mal formaté.' };
    }
    
    // Écrire les données dans la feuille temporaire
    tempSheet.getRange(1, 1, csvData.length, csvData[0].length).setValues(csvData);
    
    // Traiter les données selon le format Shopify
    const shopifyData = [];
    
    // Récupérer les indices des colonnes pertinentes (à adapter selon le format CSV de Shopify)
    const orderIdCol = findColumnIndex(csvData[0], 'Name', 'Order ID', 'Order Number');
    const customerNameCol = findColumnIndex(csvData[0], 'Shipping Name', 'Customer Name', 'Billing Name');
    const phoneCol = findColumnIndex(csvData[0], 'Phone', 'Shipping Phone', 'Billing Phone');
    const addressCol = findColumnIndex(csvData[0], 'Shipping Address', 'Shipping Address1', 'Address');
    const cityCol = findColumnIndex(csvData[0], 'Shipping City', 'City', 'Billing City');
    const productCol = findColumnIndex(csvData[0], 'Lineitem name', 'Product Name', 'Item Name');
    const quantityCol = findColumnIndex(csvData[0], 'Lineitem quantity', 'Quantity', 'Item Quantity');
    const priceCol = findColumnIndex(csvData[0], 'Subtotal', 'Total', 'Price');
    const dateCol = findColumnIndex(csvData[0], 'Created at', 'Order Date', 'Date');
    
    // Vérifier que toutes les colonnes nécessaires sont présentes
    if (orderIdCol < 0 || customerNameCol < 0 || productCol < 0) {
      ss.deleteSheet(tempSheet);
      return { success: false, message: 'Format CSV non reconnu. Colonnes manquantes.' };
    }
    
    // Parcourir les données (en sautant l'en-tête)
    for (let i = 1; i < csvData.length; i++) {
      const row = csvData[i];
      
      // Vérifier que les données minimales sont présentes
      if (!row[orderIdCol] || !row[customerNameCol]) {
        continue;
      }
      
      shopifyData.push({
        idCommande: 'S-' + row[orderIdCol],
        client: row[customerNameCol] || '',
        telephone: row[phoneCol] || '',
        adresse: row[addressCol] || '',
        ville: row[cityCol] || '',
        produit: row[productCol] || '',
        quantite: row[quantityCol] || '1',
        prix: row[priceCol] || '0',
        date: row[dateCol] || new Date().toISOString(),
        source: 'S'
      });
    }
    
    // Supprimer la feuille temporaire
    ss.deleteSheet(tempSheet);
    
    // Enregistrer les données dans la feuille d'importation Shopify
    saveImportedData(shopifyData, 'Shopify');
    
    return { 
      success: true, 
      message: `${shopifyData.length} commandes importées depuis Shopify.`,
      count: shopifyData.length
    };
  }
  
  /**
   * Importe les données CSV de Youcan
   * 
   * @param {string} csvContent Le contenu du fichier CSV
   * @return {Object} Résultat de l'importation
   */
  function importYoucanCSV(csvContent) {
    // Créer une feuille temporaire pour l'importation
    const ss = getSpreadsheet();
    let tempSheet = ss.getSheetByName('TempImport');
    
    if (tempSheet) {
      ss.deleteSheet(tempSheet);
    }
    
    tempSheet = ss.insertSheet('TempImport');
    
    // Convertir le CSV en données tabulaires
    const csvData = Utilities.parseCsv(csvContent);
    
    // Si le CSV est vide, retourner une erreur
    if (csvData.length <= 1) {
      ss.deleteSheet(tempSheet);
      return { success: false, message: 'Le fichier CSV est vide ou mal formaté.' };
    }
    
    // Écrire les données dans la feuille temporaire
    tempSheet.getRange(1, 1, csvData.length, csvData[0].length).setValues(csvData);
    
    // Traiter les données selon le format Youcan
    const youcanData = [];
    
    // Récupérer les indices des colonnes pertinentes (à adapter selon le format CSV de Youcan)
    const orderIdCol = findColumnIndex(csvData[0], 'Order ID', 'ID', 'Order Number');
    const customerNameCol = findColumnIndex(csvData[0], 'Customer Name', 'Client', 'Full Name');
    const phoneCol = findColumnIndex(csvData[0], 'Phone', 'Téléphone', 'Contact');
    const addressCol = findColumnIndex(csvData[0], 'Shipping Address', 'Adresse', 'Address');
    const cityCol = findColumnIndex(csvData[0], 'City', 'Ville', 'Shipping City');
    const productCol = findColumnIndex(csvData[0], 'Product Name', 'Produit', 'Item');
    const quantityCol = findColumnIndex(csvData[0], 'Quantity', 'Quantité', 'Item Quantity');
    const priceCol = findColumnIndex(csvData[0], 'Total', 'Prix', 'Amount');
    const dateCol = findColumnIndex(csvData[0], 'Created At', 'Date', 'Order Date');
    
    // Vérifier que toutes les colonnes nécessaires sont présentes
    if (orderIdCol < 0 || customerNameCol < 0 || productCol < 0) {
      ss.deleteSheet(tempSheet);
      return { success: false, message: 'Format CSV non reconnu. Colonnes manquantes.' };
    }
    
    // Parcourir les données (en sautant l'en-tête)
    for (let i = 1; i < csvData.length; i++) {
      const row = csvData[i];
      
      // Vérifier que les données minimales sont présentes
      if (!row[orderIdCol] || !row[customerNameCol]) {
        continue;
      }
      
      youcanData.push({
        idCommande: 'Y-' + row[orderIdCol],
        client: row[customerNameCol] || '',
        telephone: row[phoneCol] || '',
        adresse: row[addressCol] || '',
        ville: row[cityCol] || '',
        produit: row[productCol] || '',
        quantite: row[quantityCol] || '1',
        prix: row[priceCol] || '0',
        date: row[dateCol] || new Date().toISOString(),
        source: 'Y'
      });
    }
    
    // Supprimer la feuille temporaire
    ss.deleteSheet(tempSheet);
    
    // Enregistrer les données dans la feuille d'importation Youcan
    saveImportedData(youcanData, 'Youcan');
    
    return { 
      success: true, 
      message: `${youcanData.length} commandes importées depuis Youcan.`,
      count: youcanData.length
    };
  }
  
  /**
   * Enregistre les données importées dans la feuille appropriée
   * 
   * @param {Array} data Les données à enregistrer
   * @param {string} source La source des données (Shopify ou Youcan)
   */
  function saveImportedData(data, source) {
    const ss = getSpreadsheet();
    let sheetName = 'Import ' + source;
    let sheet = ss.getSheetByName(sheetName);
    
    // Créer la feuille si elle n'existe pas
    if (!sheet) {
      sheet = ss.insertSheet(sheetName);
      sheet.appendRow([
        'ID Commande', 'Client', 'Téléphone', 'Adresse', 'Ville', 
        'Produit', 'Quantité', 'Prix', 'Date', 'Source'
      ]);
    } else {
      // Effacer les données existantes (sauf l'en-tête)
      if (sheet.getLastRow() > 1) {
        sheet.deleteRows(2, sheet.getLastRow() - 1);
      }
    }
    
    // Ajouter les nouvelles données
    if (data.length > 0) {
      const rowsToAdd = data.map(function(item) {
        return [
          item.idCommande,
          item.client,
          item.telephone,
          item.adresse,
          item.ville,
          item.produit,
          item.quantite,
          item.prix,
          item.date,
          item.source
        ];
      });
      
      sheet.getRange(2, 1, rowsToAdd.length, 10).setValues(rowsToAdd);
    }
  }
  
  /**
   * Trouve l'index d'une colonne dans un tableau d'en-têtes
   * 
   * @param {Array} headers Le tableau d'en-têtes
   * @param {...string} possibleNames Les noms possibles de la colonne
   * @return {number} L'index de la colonne trouvée ou -1 si non trouvée
   */
  function findColumnIndex(headers, ...possibleNames) {
    for (let name of possibleNames) {
      for (let i = 0; i < headers.length; i++) {
        if (headers[i] && headers[i].toString().toLowerCase() === name.toLowerCase()) {
          return i;
        }
      }
    }
    return -1;
  }
  
  /**
   * Récupère la liste des commandes importées
   * 
   * @return {Array} La liste des commandes importées
   */
  function getImportedCommandes() {
    try {
      Logger.log("Début de getImportedCommandes");
      const ss = getSpreadsheet();
      
      // Vérifier et créer les feuilles si elles n'existent pas
      let sheetShopify = ss.getSheetByName('Import Shopify');
      let sheetYoucan = ss.getSheetByName('Import Youcan');
      
      if (!sheetShopify) {
        Logger.log("Création de la feuille 'Import Shopify'");
        sheetShopify = ss.insertSheet('Import Shopify');
        sheetShopify.appendRow([
          'ID Commande', 'Client', 'Téléphone', 'Adresse', 'Ville', 
          'Produit', 'Quantité', 'Prix', 'Date', 'Source'
        ]);
        // Mettre en forme l'en-tête
        sheetShopify.getRange(1, 1, 1, 10).setFontWeight('bold').setBackground('#f3f3f3');
      }
      
      if (!sheetYoucan) {
        Logger.log("Création de la feuille 'Import Youcan'");
        sheetYoucan = ss.insertSheet('Import Youcan');
        sheetYoucan.appendRow([
          'ID Commande', 'Client', 'Téléphone', 'Adresse', 'Ville', 
          'Produit', 'Quantité', 'Prix', 'Date', 'Source'
        ]);
        // Mettre en forme l'en-tête
        sheetYoucan.getRange(1, 1, 1, 10).setFontWeight('bold').setBackground('#f3f3f3');
      }
      
      let commandes = [];
      
      // Récupérer les commandes Shopify
      if (sheetShopify && sheetShopify.getLastRow() > 1) {
        Logger.log("Récupération des données Shopify");
        const shopifyData = sheetShopify.getRange(2, 1, sheetShopify.getLastRow() - 1, 10).getValues();
        shopifyData.forEach(function(row) {
          commandes.push({
            idCommande: row[0],
            client: row[1],
            telephone: row[2],
            adresse: row[3],
            ville: row[4],
            produit: row[5],
            quantite: row[6],
            prix: row[7],
            date: row[8],
            source: row[9],
            statut: 'non_verifie'
          });
        });
      } else {
        Logger.log("Pas de données Shopify à récupérer");
      }
      
      // Récupérer les commandes Youcan
      if (sheetYoucan && sheetYoucan.getLastRow() > 1) {
        Logger.log("Récupération des données Youcan");
        const youcanData = sheetYoucan.getRange(2, 1, sheetYoucan.getLastRow() - 1, 10).getValues();
        youcanData.forEach(function(row) {
          commandes.push({
            idCommande: row[0],
            client: row[1],
            telephone: row[2],
            adresse: row[3],
            ville: row[4],
            produit: row[5],
            quantite: row[6],
            prix: row[7],
            date: row[8],
            source: row[9],
            statut: 'non_verifie'
          });
        });
      } else {
        Logger.log("Pas de données Youcan à récupérer");
      }
      
      Logger.log("Vérification des commandes");
      // Détecter les doublons et les problèmes
      return verifierCommandes(commandes);
    } catch (err) {
      Logger.log("Erreur dans getImportedCommandes: " + err.toString());
      // Renvoyer un tableau vide en cas d'erreur
      return [];
    }
  }
  
  /**
   * Vérifie les commandes pour détecter les doublons et les problèmes
   * 
   * @param {Array} commandes La liste des commandes à vérifier
   * @return {Array} La liste des commandes avec leur statut
   */
  function verifierCommandes(commandes) {
    const ss = getSpreadsheet();
    const sheetInitiale = ss.getSheetByName(CONFIG.SHEETS.INITIALE);
    
    let commandesExistantes = [];
    
    // Récupérer les ID des commandes existantes
    if (sheetInitiale && sheetInitiale.getLastRow() > 1) {
      const idCol = 2; // Colonne B, ID source
      commandesExistantes = sheetInitiale.getRange(2, idCol, sheetInitiale.getLastRow() - 1, 1).getValues().flat();
    }
    
    // Créer un index des téléphones pour détecter les doublons
    const telephonesIndex = {};
    
    // Premier passage : détecter les commandes déjà existantes dans le système
    commandes.forEach(function(commande) {
      // Vérifier si la commande existe déjà dans le système
      if (commandesExistantes.includes(commande.idCommande)) {
        commande.statut = 'doublon';
        commande.doublonType = 'systeme';
        return;
      }
      
      // Formater le téléphone pour la détection de doublons
      const telephone = formaterTelephone(commande.telephone);
      
      // Enregistrer le téléphone dans l'index
      if (telephone) {
        if (!telephonesIndex[telephone]) {
          telephonesIndex[telephone] = [];
        }
        telephonesIndex[telephone].push(commande);
      }
      
      // Vérifier les données obligatoires
      if (!commande.client || !commande.telephone || !commande.produit) {
        commande.statut = 'verification';
        return;
      }
      
      commande.statut = 'ok';
    });
    
    // Deuxième passage : détecter les doublons de téléphone dans le lot actuel
    for (const telephone in telephonesIndex) {
      if (telephonesIndex[telephone].length > 1) {
        // Marquer tous les doublons sauf le premier
        for (let i = 1; i < telephonesIndex[telephone].length; i++) {
          telephonesIndex[telephone][i].statut = 'doublon';
          telephonesIndex[telephone][i].doublonType = 'telephone';
        }
      }
    }
    
    return commandes;
  }
  
  /**
   * Traite toutes les données importées
   * 
   * @return {Object} Résultat du traitement
   */
  function processImportedData() {
    const commandes = getImportedCommandes();
    
    // Compter les différents types de commandes
    let totalCount = commandes.length;
    let validCount = 0;
    let doublonCount = 0;
    let verificationCount = 0;
    
    commandes.forEach(function(commande) {
      if (commande.statut === 'doublon') {
        doublonCount++;
      } else if (commande.statut === 'verification') {
        verificationCount++;
      } else {
        validCount++;
      }
    });
    
    return {
      success: true,
      message: `Traitement terminé. ${totalCount} commandes traitées: ${validCount} valides, ${doublonCount} doublons, ${verificationCount} à vérifier.`,
      total: totalCount,
      valid: validCount,
      doublon: doublonCount,
      verification: verificationCount
    };
  }
  
  /**
   * Valide une commande et l'envoie au système
   * 
   * @param {string} idCommande L'ID de la commande à valider
   * @return {Object} Résultat de la validation
   */
  function validateCommande(idCommande) {
    const ss = getSpreadsheet();
    const sheetShopify = ss.getSheetByName('Import Shopify');
    const sheetYoucan = ss.getSheetByName('Import Youcan');
    const sheetInitiale = ss.getSheetByName(CONFIG.SHEETS.INITIALE);
    
    if (!sheetInitiale) {
      return { success: false, message: 'La feuille des commandes n\'existe pas.' };
    }
    
    let commande = null;
    let sheetSource = null;
    let rowIndex = 0;
    
    // Rechercher la commande dans Shopify
    if (sheetShopify) {
      const shopifyData = sheetShopify.getDataRange().getValues();
      for (let i = 1; i < shopifyData.length; i++) {
        if (shopifyData[i][0] === idCommande) {
          commande = {
            idCommande: shopifyData[i][0],
            client: shopifyData[i][1],
            telephone: shopifyData[i][2],
            adresse: shopifyData[i][3],
            ville: shopifyData[i][4],
            produit: shopifyData[i][5],
            quantite: shopifyData[i][6],
            prix: shopifyData[i][7],
            date: shopifyData[i][8],
            source: shopifyData[i][9]
          };
          sheetSource = sheetShopify;
          rowIndex = i + 1;
          break;
        }
      }
    }
    
    // Si non trouvée, rechercher dans Youcan
    if (!commande && sheetYoucan) {
      const youcanData = sheetYoucan.getDataRange().getValues();
      for (let i = 1; i < youcanData.length; i++) {
        if (youcanData[i][0] === idCommande) {
          commande = {
            idCommande: youcanData[i][0],
            client: youcanData[i][1],
            telephone: youcanData[i][2],
            adresse: youcanData[i][3],
            ville: youcanData[i][4],
            produit: youcanData[i][5],
            quantite: youcanData[i][6],
            prix: youcanData[i][7],
            date: youcanData[i][8],
            source: youcanData[i][9]
          };
          sheetSource = sheetYoucan;
          rowIndex = i + 1;
          break;
        }
      }
    }
    
    if (!commande) {
      return { success: false, message: 'Commande non trouvée.' };
    }
    
    // Formater le téléphone
    const telephone = formaterTelephone(commande.telephone);
    
    // Générer un numéro de commande unique
    const numeroCommande = genererNumeroCommande(commande.source);
    
    // Ajouter la commande à la feuille initiale
    sheetInitiale.appendRow([
      numeroCommande,
      commande.idCommande,
      CONFIG.STATUTS.NON_AFFECTEE,
      '',
      commande.client,
      telephone,
      commande.adresse,
      commande.ville,
      commande.produit,
      commande.quantite,
      commande.prix,
      new Date(),
      formaterDate(commande.date || new Date())
    ]);
    
    // Supprimer la commande de la feuille source
    sheetSource.deleteRow(rowIndex);
    
    // Enregistrer l'action dans le journal
    enregistrerLog(
      Session.getActiveUser().getEmail(),
      numeroCommande,
      'Validation de la commande ' + commande.idCommande
    );
    
    return { 
      success: true, 
      message: 'Commande validée et ajoutée au système.' 
    };
  }
  
  /**
   * Valide toutes les commandes sans erreur
   * 
   * @return {Object} Résultat de la validation
   */
  function validateAllCommandes() {
    const commandes = getImportedCommandes();
    let validCount = 0;
    
    // Filtrer les commandes valides
    const validCommandes = commandes.filter(commande => commande.statut === 'ok');
    
    // Valider chaque commande
    validCommandes.forEach(function(commande) {
      const result = validateCommande(commande.idCommande);
      if (result.success) {
        validCount++;
      }
    });
    
    return {
      success: true,
      message: `${validCount} commandes validées et ajoutées au système.`,
      count: validCount
    };
  }
  
  /**
   * Supprime une commande importée
   * 
   * @param {string} idCommande L'ID de la commande à supprimer
   * @return {Object} Résultat de la suppression
   */
  function deleteImportedCommande(idCommande) {
    const ss = getSpreadsheet();
    const sheetShopify = ss.getSheetByName('Import Shopify');
    const sheetYoucan = ss.getSheetByName('Import Youcan');
    
    let deleted = false;
    
    // Rechercher et supprimer dans Shopify
    if (sheetShopify) {
      const shopifyData = sheetShopify.getDataRange().getValues();
      for (let i = 1; i < shopifyData.length; i++) {
        if (shopifyData[i][0] === idCommande) {
          sheetShopify.deleteRow(i + 1);
          deleted = true;
          break;
        }
      }
    }
    
    // Si non trouvée, rechercher et supprimer dans Youcan
    if (!deleted && sheetYoucan) {
      const youcanData = sheetYoucan.getDataRange().getValues();
      for (let i = 1; i < youcanData.length; i++) {
        if (youcanData[i][0] === idCommande) {
          sheetYoucan.deleteRow(i + 1);
          deleted = true;
          break;
        }
      }
    }
    
    if (!deleted) {
      return { success: false, message: 'Commande non trouvée.' };
    }
    
    return { 
      success: true, 
      message: 'Commande supprimée.' 
    };
  }

  /**
   * Fonction simple pour tester que le module est correctement chargé
   * Elle peut être appelée depuis CMDInitPanel.html pour vérifier que le module fonctionne
   * 
   * @return {string} Message de confirmation
   */
  function testCMDInitModule() {
    try {
      Logger.log("Fonction testCMDInitModule appelée avec succès");
      return "Le module CMDInit fonctionne correctement.";
    } catch (e) {
      Logger.log("Erreur dans testCMDInitModule: " + e.toString());
      return "Erreur dans le module CMDInit: " + e.toString();
    }
  }