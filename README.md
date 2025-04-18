# Système de Gestion des Commandes Yoozak

## Vue d'ensemble

Ce projet implémente une solution provisoire pour la gestion des commandes Yoozak dans Google Apps Script. Le système permet de suivre les commandes depuis leur création jusqu'à leur livraison, en passant par différentes étapes de traitement.

## Installation

1. Créez un nouveau classeur Google Sheets
2. Dans le menu Extensions, sélectionnez Apps Script
3. Supprimez le code existant dans `Code.gs`
4. Créez de nouveaux fichiers pour chaque module du système (Main.js, AdminModule.js, etc.)
5. Copiez le code correspondant dans chaque fichier
6. Enregistrez et fermez l'éditeur de script
7. Rechargez le classeur Google Sheets
8. Utilisez le menu "Yoozak" qui apparaît pour initialiser le système

## Structure du système

Le système est composé des fichiers suivants :

- **Main.js** : Configuration globale et fonctions principales
- **AdminModule.js** : Fonctions pour les administrateurs
- **OperateurModule.js** : Fonctions pour les opérateurs
- **LogistiqueModule.js** : Fonctions pour le service logistique
- **ImportData.js** : Gestion de l'importation des données depuis Shopify et Youcan
- **FormUtils.js** : Fonctions utilitaires pour les formulaires
- **InitSetup.js** : Initialisation des feuilles de calcul
- **CreerCommandeForm.html** : Formulaire pour créer une commande
- **ModifierCommandeForm.html** : Formulaire pour modifier une commande

## Feuilles de calcul

Le système gère plusieurs feuilles de calcul :

1. **CMD config** : Configuration générale (produits, régions, utilisateurs)
2. **CMD initiale** : Liste des commandes (affectées, non affectées, problèmes)
3. **CMD TMP** : Traitement temporaire des commandes
4. **CMD produits** : Détails des commandes
5. **CMD TMP LOG** : Historique des actions
6. **CMD confirme** : Commandes confirmées
7. **CMD confirme LOG** : Historique des confirmations
8. **CMD Annulée** : Commandes annulées
9. **CMD Retournée** : Commandes retournées
10. **CMD Problem** : Commandes problématiques
11. **Import Shopify** : Données importées de Shopify
12. **Import Youcan** : Données importées de Youcan

## Guide d'utilisation

### Initialisation

1. Ouvrez le classeur Google Sheets
2. Cliquez sur le menu "Yoozak"
3. Sélectionnez "Initialiser le système"
4. Confirmez l'initialisation

### Administrateur

#### Affecter des commandes

1. Cliquez sur "Yoozak" > "Administrateur" > "Affecter commandes"
2. Sélectionnez un opérateur dans la liste
3. Indiquez le nombre de commandes à affecter
4. Confirmez l'affectation

#### Désaffecter une commande

1. Cliquez sur "Yoozak" > "Administrateur" > "Désaffecter commandes"
2. Entrez le numéro de la commande à désaffecter
3. Confirmez la désaffectation

#### Résoudre un problème

1. Cliquez sur "Yoozak" > "Administrateur" > "Résoudre problèmes"
2. Sélectionnez une commande problématique dans la liste
3. Suivez les instructions pour résoudre le problème

### Opérateur

#### Créer une commande

1. Cliquez sur "Yoozak" > "Opérateur" > "Créer commande"
2. Remplissez le formulaire avec les informations de la commande
3. Cliquez sur "Créer la commande"

#### Modifier une commande

1. Cliquez sur "Yoozak" > "Opérateur" > "Modifier commande"
2. Entrez le numéro de la commande à modifier
3. Modifiez les informations nécessaires
4. Cliquez sur "Enregistrer les modifications"

#### Confirmer une commande

1. Cliquez sur "Yoozak" > "Opérateur" > "Confirmer commande"
2. Entrez le numéro de la commande à confirmer
3. Confirmez l'action

#### Annuler une commande

1. Cliquez sur "Yoozak" > "Opérateur" > "Annuler commande"
2. Entrez le numéro de la commande à annuler
3. Indiquez le motif d'annulation
4. Confirmez l'annulation

### Service Logistique

#### Changer le statut d'une commande

1. Cliquez sur "Yoozak" > "Logistique" > "Changer statut commande"
2. Entrez le numéro de la commande
3. Sélectionnez le nouveau statut
4. Confirmez le changement

#### Imprimer des tickets

1. Cliquez sur "Yoozak" > "Logistique" > "Imprimer tickets"
2. Choisissez entre impression unitaire ou en masse
3. Suivez les instructions pour générer les tickets

## Flux de travail

1. Les commandes sont importées depuis Shopify ou Youcan, ou créées manuellement
2. Les administrateurs affectent les commandes aux opérateurs
3. Les opérateurs traitent les commandes et les confirment
4. Le service logistique suit les commandes (en préparation, expédiées, livrées)
5. Les commandes peuvent être annulées ou retournées à différentes étapes

## Notes importantes

- Les numéros de commande sont générés automatiquement
- Le statut d'une commande suit un flux précis : Confirmé > En préparation > Expédié > Livré
- Une commande peut être retournée à partir du statut "Expédié" ou "Livré"
- Le système enregistre automatiquement toutes les actions dans les journaux