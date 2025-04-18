> **Conception technique pour préparer une solution provisoire Yoozak**

Contexte :\
Afin de concevoir une solution temporaire pour aider l\'équipe
**Yoozak** dans la gestion de leurs commandes, depuis la prise de
commande jusqu\'à l\'étape finale de livraison, voici une conception
technique proposée pour le développement de l\'application.

Cas de l'utilisation :

+-----------------------------------+-----------------------------------+
| **Acteur**                        | **Fonctionnalité**                |
+===================================+===================================+
| > Trigger                         | > Insérer les lignes de fichier   |
|                                   | > Shoopify et Youcan dans le      |
|                                   | > fichier initiale.               |
+-----------------------------------+-----------------------------------+
|                                   | > Formater les données téléphone  |
|                                   | > et date.                        |
+-----------------------------------+-----------------------------------+
|                                   | > Repérer les doublons.           |
+-----------------------------------+-----------------------------------+
|                                   | > Exporter les cas d'erreur dans  |
|                                   | > le fichier CMD problème.        |
+-----------------------------------+-----------------------------------+
| > Administrateur                  | > Affecter des commandes par      |
|                                   | > operateur.                      |
+-----------------------------------+-----------------------------------+
|                                   | > Désaffecter d'une commande par  |
|                                   | > operateur.                      |
+-----------------------------------+-----------------------------------+
|                                   | > Résoudre un problème d'une      |
|                                   | > commande.                       |
+-----------------------------------+-----------------------------------+
| > Operateur                       | > Créer une commande unitaire.    |
+-----------------------------------+-----------------------------------+
|                                   | > Modifier les informations d'une |
|                                   | > commande.                       |
+-----------------------------------+-----------------------------------+
|                                   | > Confirmer ou annuler une        |
|                                   | > commande.                       |
+-----------------------------------+-----------------------------------+
| > Service\                        | > Changer le statut d'une         |
| > logistique                      | > commande                        |
|                                   | > expédié/livre/retourne.         |
+-----------------------------------+-----------------------------------+
|                                   | > Imprimer les tickets de         |
|                                   | > packages en maas ou unitaire.   |
+-----------------------------------+-----------------------------------+

Paramétrage :

+-----------------------------------+-----------------------------------+
| **Fichier**                       | **Description**                   |
+===================================+===================================+
| CMD config                        | > Fichier de configuration        |
|                                   | > générale de                     |
|                                   | >                                 |
|                                   | > l\'application :                |
|                                   | >                                 |
|                                   | > • CMD Products : Liste des      |
|                                   | > produits.                       |
|                                   | >                                 |
|                                   | > • CMD Region : Liste des villes |
|                                   | > avec les                        |
|                                   | >                                 |
|                                   | > régions.                        |
|                                   | >                                 |
|                                   | > • CMD AllowedUsers : Liste des  |
|                                   | >                                 |
|                                   | > utilisateurs autorisés.         |
|                                   | >                                 |
|                                   | > • CMD SheetUsers : Liste des    |
|                                   | > feuilles                        |
|                                   | >                                 |
|                                   | > temporaires pour chaque         |
|                                   | > opérateur.                      |
+-----------------------------------+-----------------------------------+
| CMD initiale                      | > Fichier qui contient la liste   |
|                                   | > des commandes globale avec      |
|                                   | > trois statuts possibles :       |
|                                   | >                                 |
|                                   | > • Affectée.                     |
|                                   | >                                 |
|                                   | > • Non affectée.\                |
|                                   | > • Problème.                     |
+-----------------------------------+-----------------------------------+
| CMD TMP                           | > Fichier temporaire de           |
|                                   | > traitement des                  |
|                                   |                                   |
|                                   | commandes, il contient les        |
|                                   | commandes en                      |
|                                   |                                   |
|                                   | cours de confirmation avec une    |
|                                   | action de                         |
|                                   |                                   |
|                                   | > confirmation :                  |
|                                   | >                                 |
|                                   | > • Aucune action                 |
+-----------------------------------+-----------------------------------+

> **Conception technique pour préparer une solution provisoire Yoozak**

+-----------------------------------+-----------------------------------+
|                                   | +--------------+--------------+   |
|                                   | | > •\         | > Appel 1.   |   |
|                                   | | > •\         | >            |   |
|                                   | | > •\         | > Envoi de   |   |
|                                   | | > •\         | > SMS et     |   |
|                                   | | > •\         | > MSG.       |   |
|                                   | | > •\         | >            |   |
|                                   | | > •\         | > Appel 2.   |   |
|                                   | | > •\         | >            |   |
|                                   | | > •\         | > Appel 3.   |   |
|                                   | | > •\         | >            |   |
|                                   | | > •          | > Appel 4.   |   |
|                                   | |              | >            |   |
|                                   | |              | > Appel 5.   |   |
|                                   | |              | >            |   |
|                                   | |              | > Appel 6.   |   |
|                                   | |              | >            |   |
|                                   | |              | > Appel 7.   |   |
|                                   | |              | >            |   |
|                                   | |              | > Appel 8.   |   |
|                                   | |              | >            |   |
|                                   | |              | > Proposer   |   |
|                                   | |              | > un         |   |
|                                   | |              | >            |   |
|                                   | |              |  abonnement. |   |
|                                   | |              | >            |   |
|                                   | |              | > Proposer   |   |
|                                   | |              | > une offre  |   |
|                                   | |              | > de         |   |
|                                   | |              | > réduction. |   |
|                                   | +==============+==============+   |
|                                   | +--------------+--------------+   |
+===================================+===================================+
| CMD produits                      | > Un fichier contient le détails  |
|                                   | > de la commande avec l'id de la  |
|                                   | > commande provisoire de Yoozak.  |
+-----------------------------------+-----------------------------------+
| CMD TMP LOG                       | > Fichier contient :\             |
|                                   | > • Nom de l'opérateur\           |
|                                   | > • Numéro de la commande         |
|                                   | > provisoire. • Intitule de       |
|                                   | > l'action.                       |
|                                   | >                                 |
|                                   | > Date de prise de l'action. •    |
+-----------------------------------+-----------------------------------+
| CMD confirme                      | > Fichier contient la liste des   |
|                                   | > commandes confirmes avec trois  |
|                                   | > statuts possible : • Confirmé.  |
|                                   | >                                 |
|                                   | > • En cours de préparation •     |
|                                   | > Expédié\                        |
|                                   | > • Livré                         |
+-----------------------------------+-----------------------------------+
| CMD confirme LOG                  | > Fichier pour les commandes      |
|                                   | > confirmées                      |
+-----------------------------------+-----------------------------------+
| CMD Annulée                       | > Liste des commandes             |
|                                   | > annules(avec une date           |
|                                   | > d'annulation)                   |
+-----------------------------------+-----------------------------------+
| CMD Retournée                     | > Liste des commandes retournée.  |
+-----------------------------------+-----------------------------------+
| CMD Problem                       | > Liste des commandes problèmes.  |
+-----------------------------------+-----------------------------------+

Diagramme de séquence :

> **Conception technique pour préparer une solution provisoire Yoozak**
>
> **Watcher Sheet CMD initial :**

![](vertopal_8f21727773c74ce38b219f151ef70efb/media/image1.png){width="7.458333333333333in"
height="4.004166666666666in"}

> **Conception technique pour préparer une solution provisoire Yoozak**
>
> **Administrateur :**

![](vertopal_8f21727773c74ce38b219f151ef70efb/media/image2.png){width="7.363888888888889in"
height="5.123611111111111in"}

> **Operateur :**
>
> **Conception technique pour préparer une solution provisoire Yoozak**
>
> ![](vertopal_8f21727773c74ce38b219f151ef70efb/media/image3.png){width="4.980555555555555in"
> height="10.35in"}
>
> **Conception technique pour préparer une solution provisoire Yoozak**
>
> **Service logistique :**
>
> Modification de statut de la commande :

![](vertopal_8f21727773c74ce38b219f151ef70efb/media/image4.png){width="6.323611111111111in"
height="5.633333333333334in"}

> **Conception technique pour préparer une solution provisoire Yoozak**
>
> Impression des tickets :

![](vertopal_8f21727773c74ce38b219f151ef70efb/media/image5.png){width="6.5375in"
height="3.401388888888889in"}

> **Note :**
>
> Le numéro de la commande généree au moment de l'insertion dans le
> sheet TMP de l'opérateur, il se compose par :
>
> • La source : **Y** =\> **Youcan**, **S** =\> **Shoopify**, **OO** =\>
> **Code de l'opérateur**. • Un numéro incrémenté.
>
> Il faut décrémenter la quantité du produit dans le cas du passage de
> la commande au statut retourné.
>
> Le statut de la commande annulation partiale n'est pas pris en compte
> dans cette
>
> version de développement.
>
> Ce numéro n'a aucune liaison avec le numéro mentionne leurs de la
> confirmation de la commande.
