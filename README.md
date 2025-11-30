# Google Tasks Manager for Sheets

![License](https://img.shields.io/badge/License-MIT-blue.svg)
![Platform](https://img.shields.io/badge/Platform-Google%20Workspace-green)
![Runtime](https://img.shields.io/badge/Script-V8%20Engine-f39f37)
![Status](https://img.shields.io/badge/Status-Production-success)

**Une solution complète de gestion de tâches synchronisée bidirectionnellement entre Google Sheets et Google Tasks.**

Ce projet transforme un tableur Google Sheets en un tableau de bord de pilotage. Il permet non seulement de visualiser vos tâches, mais aussi de les créer, de modifier leur statut et de recevoir des rapports matinaux, le tout sans quitter votre tableur.

---

## 🚀 Fonctionnalités Clés

* **Synchronisation Bidirectionnelle (Two-Way Sync) :**
    * 📥 **Tasks vers Sheets :** Récupération de toutes les listes et tâches (Titre, Date, Notes, Statut).
    * 📤 **Sheets vers Tasks :** Création de tâches via une simple case à cocher.
    * 🔄 **Mise à jour d'état :** Changer "À faire" en "Terminée" dans le Sheets met à jour Google Tasks instantanément.
* **Reporting Automatisé :** Envoi quotidien (08h00) d'un email récapitulatif HTML ergonomique des tâches en attente.
* **Archivage Intelligent :** Fonction de nettoyage pour retirer les tâches terminées du tableur tout en conservant l'intégrité des données.
* **Performance :** Utilisation d'opérations par lots (Batch Operations) et gestion des IDs techniques en arrière-plan.

## 🛠️ Prérequis Techniques

* Un compte Google Workspace ou Gmail.
* Un fichier Google Sheets.
* L'accès activé au service **Google Tasks API** dans le projet Apps Script.

## 📦 Installation

### 1. Configuration du Script
1.  Créez un nouveau Google Sheet.
2.  Allez dans `Extensions` > `Apps Script`.
3.  Copiez le contenu du fichier `Code.js` de ce dépôt dans l'éditeur.
4.  Créez un fichier HTML nommé `Mail_Modele.html` et copiez-y le code correspondant.

### 2. Activation des Services (Important)
Dans la barre latérale gauche de l'éditeur Apps Script :
1.  Cliquez sur le `+` à côté de **Services**.
2.  Recherchez et sélectionnez **Google Tasks API**.
3.  Cliquez sur **Ajouter** (L'identifiant doit être `Tasks`).

### 3. Configuration des Déclencheurs (Triggers)
Pour que la création et la modification fonctionnent, vous devez installer un déclencheur manuel (les déclencheurs simples `onEdit` n'ont pas les permissions API requises).

1.  Allez dans le menu **Déclencheurs** (icône réveil) à gauche.
2.  Cliquez sur **Ajouter un déclencheur**.
3.  Configurez comme suit :
    * Fonction : `gererEditionTableur`
    * Source de l'événement : `Dans la feuille de calcul`
    * Type d'événement : `Lors de la modification` (On Change)
4.  Sauvegardez et autorisez l'accès.

## 📖 Guide d'Utilisation

Une fois le script installé, rechargez votre Google Sheet. Un menu **"Gestion des tâches"** apparaît.

1.  **Initialisation :** Cliquez sur `Gestion des tâches` > `🔄 Tout synchroniser`. Cela va formater le tableau et charger vos tâches existantes.
2.  **Créer une tâche :**
    * Remplissez une ligne (Titre, Date, Notes).
    * Cochez la case en colonne **F ("Créer ?")**.
    * Attendez que la case se transforme en `✅`.
3.  **Terminer une tâche :**
    * Changez le statut en colonne **E** de "À faire" à "Terminée".
    * La tâche sera barrée dans Google Tasks.
4.  **Nettoyage :** Utilisez le menu `🧹 Archiver les tâches terminées` pour retirer les lignes finies du tableau.

## 📂 Structure du Projet

* `Code.js` : Logique serveur (GAS), gestion API, Triggers et Menu.
* `Mail_Modele.html` : Template HTML responsive pour le rapport email.

## 👤 Auteur

**Fabrice Faucheux**
* Expert Google Apps Script & Workspace
* [Atelier Informatique](https://atelier-informatique.com)

## 📄 Licence

Distribué sous la licence MIT. Voir `LICENSE` pour plus d'informations.
