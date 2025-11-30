# Gmail vers PDF et Drive (Automatisé)

![Runtime](https://img.shields.io/badge/Google%20Apps%20Script-V8-green)
![Author](https://img.shields.io/badge/Auteur-Fabrice%20Faucheux-orange)

## 📋 Description

Ce projet permet d'automatiser l'archivage d'e-mails Gmail critiques (factures, contrats, administratif) vers Google Drive. Le script surveille des libellés spécifiques, convertit le contenu des e-mails en fichiers PDF propres, sauvegarde les pièces jointes et archive automatiquement le courrier traité.

## ✨ Fonctionnalités Clés

* **Conversion PDF Intelligente :** Transforme le corps de l'e-mail en PDF incluant les images in-line (CID).
* **Gestion des Pièces Jointes :** Sauvegarde automatique des fichiers joints dans des sous-dossiers structurés.
* **Anti-Doublons :** Utilise `PropertiesService` pour s'assurer qu'un e-mail n'est jamais traité deux fois.
* **Nettoyage Automatique :** Change le libellé de l'e-mail (ex: `PDF` -> `PDF/Traité`) et archive la conversation.
* **Mode Simulation :** Permet de tester le script sans effectuer de modifications réelles (Dry Run).

## ⚙️ Configuration

Ouvrez le fichier `Code.js` et modifiez l'objet `CONFIG` au début du script :

| Paramètre | Type | Description |
| :--- | :--- | :--- |
| `libellesATraiter` | `Array<String>` | Liste des libellés Gmail à surveiller (ex: `['Factures', 'Devis']`). |
| `idDossierRacine` | `String` | ID du dossier Drive de destination. Mettre `null` pour la racine. |
| `sauvegarderPiecesJointes` | `Boolean` | `true` pour extraire les PJ dans un dossier séparé. |
| `sousLibelleTraite` | `String` | Nom du sous-libellé ajouté après traitement. |
| `modeSimulation` | `Boolean` | `true` pour tester le script sans écrire de fichiers. |

## 🚀 Installation Manuelle

1.  Accédez à [script.google.com](https://script.google.com/home).
2.  Créez un **Nouveau projet**.
3.  Copiez le contenu du fichier `Code.js` dans l'éditeur.
4.  Renommez le projet (ex: *Gmail2Drive-PDF*).
5.  Exécutez la fonction `executer()` une première fois manuellement pour valider les autorisations (GmailApp, DriveApp).

## ⏰ Automatisation

Pour activer l'exécution automatique toutes les 5 minutes :

1.  Sélectionnez la fonction `creerDeclencheur5Minutes` dans la barre d'outils.
2.  Cliquez sur **Exécuter**.
3.  Vérifiez dans le menu de gauche **Déclencheurs** (icône réveil) que le trigger est bien présent.

## 📝 Licence

Distribué sous licence MIT. Copyright (c) 2025 Fabrice Faucheux.
