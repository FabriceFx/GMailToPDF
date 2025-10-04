# 📬 Gmail → PDF → Google Drive (Apps Script)

[![Licence MIT avec attribution](https://img.shields.io/badge/Licence-MIT%20%2B%20Attribution-blue.svg)](./LICENSE)
[![Langue](https://img.shields.io/badge/langue-Français-lightgrey.svg)](#)
[![Google Apps Script](https://img.shields.io/badge/plateforme-Google%20Apps%20Script-yellow.svg)](#)

---

## 🧩 Description

Ce projet **automatise la sauvegarde de vos e-mails Gmail en fichiers PDF dans Google Drive** à partir d’un libellé (par exemple `PDF`).

Grâce à **Google Apps Script**, le script :

- détecte automatiquement les e-mails portant un libellé donné,  
- convertit chaque e-mail en **PDF complet** (texte, images, entête),  
- sauvegarde le fichier dans un **dossier Drive du même nom**,  
- enregistre les **pièces jointes** (optionnel),  
- marque les e-mails comme **traités** et les **archive**.

---

## ⚙️ Fonctionnalités principales

✅ Export automatique des e-mails Gmail vers Google Drive  
✅ Conversion HTML → PDF avec rendu fidèle  
✅ Sauvegarde optionnelle des pièces jointes  
✅ Aucun doublon : chaque message est traité une seule fois  
✅ Archivage automatique après export  
✅ Entièrement commenté et configuré en **français**  
✅ Licence **MIT avec attribution obligatoire**

---

## 📁 Structure du projet
📦 gmail-pdf-drive
├── README.md              → ce fichier
├── LICENSE                → licence MIT + attribution
└── src/
└── gmail_pdf_drive.js → code principal du script (Apps Script)

## 🚀 Installation rapide

1️⃣ Ouvrir Google Apps Script

- Va sur [https://script.google.com/](https://script.google.com/)
- Clique sur **Nouveau projet**

2️⃣ Copier le code

- Colle le contenu du fichier `gmail_pdf_drive.js` dans l’éditeur
- Sauvegarde le projet

3️⃣ Configurer

Dans le haut du script, adapte la section `CONFIG` :

```javascript
const CONFIG = {
  libellesATraiter: ['PDF'],           // Libellé Gmail à surveiller
  idDossierRacine: null,               // ID du dossier Drive (null = racine)
  sauvegarderPiecesJointes: true,      // Enregistrer les PJ
  sousLibelleTraite: 'Traité',         // Sous-libellé pour marquer comme traité
  archiverConversation: true,          // Archiver après traitement
  seulementNonLus: false,              // Option non lus uniquement
  joursDeRecherche: 30,                // Limite de recherche Gmail
  modeSimulation: false                // Mode test sans écriture
};

4️⃣ Exécuter manuellement une première fois
	•	Sélectionne la fonction executer()
	•	Clique sur ▶️ Exécuter
	•	Autorise le script à accéder à Gmail et Drive

5️⃣ Automatiser (optionnel)

Lance une seule fois :

creerDeclencheur5Minutes();

Cela crée un déclencheur automatique toutes les 5 minutes.

🧠 Exemple de fonctionnement
	1.	Ajoute le libellé PDF à un e-mail dans Gmail.
	2.	Le script le convertit en PDF et l’enregistre dans ton Drive :

📂 Google Drive
└── PDF/
    ├── 2025-10-04_Facture_Orange_XXXXXX.pdf
    └── PiecesJointes/
        └── 2025-10-04_Facture_Orange_XXXXXX/
            ├── facture.pdf
            └── details.csv
	3.	L’e-mail original est marqué comme PDF/Traité et archivé.
