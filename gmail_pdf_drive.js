/********************************************************************************************
 * NOM DU SCRIPT : Gmail vers PDF et Drive par Libellé
 * AUTEUR : Fabrice Faucheux
 * VERSION : 1.0
 * DATE : 4/10/2025
 *
 * DESCRIPTION GÉNÉRALE :
 * ------------------------------------------------------------------------------------------
 * Ce script Google Apps Script automatise l’exportation d’e-mails Gmail portant un libellé
 * spécifique (ex. "PDF") vers Google Drive sous forme de fichiers PDF.
 *
 * Pour chaque message trouvé :
 *   - Le contenu de l’e-mail est converti en PDF (avec entête, corps, images, etc.)
 *   - Le PDF est enregistré dans un dossier Drive portant le nom du libellé
 *   - Les pièces jointes sont également sauvegardées (optionnel)
 *   - L’e-mail est marqué comme traité (ajout d’un sous-libellé “Traité”) et archivé
 *   - Le script évite les doublons grâce à un stockage interne des identifiants déjà traités
 *
 * L’objectif est de faciliter la conservation et la classification automatique
 * des e-mails importants (factures, contrats, etc.) sous forme de documents PDF.
 *
 * ------------------------------------------------------------------------------------------
 * PRINCIPALES FONCTIONNALITÉS :
 *   ✅ Conversion automatique des e-mails en PDF
 *   ✅ Sauvegarde des pièces jointes dans des sous-dossiers dédiés
 *   ✅ Gestion de plusieurs libellés simultanément
 *   ✅ Archivage et marquage automatique des e-mails traités
 *   ✅ Aucune duplication : chaque message n’est traité qu’une seule fois
 *   ✅ Mode “simulation” pour tester sans rien modifier
 *
 * ------------------------------------------------------------------------------------------
 * DÉPENDANCES :
 *   - Nécessite l’accès aux services :
 *       • GmailApp      → lecture et modification des e-mails
 *       • DriveApp      → création et écriture de fichiers PDF
 *       • PropertiesService → stockage des identifiants traités
 *   - Aucune bibliothèque externe requise
 *
 * ------------------------------------------------------------------------------------------
 * CONFIGURATION PRINCIPALE (objet CONFIG) :
 *   • libellesATraiter : liste des libellés Gmail à surveiller
 *   • idDossierRacine : ID du dossier Drive parent (null = racine)
 *   • sauvegarderPiecesJointes : true/false pour inclure les PJ
 *   • sousLibelleTraite : sous-libellé ajouté après traitement (ex: "Traité")
 *   • archiverConversation : archive automatiquement la conversation
 *   • seulementNonLus : si true, ne traite que les e-mails non lus
 *   • joursDeRecherche : limite de recherche dans Gmail (performances)
 *   • modeSimulation : si true, exécute sans écrire ni archiver
 *
 * ------------------------------------------------------------------------------------------
 * AUTOMATISATION :
 *   Une fonction “creerDeclencheur5Minutes()” permet d’exécuter
 *   automatiquement le script toutes les 5 minutes via un déclencheur Apps Script.
 *
 * ------------------------------------------------------------------------------------------
 * CONSEILS :
 *   - Tester d’abord en modeSimulation: true
 *   - Vérifier les autorisations Gmail et Drive lors de la première exécution
 *   - Adapter les noms de libellés et les paramètres selon votre usage
 *
 * ------------------------------------------------------------------------------------------
/********************************************************************************************
 * LICENCE : MIT avec attribution obligatoire
 *
 * Copyright (c) 2025 Fabrice Faucheux - L'atelier informatique
 *
 * Permission est accordée, gratuitement, à toute personne obtenant une copie de ce script
 * et des fichiers de documentation associés (le "Logiciel"), de l'utiliser, le copier,
 * le modifier, le fusionner, le publier, le distribuer, le sous-licencier et/ou de vendre
 * des copies du Logiciel, sous réserve des conditions suivantes :
 *
 * ➤ Le présent avis de copyright et la mention suivante doivent apparaître clairement
 *   dans toute copie ou utilisation du Logiciel :
 *     "Basé sur le script 'Gmail vers PDF et Drive par Libellé' développé par Fabrice Faucheux"
 *
 * LE LOGICIEL EST FOURNI "EN L'ÉTAT", SANS GARANTIE D'AUCUNE SORTE,
 * EXPRESSE OU IMPLICITE, Y COMPRIS MAIS SANS S'Y LIMITER LES GARANTIES
 * DE QUALITÉ MARCHANDE, D'ADÉQUATION À UN USAGE PARTICULIER ET D'ABSENCE DE CONTREFAÇON.
 * EN AUCUN CAS LES AUTEURS OU TITULAIRES DU COPYRIGHT NE POURRONT ÊTRE TENUS
 * POUR RESPONSABLES DE TOUT DOMMAGE OU AUTRE RÉCLAMATION DÉCOULANT DE L'UTILISATION
 * OU DE LA DISTRIBUTION DU LOGICIEL.
 ********************************************************************************************/


const CONFIG = {
  libellesATraiter: ['PDF'],           // Libellés Gmail à surveiller
  idDossierRacine: null,               // ID du dossier Drive parent (null = Racine de Drive)
  sauvegarderPiecesJointes: true,      // Enregistrer aussi les pièces jointes
  sousLibelleTraite: 'Traité',         // Sous-libellé pour marquer comme traité (ex : PDF/Traité)
  archiverConversation: true,          // Archiver les conversations après export
  seulementNonLus: false,              // Ne traiter que les messages non lus
  joursDeRecherche: 30,                // Limiter la recherche (X derniers jours)
  modeSimulation: false                // true = test sans création ni archivage
};

/**
 * Point d’entrée principal : traite tous les libellés configurés
 */
function executer() {
  for (const nomLibelle of CONFIG.libellesATraiter) {
    traiterLibelle(nomLibelle);
  }
}

/**
 * Crée un déclencheur automatique toutes les 5 minutes
 */
function creerDeclencheur5Minutes() {
  ScriptApp.newTrigger('executer').timeBased().everyMinutes(5).create();
  console.log('Déclencheur créé : toutes les 5 minutes.');
}

/* ================================================================
   Traitement principal
   ================================================================ */

function traiterLibelle(nomLibelle) {
  const libelle = GmailApp.getUserLabelByName(nomLibelle) || GmailApp.createLabel(nomLibelle);
  const libelleTraite = (CONFIG.sousLibelleTraite && CONFIG.sousLibelleTraite.trim())
    ? GmailApp.createLabel(`${nomLibelle}/${CONFIG.sousLibelleTraite}`)
    : null;

  const stockage = PropertiesService.getUserProperties();
  const cleStockage = `traite:${nomLibelle}`;

  const depuis = formaterDate(ajouterJours(new Date(), -CONFIG.joursDeRecherche));
  let requete = `label:${mettreEntreGuillemets(nomLibelle)} after:${depuis}`;
  if (CONFIG.seulementNonLus) requete += ' is:unread';

  console.log(`Recherche Gmail : ${requete}`);

  const fils = GmailApp.search(requete) || [];
  console.log(`Conversations trouvées : ${fils.length}`);

  const dossierRacine = CONFIG.idDossierRacine ? DriveApp.getFolderById(CONFIG.idDossierRacine) : DriveApp.getRootFolder();
  const dossierLibelle = obtenirOuCreerSousDossier_(dossierRacine, nomLibelle);

  for (const fil of fils) {
    const messages = fil.getMessages();
    for (const message of messages) {
      const idMessage = message.getId();
      const cleMessage = `${cleStockage}:${idMessage}`;

      if (stockage.getProperty(cleMessage)) continue; // déjà traité
      if (CONFIG.seulementNonLus && !message.isUnread()) continue;

      try {
        const pdf = convertirEmailEnPdf_(message);
        const sujetSain = nettoyerNomFichier_(message.getSubject() || 'Sans sujet');
        const dateTexte = Utilities.formatDate(message.getDate(), Session.getScriptTimeZone(), 'yyyy-MM-dd_HH-mm');
        const nomBase = `${dateTexte} - ${sujetSain} - ${idMessage.substring(0, 8)}`;

        if (!CONFIG.modeSimulation) {
          // Sauvegarder le PDF principal
          const fichierPdf = dossierLibelle.createFile(pdf.setName(`${nomBase}.pdf`));
          fichierPdf.setDescription(`Exporté depuis Gmail (libellé : ${nomLibelle})`);

          // Sauvegarde des pièces jointes (optionnel)
          if (CONFIG.sauvegarderPiecesJointes) {
            enregistrerPiecesJointes_(message, dossierLibelle, nomBase);
          }

          // Marquer et archiver
          if (libelleTraite) {
            fil.addLabel(libelleTraite);
            fil.removeLabel(libelle);
          }
          if (CONFIG.archiverConversation) {
            fil.moveToArchive();
          }

          stockage.setProperty(cleMessage, '1'); // marquer comme traité
        }

        console.log(`OK : ${nomBase}`);
      } catch (erreur) {
        console.error(`Erreur sur ${idMessage} : ${erreur && erreur.stack ? erreur.stack : erreur}`);
      }
    }
  }
}

/* ================================================================
   Conversion email → PDF
   ================================================================ */

function convertirEmailEnPdf_(message) {
  let corpsHtml = message.getBody() || '';
  const sujet = message.getSubject() || 'Sans sujet';
  const expediteur = message.getFrom();
  const destinataire = message.getTo();
  const copie = message.getCc();
  const date = Utilities.formatDate(message.getDate(), Session.getScriptTimeZone(), 'yyyy-MM-dd HH:mm');

  const styleHtml = `
    <style>
      body { font-family: Arial, sans-serif; font-size: 12pt; color: #111; }
      .entete { font-size: 10pt; color: #555; margin-bottom: 12px; }
      .entete div { margin: 2px 0; }
      hr { border: 0; border-top: 1px solid #ddd; margin: 12px 0; }
      img { max-width: 100%; height: auto; }
      table { border-collapse: collapse; }
      table, td, th { border: 1px solid #ccc; }
      td, th { padding: 4px 6px; }
    </style>
  `;

  // Convertir les images cid:xxx en data URI
  corpsHtml = insererImagesCid_(corpsHtml, message);

  const contenuComplet = `
    <html>
    <head><meta charset="UTF-8">${styleHtml}</head>
    <body>
      <div class="entete">
        <div><strong>Objet :</strong> ${echapperHtml_(sujet)}</div>
        <div><strong>De :</strong> ${echapperHtml_(expediteur)}</div>
        <div><strong>À :</strong> ${echapperHtml_(destinataire || '')}</div>
        ${copie ? `<div><strong>Cc :</strong> ${echapperHtml_(copie)}</div>` : ''}
        <div><strong>Date :</strong> ${echapperHtml_(date)}</div>
      </div>
      <hr/>
      ${corpsHtml}
    </body>
    </html>
  `;

  const blobHtml = Utilities.newBlob(contenuComplet, 'text/html', sujet);
  const pdf = blobHtml.getAs(MimeType.PDF);
  return pdf;
}

/**
 * Remplace les images cid:xxx par des data URI base64 dans le HTML
 */
function insererImagesCid_(html, message) {
  const pieces = message.getAttachments({includeInlineImages: true, includeAttachments: true}) || [];
  const indexCid = {};

  pieces.forEach(piece => {
    const enTetes = piece.getAllHeaders ? piece.getAllHeaders() : {};
    const cid = (enTetes['Content-Id'] || enTetes['Content-ID'] || '').toString().replace(/[<>]/g, '').trim();
    if (cid) indexCid[cid] = piece;
  });

  return html.replace(/src=["']cid:([^"']+)["']/gi, (m, cid) => {
    const att = indexCid[cid];
    if (!att) return m;
    const type = att.getContentType() || 'image/png';
    const base64 = Utilities.base64Encode(att.getBytes());
    return `src="data:${type};base64,${base64}"`;
  });
}

/**
 * Sauvegarde les pièces jointes dans un sous-dossier Attachments/<nomBase>/
 */
function enregistrerPiecesJointes_(message, dossierLibelle, nomBase) {
  const pieces = message.getAttachments({includeInlineImages: false, includeAttachments: true}) || [];
  if (!pieces.length) return;

  const dossierAttachements = obtenirOuCreerSousDossier_(dossierLibelle, 'PiecesJointes');
  const dossierCourant = obtenirOuCreerSousDossier_(dossierAttachements, nomBase);

  pieces.forEach(piece => {
    const nom = nettoyerNomFichier_(piece.getName() || 'piece');
    const blob = piece.copyBlob();
    if (blob.getBytes().length === 0) return;
    if (!CONFIG.modeSimulation) {
      dossierCourant.createFile(blob.setName(nom));
    }
  });
}

/* ================================================================
   Fonctions utilitaires
   ================================================================ */

function obtenirOuCreerSousDossier_(parent, nom) {
  const iterateur = parent.getFoldersByName(nom);
  return iterateur.hasNext() ? iterateur.next() : parent.createFolder(nom);
}

function nettoyerNomFichier_(nom) {
  return nom.replace(/[\\/:*?"<>|]+/g, ' ')
            .replace(/\s+/g, ' ')
            .trim()
            .substring(0, 180);
}

function mettreEntreGuillemets(texte) {
  return `"${texte.replace(/"/g, '\\"')}"`;
}

function echapperHtml_(texte) {
  return String(texte)
    .replace(/&/g,'&amp;')
    .replace(/</g,'&lt;')
    .replace(/>/g,'&gt;');
}

function ajouterJours(date, nbJours) {
  const copie = new Date(date);
  copie.setDate(copie.getDate() + nbJours);
  return copie;
}

function formaterDate(date) {
  const an = date.getFullYear();
  const mois = String(date.getMonth() + 1).padStart(2, '0');
  const jour = String(date.getDate()).padStart(2, '0');
  return `${an}/${mois}/${jour}`; // format compatible recherche Gmail
}
