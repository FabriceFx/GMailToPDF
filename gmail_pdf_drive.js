/**
 * @fileoverview Script d'archivage automatisé Gmail vers Drive (PDF).
 * @author Fabrice Faucheux
 * @version 1.0.0
 */

// CONFIGURATION UTILISATEUR
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
 * Point d’entrée principal : traite tous les libellés configurés.
 * @public
 */
function executer() {
  console.time('ExecutionScript');
  try {
    CONFIG.libellesATraiter.forEach(nomLibelle => {
      traiterLibelle_(nomLibelle);
    });
  } catch (e) {
    console.error(`Erreur critique lors de l'exécution : ${e.stack}`);
  }
  console.timeEnd('ExecutionScript');
}

/**
 * Crée un déclencheur automatique (Time-driven trigger).
 * Exécute la fonction 'executer' toutes les 5 minutes.
 * @public
 */
function creerDeclencheur5Minutes() {
  const triggers = ScriptApp.getProjectTriggers();
  const existe = triggers.some(t => t.getHandlerFunction() === 'executer');
  
  if (!existe) {
    ScriptApp.newTrigger('executer').timeBased().everyMinutes(5).create();
    console.log('Déclencheur créé : exécution toutes les 5 minutes.');
  } else {
    console.warn('Le déclencheur existe déjà.');
  }
}

/* ================================================================
   LOGIQUE MÉTIER (PRIVÉE)
   ================================================================ */

/**
 * Traite un libellé spécifique pour l'export PDF.
 * @param {string} nomLibelle - Le nom du libellé Gmail.
 * @private
 */
function traiterLibelle_(nomLibelle) {
  const libelle = GmailApp.getUserLabelByName(nomLibelle) || GmailApp.createLabel(nomLibelle);
  
  // Création conditionnelle du sous-libellé "Traité"
  const libelleTraite = (CONFIG.sousLibelleTraite?.trim())
    ? (GmailApp.getUserLabelByName(`${nomLibelle}/${CONFIG.sousLibelleTraite}`) || GmailApp.createLabel(`${nomLibelle}/${CONFIG.sousLibelleTraite}`))
    : null;

  const stockage = PropertiesService.getUserProperties();
  const cleStockage = `traite:${nomLibelle}`;

  // Construction de la requête de recherche optimisée
  const dateDebut = formaterDate_(ajouterJours_(new Date(), -CONFIG.joursDeRecherche));
  let requete = `label:${mettreEntreGuillemets_(nomLibelle)} after:${dateDebut}`;
  if (CONFIG.seulementNonLus) requete += ' is:unread';

  console.log(`🔍 Recherche pour "${nomLibelle}" : ${requete}`);

  const fils = GmailApp.search(requete);
  if (fils.length === 0) {
    console.log('Aucune conversation correspondante trouvée.');
    return;
  }

  // Préparation du dossier de destination
  const dossierRacine = CONFIG.idDossierRacine ? DriveApp.getFolderById(CONFIG.idDossierRacine) : DriveApp.getRootFolder();
  const dossierLibelle = obtenirOuCreerSousDossier_(dossierRacine, nomLibelle);

  // Itération sur les conversations (Threads)
  fils.forEach(fil => {
    const messages = fil.getMessages();
    
    // Itération sur les messages individuels
    messages.forEach(message => {
      const idMessage = message.getId();
      const cleMessage = `${cleStockage}:${idMessage}`;

      // Vérification anti-doublon et critère "Non lu"
      if (stockage.getProperty(cleMessage)) return; 
      if (CONFIG.seulementNonLus && !message.isUnread()) return;

      try {
        // 1. Conversion
        const pdf = convertirEmailEnPdf_(message);
        const sujetSain = nettoyerNomFichier_(message.getSubject() || 'Sans sujet');
        const dateTexte = Utilities.formatDate(message.getDate(), Session.getScriptTimeZone(), 'yyyy-MM-dd_HH-mm');
        const nomBase = `${dateTexte} - ${sujetSain} - ${idMessage.substring(0, 8)}`;

        if (!CONFIG.modeSimulation) {
          // 2. Sauvegarde PDF
          const fichierPdf = dossierLibelle.createFile(pdf.setName(`${nomBase}.pdf`));
          fichierPdf.setDescription(`Exporté depuis Gmail (libellé : ${nomLibelle}) | ID: ${idMessage}`);

          // 3. Sauvegarde PJ (Optionnel)
          if (CONFIG.sauvegarderPiecesJointes) {
            enregistrerPiecesJointes_(message, dossierLibelle, nomBase);
          }

          // 4. Marquage et Archivage
          if (libelleTraite) {
            fil.addLabel(libelleTraite);
            fil.removeLabel(libelle);
          }
          if (CONFIG.archiverConversation) {
            fil.moveToArchive();
          }

          // 5. Enregistrement de l'état
          stockage.setProperty(cleMessage, '1');
          console.log(`✅ Succès : ${nomBase}`);
        } else {
          console.log(`[SIMULATION] Traitement de : ${nomBase}`);
        }

      } catch (erreur) {
        console.error(`❌ Erreur sur message ${idMessage} : ${erreur.stack}`);
      }
    });
  });
}

/**
 * Convertit un objet GmailMessage en Blob PDF.
 * @param {GoogleAppsScript.Gmail.GmailMessage} message 
 * @return {GoogleAppsScript.Base.Blob}
 * @private
 */
function convertirEmailEnPdf_(message) {
  let corpsHtml = message.getBody() || '';
  const sujet = message.getSubject() || 'Sans sujet';
  const expediteur = message.getFrom();
  const destinataire = message.getTo();
  const copie = message.getCc();
  const date = Utilities.formatDate(message.getDate(), Session.getScriptTimeZone(), 'yyyy-MM-dd HH:mm');

  // Styles CSS inline pour le rendu PDF
  const styleHtml = `
    <style>
      body { font-family: 'Helvetica', sans-serif; font-size: 11pt; color: #333; line-height: 1.4; }
      .entete { background-color: #f8f9fa; padding: 15px; border-bottom: 2px solid #e9ecef; margin-bottom: 20px; }
      .entete div { margin-bottom: 5px; font-size: 0.95em; }
      strong { color: #495057; }
      img { max-width: 100%; height: auto; }
      hr { border: 0; border-top: 1px solid #ddd; margin: 20px 0; }
    </style>
  `;

  // Gestion des images embarquées (CID)
  corpsHtml = insererImagesCid_(corpsHtml, message);

  const contenuComplet = `
    <html>
    <head><meta charset="UTF-8">${styleHtml}</head>
    <body>
      <div class="entete">
        <div><strong>Objet :</strong> ${echapperHtml_(sujet)}</div>
        <div><strong>De :</strong> ${echapperHtml_(expediteur)}</div>
        <div><strong>À :</strong> ${echapperHtml_(destinataire)}</div>
        ${copie ? `<div><strong>Cc :</strong> ${echapperHtml_(copie)}</div>` : ''}
        <div><strong>Date :</strong> ${date}</div>
      </div>
      ${corpsHtml}
    </body>
    </html>
  `;

  return Utilities.newBlob(contenuComplet, 'text/html', sujet).getAs(MimeType.PDF);
}

/**
 * Remplace les références 'cid:' par des données Base64 pour l'inclusion dans le PDF.
 * @param {string} html 
 * @param {GoogleAppsScript.Gmail.GmailMessage} message 
 * @return {string} HTML modifié
 * @private
 */
function insererImagesCid_(html, message) {
  const pieces = message.getAttachments({includeInlineImages: true, includeAttachments: false});
  const indexCid = {};

  pieces.forEach(piece => {
    const headers = piece.getAllHeaders();
    // Nettoyage de l'ID content (retrait des < >)
    const cid = (headers['Content-Id'] || headers['Content-ID'] || '').toString().replace(/[<>]/g, '').trim();
    if (cid) indexCid[cid] = piece;
  });

  return html.replace(/src=["']cid:([^"']+)["']/gi, (match, cid) => {
    const att = indexCid[cid];
    if (!att) return match;
    const type = att.getContentType() || 'image/png';
    const base64 = Utilities.base64Encode(att.getBytes());
    return `src="data:${type};base64,${base64}"`;
  });
}

/**
 * Sauvegarde les pièces jointes classiques dans un sous-dossier dédié.
 * @private
 */
function enregistrerPiecesJointes_(message, dossierLibelle, nomBase) {
  const pieces = message.getAttachments({includeInlineImages: false, includeAttachments: true});
  if (pieces.length === 0) return;

  const dossierAttachements = obtenirOuCreerSousDossier_(dossierLibelle, 'PiecesJointes');
  const dossierCourant = obtenirOuCreerSousDossier_(dossierAttachements, nomBase);

  pieces.forEach(piece => {
    const blob = piece.copyBlob();
    // Sécurité pour éviter les fichiers vides corrompus
    if (blob.getBytes().length > 0) {
      const nomNettoye = nettoyerNomFichier_(piece.getName());
      dossierCourant.createFile(blob.setName(nomNettoye));
    }
  });
}

/* ================================================================
   UTILITAIRES (HELPERS)
   ================================================================ */

const obtenirOuCreerSousDossier_ = (parent, nom) => {
  const iterateur = parent.getFoldersByName(nom);
  return iterateur.hasNext() ? iterateur.next() : parent.createFolder(nom);
};

const nettoyerNomFichier_ = (nom) => {
  return nom.replace(/[\\/:*?"<>|]+/g, '_') // Remplacement caractères interdits
            .replace(/\s+/g, ' ')
            .trim()
            .substring(0, 150); // Limite Drive
};

const mettreEntreGuillemets_ = (texte) => `"${texte.replace(/"/g, '\\"')}"`;

const echapperHtml_ = (str) => {
  return String(str || '')
    .replace(/&/g, '&amp;')
    .replace(/</g, '&lt;')
    .replace(/>/g, '&gt;')
    .replace(/"/g, '&quot;');
};

const ajouterJours_ = (date, jours) => {
  const d = new Date(date);
  d.setDate(d.getDate() + jours);
  return d;
};

const formaterDate_ = (date) => {
  return Utilities.formatDate(date, Session.getScriptTimeZone(), 'yyyy/MM/dd');
};
