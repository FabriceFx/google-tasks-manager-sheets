/**
 * ================================================================
 * SYSTEME COMPLET DE GESTION TASKS ↔ SHEETS (BIDIRECTIONNEL + EMAIL)
 * ================================================================
 *
 * @description
 * 1. GET : Récupère les tâches Tasks -> Sheets.
 * 2. POST : Crée une tâche Sheets -> Tasks (via Checkbox).
 * 3. PATCH : Met à jour le statut (À faire/Terminée) Sheets -> Tasks.
 * 4. REPORT : Envoie un email HTML quotidien.
 * 5. CLEAN : Archive les lignes terminées.
 *
 * @author Fabrice Faucheux <https://atelier-informatique.com>
 * @version 3.1.0 - Production Release
 */

// --- CONFIGURATION ---
const TIMEZONE = Session.getScriptTimeZone() || 'Europe/Paris';
const NOM_FEUILLE = 'Tâches';
const NOM_TEMPLATE_HTML = 'Mail_Modele';

// --- MAPPING DES COLONNES (Index base 1) ---
const COLS = {
  LISTE: 1,
  TITRE: 2,
  DATE: 3,
  NOTES: 4,
  STATUT: 5,
  ACTION_CREER: 6,
  ID_TACHE: 7, // Masquée
  ID_LISTE: 8  // Masquée
};

/**
 * Initialisation du menu.
 */
const onOpen = (e) => {
  SpreadsheetApp.getUi()
    .createMenu('Gestion des tâches')
    .addItem('🔄 Tout synchroniser (Reset)', 'remplirTableurAvecTaches')
    .addItem('📨 Envoyer rapport quotidien', 'envoyerRapportQuotidien')
    .addSeparator()
    .addItem('🧹 Archiver les tâches terminées', 'archiverTachesTerminees')
    .addItem('⚙️ Configurer automatisation', 'creerDeclencheurQuotidien')
    .addToUi();
};

const onInstall = (e) => onOpen(e);

/**
 * ================================================================
 * 1. SYNC & LECTURE (API -> SHEETS)
 * ================================================================
 */

/**
 * Récupère toutes les tâches et refait le tableau à neuf.
 */
const remplirTableurAvecTaches = () => {
  try {
    const toutes = obtenirToutesLesTachesDeToutesLesListes();
    const classeur = SpreadsheetApp.getActiveSpreadsheet();
    const feuille = classeur.getSheetByName(NOM_FEUILLE) || classeur.insertSheet(NOM_FEUILLE);

    feuille.clear(); 
    feuille.setFrozenRows(1);

    const enTetes = ['Liste', 'Élément de travail', 'Date d’échéance', 'Notes', 'Statut', 'Créer ?', 'ID_TASK', 'ID_LIST'];
    
    const lignes = toutes.map(o => [
      o.listeNom || '',
      o.title || '',
      formaterDateTableur(o.due),
      o.notes || '',
      traduireStatut(o.status),
      false,
      o.id,
      o.listeId
    ]);

    // Style En-têtes
    feuille.getRange(1, 1, 1, enTetes.length)
      .setValues([enTetes])
      .setFontWeight('bold')
      .setBackground('#4285f4')
      .setFontColor('white');

    // Injection des données
    if (lignes.length > 0) {
      feuille.getRange(2, 1, lignes.length, enTetes.length).setValues(lignes);
    }
    
    // Validations & Formatage
    const regleStatut = SpreadsheetApp.newDataValidation().requireValueInList(['À faire', 'Terminée']).setAllowInvalid(false).build();
    feuille.getRange(2, COLS.STATUT, 999).setDataValidation(regleStatut);

    const regleCoche = SpreadsheetApp.newDataValidation().requireCheckbox().build();
    feuille.getRange(2, COLS.ACTION_CREER, 999).setDataValidation(regleCoche);

    feuille.autoResizeColumns(1, 5);
    feuille.setColumnWidth(COLS.NOTES, 300);
    feuille.setColumnWidth(COLS.ACTION_CREER, 80);
    feuille.hideColumns(COLS.ID_TACHE);
    feuille.hideColumns(COLS.ID_LISTE);

    classeur.toast(`Synchronisation terminée : ${lignes.length} tâches.`, "Succès");

  } catch (erreur) {
    console.error(`Erreur Reset : ${erreur.message}`);
    SpreadsheetApp.getActiveSpreadsheet().toast("Erreur critique lors de la synchronisation.", "Erreur");
  }
};

/**
 * ================================================================
 * 2. MODIFICATION (SHEETS -> API)
 * ================================================================
 */

/**
 * TRIGGER (ON EDIT) : Gère création et changement de statut.
 */
const gererEditionTableur = (e) => {
  // 1. Vérification de base de l'événement
  if (!e) {
    console.warn("Fonction lancée manuellement sans événement (e). Impossible de continuer.");
    return;
  }

  const feuille = e.range.getSheet();
  if (feuille.getName() !== NOM_FEUILLE) return;
  
  const range = e.range;
  const colIndex = range.getColumn();
  const rowIndex = range.getRow();
  
  // Ignorer l'en-tête
  if (rowIndex === 1) return;

  console.log(`Modification détectée - Ligne: ${rowIndex}, Col: ${colIndex}, Valeur: ${e.value}`);

  // CAS A : Changement de Statut (Colonne 5)
  if (colIndex === COLS.STATUT) {
    mettreAJourStatutTache(e, rowIndex);
  }

  // CAS B : Demande de création (Colonne 6)
  // On vérifie si c'est la bonne colonne
  if (colIndex === COLS.ACTION_CREER) {
    // Vérification robuste de la case cochée
    // e.value est parfois 'TRUE', parfois TRUE (bool), parfois undefined selon le contexte
    const valeurActuelle = range.getValue(); 
    
    if (valeurActuelle === true || e.value === 'TRUE') {
      console.log("Demande de création validée. Lancement de creerNouvelleTache...");
      creerNouvelleTache(e, rowIndex);
    } else {
      console.log("Case décochée ou valeur invalide, pas d'action.");
    }
  }
};

const mettreAJourStatutTache = (e, row) => {
  const feuille = e.source.getActiveSheet();
  const ids = feuille.getRange(row, COLS.ID_TACHE, 1, 2).getValues()[0];
  const [idTache, idListe] = ids;
  
  if (!idTache || !idListe) {
    e.source.toast("Impossible de mettre à jour : Tâche non synchronisée.", "Erreur");
    return;
  }

  const statutApi = (e.value === 'Terminée') ? 'completed' : 'needsAction';

  try {
    Tasks.Tasks.patch({ status: statutApi }, idListe, idTache);
    e.source.toast("Statut mis à jour dans Google Tasks.", "Succès");
  } catch (err) {
    console.error(err);
    e.source.toast("Erreur API.", "Erreur");
  }
};

const creerNouvelleTache = (e, row) => {
  const feuille = e.source.getActiveSheet();
  const valeurs = feuille.getRange(row, 1, 1, 4).getValues()[0];
  const [nomListe, titre, dateEcheance, notes] = valeurs;

  if (!titre) {
    e.range.setValue(false);
    e.source.toast("Le titre est obligatoire.", "Attention");
    return;
  }

  try {
    const idListeCible = resoudreIdListe(nomListe);
    
    const payload = {
      title: titre,
      notes: notes,
      status: 'needsAction'
    };
    if (dateEcheance instanceof Date) payload.due = dateEcheance.toISOString();

    const tacheCreee = Tasks.Tasks.insert(payload, idListeCible);

    e.range.setValue("✅"); 
    feuille.getRange(row, COLS.ID_TACHE).setValue(tacheCreee.id);
    feuille.getRange(row, COLS.ID_LISTE).setValue(idListeCible);
    if (!nomListe) feuille.getRange(row, COLS.LISTE).setValue(tacheCreee.title ? "Défaut" : "Tasks");

    e.source.toast("Tâche créée.", "Succès");
  } catch (err) {
    console.error(err);
    e.range.setValue(false);
    e.source.toast(`Erreur création : ${err.message}`, "Échec");
  }
};

/**
 * ================================================================
 * 3. REPORTING EMAIL (VRAIE FONCTION)
 * ================================================================
 */

/**
 * Envoie le rapport quotidien basé sur le template HTML.
 */
const envoyerRapportQuotidien = () => {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet(); // On récupère l'objet tableur
    const urlTableur = ss.getUrl(); // On récupère l'URL réelle du fichier
    
    const toutesLesTaches = obtenirToutesLesTachesDeToutesLesListes();
    
    // Filtre : uniquement 'needsAction'
    const tachesActives = toutesLesTaches
      .filter(t => t.status === 'needsAction')
      .sort((a, b) => (new Date(a.due || 0)) - (new Date(b.due || 0)));

    if (tachesActives.length === 0) {
      console.info('Aucune tâche active à signaler.');
      return;
    }

    // Regroupement par liste
    const parListe = tachesActives.reduce((acc, t) => {
      if (!acc[t.listeNom]) acc[t.listeNom] = [];
      acc[t.listeNom].push({
        titre: t.title,
        echeance: formaterDate(t.due),
        notes: t.notes || 'Pas de notes'
      });
      return acc;
    }, {});

    // Génération HTML
    const templateHtml = HtmlService.createTemplateFromFile(NOM_TEMPLATE_HTML);
    templateHtml.donnees = parListe;
    templateHtml.urlTableur = urlTableur; // <--- ON PASSE L'URL ICI

    const emailUtilisateur = Session.getEffectiveUser().getEmail();
    
    MailApp.sendEmail({
      to: emailUtilisateur,
      subject: `[Tâches] Résumé quotidien - ${Object.keys(parListe).length} liste(s) active(s)`,
      htmlBody: templateHtml.evaluate().getContent()
    });
    
    console.info(`Rapport envoyé à ${emailUtilisateur}`);
    ss.toast("Rapport envoyé par email.", "Succès");

  } catch (err) {
    console.error(`Erreur email : ${err.message}`);
    SpreadsheetApp.getActiveSpreadsheet().toast("Erreur lors de l'envoi du rapport.", "Erreur");
  }
};

/**
 * Fonction exécutée par le Trigger horaire.
 */
const executerQuotidien = () => {
  console.time('ExecutionQuotidienne');
  try {
    envoyerRapportQuotidien(); // Envoie l'email
    remplirTableurAvecTaches(); // Met à jour le tableur
  } catch (error) {
    console.error(`Erreur critique : ${error.message}`);
  }
  console.timeEnd('ExecutionQuotidienne');
};

/**
 * ================================================================
 * 4. UTILITAIRES & HELPERS
 * ================================================================
 */

const archiverTachesTerminees = () => {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const feuille = ss.getSheetByName(NOM_FEUILLE);
  if(!feuille) return;

  const lastRow = feuille.getLastRow();
  if (lastRow < 2) return;

  const range = feuille.getRange(2, 1, lastRow - 1, feuille.getLastColumn());
  const values = range.getValues();
  // Filtre local sur la colonne STATUT (index COLS.STATUT - 1 car tableau base 0)
  const rowsToKeep = values.filter(r => r[COLS.STATUT - 1] !== 'Terminée');
  
  if (values.length > rowsToKeep.length) {
    range.clearContent();
    if (rowsToKeep.length > 0) {
      feuille.getRange(2, 1, rowsToKeep.length, rowsToKeep[0].length).setValues(rowsToKeep);
    }
    ss.toast(`${values.length - rowsToKeep.length} archivées.`, "Nettoyage");
  } else {
    ss.toast("Rien à archiver.", "Info");
  }
};

const obtenirToutesLesListesDeTaches = () => {
  const items = [];
  let pageToken;
  do {
    const res = Tasks.Tasklists.list({ maxResults: 100, pageToken });
    if (res.items) items.push(...res.items);
    pageToken = res.nextPageToken;
  } while (pageToken);
  return items;
};

const obtenirTachesDepuisListe = (idListe) => {
  const allTasks = [];
  let pageToken;
  try {
    do {
      const res = Tasks.Tasks.list(idListe, {
        maxResults: 100,
        showCompleted: false,
        showDeleted: false,
        showHidden: true,
        pageToken
      });
      if (res.items) allTasks.push(...res.items);
      pageToken = res.nextPageToken;
    } while (pageToken);
  } catch (e) { console.warn(e); }
  return allTasks;
};

const obtenirToutesLesTachesDeToutesLesListes = () => {
  const listes = obtenirToutesLesListesDeTaches();
  return listes.flatMap(liste => {
    const taches = obtenirTachesDepuisListe(liste.id);
    return taches.map(t => ({ ...t, listeNom: liste.title, listeId: liste.id }));
  });
};

const resoudreIdListe = (nomRecherche) => {
  const listes = obtenirToutesLesListesDeTaches();
  if (!nomRecherche) return listes[0].id;
  const trouvee = listes.find(l => l.title.toLowerCase() === nomRecherche.toLowerCase());
  return trouvee ? trouvee.id : listes[0].id;
};

const formaterDateTableur = (iso) => iso ? new Date(iso) : '';
const traduireStatut = (s) => (s === 'completed' ? 'Terminée' : 'À faire');
const formaterDate = (iso) => iso ? new Intl.DateTimeFormat('fr-FR', { dateStyle: 'long' }).format(new Date(iso)) : 'Non spécifiée';

const creerDeclencheurQuotidien = () => {
  const fn = 'executerQuotidien';
  const triggers = ScriptApp.getProjectTriggers();
  triggers.forEach(t => { if(t.getHandlerFunction() === fn) ScriptApp.deleteTrigger(t); });
  ScriptApp.newTrigger(fn).timeBased().everyDays(1).atHour(8).create();
  SpreadsheetApp.getActiveSpreadsheet().toast("Rapport quotidien activé pour 8h00.", "Info");
};
