/**
 * CYBERPUNK 3030 // API d'inscription
 * -----------------------------------
 * Backend Google Apps Script lié au Google Sheet des inscriptions.
 *
 * Mise en place (une seule fois) :
 *   1. Dans le Google Sheet : Extensions > Apps Script, coller ce fichier à la place du code existant.
 *   2. Choisir la fonction "setup" dans la barre d'outils puis cliquer "Exécuter" et accepter les autorisations.
 *   3. Déployer > Nouveau déploiement > type "Application Web"
 *        - Exécuter en tant que : Moi
 *        - Qui a accès : Tout le monde
 *   4. Copier l'URL qui finit par /exec et la coller dans index.html (CONFIG.apiUrl).
 *
 * Après une modification de ce script : Déployer > Gérer les déploiements > crayon > Version : Nouvelle version.
 * (l'URL reste la même)
 */

const SHEET_INSCRIPTIONS = 'Inscriptions';
const SHEET_RECAP = 'Récap';
const SHEET_JOURNAL = 'Journal';

const HEADERS = ['groupe_id', 'personne_id', 'prenom', 'nom', 'laser', 'dejeuner', 'soiree', 'cree_le', 'maj_le'];
const ACTIVITIES = ['laser', 'dejeuner', 'soiree'];

const MAX_MEMBERS_PER_GROUP = 10;
const MAX_NAME_LENGTH = 40;

/* ===================== POINTS D'ENTRÉE HTTP ===================== */

function doGet(e) {
  try {
    return json_({ ok: true, groups: readGroups_(), serverTime: new Date().toISOString() });
  } catch (err) {
    return json_({ ok: false, error: 'Erreur serveur : ' + err.message });
  }
}

function doPost(e) {
  let payload;
  try {
    payload = JSON.parse((e && e.postData && e.postData.contents) || '{}');
  } catch (err) {
    return json_({ ok: false, error: 'Requête illisible.' });
  }

  // Champ piège anti-robot : un humain ne le remplit jamais.
  if (payload.website) {
    return json_({ ok: true, groups: readGroups_() });
  }

  const lock = LockService.getScriptLock();
  try {
    lock.waitLock(15000);
  } catch (err) {
    return json_({ ok: false, error: 'Serveur occupé, réessaie dans quelques secondes.' });
  }

  try {
    let result;
    if (payload.action === 'save') {
      result = saveGroup_(payload.group);
    } else if (payload.action === 'delete') {
      result = deleteGroup_(payload.groupId);
    } else {
      return json_({ ok: false, error: 'Action inconnue.' });
    }
    SpreadsheetApp.flush();
    return json_(Object.assign({ ok: true, groups: readGroups_() }, result));
  } catch (err) {
    return json_({ ok: false, error: err.message });
  } finally {
    lock.releaseLock();
  }
}

/* ===================== LOGIQUE MÉTIER ===================== */

function saveGroup_(group) {
  if (!group || !Array.isArray(group.members)) throw new Error('Inscription invalide.');

  const members = group.members
    .map(function (m) {
      return {
        prenom: cleanName_(m.prenom),
        nom: cleanName_(m.nom),
        laser: m.laser === true,
        dejeuner: m.dejeuner === true,
        soiree: m.soiree === true
      };
    })
    .filter(function (m) { return m.prenom || m.nom; });

  if (members.length === 0) throw new Error('Ajoute au moins une personne.');
  if (members.length > MAX_MEMBERS_PER_GROUP) throw new Error('Maximum ' + MAX_MEMBERS_PER_GROUP + ' personnes par inscription.');
  members.forEach(function (m) {
    if (!m.prenom || !m.nom) throw new Error('Prénom et nom obligatoires pour chaque personne.');
    if (!m.laser && !m.dejeuner && !m.soiree) {
      throw new Error(m.prenom + ' doit être inscrit(e) à au moins une activité.');
    }
  });

  const sheet = getSheet_(SHEET_INSCRIPTIONS);
  const now = new Date();
  const isNew = !group.id;
  const groupId = isNew ? 'g' + Utilities.getUuid().replace(/-/g, '').slice(0, 8) : String(group.id);

  const rows = readRows_(sheet);
  const existing = rows.filter(function (r) { return r.groupe_id === groupId; });
  if (!isNew && existing.length === 0) throw new Error('Cette inscription n\'existe plus (supprimée entre temps ?). Recharge la page.');
  const createdAt = existing.length ? (existing[0].cree_le || now) : now;

  // Le groupe modifié garde sa place dans le tableau (ou va à la fin s'il est nouveau).
  const newRows = members.map(function (m, i) {
    return [groupId, groupId + '-' + (i + 1), safe_(m.prenom), safe_(m.nom), m.laser, m.dejeuner, m.soiree, createdAt, now];
  });
  const out = [];
  let inserted = false;
  rows.forEach(function (r) {
    if (r.groupe_id === groupId) {
      if (!inserted) { Array.prototype.push.apply(out, newRows); inserted = true; }
    } else {
      out.push(rowToArray_(r));
    }
  });
  if (!inserted) Array.prototype.push.apply(out, newRows);
  writeAll_(sheet, out);

  log_(isNew ? 'CRÉATION' : 'MODIFICATION', groupId, {
    avant: existing.map(describeRow_),
    apres: members.map(describeRow_)
  });

  return { groupId: groupId };
}

function deleteGroup_(groupId) {
  if (!groupId) throw new Error('Inscription introuvable.');
  const sheet = getSheet_(SHEET_INSCRIPTIONS);
  const existing = readRows_(sheet).filter(function (r) { return r.groupe_id === String(groupId); });
  if (existing.length === 0) return { groupId: groupId };
  writeAll_(sheet, readRows_(sheet).filter(function (r) { return r.groupe_id !== String(groupId); }).map(rowToArray_));
  log_('SUPPRESSION', String(groupId), { avant: existing.map(describeRow_) });
  return { groupId: groupId };
}

function readGroups_() {
  const rows = readRows_(getSheet_(SHEET_INSCRIPTIONS));
  const byId = {};
  const order = [];
  rows.forEach(function (r) {
    if (!r.groupe_id) return;
    if (!byId[r.groupe_id]) {
      byId[r.groupe_id] = { id: r.groupe_id, createdAt: toIso_(r.cree_le), updatedAt: toIso_(r.maj_le), members: [] };
      order.push(r.groupe_id);
    }
    byId[r.groupe_id].members.push({
      prenom: String(r.prenom),
      nom: String(r.nom),
      laser: r.laser === true,
      dejeuner: r.dejeuner === true,
      soiree: r.soiree === true
    });
  });
  return order
    .map(function (id) { return byId[id]; })
    .sort(function (a, b) { return String(a.createdAt).localeCompare(String(b.createdAt)); });
}

/* ===================== OUTILS SHEET ===================== */

function readRows_(sheet) {
  const last = sheet.getLastRow();
  if (last < 2) return [];
  const data = sheet.getRange(2, 1, last - 1, HEADERS.length).getValues();
  return data.map(function (row) {
    const o = {};
    HEADERS.forEach(function (h, i) { o[h] = row[i]; });
    o.groupe_id = String(o.groupe_id || '');
    return o;
  });
}

function rowToArray_(r) {
  return HEADERS.map(function (h) { return r[h]; });
}

// Réécrit tout le tableau (80 lignes max, c'est instantané) : pas de lignes supprimées, pas de décalage.
function writeAll_(sheet, values) {
  const last = sheet.getLastRow();
  if (last >= 2) {
    const old = sheet.getRange(2, 1, last - 1, HEADERS.length);
    old.clearDataValidations();
    old.clearContent();
  }
  if (values.length === 0) return;
  // Colonnes identifiants et noms en texte brut (évite qu'un "007" devienne 7).
  sheet.getRange(2, 1, values.length, 4).setNumberFormat('@');
  sheet.getRange(2, 1, values.length, HEADERS.length).setValues(values);
  sheet.getRange(2, 5, values.length, 3).insertCheckboxes();
  sheet.getRange(2, 8, values.length, 2).setNumberFormat('dd/MM/yyyy HH:mm');
}

function getSheet_(name) {
  const sheet = SpreadsheetApp.getActive().getSheetByName(name);
  if (!sheet) throw new Error('Onglet "' + name + '" manquant : lance la fonction setup dans Apps Script.');
  return sheet;
}

function log_(action, groupId, details) {
  getSheet_(SHEET_JOURNAL).appendRow([new Date(), action, groupId, JSON.stringify(details)]);
}

function describeRow_(m) {
  const acts = ACTIVITIES.filter(function (a) { return m[a] === true; });
  return m.prenom + ' ' + m.nom + ' [' + acts.join(', ') + ']';
}

function cleanName_(v) {
  return String(v || '').replace(/\s+/g, ' ').trim().slice(0, MAX_NAME_LENGTH);
}

// Empêche qu'un nom commençant par =, +, - ou @ soit interprété comme une formule.
function safe_(v) {
  return /^[=+\-@]/.test(v) ? "'" + v : v;
}

function toIso_(d) {
  return d instanceof Date ? d.toISOString() : String(d || '');
}

function json_(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj)).setMimeType(ContentService.MimeType.JSON);
}

/* ===================== INSTALLATION ===================== */

/**
 * Crée ou remet en forme les onglets. Sans risque : peut être relancé, ne supprime aucune inscription.
 */
function setup() {
  const ss = SpreadsheetApp.getActive();

  // --- Inscriptions
  let ins = ss.getSheetByName(SHEET_INSCRIPTIONS);
  if (!ins) {
    const first = ss.getSheets()[0];
    if (ss.getSheets().length === 1 && first.getLastRow() === 0) {
      ins = first.setName(SHEET_INSCRIPTIONS);
    } else {
      ins = ss.insertSheet(SHEET_INSCRIPTIONS, 0);
    }
  }
  ins.getRange(1, 1, 1, HEADERS.length).setValues([HEADERS]).setFontWeight('bold').setBackground('#1a1033').setFontColor('#00f0ff');
  ins.setFrozenRows(1);
  ins.setColumnWidths(1, 2, 90);
  ins.setColumnWidths(3, 2, 140);

  // --- Récap
  let recap = ss.getSheetByName(SHEET_RECAP) || ss.insertSheet(SHEET_RECAP, 1);
  recap.clear();
  recap.getRange('A1:B1').setValues([['CYBERPUNK 3030 // RÉCAP', '']]).setFontWeight('bold');
  recap.getRange('A3:B9').setValues([
    ['Activité', 'Inscrits'],
    ['Laser game (11h)', '=SUMPRODUCT(--Inscriptions!E2:E)'],
    ['Déjeuner (14h)', '=SUMPRODUCT(--Inscriptions!F2:F)'],
    ['Soirée bar + DJ (19h)', '=SUMPRODUCT(--Inscriptions!G2:G)'],
    ['', ''],
    ['Personnes inscrites', '=COUNTA(Inscriptions!B2:B)'],
    ['Inscriptions (groupes)', '=COUNTUNIQUE(Inscriptions!A2:A)']
  ]);
  recap.getRange('A3:B3').setFontWeight('bold').setBackground('#1a1033').setFontColor('#ff2bd6');
  recap.setColumnWidth(1, 220);

  // --- Journal
  let journal = ss.getSheetByName(SHEET_JOURNAL) || ss.insertSheet(SHEET_JOURNAL, 2);
  if (journal.getLastRow() === 0) {
    journal.appendRow(['horodatage', 'action', 'groupe_id', 'details (avant / après)']);
  }
  journal.getRange('1:1').setFontWeight('bold');
  journal.setFrozenRows(1);
  journal.getRange('A:A').setNumberFormat('dd/MM/yyyy HH:mm:ss');
  journal.setColumnWidth(4, 600);
}
