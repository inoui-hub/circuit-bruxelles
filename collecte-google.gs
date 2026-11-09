/**
 * Serveur de la classe — Google Apps Script + Google Sheet.
 *
 * Déroulement de l'exercice :
 *  1. l'élève se connecte (prénom, nom, classe cochée dans la liste du professeur) ;
 *  2. il écrit une demande en mots-clés (comme dans une barre de recherche) ;
 *  3. toutes les demandes de la classe sont visibles par tous ;
 *  4. un élève choisit une demande (une seule, et une demande ne peut être choisie que par un élève)
 *     et construit l'itinéraire correspondant (6 points maximum) ; les points apparaissent en direct
 *     sur l'écran de projection ;
 *  5. il renvoie l'itinéraire à l'élève qui avait fait la demande ;
 *  6. le demandeur évalue l'itinéraire selon 3 critères.
 *
 * MISE EN PLACE (une seule fois), avec un compte Google qui ne vous identifie pas :
 *  1. Créez un Google Sheet vide (ou un projet Apps Script autonome + SHEET_ID).
 *  2. Extensions > Apps Script. Collez tout ce fichier à la place du code existant.
 *  3. Remplacez CLE_ENSEIGNANT par un mot de passe de votre choix.
 *  4. Déployer > Nouveau déploiement > « Application web » :
 *       - Exécuter en tant que : Moi
 *       - Qui a accès : Tout le monde
 *     Autorisez l'accès, puis copiez l'URL qui se termine par /exec.
 *  5. Page élève (index.html) : collez l'URL dans  const ENDPOINT = '...'
 *  6. Page enseignant (prof.html) : « Paramètres de collecte » -> collez l'URL et la clé.
 *
 * MODIFIER LE SCRIPT PLUS TARD : Déployer > Gérer les déploiements > crayon >
 * Version : « Nouvelle version » > Déployer. L'URL reste la même.
 *
 * LIMITES : c'est un outil de classe, pas un système sécurisé.
 *  - Les prénoms/noms et les demandes d'une classe sont lisibles par quiconque connaît l'URL et la classe.
 *  - Un élève qui retape le nom et la classe d'un camarade se reconnecte à sa place.
 *  - La clé enseignant protège la lecture de tous les itinéraires et les actions d'administration.
 *  - Pas de données sensibles : seulement des prénoms/noms et des lieux de Bruxelles.
 */
const CLE_ENSEIGNANT = 'changez-moi';
// Laisser vide si le script est créé depuis le tableur (Extensions > Apps Script).
// Si le script est autonome : identifiant du Google Sheet (entre /d/ et /edit dans son adresse).
const SHEET_ID = '';
const ONLINE_SECONDS = 45; // un élève est « en ligne » s'il a donné signe de vie depuis moins de 45 s

const T = {
  Participants: ['id', 'classe', 'nom', 'prenom', 'token', 'joinedAt'],
  Demandes: ['id', 'authorId', 'classe', 'text', 'createdAt', 'takenBy', 'takenAt', 'itin', 'itinAt', 'sentAt', 'ev'],
  Config: ['k', 'v'],
};

/* ---------- accès au tableur ---------- */
function sh_(name) {
  const ss = SHEET_ID ? SpreadsheetApp.openById(SHEET_ID) : SpreadsheetApp.getActiveSpreadsheet();
  let sh = ss.getSheetByName(name);
  if (!sh) {
    sh = ss.insertSheet(name);
    sh.appendRow(T[name]);
    sh.getRange(1, 1, sh.getMaxRows(), T[name].length).setNumberFormat('@'); // tout en texte
  }
  return sh;
}
function rows_(name) {
  const cols = T[name], v = sh_(name).getDataRange().getValues(), out = [];
  for (let i = 1; i < v.length; i++) {
    const o = { _row: i + 1 };
    cols.forEach((c, j) => { o[c] = String(v[i][j] === undefined || v[i][j] === null ? '' : v[i][j]); });
    if (o[cols[0]]) out.push(o);
  }
  return out;
}
function put_(name, o) {
  const cols = T[name], sh = sh_(name);
  if (!o._row) o._row = sh.getLastRow() + 1;
  sh.getRange(o._row, 1, 1, cols.length).setValues([cols.map(c => o[c] === undefined ? '' : String(o[c]))]);
}
function del_(name, list) { // suppression de plusieurs lignes, de bas en haut
  list.slice().sort((a, b) => b._row - a._row).forEach(o => sh_(name).deleteRow(o._row));
}

/* ---------- utilitaires ---------- */
const clean_ = (s, n) => String(s == null ? '' : s).replace(/\s+/g, ' ').trim().slice(0, n || 40);
const key_ = (nom, prenom, classe) => [nom, prenom, classe].map(x => String(x).trim().toLowerCase()).join('|');
const ok_ = o => out_(Object.assign({ ok: true }, o || {}));
const err_ = m => out_({ ok: false, error: m });
function out_(o) { return ContentService.createTextOutput(JSON.stringify(o)).setMimeType(ContentService.MimeType.JSON); }
function parse_(s) { try { return s ? JSON.parse(s) : null; } catch (e) { return null; } }
const name_ = p => (p ? p.prenom + ' ' + p.nom : '');

function auth_(b) {
  if (!b || !b.id || !b.token) return null;
  const r = rows_('Participants').find(x => x.id === String(b.id));
  return r && r.token === String(b.token) ? r : null;
}
function classes_() {
  const r = rows_('Config').find(x => x.k === 'classes');
  const l = parse_(r && r.v);
  return Array.isArray(l) ? l : [];
}
function ping_(id) { CacheService.getScriptCache().put('p_' + id, '1', ONLINE_SECONDS); }
function online_(ids) {
  if (!ids.length) return {};
  const got = CacheService.getScriptCache().getAll(ids.map(i => 'p_' + i)), o = {};
  ids.forEach(i => { o[i] = !!got['p_' + i]; });
  return o;
}
const status_ = q => (q.sentAt ? 'sent' : q.takenBy ? 'taken' : 'free');
function reqView_(q, byId) {
  return { id: q.id, authorId: q.authorId, author: name_(byId[q.authorId]), text: q.text,
    takenBy: q.takenBy, taker: q.takenBy ? name_(byId[q.takenBy]) : '', status: status_(q), evaluated: !!q.ev };
}

function cleanItin_(d, min) {
  if (!d || !Array.isArray(d.points) || d.points.length < min || d.points.length > 12) return null;
  const pts = [];
  for (const p of d.points) {
    const lat = Number(p.lat), lng = Number(p.lng);
    if (!isFinite(lat) || !isFinite(lng) || lat < 50.5 || lat > 51.2 || lng < 4.0 || lng > 4.8) return null;
    pts.push({
      lat: lat, lng: lng, name: clean_(p.name, 60), cat: clean_(p.cat, 12),
      time: /^\d{1,2}:\d{2}$/.test(String(p.time || '')) ? String(p.time) : '',
      dur: Math.max(0, Math.min(600, Number(p.dur) || 0)), comment: clean_(p.comment, 500),
    });
  }
  return { version: 3, titre: clean_(d.titre, 80), intro: clean_(d.intro, 200), points: pts };
}

/* ---------- lecture : classes, état de l'élève ---------- */
function classes_out_() { return ok_({ classes: classes_() }); }

function join_(b) {
  const nom = clean_(b.nom), prenom = clean_(b.prenom), classe = clean_(b.classe, 20);
  if (!nom || !prenom || !classe) return err_('Le nom, le prénom et la classe sont obligatoires.');
  const list = classes_();
  if (!list.length) return err_('Aucune classe n\'est ouverte pour le moment. Demandez à votre professeur.');
  if (list.indexOf(classe) < 0) return err_('Classe inconnue : cochez une classe de la liste.');
  const k = key_(nom, prenom, classe);
  let r = rows_('Participants').find(x => key_(x.nom, x.prenom, x.classe) === k);
  if (!r) {
    r = { id: Utilities.getUuid().slice(0, 8), classe: classe, nom: nom, prenom: prenom, token: Utilities.getUuid(), joinedAt: new Date().toISOString() };
    put_('Participants', r);
  }
  ping_(r.id);
  return ok_({ me: { id: r.id, token: r.token, nom: r.nom, prenom: r.prenom, classe: r.classe } });
}

function state_(b) {
  const r = auth_(b); if (!r) return err_('Session invalide : reconnectez-vous.');
  ping_(r.id);
  const ps = rows_('Participants'), byId = {};
  ps.forEach(p => { byId[p.id] = p; });
  const mates = ps.filter(p => p.classe === r.classe), on = online_(mates.map(p => p.id));
  const qs = rows_('Demandes').filter(q => q.classe === r.classe);
  const mine = qs.find(q => q.authorId === r.id), task = qs.find(q => q.takenBy === r.id);
  return ok_({
    me: { id: r.id, nom: r.nom, prenom: r.prenom, classe: r.classe },
    online: mates.map(p => ({ id: p.id, prenom: p.prenom, nom: p.nom, online: !!on[p.id] })),
    requests: qs.map(q => reqView_(q, byId)),
    mine: mine ? Object.assign(reqView_(mine, byId), {
      received: mine.sentAt ? { from: name_(byId[mine.takenBy]), itin: parse_(mine.itin), sentAt: mine.sentAt } : null,
      ev: parse_(mine.ev) }) : null,
    task: task ? { id: task.id, text: task.text, author: name_(byId[task.authorId]), itin: parse_(task.itin), sent: !!task.sentAt } : null,
  });
}

/* ---------- actions élèves ---------- */
function request_(b) {
  const r = auth_(b); if (!r) return err_('Session invalide : reconnectez-vous.');
  const text = clean_(b.text, 100);
  if (text.length < 2) return err_('Écrivez votre demande en quelques mots-clés.');
  const q = rows_('Demandes').find(x => x.authorId === r.id);
  if (q) {
    if (q.takenBy) return err_('Votre demande a déjà été choisie par un camarade : elle ne peut plus être modifiée.');
    q.text = text; put_('Demandes', q); return ok_({});
  }
  put_('Demandes', { id: Utilities.getUuid().slice(0, 8), authorId: r.id, classe: r.classe, text: text, createdAt: new Date().toISOString() });
  return ok_({});
}
function delrequest_(b) {
  const r = auth_(b); if (!r) return err_('Session invalide : reconnectez-vous.');
  const q = rows_('Demandes').find(x => x.authorId === r.id);
  if (!q) return ok_({});
  if (q.takenBy) return err_('Votre demande a déjà été choisie : elle ne peut plus être retirée.');
  del_('Demandes', [q]); return ok_({});
}
function take_(b) {
  const r = auth_(b); if (!r) return err_('Session invalide : reconnectez-vous.');
  const qs = rows_('Demandes');
  if (qs.find(x => x.takenBy === r.id)) return err_('Vous avez déjà choisi une demande.');
  const q = qs.find(x => x.id === String(b.requestId));
  if (!q) return err_('Demande introuvable (elle a peut-être été retirée).');
  if (q.classe !== r.classe) return err_('Cette demande n\'est pas dans votre classe.');
  if (q.authorId === r.id) return err_('Vous ne pouvez pas choisir votre propre demande.');
  if (q.takenBy) return err_('Cette demande vient d\'être choisie par un camarade. Choisissez-en une autre.');
  q.takenBy = r.id; q.takenAt = new Date().toISOString(); put_('Demandes', q);
  return ok_({});
}
function release_(b) {
  const r = auth_(b); if (!r) return err_('Session invalide : reconnectez-vous.');
  const q = rows_('Demandes').find(x => x.takenBy === r.id);
  if (!q) return ok_({});
  if (q.sentAt) return err_('Votre itinéraire est déjà envoyé.');
  q.takenBy = ''; q.takenAt = ''; q.itin = ''; q.itinAt = ''; put_('Demandes', q);
  return ok_({});
}
function draft_(b) {
  const r = auth_(b); if (!r) return err_('Session invalide : reconnectez-vous.');
  const q = rows_('Demandes').find(x => x.takenBy === r.id);
  if (!q) return err_('Choisissez d\'abord une demande.');
  if (q.sentAt) return err_('Votre itinéraire est déjà envoyé : il ne peut plus être modifié.');
  const it = cleanItin_(b.itin, 0); if (!it) return err_('Itinéraire invalide (points à placer à Bruxelles).');
  q.itin = JSON.stringify(it); q.itinAt = new Date().toISOString(); put_('Demandes', q);
  return ok_({});
}
function send_(b) {
  const r = auth_(b); if (!r) return err_('Session invalide : reconnectez-vous.');
  const q = rows_('Demandes').find(x => x.takenBy === r.id);
  if (!q) return err_('Choisissez d\'abord une demande.');
  if (q.sentAt) return err_('Votre itinéraire est déjà envoyé.');
  const it = cleanItin_(b.itin, 1); if (!it) return err_('Itinéraire invalide : placez au moins un point à Bruxelles.');
  q.itin = JSON.stringify(it); q.itinAt = new Date().toISOString(); q.sentAt = q.itinAt; put_('Demandes', q);
  return ok_({});
}
function evaluate_(b) {
  const r = auth_(b); if (!r) return err_('Session invalide : reconnectez-vous.');
  const q = rows_('Demandes').find(x => x.authorId === r.id);
  if (!q || !q.sentAt) return err_('Vous n\'avez pas encore reçu d\'itinéraire.');
  const n = v => { v = Math.round(Number(v)); return v >= 1 && v <= 5 ? v : 0; };
  const ev = { c1: n(b.c1), c2: n(b.c2), c3: n(b.c3), comment: clean_(b.comment, 400), at: new Date().toISOString() };
  if (!ev.c1 || !ev.c2 || !ev.c3) return err_('Répondez aux trois critères (de 1 à 5).');
  q.ev = JSON.stringify(ev); put_('Demandes', q);
  return ok_({});
}

/* ---------- actions enseignant (clé requise) ---------- */
function all_() {
  const ps = rows_('Participants'), qs = rows_('Demandes'), on = online_(ps.map(p => p.id)), byId = {};
  ps.forEach(p => { byId[p.id] = p; });
  return ok_({
    classes: classes_(), now: new Date().toISOString(),
    students: ps.map(p => ({ id: p.id, classe: p.classe, nom: p.nom, prenom: p.prenom, joinedAt: p.joinedAt, online: !!on[p.id] })),
    requests: qs.map(q => Object.assign(reqView_(q, byId), { classe: q.classe, itin: parse_(q.itin), itinAt: q.itinAt, sentAt: q.sentAt, ev: parse_(q.ev) })),
  });
}
function setclasses_(b) {
  const seen = {}, list = [];
  (Array.isArray(b.classes) ? b.classes : []).forEach(c => { c = clean_(c, 20); if (c && !seen[c.toLowerCase()]) { seen[c.toLowerCase()] = 1; list.push(c); } });
  const row = rows_('Config').find(x => x.k === 'classes') || { k: 'classes' };
  row.v = JSON.stringify(list.slice(0, 30)); put_('Config', row);
  return ok_({ classes: list.slice(0, 30) });
}
function unsend_(b) {
  const q = rows_('Demandes').find(x => x.id === String(b.requestId));
  if (!q) return err_('Demande introuvable.');
  q.sentAt = ''; q.ev = ''; put_('Demandes', q); return ok_({});
}
function delreq_(b) {
  const q = rows_('Demandes').find(x => x.id === String(b.requestId));
  if (!q) return err_('Demande introuvable.');
  del_('Demandes', [q]); return ok_({});
}
function remove_(b) {
  const p = rows_('Participants').find(x => x.id === String(b.id));
  if (!p) return err_('Élève introuvable.');
  const qs = rows_('Demandes');
  qs.filter(q => q.takenBy === p.id).forEach(q => { q.takenBy = ''; q.takenAt = ''; q.itin = ''; q.itinAt = ''; q.sentAt = ''; q.ev = ''; put_('Demandes', q); });
  del_('Demandes', qs.filter(q => q.authorId === p.id));
  del_('Participants', [p]); return ok_({});
}
function purge_(b) {
  ['Participants', 'Demandes'].forEach(n => {
    const sh = sh_(n), last = sh.getLastRow();
    if (last > 1) sh.deleteRows(2, last - 1);
  });
  return ok_({});
}

/* ---------- points d'entrée ---------- */
const READ = { state: state_ };
const WRITE = { join: join_, request: request_, delrequest: delrequest_, take: take_, release: release_, draft: draft_, send: send_, evaluate: evaluate_ };
const ADMIN = { setclasses: setclasses_, unsend: unsend_, delreq: delreq_, remove: remove_, purge: purge_ };

function doPost(e) {
  let b; try { b = JSON.parse(e.postData.contents); } catch (x) { return err_('Requête illisible.'); }
  const admin = b.key === CLE_ENSEIGNANT;
  try {
    if (b.action === 'classes') return classes_out_();
    if (b.action === 'all') return admin ? all_() : err_('Clé enseignant incorrecte.');
    if (READ[b.action]) return READ[b.action](b); // lectures : sans verrou, pour rester rapide
    const fn = WRITE[b.action] || (ADMIN[b.action] && (admin ? ADMIN[b.action] : null));
    if (!fn) return err_(ADMIN[b.action] ? 'Clé enseignant incorrecte.' : 'Action inconnue.');
    const lock = LockService.getScriptLock();
    lock.waitLock(20000);
    try { return fn(b); } finally { lock.releaseLock(); }
  } catch (x) {
    return err_('Erreur serveur.');
  }
}
function doGet(e) {
  const p = (e && e.parameter) || {};
  if (p.action === 'classes') return classes_out_();
  if (p.action === 'all') return p.key === CLE_ENSEIGNANT ? all_() : err_('Clé enseignant incorrecte.');
  return ok_({ service: 'itineraires-bruxelles' });
}
