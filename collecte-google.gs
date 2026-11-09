/**
 * Serveur de la classe — Google Apps Script + Google Sheet.
 *
 * Ce que fait ce script :
 *  - les élèves se connectent avec nom, prénom, classe (liste visible par le professeur) ;
 *  - ils envoient leur itinéraire au professeur ;
 *  - ils peuvent envoyer leur itinéraire à UN camarade de la classe ; un camarade ne peut
 *    être choisi qu'une seule fois (celui qui est déjà choisi n'est plus sélectionnable) ;
 *  - chaque élève retrouve l'itinéraire qu'il a reçu, affiché en feuille de route.
 *
 * MISE EN PLACE (une seule fois, ~5 minutes), avec un compte Google qui ne vous identifie pas :
 *  1. Créez un Google Sheet vide.
 *  2. Extensions > Apps Script. Collez tout ce fichier à la place du code existant.
 *  3. Remplacez CLE_ENSEIGNANT par un mot de passe de votre choix.
 *  4. Déployer > Nouveau déploiement > type « Application web » :
 *       - Exécuter en tant que : Moi
 *       - Qui a accès : Tout le monde
 *     Autorisez l'accès, puis copiez l'URL qui se termine par /exec.
 *  5. Page élève (index.html) : collez l'URL dans  const ENDPOINT = '...'
 *  6. Page enseignant (prof.html) : « Paramètres de collecte » -> collez l'URL et la clé.
 *
 * MODIFIER LE SCRIPT PLUS TARD : après chaque modification, Déployer > Gérer les déploiements >
 * crayon > Version : « Nouvelle version » > Déployer. L'URL reste la même.
 *
 * LIMITES : c'est un outil de classe, pas un système sécurisé.
 *  - La liste des élèves d'une classe (prénom, nom) est lisible par quiconque connaît l'URL.
 *  - Un élève qui retape le nom et la classe d'un camarade se reconnecte à sa place.
 *  - La clé enseignant protège la lecture des itinéraires et les actions d'administration.
 *  - Pas de données sensibles : seulement des prénoms/noms et des lieux de Bruxelles.
 */
const CLE_ENSEIGNANT = 'changez-moi';
const ONGLET = 'Eleves';
const COLS = ['id', 'classe', 'nom', 'prenom', 'token', 'joinedAt', 'itin', 'itinAt', 'sendTo', 'recvFrom'];

/* ---------- accès au tableur ---------- */
function sheet_() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sh = ss.getSheetByName(ONGLET);
  if (!sh) {
    sh = ss.insertSheet(ONGLET);
    sh.appendRow(COLS);
    sh.getRange('A:J').setNumberFormat('@'); // tout en texte
  }
  return sh;
}
function rows_() {
  const v = sheet_().getDataRange().getValues();
  const out = [];
  for (let i = 1; i < v.length; i++) {
    const o = { _row: i + 1 };
    COLS.forEach((c, j) => { o[c] = String(v[i][j] === undefined || v[i][j] === null ? '' : v[i][j]); });
    if (o.id) out.push(o);
  }
  return out;
}
function write_(o) {
  sheet_().getRange(o._row, 1, 1, COLS.length).setValues([COLS.map(c => o[c] === undefined ? '' : String(o[c]))]);
}

/* ---------- utilitaires ---------- */
const clean_ = (s, n) => String(s == null ? '' : s).replace(/\s+/g, ' ').trim().slice(0, n || 40);
const key_ = (nom, prenom, classe) => [nom, prenom, classe].map(x => String(x).trim().toLowerCase()).join('|');
const ok_ = o => out_(Object.assign({ ok: true }, o || {}));
const err_ = m => out_({ ok: false, error: m });
function out_(o) {
  return ContentService.createTextOutput(JSON.stringify(o)).setMimeType(ContentService.MimeType.JSON);
}
function auth_(b) {
  if (!b || !b.id || !b.token) return null;
  const r = rows_().find(x => x.id === String(b.id));
  return r && r.token === String(b.token) ? r : null;
}
function byId_(id) { return rows_().find(x => x.id === String(id)) || null; }
const who_ = r => ({ id: r.id, prenom: r.prenom, nom: r.nom });
function parse_(s) { try { return s ? JSON.parse(s) : null; } catch (e) { return null; } }

function cleanItin_(d) {
  if (!d || !Array.isArray(d.points) || d.points.length < 1 || d.points.length > 12) return null;
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
  return { version: 2, titre: clean_(d.titre, 80), pour: clean_(d.pour, 80), intro: clean_(d.intro, 200), points: pts };
}

/* ---------- actions élèves ---------- */
function join_(b) {
  const nom = clean_(b.nom), prenom = clean_(b.prenom), classe = clean_(b.classe, 12);
  if (!nom || !prenom || !classe) return err_('Le nom, le prénom et la classe sont obligatoires.');
  const k = key_(nom, prenom, classe);
  let r = rows_().find(x => key_(x.nom, x.prenom, x.classe) === k);
  if (!r) {
    r = { _row: sheet_().getLastRow() + 1, id: Utilities.getUuid().slice(0, 8), classe: classe, nom: nom, prenom: prenom,
      token: Utilities.getUuid(), joinedAt: new Date().toISOString(), itin: '', itinAt: '', sendTo: '', recvFrom: '' };
    write_(r);
  }
  const to = r.sendTo ? byId_(r.sendTo) : null;
  return ok_({ me: { id: r.id, token: r.token, nom: r.nom, prenom: r.prenom, classe: r.classe },
    itin: parse_(r.itin), sentTo: to ? who_(to) : null });
}
function submit_(b) {
  const r = auth_(b); if (!r) return err_('Session invalide : reconnectez-vous.');
  if (r.sendTo) return err_('Votre itinéraire a déjà été envoyé à un camarade : il ne peut plus être modifié.');
  const itin = cleanItin_(b.itin); if (!itin) return err_('Itinéraire invalide (1 à 12 points placés à Bruxelles).');
  r.itin = JSON.stringify(itin); r.itinAt = new Date().toISOString(); write_(r);
  return ok_({});
}
function send_(b) {
  const r = auth_(b); if (!r) return err_('Session invalide : reconnectez-vous.');
  if (!r.itin) return err_('Envoyez d\'abord votre itinéraire au professeur.');
  if (r.sendTo) return err_('Vous avez déjà envoyé votre itinéraire.');
  const t = byId_(b.toId);
  if (!t) return err_('Camarade introuvable.');
  if (t.id === r.id) return err_('Vous ne pouvez pas vous envoyer votre propre itinéraire.');
  if (t.classe !== r.classe) return err_('Ce camarade n\'est pas dans votre classe.');
  if (t.recvFrom) return err_('Ce camarade a déjà été choisi par quelqu\'un d\'autre. Actualisez la liste.');
  r.sendTo = t.id; t.recvFrom = r.id; write_(r); write_(t);
  return ok_({ sentTo: who_(t) });
}
function inbox_(b) {
  const r = auth_(b); if (!r) return err_('Session invalide : reconnectez-vous.');
  const from = r.recvFrom ? byId_(r.recvFrom) : null;
  const to = r.sendTo ? byId_(r.sendTo) : null;
  return ok_({
    received: from && from.itin ? { from: { prenom: from.prenom, nom: from.nom }, itin: parse_(from.itin) } : null,
    sentTo: to ? who_(to) : null,
    hasItin: !!r.itin,
  });
}
function list_(classe) {
  const c = clean_(classe, 12);
  const rows = rows_().filter(r => !c || r.classe === c);
  return ok_({ students: rows.map(r => ({ id: r.id, prenom: r.prenom, nom: r.nom, classe: r.classe, hasItin: !!r.itin, taken: !!r.recvFrom })) });
}

/* ---------- actions enseignant (clé requise) ---------- */
function all_() {
  return ok_({ students: rows_().map(r => ({
    id: r.id, classe: r.classe, nom: r.nom, prenom: r.prenom, joinedAt: r.joinedAt, itinAt: r.itinAt,
    itin: parse_(r.itin), sendTo: r.sendTo, recvFrom: r.recvFrom })) });
}
function unlink_(r) {
  if (r.sendTo) { const t = byId_(r.sendTo); if (t) { t.recvFrom = ''; write_(t); } r.sendTo = ''; }
}
function unsend_(b) {
  const r = byId_(b.id); if (!r) return err_('Élève introuvable.');
  unlink_(r); write_(r); return ok_({});
}
function remove_(b) {
  const r = byId_(b.id); if (!r) return err_('Élève introuvable.');
  unlink_(r);
  if (r.recvFrom) { const s = byId_(r.recvFrom); if (s) { s.sendTo = ''; write_(s); } }
  sheet_().deleteRow(r._row); return ok_({});
}

/* ---------- points d'entrée ---------- */
function doPost(e) {
  const lock = LockService.getScriptLock();
  lock.waitLock(20000);
  try {
    let b; try { b = JSON.parse(e.postData.contents); } catch (x) { return err_('Requête illisible.'); }
    const admin = b.key === CLE_ENSEIGNANT;
    switch (b.action) {
      case 'join': return join_(b);
      case 'submit': return submit_(b);
      case 'send': return send_(b);
      case 'inbox': return inbox_(b);
      case 'list': return list_(b.classe);
      case 'all': return admin ? all_() : err_('Clé enseignant incorrecte.');
      case 'unsend': return admin ? unsend_(b) : err_('Clé enseignant incorrecte.');
      case 'remove': return admin ? remove_(b) : err_('Clé enseignant incorrecte.');
      default: return err_('Action inconnue.');
    }
  } catch (x) {
    return err_('Erreur serveur.');
  } finally {
    lock.releaseLock();
  }
}
function doGet(e) {
  const p = (e && e.parameter) || {};
  if (p.action === 'list') return list_(p.classe);
  if (p.action === 'all') return p.key === CLE_ENSEIGNANT ? all_() : err_('Clé enseignant incorrecte.');
  return ok_({ service: 'itineraires-bruxelles' });
}
