/**
 * Collecte des itinéraires d'élèves — Google Apps Script (facultatif).
 *
 * MISE EN PLACE (une seule fois, ~5 minutes) :
 *  1. Créez un Google Sheet vide (par ex. « Itinéraires Bruxelles »).
 *  2. Menu Extensions > Apps Script. Collez tout ce fichier à la place du code existant.
 *  3. Remplacez CLE_ENSEIGNANT ci-dessous par un mot de passe de votre choix.
 *  4. Déployer > Nouveau déploiement > type « Application web » :
 *       - Exécuter en tant que : Moi
 *       - Qui a accès : Tout le monde
 *     Autorisez l'accès, puis copiez l'URL qui se termine par /exec.
 *  5. Page élève (index.html) : collez cette URL dans  const ENDPOINT = '...'
 *  6. Page enseignant (prof.html) : « Paramètres de collecte » -> collez l'URL et la clé.
 *
 * Les envois s'accumulent dans l'onglet « Envois » du Sheet (une ligne par envoi).
 * Un élève qui renvoie son itinéraire remplace le précédent (même nom + prénom + classe) à la lecture.
 *
 * LIMITES : la clé enseignant protège la lecture, pas l'écriture : n'importe qui connaissant
 * l'URL peut envoyer une ligne. Pour un usage en classe c'est suffisant ; ne mettez pas de données sensibles.
 */
const CLE_ENSEIGNANT = 'changez-moi';
const ONGLET = 'Envois';

function feuille_() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  return ss.getSheetByName(ONGLET) || ss.insertSheet(ONGLET);
}

// Réception d'un itinéraire envoyé par un élève
function doPost(e) {
  try {
    const d = JSON.parse(e.postData.contents);
    if (!d || !d.student || !Array.isArray(d.points) || d.points.length > 12) return texte_('refus');
    const json = JSON.stringify(d);
    if (json.length > 40000) return texte_('trop volumineux');
    feuille_().appendRow([new Date(), d.student.classe, d.student.nom, d.student.prenom, json]);
    return texte_('ok');
  } catch (err) {
    return texte_('erreur');
  }
}

// Lecture par l'enseignant : ?key=CLE_ENSEIGNANT
function doGet(e) {
  if (!e || !e.parameter || e.parameter.key !== CLE_ENSEIGNANT) return texte_('[]');
  const rows = feuille_().getDataRange().getValues();
  const out = [];
  rows.forEach(r => { try { out.push(JSON.parse(r[4])); } catch (_) {} });
  return ContentService.createTextOutput(JSON.stringify(out)).setMimeType(ContentService.MimeType.JSON);
}

function texte_(s) {
  return ContentService.createTextOutput(s);
}
