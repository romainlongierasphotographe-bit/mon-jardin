/**
 * Mon Jardin — script compagnon (photos + fiches automatiques)
 * ------------------------------------------------------------------
 * Enregistre les photos dans Google Drive (dossier « Mon Jardin – Photos »)
 * et note chaque photo dans l'onglet « Photos » du Google Sheet :
 *   ID | Type (plante / massif) | PlantID | Massif | Date | FileID | URL | Note
 * Peut aussi définir la photo de profil d'une plante (colonne « Photo »
 * de l'onglet « Plantes »).
 *
 * Installation : voir apps-script/LISEZMOI.md
 * Ce script est indépendant du script existant de l'appli : il ne le remplace pas.
 */

const SHEET_ID      = '1Hp0J2iwvFoSwqwd2nrlR0ILUfDx57y27-mMLxOeG2-U';
const FOLDER_NAME   = 'Mon Jardin – Photos';
const PHOTOS_SHEET  = 'Photos';
const PLANTES_SHEET = 'Plantes';
const HEADERS       = ['ID', 'Type', 'PlantID', 'Massif', 'Date', 'FileID', 'URL', 'Note'];

function doGet() {
  return json_({ ok: true, service: 'Mon Jardin – photos' });
}

function doPost(e) {
  let res;
  try {
    const p = JSON.parse(e.postData.contents);
    switch (p.action) {
      case 'upload':     res = upload_(p); break;
      case 'setProfile': res = withLock_(() => setProfile_(p.plantId, p.url || '')); break;
      case 'delete':     res = withLock_(() => deletePhoto_(p.id)); break;
      case 'fiche':      res = fiche_(p.nom || '', p.latin || ''); break;
      default: throw new Error('Action inconnue : ' + p.action);
    }
    res.ok = true;
  } catch (err) {
    res = { ok: false, error: String(err && err.message ? err.message : err) };
  }
  return json_(res);
}

/** À lancer une fois depuis l'éditeur pour autoriser Drive et Sheets. */
function autoriser() {
  getFolder_();
  getPhotosSheet_();
  Logger.log('OK : dossier Drive et onglet Photos prêts.');
}

// ── Actions ─────────────────────────────────────────────────────────

function upload_(p) {
  const m = /^data:(image\/[\w+.-]+);base64,(.+)$/.exec(p.data || '');
  if (!m) throw new Error('Image invalide');
  const type = p.type === 'massif' ? 'massif' : 'plante';
  if (type === 'plante' && !p.plantId) throw new Error('Plante manquante');
  if (type === 'massif' && !p.massif) throw new Error('Massif manquant');

  const date = /^\d{4}-\d{2}-\d{2}$/.test(p.date || '') ? p.date : today_();
  const id = Date.now().toString(36) + Math.floor(Math.random() * 1296).toString(36);
  const label = type === 'massif' ? p.massif : (p.plantNom || p.plantId);
  const ext = m[1] === 'image/png' ? 'png' : 'jpg';
  const blob = Utilities.newBlob(Utilities.base64Decode(m[2]), m[1], `${date} – ${label} – ${id}.${ext}`);

  const file = getFolder_().createFile(blob);
  file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
  const fileId = file.getId();
  const url = 'https://drive.google.com/thumbnail?id=' + fileId + '&sz=w1600';
  const photo = {
    id, type,
    plantId: type === 'plante' ? p.plantId : '',
    massif: type === 'massif' ? p.massif : '',
    date, fileId, url,
    note: p.note || ''
  };

  withLock_(() => {
    getPhotosSheet_().appendRow(HEADERS.map(h => ({
      ID: photo.id, Type: photo.type, PlantID: photo.plantId, Massif: photo.massif,
      Date: photo.date, FileID: photo.fileId, URL: photo.url, Note: photo.note
    })[h]));
    if (type === 'plante' && p.setProfile) setProfile_(p.plantId, url);
  });
  return { photo };
}

function setProfile_(plantId, url) {
  const sh = SpreadsheetApp.openById(SHEET_ID).getSheetByName(PLANTES_SHEET);
  if (!sh) throw new Error('Onglet « ' + PLANTES_SHEET + ' » introuvable');
  const values = sh.getDataRange().getValues();
  const headers = values[0].map(h => String(h).trim());
  const cId = headers.indexOf('ID');
  const cPhoto = headers.indexOf('Photo');
  if (cId < 0 || cPhoto < 0) throw new Error('Colonnes « ID » ou « Photo » introuvables');
  for (let r = 1; r < values.length; r++) {
    if (String(values[r][cId]) === String(plantId)) {
      sh.getRange(r + 1, cPhoto + 1).setValue(url);
      return { plantId, url };
    }
  }
  throw new Error('Plante introuvable : ' + plantId);
}

function deletePhoto_(id) {
  const sh = getPhotosSheet_();
  const values = sh.getDataRange().getValues();
  for (let r = 1; r < values.length; r++) {
    if (String(values[r][0]) === String(id)) {
      const row = values[r];
      const fileId = row[HEADERS.indexOf('FileID')];
      const url = row[HEADERS.indexOf('URL')];
      const plantId = row[HEADERS.indexOf('PlantID')];
      // Mise à la corbeille Drive (récupérable pendant 30 jours)
      try { if (fileId) DriveApp.getFileById(fileId).setTrashed(true); } catch (e) {}
      sh.deleteRow(r + 1);
      // Si c'était la photo de profil, on la retire
      let profileCleared = false;
      if (plantId) {
        const pl = SpreadsheetApp.openById(SHEET_ID).getSheetByName(PLANTES_SHEET);
        const pv = pl.getDataRange().getValues();
        const h = pv[0].map(x => String(x).trim());
        const cId = h.indexOf('ID'), cPhoto = h.indexOf('Photo');
        for (let i = 1; i < pv.length; i++) {
          if (String(pv[i][cId]) === String(plantId) && String(pv[i][cPhoto]) === String(url)) {
            pl.getRange(i + 1, cPhoto + 1).setValue('');
            profileCleared = true;
          }
        }
      }
      return { id, profileCleared };
    }
  }
  throw new Error('Photo introuvable : ' + id);
}

// ── Fiche plante automatique (IA) ───────────────────────────────────
// Clé à ajouter dans Paramètres du projet → Propriétés du script :
//   GEMINI_API_KEY    (gratuit : https://aistudio.google.com/apikey)
//   ou ANTHROPIC_API_KEY (Claude, payant à l'usage : https://console.anthropic.com)
// Facultatif : GEMINI_MODEL (défaut gemini-flash-latest), CLAUDE_MODEL (défaut claude-haiku-5-5)

function fiche_(nom, latin) {
  nom = String(nom).trim(); latin = String(latin).trim();
  if (!nom && !latin) throw new Error('Nom de plante manquant');
  const cacheKey = 'fiche:' + (nom + '|' + latin).toLowerCase().slice(0, 200);
  const cache = CacheService.getScriptCache();
  const hit = cache.get(cacheKey);
  if (hit) return { fiche: JSON.parse(hit), cache: true };

  const props = PropertiesService.getScriptProperties();
  const gemini = props.getProperty('GEMINI_API_KEY');
  const claude = props.getProperty('ANTHROPIC_API_KEY');
  if (!gemini && !claude) throw new Error('aucune clé IA configurée dans le script (GEMINI_API_KEY ou ANTHROPIC_API_KEY)');

  const prompt = fichePrompt_(nom, latin);
  const text = gemini
    ? callGemini_(gemini, props.getProperty('GEMINI_MODEL') || 'gemini-flash-latest', prompt)
    : callClaude_(claude, props.getProperty('CLAUDE_MODEL') || 'claude-haiku-5-5', prompt);

  const json = String(text).replace(/^[\s\S]*?(\{)/, '$1').replace(/\}[^}]*$/, '}');
  let fiche;
  try { fiche = JSON.parse(json); } catch (e) { throw new Error('réponse de l\'IA illisible'); }
  cache.put(cacheKey, JSON.stringify(fiche), 21600);
  return { fiche };
}

function fichePrompt_(nom, latin) {
  return `Tu es un jardinier professionnel français. Donne la fiche pratique de la plante « ${nom}${latin ? ` » (nom latin indiqué : ${latin})` : ' »'}, pour un jardin en Dordogne (Périgord, sud-ouest de la France : climat océanique doux, étés chauds et secs, gelées d'hiver jusqu'à -10 °C environ). Si le nom désigne une variété ou un cultivar, adapte la fiche à cette variété. Les mois et conseils doivent correspondre à ce climat.

Réponds UNIQUEMENT avec un objet JSON valide, sans texte autour, avec exactement ces clés :
{
  "trouve": true si tu reconnais la plante avec certitude, sinon false,
  "nomCommun": "nom commun français usuel",
  "nomLatin": "Genre espèce 'Cultivar' si connu",
  "type": "catégorie courte : Arbuste, Arbuste persistant, Vivace, Annuelle, Bisannuelle, Bulbe, Arbre, Arbre fruitier, Grimpante, Graminée, Aromatique, Agrume, Couvre-sol, Fougère, Potagère…",
  "croissance": "Lente" | "Moyenne" | "Rapide",
  "longevite": "court : Annuelle, Bisannuelle, Vivace, 10 à 20 ans, Plusieurs décennies…",
  "hauteur": "hauteur adulte, ex. 60 cm – 1 m",
  "etalement": "largeur adulte, ex. 50 cm",
  "rusticite": "une catégorie suivie de la température entre parenthèses : Très rustique (-20 °C), Rustique (-12 °C), Semi-rustique (-5 °C) ou Sensible au gel (0 °C)",
  "parfum": "Oui" | "Léger" | "Non",
  "exposition": "Plein soleil" | "Mi-ombre" | "Ombre",
  "arrosage": "Faible" | "Modéré" | "Régulier" | "Abondant",
  "floraison": "période courte, ex. Juin – Sept. (vide si pas de floraison décorative)",
  "floraisonDebut": numéro du mois de début de floraison (1 à 12, 0 si aucune),
  "floraisonFin": numéro du mois de fin de floraison (1 à 12, 0 si aucune),
  "couleur": "couleurs des fleurs, ex. Rose, blanc",
  "sol": "type de sol en quelques mots, ex. Drainant, calcaire toléré",
  "arrosageDetail": "2 à 4 conseils courts séparés par des points-virgules",
  "solDetail": "2 à 4 conseils courts sur le sol et l'exposition, séparés par des points-virgules",
  "plantationDetail": "2 à 4 conseils de plantation séparés par des points-virgules (période, espacement, profondeur…)",
  "tailleDetail": "2 à 4 conseils de taille séparés par des points-virgules (vide si aucune taille)",
  "engrais": "1 à 3 conseils de fertilisation séparés par des points-virgules",
  "bonSavoir": "1 à 3 astuces utiles séparées par des points-virgules",
  "toxicite": "vide si non toxique, sinon pour qui et quelles parties (ex. Toxique pour chats et chiens : feuilles et baies)",
  "multiplication": [{"methode": "semis" | "bouturage" | "division" | "marcottage" | "greffe", "debut": "mois en toutes lettres, ex. Mars", "fin": "mois en toutes lettres", "conseil": "phrase courte"}],
  "actions": [{"type": "taille" | "plantation" | "engrais" | "arrosage" | "semis" | "bouturage" | "division" | "autre", "mois": [numéros des mois concernés, 1 à 12], "conseil": "phrase courte"}]
}
"actions" est le calendrier d'entretien de l'année (3 à 6 tâches : taille, plantation, engrais, paillage, protection hivernale…). Tous les textes en français, concis.`;
}

function callGemini_(key, model, prompt) {
  const res = UrlFetchApp.fetch('https://generativelanguage.googleapis.com/v1beta/models/' + encodeURIComponent(model) + ':generateContent', {
    method: 'post', contentType: 'application/json', muteHttpExceptions: true,
    headers: { 'x-goog-api-key': key },
    payload: JSON.stringify({
      contents: [{ role: 'user', parts: [{ text: prompt }] }],
      generationConfig: { responseMimeType: 'application/json', temperature: 0.2 }
    })
  });
  const d = JSON.parse(res.getContentText() || '{}');
  if (res.getResponseCode() !== 200) throw new Error('Gemini : ' + ((d.error && d.error.message) || ('erreur ' + res.getResponseCode())));
  const parts = (((d.candidates || [])[0] || {}).content || {}).parts || [];
  const text = parts.map(x => x.text || '').join('');
  if (!text) throw new Error('Gemini : réponse vide');
  return text;
}

function callClaude_(key, model, prompt) {
  const res = UrlFetchApp.fetch('https://api.anthropic.com/v1/messages', {
    method: 'post', contentType: 'application/json', muteHttpExceptions: true,
    headers: { 'x-api-key': key, 'anthropic-version': '2023-06-01' },
    payload: JSON.stringify({ model, max_tokens: 2000, temperature: 0.2, messages: [{ role: 'user', content: prompt }] })
  });
  const d = JSON.parse(res.getContentText() || '{}');
  if (res.getResponseCode() !== 200) throw new Error('Claude : ' + ((d.error && d.error.message) || ('erreur ' + res.getResponseCode())));
  const text = (d.content || []).filter(c => c.type === 'text').map(c => c.text).join('');
  if (!text) throw new Error('Claude : réponse vide');
  return text;
}

/** Test depuis l'éditeur : choisis testerFiche puis « Exécuter », le résultat s'affiche dans le journal. */
function testerFiche() {
  Logger.log(JSON.stringify(fiche_('Rosier Pierre de Ronsard', ''), null, 2));
}

// ── Utilitaires ─────────────────────────────────────────────────────

function getFolder_() {
  const it = DriveApp.getFoldersByName(FOLDER_NAME);
  return it.hasNext() ? it.next() : DriveApp.createFolder(FOLDER_NAME);
}

function getPhotosSheet_() {
  const ss = SpreadsheetApp.openById(SHEET_ID);
  let sh = ss.getSheetByName(PHOTOS_SHEET);
  if (!sh) {
    sh = ss.insertSheet(PHOTOS_SHEET);
    sh.getRange(1, 1, sh.getMaxRows(), HEADERS.length).setNumberFormat('@'); // texte brut (dates en AAAA-MM-JJ)
    sh.appendRow(HEADERS);
    sh.setFrozenRows(1);
  }
  return sh;
}

function withLock_(fn) {
  const lock = LockService.getScriptLock();
  lock.waitLock(20000);
  try { return fn(); } finally { lock.releaseLock(); }
}

function today_() {
  return Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyy-MM-dd');
}

function json_(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj)).setMimeType(ContentService.MimeType.JSON);
}
