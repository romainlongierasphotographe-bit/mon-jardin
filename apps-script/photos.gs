/**
 * Mon Jardin — script des photos
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
