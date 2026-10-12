/**
 * LA MÉMOIRE — Rangement CRS (2e étape, après rattrapageCRS)
 * À mettre dans un NOUVEAU FICHIER du même projet Apps Script (le + à côté de « Fichiers »).
 * Fonction à exécuter : rangerCRS
 *
 * Pour chaque facture électronique CRS (La mémoire › CRS › À lire) :
 *   -> UN SEUL PDF dans le dossier de la job : la facture en premier, puis le ou les appels
 *      de service papier qui y correspondent. Ex. « Facture 2590 + appel 1249.pdf ».
 *   -> Facture sans appel papier trouvé : « Facture 2590.pdf », seule, dans le dossier de la job.
 *   -> La facture d'origine va dans La mémoire › CRS › « Originaux factures » (rien n'est effacé).
 * Appels de service et OT papier qui ne vont avec aucune facture envoyée :
 *   -> copie en PDF dans le dossier de la job. Ex. « Appel 1740 (papier).pdf », « OT 1253 (papier).pdf ».
 * Les photos originales restent dans le Drive de l'équipe : on fait des copies, on ne déplace rien.
 *
 * RÈGLE : aucune devinette. Un document est rangé seulement si son PO est certain :
 * écrit dans le nom du fichier (P0xxxx) ou trouvé par le numéro de facture dans le registre.
 * Sinon il est inscrit « À vérifier » au registre et n'est rangé nulle part.
 *
 * RANGER_ESSAI = true  -> écrit seulement dans le journal ce qu'il ferait
 * RANGER_ESSAI = false -> range pour vrai
 * Si Google arrête le script (« temps dépassé »), relance-le : il saute ce qui est déjà fait.
 */

const RANGER_ESSAI = false;

const R_DOSSIER_CRS  = '1WxSq4thx03QMpUeCEPfo4QcwHfBZX2Hq';            // La mémoire / CRS
const R_SHEET_FACT   = '1I6QfEObkpfrSAqz2vSO_XdCKABWYoe2e5TOG3so3Uks'; // La mémoire — Factures
const R_SHEET_APPELS = '1FJvH7oh82oRlEAyg2ddHOrJEDc-NGI6849m7StONcD8'; // La mémoire — Appels de service
const R_SOURCES = {                                                     // Drive partagé de l'équipe
  'A- REF FACTURE': '1KhCz_ouobsvzzSxezMdBtBdCCCpWIeAT',
  'B- REF OT': '1z7SN2qUHHtily5LiyE5QWwKtIdlCDbEI',
  '1-PAD SERVICE À TRAITER': '1CI-OuXIuiSKUksL7BUyGAN__tRPxbklB'
};
const R_PDFLIB = 'https://cdn.jsdelivr.net/npm/pdf-lib@1.17.1/dist/pdf-lib.min.js';

// pdf-lib (outil pour assembler des PDF) a besoin de setTimeout, qui n'existe pas dans Apps Script.
var setTimeout = function (f, t) { Utilities.sleep(t || 0); return f(); };

async function rangerCRS() {
  const debut = Date.now();
  const crs = DriveApp.getFolderById(R_DOSSIER_CRS);
  const stats = { fusions: 0, seules: 0, copies: 0, dejaFait: 0, aVerifier: 0, horsCRS: 0, erreurs: 0 };

  // --- Dossiers de jobs CRS : PO -> dossier ---
  const jobs = {};
  const it = crs.getFolders();
  while (it.hasNext()) { const f = it.next(); const m = f.getName().match(/(P0\d{4})\s*$/); if (m) jobs[m[1]] = f; }
  const aLire = sousDossier_(crs, 'À lire');
  const originaux = RANGER_ESSAI ? null : sousDossier_(crs, 'Originaux factures');

  // --- Registre des factures : numéro -> PO ---
  const facPO = {}, facNom = {}, facConnue = new Set();
  const fRows = SpreadsheetApp.openById(R_SHEET_FACT).getSheets()[0].getDataRange().getDisplayValues();
  const fh = fRows[0].map(h => h.toLowerCase());
  const iNum = fh.findIndex(h => h.startsWith('facture')), iPO = fh.findIndex(h => h === 'po'), iNom = fh.findIndex(h => h === 'nom');
  fRows.slice(1).forEach(r => {
    if (r[iNum]) facConnue.add(String(r[iNum]).trim());
    if (r[iNum] && r[iPO] && !/^AUCUN$/i.test(String(r[iPO]).trim())) { facPO[String(r[iNum]).trim()] = String(r[iPO]).trim(); facNom[String(r[iNum]).trim()] = String(r[iNom] || '').trim(); }
  });

  // --- Factures électroniques présentes dans À lire : numéro -> fichier ---
  const facFichier = {};
  const fl = aLire.getFiles();
  while (fl.hasNext()) { const f = fl.next(); const m = f.getName().match(/^Facture\s+(\d+)/i); if (m) facFichier[m[1]] = f; }

  // Factures déjà rangées lors d'une passe précédente (dans « Originaux factures ») : on les relit
  // pour ajouter un appel papier trouvé plus tard dans le PDF déjà fait.
  const facOrig = new Set();
  const dOrig = originaux || (crs.getFoldersByName('Originaux factures').hasNext() ? crs.getFoldersByName('Originaux factures').next() : null);
  if (dOrig) { const fo = dOrig.getFiles(); while (fo.hasNext()) { const f = fo.next(); const m = f.getName().match(/^Facture\s+(\d+)/i); if (m && !facFichier[m[1]]) { facFichier[m[1]] = f; facOrig.add(m[1]); } else if (m && !RANGER_ESSAI) f.setTrashed(true); /* doublon : la facture est aussi dans À lire */ } }

  // --- Appels de service / OT papier du Drive de l'équipe ---
  const papiers = [];
  Object.keys(R_SOURCES).forEach(src => {
    const files = DriveApp.getFolderById(R_SOURCES[src]).getFiles();
    while (files.hasNext()) {
      const file = files.next();
      const p = analyserNom_(file.getName());
      if (!p) continue;
      let po = p.po;
      if (!po) {
        const viaFac = p.factures.map(f => facPO[f]).filter(Boolean);
        if (viaFac.length && viaFac.every(x => x === viaFac[0])) po = viaFac[0];
      }
      // Un OT n'a pas toujours de PO : on le relie au job par le nom du client, seulement si un seul dossier correspond.
      if (!po && p.type === 'OT' && /(^|[^a-z])CRS([^a-z]|$)/i.test(file.getName())) po = parNom_(p.client, jobs);
      papiers.push(Object.assign(p, { file, src, poFinal: po, dest: po ? jobs[po] : null }));
    }
  });

  // Outil PDF chargé seulement si on range pour vrai.
  if (!RANGER_ESSAI) eval(UrlFetchApp.fetch(R_PDFLIB).getContentText());

  // --- 1) Une facture = un PDF (facture + ses appels papier) ---
  const papierUtilise = new Set();
  let arrete = false;
  for (const num of Object.keys(facFichier).sort()) {
    if (Date.now() - debut > 5 * 60 * 1000) { arrete = true; Logger.log('Presque 6 minutes : je m\'arrête ici. Relance le script pour continuer.'); break; }
    const po = facPO[num] || '';
    // Facture pas encore au registre : on attend, on ne devine pas son PO.
    if (!facConnue.has(num)) { stats.enAttente = (stats.enAttente || 0) + 1; Logger.log('En attente : facture ' + num + ' pas encore au registre des factures'); continue; }
    if (!po) {
      // Service CRS sans PO (ex. une borne) : rangé dans « Services sans PO » avec ses appels papier.
      if (!jobs['']) {
        const it2 = crs.getFoldersByName('Services sans PO');
        jobs[''] = it2.hasNext() ? it2.next() : RANGER_ESSAI ? { essai: true, getName: () => 'Services sans PO', getFilesByName: () => ({ hasNext: () => false }) } : crs.createFolder('Services sans PO');
      }
    } else if (!jobs[po]) {
      // PO certain (écrit sur la facture) mais pas encore de dossier : on le crée avec le nom de la facture.
      const nomJob = (facNom[num] || 'SANS NOM').split(/[,]/)[0].replace(/\b(EXTRA|GARAGE|APP\.?\s*\d+)\b.*$/i, '').trim().toUpperCase() + ' ' + po;
      Logger.log((RANGER_ESSAI ? '[essai] ' : '') + 'Nouveau dossier : ' + nomJob);
      jobs[po] = RANGER_ESSAI ? { essai: true, getName: () => nomJob, getFilesByName: () => ({ hasNext: () => false }) } : crs.createFolder(nomJob);
    }
    const dest = jobs[po];
    const aps = papiers.filter(p => p.type === 'AP' && p.factures.indexOf(num) >= 0 && p.poFinal === po);
    aps.forEach(p => papierUtilise.add(p.file.getId()));
    const nom = 'Facture ' + num + (aps.length ? ' + appel' + (aps.length > 1 ? 's ' : ' ') + aps.map(p => p.numero).join(', ') : '') + '.pdf';
    if (contenu_(dest).has(nom)) { stats.dejaFait++; continue; }
    // Anciens PDF de cette facture (faits sans tous les appels) : remplacés par le nouveau.
    const anciens = [];
    contenu_(dest).forEach((f, n) => { if (f && new RegExp('^Facture ' + num + '( \\+|\\.pdf)').test(n)) anciens.push(f); });
    if (anciens.length) Logger.log((RANGER_ESSAI ? '[essai] ' : '') + 'Remplace : ' + anciens.map(f => f.getName()).join(', '));
    Logger.log((RANGER_ESSAI ? '[essai] ' : '') + dest.getName() + '  <-  ' + nom);
    if (RANGER_ESSAI) { aps.length ? stats.fusions++ : stats.seules++; continue; }
    try {
      const facture = facFichier[num];
      if (aps.length) {
        const blob = await assembler_([facture].concat(aps.map(p => p.file)), nom);
        contenu_(dest).set(nom, dest.createFile(blob));
        stats.fusions++;
      } else {
        contenu_(dest).set(nom, facture.makeCopy(nom, dest));
        stats.seules++;
      }
      // La facture reste dans À lire (archive) : le rattrapage de nuit la reconnaît et ne la re-télécharge pas.
      anciens.forEach(f => { f.setTrashed(true); contenu_(dest).delete(f.getName()); });
    } catch (e) { stats.erreurs++; Logger.log('ERREUR facture ' + num + ' : ' + e); }
  }

  // --- 2) Appels / OT papier qui ne vont avec aucune facture envoyée : copie PDF seule ---
  for (const p of papiers) {
    if (papierUtilise.has(p.file.getId())) continue;
    // Le script s'est arrêté avant la fin des factures : on attend la prochaine passe pour les papiers.
    if (arrete) continue;
    if (!p.dest) {
      if (/(^|[^a-z])CRS([^a-z]|$)/i.test(p.file.getName()) || p.poFinal) { p.statut = p.factures.length && !p.factures.some(f => facPO[f]) ? 'En attente : facture ' + p.factures.join(', ') + ' pas encore au registre' : 'À vérifier : PO introuvable'; stats.aVerifier++; }
      else { p.statut = 'Hors CRS (pas encore rangé)'; stats.horsCRS++; }
      continue;
    }
    // Une facture de ce papier est déjà assemblée (envoyée plus tôt) ? Rien à faire de plus.
    if (p.factures.some(f => contenu_(p.dest).has('Facture ' + f + '.pdf'))) { p.statut = 'Rangé'; continue; }
    const nom = (p.type === 'AP' ? 'Appel ' : 'OT ') + p.numero + (p.factures.length ? ' — facture ' + p.factures.join(', ') : '') + ' (papier).pdf';
    p.statut = 'Rangé';
    if (contenu_(p.dest).has(nom)) { stats.dejaFait++; continue; }
    if (Date.now() - debut > 5 * 60 * 1000) { Logger.log('Presque 6 minutes : relance le script pour continuer.'); break; }
    Logger.log((RANGER_ESSAI ? '[essai] ' : '') + p.dest.getName() + '  <-  ' + nom);
    if (RANGER_ESSAI) { stats.copies++; continue; }
    try { contenu_(p.dest).set(nom, p.dest.createFile(await assembler_([p.file], nom))); stats.copies++; }
    catch (e) { stats.erreurs++; Logger.log('ERREUR ' + p.file.getName() + ' : ' + e); }
  }

  // --- 3) Registre « La mémoire — Appels de service » ---
  const lignes = papiers.map(p => [
    p.type === 'AP' ? p.numero : '', p.factures.join(', '), p.type === 'OT' ? p.numero : (p.ot || ''),
    p.poFinal || '', p.client, p.file.getName(), p.file.getId(), p.src, p.dest ? p.dest.getName() : '',
    p.statut || (papierUtilise.has(p.file.getId()) ? 'Rangé avec la facture' : '')
  ]).sort((a, b) => String(b[0] || b[2]).localeCompare(String(a[0] || a[2]), undefined, { numeric: true }));
  if (!RANGER_ESSAI && lignes.length) {
    const sh = SpreadsheetApp.openById(R_SHEET_APPELS).getSheets()[0];
    if (sh.getLastRow() > 1) sh.getRange(2, 1, sh.getLastRow() - 1, sh.getLastColumn()).clearContent();
    sh.getRange(2, 1, lignes.length, lignes[0].length).setValues(lignes);
  }

  Logger.log('---------------------------------');
  Logger.log((RANGER_ESSAI ? 'MODE ESSAI — rien créé. ' : '') +
    'Factures avec appel(s) papier : ' + stats.fusions + ' | Factures seules : ' + stats.seules +
    ' | Appels/OT papier copiés seuls : ' + stats.copies + ' | Déjà faits : ' + stats.dejaFait +
    ' | En attente : ' + (stats.enAttente || 0) + ' | À vérifier : ' + stats.aVerifier + ' | Hors CRS : ' + stats.horsCRS + ' | Erreurs : ' + stats.erreurs);
}

// Assemble des fichiers (PDF, JPG, PNG) en un seul PDF, dans l'ordre donné.
async function assembler_(fichiers, nom) {
  const doc = await PDFLib.PDFDocument.create();
  for (const f of fichiers) {
    let blob = f.getBlob();
    const type = blob.getContentType();
    if (type === 'application/pdf') {
      const src = await PDFLib.PDFDocument.load(new Uint8Array(blob.getBytes()), { ignoreEncryption: true });
      const pages = await doc.copyPages(src, src.getPageIndices());
      pages.forEach(pg => doc.addPage(pg));
      continue;
    }
    if (type !== 'image/jpeg' && type !== 'image/png') blob = blob.getAs('image/jpeg'); // ex. HEIC
    const bytes = new Uint8Array(blob.getBytes());
    const img = blob.getContentType() === 'image/png' ? await doc.embedPng(bytes) : await doc.embedJpg(bytes);
    // Page format lettre, photo centrée et réduite pour tenir dans la page.
    const W = 612, H = 792, marge = 18;
    const s = Math.min((W - 2 * marge) / img.width, (H - 2 * marge) / img.height, 1);
    const page = doc.addPage([W, H]);
    page.drawImage(img, { x: (W - img.width * s) / 2, y: (H - img.height * s) / 2, width: img.width * s, height: img.height * s });
  }
  const out = await doc.save();
  return Utilities.newBlob(Array.from(new Int8Array(out.buffer, out.byteOffset, out.byteLength)), 'application/pdf', nom);
}

// Contenu d'un dossier (nom -> fichier), lu une seule fois par passe : beaucoup plus rapide
// que de demander à Drive fichier par fichier.
const _contenu = {};
function contenu_(dossier) {
  if (dossier.essai || !dossier.getId) return new Map();
  const id = dossier.getId();
  if (!_contenu[id]) { const m = new Map(); const it = dossier.getFiles(); while (it.hasNext()) { const f = it.next(); m.set(f.getName(), f); } _contenu[id] = m; }
  return _contenu[id];
}

function sousDossier_(parent, nom) {
  const it = parent.getFoldersByName(nom);
  return it.hasNext() ? it.next() : parent.createFolder(nom);
}

// Trouve le PO d'un job par le nom du client (« Ot 1741 CRS Andriuk » -> ANDRIUK P03289).
// Rien si aucun ou plusieurs dossiers correspondent : on ne range jamais sur un doute.
function parNom_(client, jobs) {
  const sansAccent = x => x.normalize('NFD').replace(/[\u0300-\u036f]/g, '').toUpperCase();
  const vides = /^(CRS|SERV|SERVICE|ROUGH|FINITION|AJUSTEMENT|AJUSTEMENTS|EXTRA|BORNE|CUISINE|SALLE|BAIN|TEMPS|REEL|APP)$/;
  const mots = sansAccent(client).split(/[^A-Z]+/).filter(m => m.length >= 4 && !vides.test(m));
  const trouves = new Set();
  Object.keys(jobs).forEach(po => {
    if (!po) return;
    const nomJob = ' ' + sansAccent(jobs[po].getName()).replace(/P0\d{4}\s*$/, '').replace(/[^A-Z]+/g, ' ') + ' ';
    if (mots.some(m => nomJob.indexOf(' ' + m + ' ') >= 0)) trouves.add(po);
  });
  return trouves.size === 1 ? [...trouves][0] : '';
}

// Lit un nom comme « AP_1249_CRS ROUGH LYNCH AU P03192_F-2590.jpg », « Ap 1740 CRS Andrick_F-2627 »,
// « OT_1253_CRS ROUGH POULIN_580 P03213 AU.jpg » ou « AP_741_PAY-RIVERMERE_FACT_2300.HEIC ».
function analyserNom_(nom) {
  const t = nom.match(/^\s*(a\s*p|o\s*t)[\s_\-]*([A-Z]{0,2}\d+)/i);
  if (!t) return null;
  const type = t[1].replace(/\s/g, '').toUpperCase();
  const reste = nom.slice(t[0].length);
  const po = (reste.match(/(?:^|[^A-Za-z0-9])P[O0]?0?(\d{4})(?![0-9])/i) || [])[1];
  const factures = [];
  const reF = /(?:^|[^A-Za-z])(?:F|FACT)[\s_\-]*(\d{4})((?:-\d{4})*)/gi;
  let m;
  while ((m = reF.exec(reste))) { factures.push(m[1]); (m[2].match(/\d{4}/g) || []).forEach(x => factures.push(x)); }
  // Collé au nom : « CRS aubinF-2469 »
  const reF2 = /[a-z’']F[\s_\-]+(\d{4})/g;
  while ((m = reF2.exec(reste))) if (factures.indexOf(m[1]) < 0) factures.push(m[1]);
  const ot = (reste.match(/OT[\s_\-]*(\d{3,4})/i) || [])[1] || '';
  const client = reste.replace(/\.[a-z0-9]{3,4}$/i, '').replace(/(?:^|[_\s])(?:F|FACT)[\s_\-]*\d{4}(?:-\d{4})*.*$/i, '')
    .replace(/[_]+/g, ' ').replace(/\s+/g, ' ').trim();
  return { type, numero: t[2], po: po ? 'P0' + po : '', factures, ot, client };
}
