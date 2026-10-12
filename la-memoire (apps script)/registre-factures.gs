/**
 * LA MÉMOIRE — Registre des factures (automatique)
 *
 * Chaque facture envoyée depuis info@ (« Facture #2644 (07-10-2026) : Saran Électrique Inc. »)
 * est lue (le PDF joint) et inscrite dans « La mémoire — Factures » :
 *   numéro, date, client, PO, nom du job, description, montant avant taxes, total, date d'envoi, PDF.
 * Le CRM (saranelectrique.net) lit cette feuille : la facture y apparaît avec son prix.
 *
 * Règle « pas d'erreur » : si le PDF ne se lit pas clairement (montants introuvables),
 * la facture n'est PAS inscrite. Elle est notée dans le journal et réessayée à la prochaine passe.
 * Une facture lue sans PO est inscrite avec PO = « AUCUN » (service sans PO).
 *
 * À ajouter comme nouveau fichier « registre » dans le même projet Apps Script.
 * Fonctions :
 *   registreFactures()  -> une passe (essai d'abord avec REG_ESSAI = true)
 *   installerNuit()     -> à lancer UNE fois : tout roule seul chaque nuit
 */

const REG_ESSAI = false;
const REG_DEPUIS = '2026/09/24';                                 // la 2601 est la dernière inscrite à la main
const REG_SHEET_FACT = '1I6QfEObkpfrSAqz2vSO_XdCKABWYoe2e5TOG3so3Uks';
const REG_DOSSIER_MEMOIRE = '1U0syTc5wf_u-0R9Q423s8ykXiou9SWAP';   // La mémoire
const REG_MOI = 'info@saranelectrique.com';

function registreFactures() {
  const debut = Date.now();
  try { bonsDeCommandeOdoo(); } catch (e) { Logger.log('Bons de commande Odoo : ' + e); }
  const sh = SpreadsheetApp.openById(REG_SHEET_FACT).getSheets()[0];
  completerLiensPdf_(sh);
  const deja = new Set(sh.getRange(1, 1, Math.max(sh.getLastRow(), 1), 1).getDisplayValues().map(r => String(r[0]).trim()));
  const autres = regSousDossier_(DriveApp.getFolderById(REG_DOSSIER_MEMOIRE), 'Factures autres clients');
  const stats = { inscrites: 0, deja: 0, illisibles: 0, ignorees: 0 };

  // Messages envoyés « Facture #N (jj-mm-aaaa) », le plus ancien d'abord (garde la 1re date d'envoi).
  const msgs = [];
  for (let i = 0; ; i += 100) {
    const lot = GmailApp.search('from:' + REG_MOI + ' subject:Facture has:attachment after:' + REG_DEPUIS, i, 100);
    lot.forEach(t => t.getMessages().forEach(m => msgs.push(m)));
    if (lot.length < 100) break;
  }
  msgs.sort((a, b) => a.getDate() - b.getDate());

  const lignes = [], articles = [];
  for (const msg of msgs) {
    if (Date.now() - debut > 5 * 60 * 1000) { Logger.log('Presque 6 minutes : relance pour continuer.'); break; }
    const s = msg.getSubject().match(/^Facture\s*#(\d+)\s*\((\d{2})-(\d{2})-(\d{4})\)/i);
    if (!s || msg.getFrom().indexOf(REG_MOI) < 0) continue;
    const num = s[1];
    if (deja.has(num)) { stats.deja++; continue; }
    const dest = (msg.getTo() + ',' + msg.getCc()).split(',').map(x => (x.match(/[\w.+-]+@[\w.-]+/) || [''])[0].toLowerCase()).filter(x => x && x !== REG_MOI);
    if (!dest.length) { stats.ignorees++; continue; }          // envoyée seulement à moi-même (essai)
    const pj = msg.getAttachments({ includeInlineImages: false }).filter(a => /\.pdf$/i.test(a.getName()) || a.getContentType() === 'application/pdf')[0];
    if (!pj) continue;

    const txt = lirePdf_(pj.copyBlob());
    const lu = analyserFacture_(txt);
    if (!lu) { stats.illisibles++; Logger.log('Illisible, réessai plus tard : facture ' + num + ' | début du texte lu : ' + txt.replace(/\s+/g, ' ').slice(0, 150)); continue; }

    const crs = dest.indexOf('facturation.crs@dext.cc') >= 0;
    let pdfId = '';
    if (!REG_ESSAI) {
      // CRS : le PDF est déjà sauvé dans CRS/À lire par rattrapageCRS. Les autres : on le garde ici.
      const nomPdf = 'Facture ' + num + ' (' + s[2] + '-' + s[3] + '-' + s[4] + ').pdf';
      if (crs) {
        const f = trouverFichier_(nomPdf);
        pdfId = f ? f.getId() : '';
      } else {
        const it = autres.getFilesByName(nomPdf);
        pdfId = (it.hasNext() ? it.next() : autres.createFile(pj.copyBlob()).setName(nomPdf)).getId();
      }
    }
    const ligne = [num, s[4] + '-' + s[3] + '-' + s[2], crs ? 'CRS' : dest[0], lu.po || 'AUCUN', lu.nom, lu.desc,
      lu.avantTaxes, lu.total, Utilities.formatDate(msg.getDate(), 'America/Toronto', 'yyyy-MM-dd'), pdfId];
    lignes.push(ligne);
    facLignes_(txt).forEach(a => articles.push([num].concat(a)));
    deja.add(num);
    stats.inscrites++;
    Logger.log((REG_ESSAI ? '[essai] ' : '') + 'Facture ' + num + ' | ' + ligne[2] + ' | ' + ligne[3] + ' | ' + lu.nom + ' | ' + lu.desc + ' | ' + lu.avantTaxes + ' $ avant taxes');
  }

  if (!REG_ESSAI && lignes.length) sh.getRange(sh.getLastRow() + 1, 1, lignes.length, lignes[0].length).setValues(lignes);
  if (!REG_ESSAI && articles.length) { const o = facOngletLignes_(); o.getRange(o.getLastRow() + 1, 1, articles.length, 6).setValues(articles); }
  Logger.log('---------------------------------');
  Logger.log((REG_ESSAI ? 'MODE ESSAI — rien inscrit. ' : '') + 'Inscrites : ' + stats.inscrites + ' | Déjà au registre : ' + stats.deja +
    ' | Illisibles (réessai) : ' + stats.illisibles + ' | Envoyées à moi seulement : ' + stats.ignorees);
}

// Lit le texte d'un PDF : Google Drive le convertit en Doc le temps de le lire, puis le Doc est jeté.
function lirePdf_(blob) {
  const jeton = ScriptApp.getOAuthToken();
  const limite = '-------lamemoire' + Date.now();
  const meta = JSON.stringify({ name: 'tmp lecture facture', mimeType: 'application/vnd.google-apps.document' });
  const corps = Utilities.newBlob(
    '--' + limite + '\r\nContent-Type: application/json; charset=UTF-8\r\n\r\n' + meta +
    '\r\n--' + limite + '\r\nContent-Type: application/pdf\r\n\r\n').getBytes()
    .concat(blob.getBytes())
    .concat(Utilities.newBlob('\r\n--' + limite + '--').getBytes());
  let rep;
  for (let essai = 0; essai < 3; essai++) {          // Google répond parfois « Internal Error » : on réessaie
    rep = UrlFetchApp.fetch('https://www.googleapis.com/upload/drive/v3/files?uploadType=multipart&ocrLanguage=fr', {
      method: 'post', contentType: 'multipart/related; boundary=' + limite, payload: corps,
      headers: { Authorization: 'Bearer ' + jeton }, muteHttpExceptions: true });
    if (rep.getResponseCode() < 500) break;
    Utilities.sleep(2000);
  }
  if (rep.getResponseCode() !== 200) { Logger.log('Lecture PDF impossible : ' + rep.getContentText().slice(0, 200)); return ''; }
  const id = JSON.parse(rep.getContentText()).id;
  const texte = UrlFetchApp.fetch('https://www.googleapis.com/drive/v3/files/' + id + '/export?mimeType=text/plain',
    { headers: { Authorization: 'Bearer ' + jeton } }).getContentText();
  DriveApp.getFileById(id).setTrashed(true);
  return texte;
}

// Tire les infos du texte de la facture. Rien (null) si les montants ne sont pas clairs.
// Les montants sont trouvés sans dépendre de la mise en page : on cherche un sous-total S
// pour lequel TPS = 5 % de S, TVQ = 9,975 % de S et S + TPS + TVQ = un total présent sur la facture.
function analyserFacture_(txt) {
  if (!txt) return null;
  const plat = txt.replace(/ /g, ' ').replace(/[ \t]+/g, ' ');
  const argent = (plat.match(/\d{1,3}(?: \d{3})+,\d{2}|\d+,\d{2}/g) || []).map(x => Number(x.replace(/ /g, '').replace(',', '.')));
  const present = v => argent.some(a => Math.abs(a - v) < 0.015);
  let meilleur = null;
  argent.forEach(st => {
    const tps = Math.round(st * 5) / 100, tvq = Math.round(st * 9.975) / 100, tot = st + tps + tvq;
    const ok = present(tps) && present(tvq) && (argent.some(a => Math.abs(a - tot) <= 0.02));
    if (ok && (!meilleur || st > meilleur.st)) meilleur = { st, tot: argent.find(a => Math.abs(a - tot) <= 0.02) };
  });
  // Facture sans taxes (ex. entre compagnies) : « MONTANT DÛ  vente  retenue  sous-total  0,00  0,00  total » avec sous-total = total.
  if (!meilleur) {
    const iM = plat.search(/MONTANT\s+D[ÛU]/i);
    const n = iM < 0 ? [] : (plat.slice(iM).match(/\d{1,3}(?: \d{3})+,\d{2}|\d+,\d{2}/g) || []).map(x => Number(x.replace(/ /g, '').replace(',', '.')));
    if (n.length >= 6 && n[3] === 0 && n[4] === 0 && Math.abs(n[2] - n[5]) < 0.015 && n[2] > 0) meilleur = { st: n[2], tot: n[5] };
  }
  if (!meilleur || meilleur.st <= 0) return null;

  const po = (plat.match(/\bP[O0]\s?-?\s?0?(\d{4})\b/i) || [])[1];
  const lignes = plat.split(/\n/).map(l => l.trim()).filter(Boolean);
  const iPage = lignes.findIndex(l => /Page\s+\d+\s+de\s+\d+/i.test(l));
  const nom = iPage >= 0 && lignes[iPage + 1] && !/Termes|cr[ée]dit|^\d/i.test(lignes[iPage + 1]) ? lignes[iPage + 1].slice(0, 60) : '';
  const iProd = lignes.findIndex(l => /Produit\s+Description/i.test(l));
  const desc = iProd > 0 && !/facturation|@|T[ée]l|Taschereau|No de contrat/i.test(lignes[iProd - 1]) ? lignes[iProd - 1].slice(0, 60) : '';
  return { po: po ? 'P0' + po : '', nom, desc, avantTaxes: meilleur.st.toFixed(2), total: meilleur.tot.toFixed(2) };
}

// Pour me montrer le texte lu d'une facture : change le numéro, lance, envoie-moi le journal.
function voirTexteFacture() {
  const num = '2633';
  const t = GmailApp.search('from:' + REG_MOI + ' subject:"Facture #' + num + '"', 0, 1)[0];
  const pj = t.getMessages()[0].getAttachments({ includeInlineImages: false })[0];
  Logger.log(lirePdf_(pj.copyBlob()).slice(0, 3000));
}

function trouverFichier_(nom) {
  const it = DriveApp.getFilesByName(nom);
  return it.hasNext() ? it.next() : null;
}

function regSousDossier_(parent, nom) {
  const it = parent.getFoldersByName(nom);
  return it.hasNext() ? it.next() : (REG_ESSAI ? parent : parent.createFolder(nom));
}

// À lancer UNE fois (on peut le relancer : il remplace les anciens horaires).
// Deux fois par jour, à midi et à minuit, avec une heure entre chaque étape :
//   12 h et 0 h  -> rattrapageCRS (nouveaux PO, plans, factures CRS)
//   13 h et 1 h  -> registreFactures (numéro, PO, montants de toutes les factures envoyées)
//   14 h et 2 h  -> rangerCRS (une facture = un PDF avec ses appels papier, dans le dossier du job)
function installerNuit() {
  const etapes = [['rattrapageCRS', 0], ['registreFactures', 1], ['rangerCRS', 2], ['agendaCRS', 3], ['courrielsJobs', 3], ['achatsFournisseurs', 1], ['adressesManquantes', 2]];
  ScriptApp.getProjectTriggers().filter(t => etapes.some(e => e[0] === t.getHandlerFunction())).forEach(t => ScriptApp.deleteTrigger(t));
  etapes.forEach(([f, decalage]) => [0, 12].forEach(h =>
    ScriptApp.newTrigger(f).timeBased().everyDays(1).atHour(h + decalage).inTimezone('America/Toronto').create()));
  Logger.log('Installé : midi et minuit — rattrapage (12 h / 0 h), registre (13 h / 1 h), rangement (14 h / 2 h), agenda et courriels (15 h / 3 h), achats fournisseurs (13 h / 1 h), adresses manquantes (14 h / 2 h).');
}

// MISE À JOUR TOUT DE SUITE : un clic, et les 3 étapes roulent une après l'autre
// (rattrapage maintenant, registre dans 5 min, rangement dans 10 min). Tout est à jour en ~15 minutes.
function miseAJourMaintenant() {
  const props = PropertiesService.getScriptProperties();
  const vieux = (props.getProperty('majIds') || '').split(',');
  ScriptApp.getProjectTriggers().filter(t => vieux.indexOf(t.getUniqueId()) >= 0).forEach(t => ScriptApp.deleteTrigger(t));
  const ids = [['rattrapageCRS', 1], ['registreFactures', 5], ['rangerCRS', 10], ['agendaCRS', 14], ['courrielsJobs', 16], ['achatsFournisseurs', 7], ['adressesManquantes', 12]]
    .map(([f, min]) => ScriptApp.newTrigger(f).timeBased().after(min * 60 * 1000).create().getUniqueId());
  props.setProperty('majIds', ids.join(','));
  Logger.log('Mise à jour lancée : rattrapage dans 1 min, registre dans 5 min, rangement dans 10 min, agenda dans 14 min, courriels dans 16 min. Tout sera à jour dans environ 20 minutes.');
}

// Page web privée (bouton « Mettre à jour » du CRM). Déployée en application Web, accès : moi uniquement.
function doGet(e) {
  // « ?quoi=adresses » (bouton « Chercher les adresses manquantes » du CRM) : seulement la recherche d'adresses, dans 1 minute.
  if (e && e.parameter && e.parameter.quoi === 'adresses') {
    ScriptApp.newTrigger('adressesManquantes').timeBased().after(60 * 1000).create();
    return HtmlService.createHtmlOutput(
      '<div style="font-family:system-ui,sans-serif;max-width:420px;margin:15vh auto;text-align:center;padding:24px">' +
      '<div style="font-size:42px">🔎</div><h2 style="margin:8px 0">Recherche d\'adresses lancée</h2>' +
      '<p style="color:#555;line-height:1.5">Je cherche l\'adresse des jobs qui n\'en ont pas, dans le registre et dans tes courriels.<br>Les propositions seront dans le CRM (Tableau de bord) dans environ <b>5 minutes</b>.</p>' +
      '<p style="color:#888;font-size:14px">Tu peux fermer cet onglet.</p></div>')
      .setTitle('La mémoire — adresses').addMetaTag('viewport', 'width=device-width, initial-scale=1');
  }
  miseAJourMaintenant();
  const heure = Utilities.formatDate(new Date(Date.now() + 20 * 60 * 1000), 'America/Toronto', 'HH:mm');
  return HtmlService.createHtmlOutput(
    '<div style="font-family:system-ui,sans-serif;max-width:420px;margin:15vh auto;text-align:center;padding:24px">' +
    '<div style="font-size:42px">✅</div><h2 style="margin:8px 0">Mise à jour lancée</h2>' +
    '<p style="color:#555;line-height:1.5">Nouveaux PO et factures, registre, rangement, documents dans l\'agenda, puis fils de courriels.<br>Tout sera à jour vers <b>' + heure + '</b>.</p>' +
    '<p style="color:#888;font-size:14px">Tu peux fermer cet onglet.</p></div>')
    .setTitle('La mémoire — mise à jour')
    .addMetaTag('viewport', 'width=device-width, initial-scale=1');
}


// Complète le lien PDF des factures CRS inscrites avant que leur PDF arrive dans CRS/À lire.
function completerLiensPdf_(sh) {
  if (REG_ESSAI || sh.getLastRow() < 2) return;
  const vals = sh.getRange(2, 1, sh.getLastRow() - 1, 10).getDisplayValues();
  let n = 0;
  vals.forEach((r, i) => {
    if (r[2] !== 'CRS' || r[9]) return;
    const d = r[1].split('-');                       // aaaa-mm-jj -> « Facture N (jj-mm-aaaa).pdf »
    const f = trouverFichier_('Facture ' + r[0] + ' (' + d[2] + '-' + d[1] + '-' + d[0] + ').pdf');
    if (f) { sh.getRange(i + 2, 10).setValue(f.getId()); n++; }
  });
  if (n) Logger.log('Liens PDF complétés : ' + n);
}

// Bons de commande que Jonathan n'a pas joints au courriel : le PDF est téléchargé depuis le lien Odoo du courriel
// et sauvé dans le dossier du job (« Bon de commande - P0xxxx.pdf »). Seulement si le dossier n'en a pas déjà un.
function bonsDeCommandeOdoo() {
  const crs = DriveApp.getFolderById('1WxSq4thx03QMpUeCEPfo4QcwHfBZX2Hq');
  const jobs = {};
  const it = crs.getFolders();
  while (it.hasNext()) { const f = it.next(); const m = f.getName().match(/(P0\d{4})\s*$/); if (m) jobs[m[1]] = f; }
  const faits = new Set();
  GmailApp.search('from:notifications@cuisinesrivesud17.odoo.com after:2026/06/01', 0, 200).forEach(t => t.getMessages().forEach(msg => {
    const m = msg.getSubject().match(/P0(\d{4})/);
    if (!m) return;
    const po = 'P0' + m[1], dossier = jobs[po];
    if (!dossier || faits.has(po)) return;
    const fs = dossier.getFiles();
    while (fs.hasNext()) if (/^Bon de commande/i.test(fs.next().getName())) { faits.add(po); return; }
    const l = msg.getBody().replace(/&amp;/g, '&').match(/(https:\/\/[\w.-]+odoo\.com)\/my\/purchase\/(\d+)\?access_token=([\w-]+)/);
    if (!l) return;
    const url = l[1] + '/my/purchase/' + l[2] + '?access_token=' + l[3] + '&report_type=pdf&download=true';
    const rep = UrlFetchApp.fetch(url, { muteHttpExceptions: true, followRedirects: true });
    const type = String(rep.getHeaders()['Content-Type'] || '');
    if (rep.getResponseCode() === 200 && /pdf/i.test(type)) {
      if (!REG_ESSAI) dossier.createFile(rep.getBlob().setName('Bon de commande - ' + po + '.pdf'));
      faits.add(po);
      Logger.log((REG_ESSAI ? '[essai] ' : '') + 'Bon de commande récupéré d\'Odoo : ' + dossier.getName());
    } else {
      Logger.log('Bon de commande Odoo non disponible pour ' + po + ' (réponse ' + rep.getResponseCode() + ')');
    }
  }));
}


// AGENDA : chaque événement CRS à venir (« CRS ROUGH LEMAY P03445 ») reçoit en pièces jointes les documents
// utiles à l'équipe qui sont dans le dossier du job : plans et OT papier.
// JAMAIS le bon de commande ni les factures (il y a des prix). Rien n'est enlevé, seulement ajouté.
// Les fichiers ajoutés sont partagés en lecture avec le compte de l'équipe (saranelectrique@gmail.com).
const AG_CALENDRIERS = ['saranelectrique@gmail.com', 'c_7d3dd0732b25a3afd30d0f2c453bef2d31a41b08a6d7112b08accea16ff49daf@group.calendar.google.com'];
const AG_EQUIPE = 'saranelectrique@gmail.com';
const AG_JOURS = 120;

// Seuls ces documents vont à l'équipe : plans et OT papier. Tout le reste (PO, bon de commande, factures,
// appels avec prix, soumissions) reste privé. Liste « permise » plutôt que « défendue » : rien ne passe par erreur.
function agPermis_(nom) { return /plan/i.test(nom) || /^OT\s*\d+.*\(papier\)/i.test(nom) ; }

// Partage silencieux (aucun courriel envoyé) avec le compte de l'équipe.
function agPartager_(fileId) {
  const rep = UrlFetchApp.fetch('https://www.googleapis.com/drive/v3/files/' + fileId + '/permissions?sendNotificationEmail=false', {
    method: 'post', contentType: 'application/json', payload: JSON.stringify({ role: 'reader', type: 'user', emailAddress: AG_EQUIPE }),
    headers: { Authorization: 'Bearer ' + ScriptApp.getOAuthToken() }, muteHttpExceptions: true });
  return rep.getResponseCode() === 200;
}

function agendaCRS() {
  const crs = DriveApp.getFolderById('1WxSq4thx03QMpUeCEPfo4QcwHfBZX2Hq');
  const jobs = {};
  const it = crs.getFolders();
  while (it.hasNext()) { const f = it.next(); const m = f.getName().match(/(P0\d{4})\s*$/); if (m) jobs[m[1]] = f; }
  const cache = {};
  const contenu = po => {
    if (!cache[po]) {
      const ok = [], prives = new Set(), fs = jobs[po].getFiles();
      while (fs.hasNext()) {
        const f = fs.next(), n = f.getName();
        if (/Tout le dossier/i.test(n)) continue;          // PDF fusionné par ton autre script : on n'y touche pas
        if (/^image\//.test(f.getMimeType())) continue;   // photos : jamais retirées (ni ajoutées automatiquement)
        if (agPermis_(n)) ok.push(f); else prives.add(f.getId());
      }
      cache[po] = { ok: ok.sort((a, b) => a.getName().localeCompare(b.getName())), prives };
    }
    return cache[po];
  };
  const debut = new Date(), fin = new Date(Date.now() + AG_JOURS * 864e5);
  let ajouts = 0, retires = 0, evts = 0;
  const aRetirerPartage = new Set();
  // Ce qu'on a déjà ajouté à chaque événement : ton autre script fusionne les pièces jointes en
  // « 📎 Tout le dossier (PDF) » et enlève les originaux — on ne les rajoute donc pas une 2e fois.
  const props = PropertiesService.getScriptProperties();
  const memo = JSON.parse(props.getProperty('agendaAjouts') || '{}'), vus = {};
  AG_CALENDRIERS.forEach(calId => {
    const cal = CalendarApp.getCalendarById(calId);
    if (!cal) { Logger.log('Agenda introuvable : ' + calId); return; }
    cal.getEvents(debut, fin).forEach(ev => {
      const m = ev.getTitle().match(/\bCRS\b.*\b(P0\d{4})\b/i);
      if (!m || !jobs[m[1].toUpperCase()]) return;
      const c = contenu(m[1].toUpperCase());
      const id = ev.getId().replace(/@google\.com$/, '');
      let pj;
      try { pj = Calendar.Events.get(calId, id, { fields: 'attachments' }).attachments || []; }
      catch (e) { Logger.log('Lecture impossible : ' + ev.getTitle() + ' — ' + String(e).slice(0, 160)); return; }
      evts++;
      const idDe = a => a.fileId || ((a.fileUrl || '').match(/\/d\/([\w-]+)/) || [])[1];
      // 1) Retire les documents avec des prix qu'un passage précédent aurait ajoutés.
      const gardees = pj.filter(a => !c.prives.has(idDe(a)));
      const enleves = pj.filter(a => c.prives.has(idDe(a)));
      enleves.forEach(a => aRetirerPartage.add(idDe(a)));
      // 2) Ajoute les plans et OT manquants.
      const deja = new Set(gardees.map(idDe).concat(gardees.map(a => a.title)));
      // Première fois : si le dossier fusionné existe déjà, on considère ses documents actuels comme déjà ajoutés.
      if (!memo[id] && gardees.some(x => /Tout le dossier/i.test(x.title || ''))) memo[id] = c.ok.map(f => f.getId());
      const cle = id, dejaAjoutes = new Set(memo[cle] || []);
      vus[cle] = memo[cle] || [];
      const nouveaux = c.ok.filter(f => !deja.has(f.getId()) && !deja.has(f.getName()) && !dejaAjoutes.has(f.getId()));
      if (!nouveaux.length && !enleves.length) return;
      nouveaux.forEach(f => agPartager_(f.getId()));
      const toutes = gardees.concat(nouveaux.map(f => ({ fileUrl: 'https://drive.google.com/file/d/' + f.getId() + '/view', fileId: f.getId(), title: f.getName(), mimeType: f.getMimeType() })));
      try {
        Calendar.Events.patch({ attachments: toutes.slice(0, 25) }, calId, id, { supportsAttachments: true });
        ajouts += nouveaux.length; retires += enleves.length;
        vus[cle] = [...dejaAjoutes, ...nouveaux.map(f => f.getId())];
        const quand = Utilities.formatDate(ev.getStartTime(), 'America/Toronto', 'yyyy-MM-dd');
        if (enleves.length) Logger.log(ev.getTitle() + ' (' + quand + ')  RETIRÉ (prix) : ' + enleves.map(a => a.title).join(', '));
        if (nouveaux.length) Logger.log(ev.getTitle() + ' (' + quand + ')  <-  ' + nouveaux.map(f => f.getName()).join(', '));
      } catch (e) { Logger.log('ERREUR agenda ' + ev.getTitle() + ' : ' + String(e).slice(0, 200)); }
    });
  });
  props.setProperty('agendaAjouts', JSON.stringify(vus));   // seulement les événements à venir
  // 3) Les documents avec des prix ne sont plus partagés avec le compte de l'équipe.
  Object.values(cache).forEach(c => c.prives.forEach(fid => { try { DriveApp.getFileById(fid).removeViewer(AG_EQUIPE); } catch (e) {} }));
  Logger.log('---------------------------------');
  Logger.log('Événements CRS vérifiés : ' + evts + ' | Documents ajoutés : ' + ajouts + ' | Documents avec prix retirés : ' + retires);
}


// COURRIELS : pour chaque job, les fils de courriels qui la concernent (numéro de PO, ou courriel du client),
// avec les dates, qui a écrit en dernier et ce qui se passe. Écrit dans l'onglet « Courriels » du registre des jobs,
// que le CRM affiche dans la fiche de chaque job.
const CO_SHEET_JOBS = '1POqSf9_TAJCr8cYsV2QEGYzcPdLOeIGeWnIX8Y_swBg';
const CO_MOI = /saranelectrique\.com|saranelectrique@gmail\.com|eclairagesaran\.com/i;
const CO_AUTO = /notifications@|noreply|no-reply|dext\.cc|mailer-daemon/i;

function courrielsJobs() {
  const debut = Date.now();
  const ss = SpreadsheetApp.openById(CO_SHEET_JOBS);
  // Jobs : registre (PO + courriel du client) + dossiers CRS de La mémoire.
  const jobs = {};
  const reg = ss.getSheets()[0].getDataRange().getDisplayValues();
  const h = reg[0].map(x => x.toLowerCase());
  const iPO = h.indexOf('po'), iNom = h.indexOf('nom'), iMail = h.findIndex(x => x.startsWith('courriel'));
  reg.slice(1).forEach(r => { const po = (r[iPO] || '').trim(); if (/^P0\d{4}$|^[A-Z0-9]{2,5}-\d{3}$/.test(po)) jobs[po] = { nom: r[iNom], mail: (r[iMail] || '').trim().toLowerCase() }; });
  const it = DriveApp.getFolderById('1WxSq4thx03QMpUeCEPfo4QcwHfBZX2Hq').getFolders();
  while (it.hasNext()) { const m = it.next().getName().match(/^(.*?)\s*(P0\d{4})\s*$/); if (m && !jobs[m[2]]) jobs[m[2]] = { nom: m[1], mail: '' }; }

  const lignes = [];
  for (const po of Object.keys(jobs).sort().reverse()) {
    if (Date.now() - debut > 5 * 60 * 1000) { Logger.log('Presque 6 minutes : arrêt, la suite à la prochaine passe.'); break; }
    const j = jobs[po], n = po.slice(3);
    let q = '(' + (/^P0/.test(po) ? po + ' OR PO' + n : '"' + po + '"') + (j.mail ? ' OR from:' + j.mail + ' OR to:' + j.mail : '') + ') newer_than:240d -in:spam -in:trash';
    GmailApp.search(q, 0, 8).forEach(t => {
      const msgs = t.getMessages();
      const interne = m => CO_MOI.test(m.getFrom()) || /via Google Drive|drive-shares|calendar-notification/i.test(m.getFrom());
      if (msgs.every(interne)) return;                       // courriels internes (horaires, punch, partages Drive) : ignorés
      const humains = msgs.filter(m => !CO_AUTO.test(m.getFrom()) && !/via Google Drive|drive-shares/i.test(m.getFrom()));
      const dernier = (humains.length ? humains : msgs)[(humains.length ? humains : msgs).length - 1];
      const de = dernier.getFrom(), deMoi = CO_MOI.test(de);
      const corps = dernier.getPlainBody().replace(/\r/g, '').split(/\n(?:Le .{5,80}a écrit|On .{5,80}wrote|-{2,}|De ?:|From:)/)[0];
      const question = /\?/.test(corps);
      const statut = !humains.length ? 'Info (automatique)'
        : deMoi ? (question ? 'En attente de leur réponse' : 'Répondu')
        : (question ? 'À toi de répondre' : 'Reçu');
      const nomDe = de.replace(/<.*>/, '').replace(/"/g, '').trim() || de;
      const extrait = corps.replace(/\s+/g, ' ').trim().slice(0, 220);
      lignes.push([po, j.nom, t.getFirstMessageSubject(), Utilities.formatDate(msgs[0].getDate(), 'America/Toronto', 'yyyy-MM-dd'),
        Utilities.formatDate(dernier.getDate(), 'America/Toronto', 'yyyy-MM-dd'), msgs.length, nomDe, statut, extrait, t.getPermalink()]);
    });
  }
  let sh = ss.getSheetByName('Courriels');
  if (!sh) sh = ss.insertSheet('Courriels');
  sh.clearContents();
  const entete = ['PO', 'Job', 'Sujet', 'Premier', 'Dernier', 'Messages', 'Dernier message de', 'Statut', 'Extrait', 'Lien'];
  sh.getRange(1, 1, 1, entete.length).setValues([entete]);
  if (lignes.length) sh.getRange(2, 1, lignes.length, entete.length).setValues(lignes);
  Logger.log('Fils de courriels inscrits : ' + lignes.length + ' (' + lignes.filter(l => l[7] === 'À toi de répondre').length + ' à répondre)');
}

// ARTICLES DE TES FACTURES : « CODE description QTÉ UNITÉ PRIX $TOTAL » (ex. « NMD2C14X150 NMD2C14 75 ME 3,82 $286,50 »).
// Rend [codes, description, qté, prix, montant]. La main-d'œuvre (SA01 Électriciens…) est marquée « MO ».
function facLignes_(txt) {
  const t = String(txt || '').replace(/ /g, ' ');
  const i = t.search(/Produit\s+Description\s+Quantit/i);
  if (i < 0) return [];
  const fin = t.slice(i).search(/Des frais d'administration|No TPS|VENTE\s+MONTANT/i);
  const zone = t.slice(i, fin > 0 ? i + fin : undefined).replace(/Produit\s+Description\s+Quantit\S*\s+Prix\s+Total/i, ' ').replace(/\s+/g, ' ');
  const argent = /(\d{1,3}(?: \d{3})*,\d{2,4})\s*\$\s?(\d{1,3}(?: \d{3})*,\d{2})/g;
  const num = x => Number(String(x).replace(/ /g, '').replace(',', '.'));
  const out = []; let m, avant = 0;
  while ((m = argent.exec(zone))) {
    const seg = zone.slice(avant, m.index).trim(); avant = m.index + m[0].length;
    const prix = num(m[1]), total = num(m[2]);
    const q = [...seg.matchAll(/(\d+(?:[.,]\d+)?)\s+(UN|ME|PI|PQT|BTE|RL|RLX|M|PC|PCE|CH|HR|H|EA|LOT|BOITE)\b/g)].pop();
    let qte = q ? num(q[1]) : (prix ? Math.round(total / prix * 100) / 100 : 0);
    if (prix && Math.abs(qte * prix - total) > 0.05 && Math.abs(total / prix - Math.round(total / prix)) < 0.01) qte = Math.round(total / prix);
    const codes = (seg.match(/\b(?=[A-Z0-9-]*\d)(?=[A-Z0-9-]*[A-Z])[A-Z0-9][A-Z0-9-]{3,}\b/g) || []).filter(c => !/^\d+$/.test(c));
    const desc = seg.replace(q ? q[0] : '', ' ').replace(/\s+/g, ' ').trim().slice(0, 90);
    const mo = /^SA\d|[ÉE]lectricien|main.d.?[oœ]uvre|temps homme/i.test(seg) || (q && /^(HR|H)$/.test(q[2]));
    out.push([(mo ? 'MO ' : '') + codes.join(' '), desc, qte, prix, total]);
  }
  return out;
}


// Onglet « Lignes » du registre des factures : Facture, Codes, Description, Qté, Prix, Montant.
function facOngletLignes_() {
  const ss = SpreadsheetApp.openById(REG_SHEET_FACT);
  let sh = ss.getSheetByName('Lignes');
  if (!sh) { sh = ss.insertSheet('Lignes'); sh.getRange(1, 1, 1, 6).setValues([['Facture', 'Codes', 'Description', 'Qté', 'Prix', 'Montant']]); }
  return sh;
}

// Rattrapage : articles de tes factures déjà au registre (lit le PDF). Relance tant qu'il en reste.
function lignesFactures() {
  const debut = Date.now();
  const reg = SpreadsheetApp.openById(REG_SHEET_FACT).getSheets()[0].getDataRange().getDisplayValues().slice(1);
  const o = facOngletLignes_();
  const faites = new Set(o.getDataRange().getDisplayValues().slice(1).map(r => r[0]));
  const ajout = []; let n = 0, reste = 0, sansPdf = 0;
  for (const r of reg) {
    const num = r[0], pdf = r[9];
    if (!num || faites.has(num)) continue;
    if (!pdf) { sansPdf++; continue; }
    if (Date.now() - debut > 5 * 60 * 1000) { reste++; continue; }
    let l = [];
    try { l = facLignes_(lirePdf_(DriveApp.getFileById(pdf).getBlob())); } catch (e) { Logger.log('PDF introuvable : facture ' + num); }
    l.forEach(x => ajout.push([num].concat(x)));
    if (!l.length) ajout.push([num, '', '(aucun article lu)', 0, 0, 0]);
    n++;
  }
  if (ajout.length) o.getRange(o.getLastRow() + 1, 1, ajout.length, 6).setValues(ajout);
  Logger.log('Factures détaillées : ' + n + ' | Articles : ' + ajout.length + ' | Sans PDF : ' + sansPdf + (reste ? ' | Il en reste ' + reste + ' : relance.' : ' | Terminé.'));
}
