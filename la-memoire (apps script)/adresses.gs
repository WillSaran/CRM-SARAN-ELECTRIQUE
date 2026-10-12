/**
 * LA MÉMOIRE — Adresses manquantes
 *
 * Pour chaque job de l'agenda (The Crew + En Attente) dont les événements n'ont pas d'adresse :
 *   1) l'adresse du registre des jobs (par le PO), sinon
 *   2) une adresse trouvée dans tes courriels (recherche par le PO / numéro de job, sinon par le nom).
 * Rien n'est écrit dans l'agenda ici : les propositions vont dans l'onglet « Adresses » du registre des jobs,
 * et c'est toi qui confirmes (✓) dans le CRM — le CRM écrit alors l'adresse dans les événements.
 *
 * Nouveau fichier « adresses » dans le même projet Apps Script (utilise achNoJob_ / achMots_ du fichier achats).
 */

const ADR_ESSAI = false;
const ADR_JOBS = '1POqSf9_TAJCr8cYsV2QEGYzcPdLOeIGeWnIX8Y_swBg';       // La mémoire — Jobs
const ADR_CALS = ['saranelectrique@gmail.com', 'c_7d3dd0732b25a3afd30d0f2c453bef2d31a41b08a6d7112b08accea16ff49daf@group.calendar.google.com'];
// Tes adresses et celles des fournisseurs : jamais proposées.
const ADR_EXCLURE = /jacques[- ]cartier|grand[- ]bernier|industriel|\bbernier\b|\bmatte\b|autoroute|sonepar|division|360 rue st-louis|taschereau|brunelle/i;   // + adresses de Cuisines Rive-Sud
const ADR_IGNORER = /^(mon plan|plan de la semaine|topo|ma journ[ée]e|cong[ée]|vacances|pr[ée]p)|rendez|banque|dentiste|m[ée]decin|r[ée]union|meeting|formation|anniversaire|camion|shop|inventaire/i;

function adressesManquantes() {
  const debut = Date.now();
  const onglet = adrOnglet_();
  const lignes = onglet.getDataRange().getDisplayValues();
  const parCle = {}; lignes.slice(1).forEach((r, i) => { if (r[0]) parCle[r[0]] = { row: i + 2, statut: r[8] }; });
  const registre = adrRegistre_();

  // 1) Événements sans adresse, regroupés par job (PO / no de job, sinon le titre sans phase ni initiales).
  const jobs = {};
  ADR_CALS.forEach(id => {
    const cal = CalendarApp.getCalendarById(id); if (!cal) return;
    cal.getEvents(new Date(Date.now() - 30 * 864e5), new Date(Date.now() + 120 * 864e5)).forEach(ev => {
      const t = ev.getTitle().trim();
      if (!t || ADR_IGNORER.test(t) || String(ev.getLocation() || '').trim()) return;
      const po = achNoJob_(t);
      const base = t.replace(/\b(ROUGH|FINITION|FINALE|SERV|SERVICE|RAPUP|AJUSTEMENTS?|D[ÉE]MO)\b/ig, '').replace(/\b[A-Z]{2}\b/g, '').replace(/\s+/g, ' ').trim();
      const cle = po || base.toUpperCase();
      const j = jobs[cle] || (jobs[cle] = { cle, po, titre: base, evs: [], prochain: '' });
      j.evs.push(ev.getId() + '|' + id);
      const d = Utilities.formatDate(ev.getStartTime(), 'America/Toronto', 'yyyy-MM-dd');
      if (!j.prochain || d < j.prochain) j.prochain = d;
    });
  });

  // 2) Une proposition par job.
  const trouvees = {}, ajout = [];
  let n = 0, rien = 0, reste = 0;
  for (const j of Object.values(jobs)) {
    const deja = parCle[j.cle];
    if (deja && /^(Propos|[ÉE]crite|Refus)/i.test(deja.statut)) continue;      // déjà traitée
    if (Date.now() - debut > 5 * 60 * 1000) { reste++; continue; }
    let prop = null;
    const reg = j.po && registre[j.po];
    if (reg) prop = { adr: reg, autres: '', source: 'Registre des jobs', lien: '' };
    else prop = adrChercherCourriels_(j);
    if (prop) { trouvees[j.cle] = prop; n++; } else rien++;
    const ligne = [j.cle, j.titre, j.prochain, j.evs.join(';'), prop ? prop.adr : '', prop ? prop.autres : '', prop ? prop.source : '', prop ? prop.lien : '',
      prop ? 'Proposée' : 'Rien trouvé', Utilities.formatDate(new Date(), 'America/Toronto', 'yyyy-MM-dd HH:mm')];
    if (deja) { if (!ADR_ESSAI) onglet.getRange(deja.row, 1, 1, 10).setValues([ligne]); } else ajout.push(ligne);
    Logger.log((ADR_ESSAI ? '[essai] ' : '') + j.titre + (j.po ? ' (' + j.po + ')' : '') + ' → ' + (prop ? prop.adr + '  [' + prop.source + ']' : 'rien trouvé'));
  }

  // 3) Une adresse proposée pour 3 jobs différentes ou plus = la signature d'un courriel (bureau du contracteur) : retirée.
  const compte = {}; Object.values(trouvees).forEach(p => { if (p.source !== 'Registre des jobs') compte[p.adr] = (compte[p.adr] || 0) + 1; });
  ajout.forEach(l => { if (compte[l[4]] >= 3) { Logger.log('Retirée (adresse de signature) : ' + l[4] + ' pour ' + l[1]); l[4] = ''; l[6] = ''; l[8] = 'Rien trouvé'; } });

  if (!ADR_ESSAI && ajout.length) onglet.getRange(onglet.getLastRow() + 1, 1, ajout.length, 10).setValues(ajout);
  Logger.log('---------------------------------');
  Logger.log((ADR_ESSAI ? 'MODE ESSAI — rien inscrit. ' : '') + 'Jobs sans adresse : ' + Object.keys(jobs).length + ' | Adresses proposées : ' + n + ' | Rien trouvé : ' + rien + (reste ? ' | Il en reste ' + reste + ' : relance.' : ''));
}

// Cherche dans les courriels des 8 derniers mois : par le PO / numéro de job, sinon par les mots du nom.
function adrChercherCourriels_(j) {
  let q = '';
  if (j.po) q = '"' + j.po + '"';
  else {
    const mots = achMots_(j.titre).slice(0, 3);
    if (!mots.length) return null;
    q = mots.map(m => '"' + m + '"').join(' ');
  }
  const fils = GmailApp.search(q + ' newer_than:240d', 0, 8);
  const votes = {}, lien = {};
  fils.forEach(t => t.getMessages().slice(-4).forEach(m => {
    const txt = m.getSubject() + '\n' + m.getPlainBody().split(/\n\s*(?:--|Le .{5,80} a écrit|On .{5,80} wrote|De :|From:)/)[0];   // sans les anciens messages cités
    adrExtraire_(txt).forEach(a => { votes[a] = (votes[a] || 0) + 1; if (!lien[a]) lien[a] = { url: 'https://mail.google.com/mail/u/0/#all/' + t.getId(), sujet: m.getSubject(), date: Utilities.formatDate(m.getDate(), 'America/Toronto', 'yyyy-MM-dd') }; });
  }));
  const tri = Object.keys(votes).sort((a, b) => votes[b] - votes[a] || (/[A-Z]\d[A-Z]/.test(b) - /[A-Z]\d[A-Z]/.test(a)));
  if (!tri.length) return null;
  const l = lien[tri[0]];
  return { adr: tri[0], autres: tri.slice(1, 3).join(' | '), source: 'Courriel « ' + l.sujet.slice(0, 60) + ' » (' + l.date + ')', lien: l.url };
}

// Adresses québécoises dans un texte : « 123 rue Dauphine, Saint-Jean-sur-Richelieu, QC J3B 1A1 ».
function adrExtraire_(txt) {
  const out = [];
  const re = /(?:^|[^\d])(\d{1,5}[A-Za-z]?)\s*,?\s+((?:rue|boul(?:evard)?\.?|bd\.?|ch(?:emin)?\.?|av(?:enue)?\.?|rang|mont[ée]e|place|croissant|route|terrasse|impasse|all[ée]e|c[ôo]te|carr[ée]|cercle|promenade|avenue)\s+[A-Za-zÀ-ÿ0-9'’.\- ]{2,40}?)(?=\s*(?:,|\n|\r|$|\(|\s(?:app|apt|unit|local|suite|#)\b|\s[A-Z]\d[A-Z]\s?\d|\s(?:QC|Qu[ée]bec)\b))((?:\s*,\s*[A-Za-zÀ-ÿ'’.\- ]{3,40}?(?=\s*(?:,|\n|\r|$|\s(?:QC|Qu[ée]bec)\b|\s[A-Z]\d[A-Z]))))?(?:\s*,?\s*(?:QC|Qu[ée]bec))?(?:\s*,?\s*([A-Z]\d[A-Z]\s?\d[A-Z]\d))?/gi;
  let m;
  while ((m = re.exec(String(txt || '')))) {
    const rue = m[2].replace(/\s+/g, ' ').trim(), ville = (m[3] || '').replace(/^\s*,\s*/, '').trim(), cp = (m[4] || '').toUpperCase();
    let a = m[1] + ' ' + rue + (ville ? ', ' + ville : '') + (cp ? ' ' + cp : '');
    a = a.replace(/\s+/g, ' ').trim();
    if (ADR_EXCLURE.test(a) || rue.split(' ').length > 7) continue;
    if (out.indexOf(a) < 0) out.push(a);
  }
  return out;
}

// Adresses du registre des jobs (colonnes Adresse + Ville), par PO.
function adrRegistre_() {
  const v = SpreadsheetApp.openById(ADR_JOBS).getSheets()[0].getDataRange().getDisplayValues();
  const h = v[0].map(x => String(x).toLowerCase()), iPO = h.indexOf('po'), iA = h.findIndex(x => /^adresse/.test(x)), iV = h.findIndex(x => /^ville/.test(x));
  const out = {};
  if (iPO < 0 || iA < 0) return out;
  v.slice(1).forEach(r => { if (r[iPO] && r[iA]) out[r[iPO]] = r[iA] + (iV >= 0 && r[iV] ? ', ' + r[iV] : ''); });
  return out;
}

// Onglet « Adresses » du registre des jobs : une ligne par job.
function adrOnglet_() {
  const ss = SpreadsheetApp.openById(ADR_JOBS);
  let sh = ss.getSheetByName('Adresses');
  if (!sh) { sh = ss.insertSheet('Adresses'); sh.getRange(1, 1, 1, 10).setValues([['Clé', 'Job', 'Prochaine date', 'Événements', 'Adresse proposée', 'Autres possibilités', 'Source', 'Lien', 'Statut', 'Le']]); }
  return sh;
}

// ─────────────────────────────────────────────────────────────────────────────
// Fiche d'une job à partir du courriel du PO : lit les PDF joints (bon de commande, plans de production),
// trouve l'adresse du chantier, téléphone et courriel du client, puis :
//   - remplit la ligne du registre des jobs (seulement les cases vides ; crée la ligne si la job n'y est pas),
//   - met l'adresse dans les événements de la job qui n'en ont pas.
// Lance ficheStLaurent() (ou remplirFicheJob avec un autre PO) et envoie-moi le journal.
function ficheStLaurent() {
  remplirFicheJob('P02324', { code: 'CRS', nom: 'St-Laurent', client: 'Micheline St-Laurent', projet: 'Modification cuisine',
    note: '20 oct. gardé en réserve avec Karine Bourbier (CRS) : 2 ventilateurs et un lustre, sur échafaudage — date à confirmer par Karine.' });
}

function remplirFicheJob(po, infos) {
  infos = infos || {};
  const num = po.replace(/^P0?/i, '');
  const fils = GmailApp.search('(P0' + num + ' OR PO' + num + ' OR "PO ' + num + '") has:attachment', 0, 10);
  const adresses = {}, tels = {}, mails = {};
  fils.forEach(t => t.getMessages().forEach(m => m.getAttachments({ includeInlineImages: false }).forEach(pj => {
    if (!/\.pdf$/i.test(pj.getName())) return;
    let txt = '';
    try { txt = lirePdf_(pj.copyBlob()); } catch (e) { Logger.log('PDF illisible : ' + pj.getName()); return; }
    Logger.log('--- ' + pj.getName() + ' : ' + String(txt).replace(/\s+/g, ' ').slice(0, 300));
    if (!/bon de commande|plan/i.test(pj.getName())) return;          // fiches techniques des fabricants : ignorées
    adrExtraire_(txt).filter(a => !/brunelle/i.test(a)).forEach(a => { adresses[a] = (adresses[a] || 0) + 1; });
    (String(txt).match(/\(?\b\d{3}\)?[\s.-]?\d{3}[\s.-]\d{4}\b/g) || []).filter(x => !/656.?8660|781.?8687|490.?3412/.test(x)).forEach(x => { tels[x] = (tels[x] || 0) + 1; });
    (String(txt).match(/[\w.+-]+@[\w-]+\.[\w.]+/g) || []).filter(x => !/cuisinesrivesud|saranelectrique|odoo/i.test(x)).forEach(x => { mails[x] = (mails[x] || 0) + 1; });
  })));
  const top = o => Object.keys(o).sort((a, b) => o[b] - o[a]);
  const adr = top(adresses)[0] || '', tel = top(tels)[0] || '', mail = top(mails)[0] || '';
  Logger.log('Adresses trouvées : ' + JSON.stringify(adresses));
  Logger.log('Téléphones : ' + JSON.stringify(tels) + ' | Courriels : ' + JSON.stringify(mails));
  if (!adr) Logger.log('⚠️ Aucune adresse trouvée dans les PDF — rien écrit pour l\'adresse.');
  const rue = adr.split(',')[0].trim(), ville = (adr.split(',')[1] || '').replace(/[A-Z]\d[A-Z]\s?\d[A-Z]\d/, '').trim();

  // Registre des jobs : seulement les cases vides.
  const sh = SpreadsheetApp.openById(ADR_JOBS).getSheets()[0];
  const v = sh.getDataRange().getValues();
  const h = v[0].map(x => String(x).toLowerCase());
  const col = re => h.findIndex(x => re.test(x));
  const c = { code: col(/^code/), po: col(/^po$/), nom: col(/^nom/), adr: col(/^adresse/), ville: col(/^ville/), cli: col(/^client/), tel: col(/^t[ée]l/), mail: col(/^courriel/), proj: col(/^projet/), note: col(/^note/) };
  let i = v.findIndex((r, k) => k > 0 && String(r[c.po]).toUpperCase() === 'P0' + num);
  if (i < 0) { sh.appendRow(h.map(() => '')); i = sh.getLastRow() - 1; v[i] = h.map(() => ''); Logger.log('Ligne créée au registre des jobs pour P0' + num); }
  const poser = (k, val) => { if (c[k] >= 0 && val && !String(v[i][c[k]] || '').trim()) { sh.getRange(i + 1, c[k] + 1).setValue(val); Logger.log('Registre ← ' + h[c[k]] + ' : ' + val); } };
  poser('code', infos.code); poser('po', 'P0' + num); poser('nom', infos.nom); poser('adr', rue); poser('ville', ville);
  poser('cli', infos.client); poser('tel', tel); poser('mail', mail); poser('proj', infos.projet); poser('note', infos.note);

  // Événements de la job sans adresse.
  if (adr) ADR_CALS.forEach(id => {
    const cal = CalendarApp.getCalendarById(id); if (!cal) return;
    cal.getEvents(new Date(Date.now() - 300 * 864e5), new Date(Date.now() + 200 * 864e5)).forEach(ev => {
      if (!new RegExp('P[O0]?\\s?0?' + num + '\\b', 'i').test(ev.getTitle() + ' ' + ev.getDescription())) return;
      if (String(ev.getLocation() || '').trim()) return;
      ev.setLocation(adr);
      Logger.log('Adresse ajoutée à l\'événement : ' + ev.getTitle() + ' (' + Utilities.formatDate(ev.getStartTime(), 'America/Toronto', 'yyyy-MM-dd') + ')');
    });
  });
  Logger.log('Terminé.');
}

// Correction St-Laurent : « 360 Rue St-Louis » était l'adresse de Cuisines Rive-Sud (expédition), pas le chantier.
// Met la vraie adresse (donnée par William) : 1072 rue des Grands-Ducs, Longueuil, J4G 2N9.
function corrigerStLaurent() {
  const sh = SpreadsheetApp.openById(ADR_JOBS).getSheets()[0];
  const v = sh.getDataRange().getValues(), h = v[0].map(x => String(x).toLowerCase());
  const iPO = h.indexOf('po'), iA = h.findIndex(x => /^adresse/.test(x));
  const iV = h.findIndex(x => /^ville/.test(x));
  v.forEach((r, i) => { if (i && String(r[iPO]).toUpperCase() === 'P02324') {
    sh.getRange(i + 1, iA + 1).setValue('1072 rue des Grands-Ducs');
    if (iV >= 0) sh.getRange(i + 1, iV + 1).setValue('Longueuil');
    Logger.log('Registre P02324 (ligne ' + (i + 1) + ') : adresse = 1072 rue des Grands-Ducs, Longueuil.'); } });
  return;   // l'adresse est connue : plus besoin de lire les plans
  const t = GmailApp.search('P02324 has:attachment', 0, 1)[0];
  t.getMessages().forEach(m => m.getAttachments({ includeInlineImages: false }).forEach(pj => {
    if (!/plan|bon de commande/i.test(pj.getName())) return;
    const txt = String(lirePdf_(pj.copyBlob())).replace(/[ \t]+/g, ' ');
    Logger.log('===== ' + pj.getName() + ' =====');
    for (let k = 0; k < Math.min(txt.length, 6000); k += 1500) Logger.log(txt.slice(k, k + 1500));
  }));
}
