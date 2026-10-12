/**
 * LA MÉMOIRE — Achats fournisseurs (Dubo, Lumen, Groupe Gagnon, Westburne)
 *
 * Chaque facture fournisseur reçue par courriel :
 *   - le PDF est sauvé dans « La mémoire / Fournisseurs / <Fournisseur> »
 *   - une ligne est ajoutée au registre « La mémoire — Achats » : numéro, date, référence de commande,
 *     sous-total, total, et la JOB PROPOSÉE (par le PO ou le nom dans tes titres d'événement).
 * La job n'est jamais confirmée toute seule : tu confirmes (ou corriges, ou envoies au stock) dans le CRM.
 *
 * Nouveau fichier « achats » dans le même projet Apps Script (il utilise lirePdf_ du fichier registre).
 * Lance achatsFournisseurs() en essai d'abord (ACH_ESSAI = true), envoie-moi le journal.
 */

const ACH_ESSAI = false;
const ACH_SHEET = '1i6DNjPmxKRubA8IcSxNoC12SEwefVLcNF39F-dvEnmw';      // La mémoire — Achats
const ACH_MEMOIRE = '1U0syTc5wf_u-0R9Q423s8ykXiou9SWAP';               // La mémoire
const ACH_JOBS = '1POqSf9_TAJCr8cYsV2QEGYzcPdLOeIGeWnIX8Y_swBg';       // La mémoire — Jobs
const ACH_DEPUIS = '2026/08/01';
const ACH_FOURNISSEURS = [
  { nom: 'Dubo', q: 'from:dubo.qc.ca -from:credit@dubo.qc.ca' },
  { nom: 'Lumen', q: 'from:lumen.ca' },
  { nom: 'Groupe Gagnon', q: 'from:gagnonlgq.com -from:paiements@gagnonlgq.com' },
  { nom: 'Westburne', q: 'from:westburne.ca' },
];

function achatsFournisseurs() {
  const debut = Date.now();
  const sh = SpreadsheetApp.openById(ACH_SHEET).getSheets()[0];
  const vals = sh.getDataRange().getDisplayValues();
  const deja = new Set(vals.slice(1).map(r => r[0] + '|' + r[1]));
  const racine = achDossier_(DriveApp.getFolderById(ACH_MEMOIRE), 'Fournisseurs');
  const jobs = achJobs_();
  const stats = { ajoutees: 0, deja: 0, illisibles: 0, proposees: 0, documents: 0 };
  const docs = achOngletDocs_(), docsDeja = new Set(docs.getDataRange().getDisplayValues().slice(1).map(r => r[0] + '|' + r[9]));
  const lignes = [], articles = [];

  for (const f of ACH_FOURNISSEURS) {
    const dossier = achDossier_(racine, f.nom);
    const fils = GmailApp.search(f.q + ' has:attachment after:' + ACH_DEPUIS, 0, 200);
    for (const t of fils) for (const msg of t.getMessages()) {
      if (Date.now() - debut > 5 * 60 * 1000) { Logger.log('Presque 6 minutes : relance pour continuer.'); return achEcrire_(sh, lignes, stats, articles); }
      if (/saranelectrique/i.test(msg.getFrom())) continue;
      const pdfs = msg.getAttachments({ includeInlineImages: false }).filter(a => /\.pdf$/i.test(a.getName()));
      for (const pj of pdfs) { try {
        // Soumissions, relevés de compte et confirmations : gardés à part (onglet « Documents » + dossier du fournisseur).
        const type = achTypeDoc_(pj.getName(), msg.getSubject(), f.nom);
        if (type === 'ignorer') continue;
        if (type !== 'facture') { achDocument_(f, type, pj, msg, docs, docsDeja, dossier, stats); continue; }
        const pre = (pj.getName().match(/(\d{6,})/) || [])[1];
        if (pre && deja.has(f.nom + '|' + pre)) { stats.deja++; continue; }     // déjà lue : on ne relit pas le PDF
        const txt = lirePdf_(pj.copyBlob());
        const t0 = String(txt).slice(0, 400);
        if (/RELEV[ÉE] DE COMPTE|[ÉE]TAT DE COMPTE|STATEMENT/i.test(t0)) { achDocument_(f, 'Relevé', pj, msg, docs, docsDeja, dossier, stats, txt); continue; }
        if (/\bSOUMISSION\b|\bQUOTATION\b/i.test(t0) && !/\bFACTURE\b/i.test(t0)) { achDocument_(f, 'Soumission', pj, msg, docs, docsDeja, dossier, stats, txt); continue; }
        // Un PDF peut contenir plusieurs factures (Lumen « LUM_MULTIPLE_FACTURES ») : chacune est lue à part.
        for (const part of achDecouper_(txt, pj.getName())) {
        const lu = achLire_(part.txt, part.nom);
        if (!lu) { stats.illisibles++; const t1 = String(part.txt).replace(/\s+/g, ' '); Logger.log('Illisible (' + f.nom + ') : ' + part.nom + ' | ' + t1.slice(0, 120) + ' … ' + t1.slice(-350)); continue; }
        const cle = f.nom + '|' + lu.num;
        if (deja.has(cle)) { stats.deja++; continue; }
        deja.add(cle);
        if (!lu.date) lu.date = Utilities.formatDate(msg.getDate(), 'America/Toronto', 'yyyy-MM-dd');
        const prop = achProposer_(lu.ref, jobs);
        if (prop) stats.proposees++;
        const nomPdf = f.nom + ' ' + lu.num + (lu.date ? ' (' + lu.date + ')' : '') + '.pdf';
        let pdfId = '';
        if (!ACH_ESSAI) { const it = dossier.getFilesByName(nomPdf); pdfId = (it.hasNext() ? it.next() : dossier.createFile(pj.copyBlob()).setName(nomPdf)).getId(); }
        lignes.push([f.nom, lu.num, lu.date, lu.ref, lu.st, lu.total, pdfId, prop, '', '', '', Utilities.formatDate(msg.getDate(), 'America/Toronto', 'yyyy-MM-dd')]);
        achLignes_(part.txt).forEach(l => articles.push([f.nom, lu.num].concat(l)));
        stats.ajoutees++;
        Logger.log((ACH_ESSAI ? '[essai] ' : '') + f.nom + ' ' + lu.num + ' | ' + lu.date + ' | réf. « ' + lu.ref + ' » | ' + lu.st + ' $ | job proposée : ' + (prop || '—'));
        }
      } catch (e) {
        // Erreur passagère de Google (code 500) : on saute ce PDF, il sera relu au prochain passage. Les autres sont quand même inscrits.
        stats.illisibles++; Logger.log('Erreur temporaire sur ' + pj.getName() + ' (' + String(e.message || e).slice(0, 80) + ') : réessayée au prochain passage.');
      } }
    }
  }
  achEcrire_(sh, lignes, stats, articles);
}

function achEcrire_(sh, lignes, stats, articles) {
  if (!ACH_ESSAI && lignes.length) sh.getRange(sh.getLastRow() + 1, 1, lignes.length, lignes[0].length).setValues(lignes);
  if (!ACH_ESSAI && articles && articles.length) achOngletLignes_().getRange(achOngletLignes_().getLastRow() + 1, 1, articles.length, 6).setValues(articles);
  Logger.log('---------------------------------');
  Logger.log((ACH_ESSAI ? 'MODE ESSAI — rien inscrit. ' : '') + 'Factures ajoutées : ' + stats.ajoutees + ' | Déjà au registre : ' + stats.deja +
    ' | Job proposée : ' + stats.proposees + ' | Soumissions/relevés ajoutés : ' + (stats.documents || 0) + ' | Illisibles : ' + stats.illisibles);
}

// Type d'un PDF d'après son nom et le sujet du courriel : facture, soumission, relevé, confirmation, ou à ignorer.
function achTypeDoc_(nom, sujet, fournisseur) {
  const t = nom + ' ' + sujet;
  if (fournisseur === 'Dubo' && /^Inv\d+/i.test(nom)) return 'facture';
  if (/demande de cr[ée]dit|formulaire|fiche technique|spec ?sheet/i.test(t)) return 'ignorer';
  if (/relev[ée]|statement/i.test(t) || /^LUM\d+\.pdf$/i.test(nom)) return 'Relevé';
  if (/soumission|\bsoum\b|quote|quotation|devis|^QT\d/i.test(t)) return 'Soumission';
  if (/confirmation/i.test(t)) return 'Confirmation';
  if (fournisseur === 'Dubo' && !/^Inv\d+/i.test(nom)) return 'ignorer';    // fiches techniques, etc.
  return 'facture';
}

// Onglet « Documents » du registre des achats : soumissions, relevés de compte, confirmations de commande.
function achOngletDocs_() {
  const ss = SpreadsheetApp.openById(ACH_SHEET);
  let sh = ss.getSheetByName('Documents');
  if (!sh) { sh = ss.insertSheet('Documents'); sh.getRange(1, 1, 1, 10).setValues([['Fournisseur', 'Type', 'Numéro', 'Date', 'Référence', 'Total', 'PDF', 'Fichier', 'Reçu le', 'Empreinte']]); }
  return sh;
}
function achDocument_(f, type, pj, msg, docs, docsDeja, dossier, stats, txt) {
  const empreinte = pj.getHash();
  if (docsDeja.has(f.nom + '|' + empreinte)) return;                       // même PDF déjà gardé (renvoyé dans un autre courriel)
  docsDeja.add(f.nom + '|' + empreinte);
  const recu = Utilities.formatDate(msg.getDate(), 'America/Toronto', 'yyyy-MM-dd');
  let num = '', ref = '', total = '';
  if (type !== 'Relevé') {
    txt = txt || lirePdf_(pj.copyBlob());
    const plat = String(txt || '').replace(/[ \t]+/g, ' ');
    num = ((plat.match(/(?:soumission|quote|r[ée]f[ée]rence rexel|n[o°º]\.?\s*(?:de\s*)?(?:soumission|commande))\s*(?:n[o°º]\.?|#)?\s*:?\s*([A-Z]{0,3}-?\d{4,}(?:-\d+)?)/i) || [])[1]) || '';
    const lu = achLire_(txt, pj.getName());
    if (lu) { ref = lu.ref; total = lu.total; }
    else { const r = plat.match(/(?:PROJET|R[ÉE]F[ÉE]RENCE|VOTRE\s*N[°ºo]\.?\s*COMMANDE|PO)\s*:[ \t]*([^\n]{2,50})/i); if (r) ref = r[1].trim(); }
    if (/^(VENDU|EXP[ÉE]DI|LIVR|FACTUR|DATE|T[ÉE]L)|:\s*$/i.test(ref)) ref = '';      // une étiquette vide, pas ta référence
  }
  if (!num) num = ((pj.getName().match(/([A-Z]{0,3}-?\d{5,}(?:-\d+)?)/i) || [])[1]) || '';
  const sous = achDossier_(dossier, type === 'Relevé' ? 'Relevés' : 'Soumissions');
  const nomPdf = f.nom + ' ' + type.toLowerCase() + ' ' + (num || pj.getName().replace(/\.pdf$/i, '')) + ' (' + recu + ').pdf';
  let pdfId = '';
  if (!ACH_ESSAI) {
    const it = sous.getFilesByName(nomPdf); pdfId = (it.hasNext() ? it.next() : sous.createFile(pj.copyBlob()).setName(nomPdf)).getId();
    docs.appendRow([f.nom, type, num, recu, ref, total, pdfId, pj.getName(), recu, empreinte]);
  }
  stats.documents = (stats.documents || 0) + 1;
  Logger.log((ACH_ESSAI ? '[essai] ' : '') + f.nom + ' — ' + type + ' ' + (num || pj.getName()) + ' | ' + recu + (ref ? ' | réf. « ' + ref + ' »' : '') + (total ? ' | ' + total + ' $' : ''));
}

// Sépare un PDF qui contient plusieurs factures Lumen (« FACTURE *** ORIGINAL *** » répété). Sinon : le PDF tel quel.
function achDecouper_(txt, nom) {
  const parts = String(txt || '').split(/(?=FACTURE\s*\*+\s*O\s?R\s?I\s?G\s?I\s?N\s?A\s?L)/i).filter(p => /\d{8}-\d{2}/.test(p));
  if (parts.length < 2) return [{ txt: txt, nom: nom }];
  return parts.map(p => { const n = p.match(/(\d{8})-(\d{2})/); return { txt: p, nom: n[1] + n[2] + '.pdf' }; });
}

// Lit numéro, date, référence de commande et montants. Rien (null) si les montants ne balancent pas.
function achLire_(txt, nomFichier) {
  if (!txt) return null;
  const plat = txt.replace(/ /g, ' ').replace(/[ \t]+/g, ' ');
  const n = x => x == null ? null : Number(String(x).replace(/[\s,]/g, '').replace(/^(-?\d+)\.(\d{2})\d*$/, '$1.$2'));
  let m = null;
  // 1) Montants étiquetés (Dubo : « SOUS-TOTAL: 148.45 TPS : 7.44 TVQ : 14.81 … MONTANT DÛ: 170.70 »).
  const lab = re => { const r = plat.match(re); return r ? n(r[1]) : null; };
  const st1 = lab(/SOUS[- ]?TOTAL\s*:?\s*(-?[\d,]+\.\d{2})/i), tps1 = lab(/\bT\.?P\.?S\.?\s*:?\s*(-?[\d,]+\.\d{2})/i),
        tvq1 = lab(/\bT\.?V\.?Q\.?\s*:?\s*(-?[\d,]+\.\d{2})/i), tot1 = lab(/MONTANT\s*D[ÛU]\s*:?\s*\$?\s*(-?[\d,]+\.\d{2})/i) ?? lab(/(?:^|[^-A-Z])TOTAL\s*(?:DE LA FACTURE|À PAYER|CAD)?\s*:?\s*\$?\s*(-?[\d,]+\.\d{2})/i);
  if (st1 != null && tps1 != null && tvq1 != null && tot1 != null && Math.abs(st1 + tps1 + tvq1 - tot1) <= 0.03) m = { st: st1, tot: tot1 };
  // 2) Sinon : un sous-total dont 5 % et 9,975 % sont sur la facture, et le total qui balance.
  if (!m) {
    const argent = (plat.match(/\d{1,3}(?:[ ,.]\d{3})*[.,]\d{2}(?!\d)/g) || []).map(x => Number(x.replace(/\D/g, '')) / 100);
    const present = v => argent.some(a => Math.abs(a - v) < 0.015);
    argent.forEach(st => {
      const tps = Math.round(st * 5) / 100, tvq = Math.round(st * 9.975) / 100, tot = st + tps + tvq;
      if (st > 0 && present(tps) && present(tvq) && argent.some(a => Math.abs(a - tot) <= 0.03) && (!m || st > m.st)) m = { st, tot: argent.find(a => Math.abs(a - tot) <= 0.03) };
    });
  }
  // 3) Sinon (Lumen : « 485.28 0.00 24.26 48.41 557.95 » = brut, frais divers, TPS, TVQ, total) :
  //    une suite de montants qui s'additionnent au dernier, avec une TPS et une TVQ qui balancent.
  if (!m) {
    // Montants entiers (« 1995.48 », « 1,995.48 ») ; « 99.77- » = crédit (négatif).
    const suite = (plat.match(/(?<![\d.,])-?\d+(?:,\d{3})*\.\d{2}-?(?![\d])/g) || []).map(x => (/-/.test(x) ? -1 : 1) * Number(x.replace(/[,-]/g, '')));
    for (let i = 0; i < suite.length && !m; i++) for (let k = 4; k <= 5 && !m; k++) {
      const w = suite.slice(i, i + k); if (w.length < k) break;
      const tot = w[k - 1], parts = w.slice(0, k - 1), somme = parts.reduce((a, b) => a + b, 0);
      if (!tot || Math.abs(somme - tot) > 0.03) continue;
      const tps = parts[k - 3], tvq = parts[k - 2], base = somme - tps - tvq;
      if (base && Math.sign(base) === Math.sign(tot) && Math.abs(tps - base * 0.05) <= 0.06 && Math.abs(tvq - base * 0.09975) <= 0.12) m = { st: Math.round(base * 100) / 100, tot };
    }
  }
  if (!m) return null;
  const num = ((nomFichier.match(/(\d{6,})/) || [])[1]) ||
    ((plat.match(/(?:facture|invoice|inv\.?)\s*(?:n[o°º]\.?|#|no|number)?\s*[:#]?\s*([A-Z]{0,3}\d{5,})/i) || [])[1]) || nomFichier.replace(/\.pdf$/i, '');
  // Date : « Date de Fact. 10/2/2026 » (Dubo, mois/jour/année), sinon aaaa-mm-jj, sinon jj/mm/aaaa.
  let date = '';
  const dd = plat.match(/Date de Fact\.?\s*(\d{1,2})\/(\d{1,2})\/(20\d{2})/i);
  const iso = plat.match(/(20\d{2})-(\d{2})-(\d{2})/);
  const dl = plat.match(/DATE DE FACTURE[\s\S]{0,60}?(\d{2})\/(\d{2})\/(20\d{2})/i);          // Lumen : mois/jour/année
  if (dd) date = dd[3] + '-' + ('0' + dd[1]).slice(-2) + '-' + ('0' + dd[2]).slice(-2);
  else if (dl) date = dl[3] + '-' + dl[1] + '-' + dl[2];
  else if (iso) date = iso[0];
  // Référence de commande : Dubo « … N° de projet Date due nette <TA RÉFÉRENCE> AR-2% … »
  let ref = '';
  const rd = plat.match(/Date due nette\s+(.*?)\s+(?:AR-|NET\b|N\/\d|\d+\s*JOURS|COMPTANT|C\.O\.D)/i);
  if (rd) ref = rd[1].trim();
  // Lumen : « N° CLIENT  VOTRE N° COMMANDE … 31441900 CRS FINITION GAGNE St-Jean-Richelieu … » (ton no client, puis ta référence).
  if (!ref) {
    const l = plat.match(/N[°ºo]\.?\s*CLIENT\s+VOTRE\s+N[°ºo]\.?\s*COMMANDE[\s\S]{0,160}?\b\d{6,10}[ \t]+([^\n]{2,60})/i);
    if (l) ref = l[1].replace(/\s+(?:St|Saint|Ste|Sainte)[- ][A-Z][\s\S]*$/i, '').replace(/\s+\d{2,5}\s+(?:Boul|Bd|Rue|Ch|Av)\b[\s\S]*$/i, '').replace(/\s+(?:TEL|T[ÉE]L)\b[\s\S]*$/i, '').trim();
  }
  if (!ref) { const r = plat.match(/(?:votre\s*(?:commande|r[ée]f[ée]rence|PO|bon\s*de\s*commande)|customer\s*(?:PO|order|ref(?:erence)?)|r[ée]f[ée]rence\s*client|purchase\s*order|PO\s*client)\s*[:#.]?\s*([^\n]{2,40})/i); if (r) ref = r[1].trim(); }
  return { num, date, ref: ref.slice(0, 50), st: m.st.toFixed(2), total: m.tot.toFixed(2) };
}

// Jobs connues : PO + nom (registre) et titres d'événements CRS/clients des 4 derniers mois et 2 prochains.
function achJobs_() {
  const out = [];
  const reg = SpreadsheetApp.openById(ACH_JOBS).getSheets()[0].getDataRange().getDisplayValues();
  const h = reg[0].map(x => x.toLowerCase()), iPO = h.indexOf('po'), iNom = h.indexOf('nom'), iCode = h.indexOf('code');
  reg.slice(1).forEach(r => { if (r[iPO]) out.push({ cle: (r[iCode] || 'CRS') + ' ' + r[iNom] + ' ' + r[iPO], po: r[iPO], mots: achMots_(r[iNom]) }); });
  const cal = CalendarApp.getCalendarById('saranelectrique@gmail.com');
  if (cal) cal.getEvents(new Date(Date.now() - 120 * 864e5), new Date(Date.now() + 60 * 864e5)).forEach(ev => {
    const t = ev.getTitle(); if (/^(mon plan|plan de la semaine)|rendez|banque|dentiste|m[ée]decin|cong[ée]|vacances|r[ée]union|meeting|formation|anniversaire/i.test(t)) return;
    const po = achNoJob_(t);
    if (po && out.some(j => j.po === po)) return;
    const base = t.replace(/\b(ROUGH|FINITION|SERV|SERVICE|RAPUP|FINALE|AJUSTEMENTS?)\b/ig, '').replace(/\s+/g, ' ').trim();
    if (!out.some(j => j.cle === base)) out.push({ cle: base, po, mots: achMots_(base) });
  });
  return out;
}
// No de job dans un texte : « PAY-012 », « ASC 004 », « PART-031 », ou le PO de CRS « P03213 ».
function achNoJob_(t) {
  const n = String(t || '').match(/\b(PAY|PY|ASC|SAR|GOU|5PIL|BOM|JOY|GGP|RROY|PART|SE)\s?-?\s?(\d{3})\b/i);
  if (n) return (n[1].toUpperCase() === 'PY' ? 'PAY' : n[1].toUpperCase()) + '-' + n[2];
  const p = String(t || '').match(/\bP[O0]?\s?0?(\d{4})\b/i);
  return p ? 'P0' + p[1] : '';
}
function achMots_(s) {
  return String(s || '').normalize('NFD').replace(/[̀-ͯ]/g, '').toUpperCase().split(/[^A-Z0-9]+/)
    .filter(m => m.length >= 4 && !/^(CRS|SERV|ROUGH|FINITION|CUISINE|SALLE|BAIN|BORNE|PANNEAU|EXTRA|CAMION|LETTRAGE|ENTREE|ELECTRIQUE|GARAGE|SERVICE|RAPUP|RUE|PHIL|PHILIPPE|MAUDE|ADELE|DANIEL|BRANDON)$/.test(m));
}
// Propose une job seulement si UNE seule correspond (PO exact, sinon un nom). Sinon : rien.
function achProposer_(ref, jobs) {
  if (!ref) return '';
  // Destinations écrites dans ta référence : « camion 14 », « inventaire », « shop ».
  const cam = ref.match(/\bcamion\s*#?\s*0?(7|14|25)\b/i); if (cam) return 'Camion ' + cam[1];
  if (/\b(inventaire|stock)\b/i.test(ref)) return 'Inventaire';
  if (/\b(shop|atelier)\b/i.test(ref)) return 'Shop';
  const po = achNoJob_(ref);
  if (po) { const j = jobs.find(x => x.po === po); if (j) return j.cle; }
  const mots = achMots_(ref);
  if (!mots.length) return '';
  // Assez de mots en commun : 2 si ta référence et la job en ont 2 ou plus (« Olivier Boudreau » ≠ « Olivier Poulin »).
  const c = jobs.filter(j => { const n = j.mots.filter(m => mots.indexOf(m) >= 0).length; return n > 0 && n >= Math.min(2, mots.length, j.mots.length); });
  if (!c.length) return '';
  // Plusieurs événements de la même job (rough, finition, initiales différentes) = une seule job.
  const avecPO = [...new Set(c.filter(x => x.po).map(x => x.po))];
  if (avecPO.length === 1 && c.every(x => x.po === avecPO[0] || !x.po)) return c.find(x => x.po === avecPO[0]).cle;
  const coeur = x => x.mots.filter(m => mots.indexOf(m) >= 0).sort().join(' ');
  const coeurs = [...new Set(c.map(coeur))];
  return coeurs.length === 1 && avecPO.length <= 1 ? (c.find(x => x.po) || c[0]).cle : '';
}

// Onglet « Lignes » du registre des achats : Fournisseur, Facture, Code, Description, Qté expédiée, Montant.
function achOngletLignes_() {
  const ss = SpreadsheetApp.openById(ACH_SHEET);
  let sh = ss.getSheetByName('Lignes');
  if (!sh) { sh = ss.insertSheet('Lignes'); sh.getRange(1, 1, 1, 6).setValues([['Fournisseur', 'Facture', 'Code', 'Description', 'Qté', 'Montant']]); }
  return sh;
}

// Articles d'une facture Dubo : « 3.0 3.0 0.0 E ABB-CI2169 C 411.0128 12.33 CONNECTEUR BX/FLEX 1PO ALUMINIUM »
// = commandé, expédié, en attente, unité, CODE, unité de prix, prix, MONTANT, description.
function achLignes_(txt) {
  const out = [];
  const re = /^\s*(-?\d+(?:\.\d+)?)\s+(-?\d+(?:\.\d+)?)\s+(-?\d+(?:\.\d+)?)\s+[A-Z]{1,3}\s+(\S+)\s+[A-Z]{1,3}\s+(-?[\d,]+\.\d+)\s+(-?[\d,]+\.\d{2})\s+(.+?)\s*$/gm;
  let m;
  while ((m = re.exec(String(txt || '')))) out.push([m[4], m[7].slice(0, 80), Number(m[2]), Number(m[6].replace(/,/g, ''))]);
  if (out.length) return out;
  // Lumen : « 1 STEMIR0502W CONV MIRAGE BLANC500W 240V 626296409145 » puis « 2 2 0 242.64 CH 242.64 485.28 »
  // = ligne, CODE, description, UPC ; commandé, expédié, à suivre, prix liste, unité, prix net, TOTAL.
  const lu = /^\s*\d{1,3}\s+([A-Z0-9][A-Z0-9\-\/.]{2,})\s+(.+?)\s+\d{11,13}\s*\n+\s*(-?\d+(?:\.\d+)?)\s+(-?\d+(?:\.\d+)?)\s+(-?\d+(?:\.\d+)?)\s+[\d,]+\.\d+\s+[A-Z]{1,3}\s+[\d,]+\.\d+\s+(-?[\d,]+\.\d{2})/gm;
  while ((m = lu.exec(String(txt || '')))) out.push([m[1], m[2].slice(0, 80), Number(m[4]), Number(m[6].replace(/,/g, ''))]);
  return out;
}

// Rattrapage : ajoute les articles des factures déjà inscrites (lit le PDF rangé dans le Drive). Relance tant qu'il en reste.
function lignesAchats() {
  const debut = Date.now();
  const reg = SpreadsheetApp.openById(ACH_SHEET).getSheets()[0].getDataRange().getDisplayValues().slice(1);
  const onglet = achOngletLignes_();
  const faites = new Set(onglet.getDataRange().getDisplayValues().slice(1).map(r => r[0] + '|' + r[1]));
  const ajout = []; let n = 0, reste = 0;
  for (const r of reg) {
    if (!r[6] || faites.has(r[0] + '|' + r[1])) continue;
    if (Date.now() - debut > 5 * 60 * 1000) { reste++; continue; }
    const l = achLignes_(lirePdf_(DriveApp.getFileById(r[6]).getBlob()));
    l.forEach(x => ajout.push([r[0], r[1]].concat(x)));
    if (!l.length) ajout.push([r[0], r[1], '', '(aucun article lu)', 0, 0]);
    n++;
  }
  if (ajout.length) onglet.getRange(onglet.getLastRow() + 1, 1, ajout.length, 6).setValues(ajout);
  Logger.log('Factures détaillées : ' + n + ' | Articles ajoutés : ' + ajout.length + (reste ? ' | Il en reste ' + reste + ' : relance.' : ' | Terminé.'));
}

// Pour me montrer le texte d'une de TES factures (articles facturés au client).
function voirTexteMaFacture() {
  const num = '2601';
  const t = GmailApp.search('from:info@saranelectrique.com subject:"Facture #' + num + '"', 0, 1)[0];
  const pj = t.getMessages()[0].getAttachments({ includeInlineImages: false })[0];
  Logger.log(lirePdf_(pj.copyBlob()).slice(0, 4000));
}

function achDossier_(parent, nom) {
  const it = parent.getFoldersByName(nom);
  return it.hasNext() ? it.next() : (ACH_ESSAI ? parent : parent.createFolder(nom));
}

// Pour me montrer le texte de factures Lumen, Gagnon, Westburne. Lance, envoie-moi le journal.
function voirTexteAchat() {
  ['3362617700.pdf'].forEach(nom => {
    const t = GmailApp.search('from:lumen.ca filename:' + nom, 0, 1)[0];
    if (!t) { Logger.log('Introuvable : ' + nom); return; }
    const pj = t.getMessages().map(m => m.getAttachments().find(a => a.getName() === nom)).find(Boolean);
    Logger.log('===== ' + nom + ' =====\n' + lirePdf_(pj.copyBlob()).slice(0, 3500));
  });
  const reg = SpreadsheetApp.openById(ACH_SHEET).getSheets()[0].getDataRange().getDisplayValues();
  ['Groupe Gagnon', 'Westburne'].forEach(fo => {
    const r = reg.find(x => x[0] === fo && x[6]);
    if (r) Logger.log('===== ' + fo + ' ' + r[1] + ' =====\n' + lirePdf_(DriveApp.getFileById(r[6]).getBlob()).slice(0, 3000));
  });
}

// Rattrapage des références : relit le PDF des factures sans référence (ex. Lumen) et propose la job.
// Ne touche jamais une facture déjà associée. Relance tant qu'il en reste.
function referencesAchats() {
  const debut = Date.now();
  const sh = SpreadsheetApp.openById(ACH_SHEET).getSheets()[0];
  const vals = sh.getDataRange().getDisplayValues();
  const jobs = achJobs_();
  let n = 0, trouvees = 0, reste = 0;
  for (let i = 1; i < vals.length; i++) {
    const r = vals[i];
    if (r[3] || !r[6] || r[8]) continue;                        // a déjà une réf., pas de PDF, ou déjà associée
    if (Date.now() - debut > 5 * 60 * 1000) { reste++; continue; }
    const txt = lirePdf_(DriveApp.getFileById(r[6]).getBlob());
    const lu = achLire_(txt, r[1] + '.pdf');
    n++;
    if (!lu || !lu.ref) { Logger.log(r[0] + ' ' + r[1] + ' : aucune référence trouvée | ' + String(txt).replace(/\s+/g, ' ').slice(0, 200)); continue; }
    const prop = achProposer_(lu.ref, jobs);
    sh.getRange(i + 1, 4).setValue(lu.ref);
    if (prop && !r[7]) sh.getRange(i + 1, 8).setValue(prop);
    trouvees++;
    Logger.log(r[0] + ' ' + r[1] + ' | réf. « ' + lu.ref + ' » | job proposée : ' + (prop || '—'));
  }
  Logger.log('Factures relues : ' + n + ' | Références trouvées : ' + trouvees + (reste ? ' | Il en reste ' + reste + ' : relance.' : ' | Terminé.'));
}

// Ménage du registre des achats : retire ce qui n'est pas une vraie facture
//  - relevés / états de compte (Lumen, fax Gagnon) et soumissions (Westburne) → déplacés dans l'onglet « Documents »
//  - « LUM_MULTIPLE_FACTURES » : ses factures sont maintenant inscrites une à une, la ligne combinée est un doublon
//  - lignes en double (même fournisseur + même numéro)
// Ne touche jamais une facture que tu as déjà associée (colonne Job) : elle est seulement signalée.
// 1re fois : NET_ESSAI = true (montre seulement ce qui serait retiré). Ensuite false pour vraiment le faire.
const NET_ESSAI = true;
function nettoyerAchats() {
  const sh = SpreadsheetApp.openById(ACH_SHEET).getSheets()[0];
  const vals = sh.getDataRange().getDisplayValues();
  const docs = achOngletDocs_();
  const aRetirer = [], vus = new Set();
  for (let i = 1; i < vals.length; i++) {
    const r = vals[i], cle = r[0] + '|' + String(r[1]).replace(/^0+/, '');
    let raison = '', type = '';
    if (vus.has(cle)) raison = 'doublon';
    vus.add(cle);
    if (!raison && /MULTIPLE/i.test(r[1])) raison = 'PDF combiné (factures déjà inscrites une à une)';
    if (!raison && r[0] !== 'Dubo' && r[6]) {
      let t0 = '';
      try { t0 = String(lirePdf_(DriveApp.getFileById(r[6]).getBlob())).slice(0, 500); } catch (e) { Logger.log('Lecture impossible : ' + r[0] + ' ' + r[1] + ' (relance plus tard)'); continue; }
      if (/RELEV[ÉE] DE COMPTE|[ÉE]TAT DE COMPTE|STATEMENT/i.test(t0)) { raison = 'relevé de compte'; type = 'Relevé'; }
      else if (/\bSOUMISSION\b|\bQUOTATION\b/i.test(t0) && !/\bFACTURE\b/i.test(t0)) { raison = 'soumission'; type = 'Soumission'; }
    }
    if (!raison) continue;
    if (r[8]) { Logger.log('À vérifier (déjà associée à « ' + r[8] + ' », pas touchée) : ' + r[0] + ' ' + r[1] + ' — ' + raison); continue; }
    if (type && !NET_ESSAI) docs.appendRow([r[0], type, r[1], r[2], r[3], r[5], r[6], '(retiré du registre des factures)', r[11], 'registre-' + r[0] + '-' + r[1]]);
    aRetirer.push(i + 1);
    Logger.log('Retirée : ' + r[0] + ' ' + r[1] + ' (' + r[2] + ', ' + r[4] + ' $) — ' + raison);
  }
  if (!NET_ESSAI) aRetirer.reverse().forEach(n => sh.deleteRow(n));
  // Articles des factures retirées
  const restants = new Set(sh.getDataRange().getDisplayValues().slice(1).map(r => r[0] + '|' + r[1]));
  const lig = achOngletLignes_(), lv = lig.getDataRange().getDisplayValues();
  let nl = 0;
  for (let i = lv.length - 1; i >= 1; i--) if (!restants.has(lv[i][0] + '|' + lv[i][1])) { if (!NET_ESSAI) lig.deleteRow(i + 1); nl++; }
  Logger.log((NET_ESSAI ? 'MODE ESSAI — rien changé. ' : '') + 'Lignes retirées du registre : ' + aRetirer.length + ' | Articles retirés : ' + nl);
}
