/**
 * LA MÉMOIRE — Test de rattrapage CRS (Jonathan Asselin)
 * Passe une seule fois sur les anciens courriels. Ne roule pas aux 15 min.
 *
 * 1) PO d'Odoo « P0xxxx (NOM) »  -> La mémoire / CRS / NOM P0xxxx / (bon de commande + plans)
 * 2) Factures envoyées à facturation.crs@dext.cc -> La mémoire / CRS / À lire / Facture xxxx (date).pdf
 *
 * MODE ESSAI : laisse ESSAI = true la première fois. Le script écrit seulement
 * dans le journal ce qu'il ferait. Quand le journal a l'air bon, mets false et relance.
 */

const ESSAI = true;
const DEPUIS = '2026/06/01';
const DOSSIER_CRS = '1WxSq4thx03QMpUeCEPfo4QcwHfBZX2Hq'; // La mémoire / CRS

function rattrapageCRS() {
  const crs = DriveApp.getFolderById(DOSSIER_CRS);
  const stats = { po: 0, factures: 0, sauves: 0, doublons: 0 };

  // ---------- 1) Les PO de Jonathan ----------
  const fils = rechercher_('from:notifications@cuisinesrivesud17.odoo.com has:attachment after:' + DEPUIS);
  fils.forEach(fil => {
    fil.getMessages().forEach(msg => {
      const m = msg.getSubject().match(/P0(\d{4})\s*\(([^)]+)/i);
      if (!m) return;
      const po = 'P0' + m[1];
      const nom = nettoyer_(m[2]);
      const pdfs = pdfs_(msg);
      if (!pdfs.length) return;
      stats.po++;
      const dossierJob = sousDossier_(crs, nom + ' ' + po);
      pdfs.forEach(pj => sauver_(dossierJob, pj, pj.getName(), stats, nom + ' ' + po));
    });
  });

  // ---------- 2) Les factures envoyées à CRS ----------
  const aLire = sousDossier_(crs, 'À lire');
  const fact = rechercher_('to:facturation.crs@dext.cc subject:Facture has:attachment after:' + DEPUIS);
  fact.forEach(fil => {
    fil.getMessages().forEach(msg => {
      const m = msg.getSubject().match(/Facture\s*#(\d+)\s*\(([^)]+)\)/i);
      if (!m) return;
      const pdfs = pdfs_(msg);
      if (!pdfs.length) return;
      stats.factures++;
      // une seule facture par courriel : on prend le premier PDF
      sauver_(aLire, pdfs[0], 'Facture ' + m[1] + ' (' + m[2] + ').pdf', stats, 'À lire');
    });
  });

  Logger.log('---------------------------------');
  Logger.log((ESSAI ? 'MODE ESSAI — rien créé. ' : '') +
    'PO trouvés : ' + stats.po + ' | Factures trouvées : ' + stats.factures +
    ' | Fichiers ' + (ESSAI ? 'à sauver' : 'sauvés') + ' : ' + stats.sauves +
    ' | Doublons sautés : ' + stats.doublons);
}

// ---------- Outils ----------

function rechercher_(requete) {
  const tout = [];
  for (let debut = 0; ; debut += 100) {
    const lot = GmailApp.search(requete, debut, 100);
    tout.push(...lot);
    if (lot.length < 100) break;
  }
  return tout;
}

// Seulement les vrais PDF. Les logos et images de signature sont ignorés.
function pdfs_(msg) {
  return msg.getAttachments({ includeInlineImages: false })
    .filter(pj => pj.getContentType() === 'application/pdf' || /\.pdf$/i.test(pj.getName()));
}

function nettoyer_(txt) {
  return txt.replace(/\s+/g, ' ').trim();
}

// Trouve le sous-dossier ou le crée (en mode essai, on fait semblant).
const _cache = {};
function sousDossier_(parent, nom) {
  const cle = parent.getId() + '/' + nom;
  if (_cache[cle]) return _cache[cle];
  const it = parent.getFoldersByName(nom);
  let d = it.hasNext() ? it.next() : null;
  if (!d && !ESSAI) d = parent.createFolder(nom);
  _cache[cle] = d || { faux: true, nom: nom, fichiers: {} };
  return _cache[cle];
}

// Anti-doublon : même nom déjà présent dans le dossier = on saute.
function sauver_(dossier, pj, nom, stats, etiquette) {
  if (dossier.faux) {
    if (dossier.fichiers[nom]) { stats.doublons++; return; }
    dossier.fichiers[nom] = true;
    stats.sauves++;
    Logger.log('[essai] ' + etiquette + '  <-  ' + nom);
    return;
  }
  if (dossier.getFilesByName(nom).hasNext()) { stats.doublons++; return; }
  if (ESSAI) {
    stats.sauves++;
    Logger.log('[essai] ' + etiquette + '  <-  ' + nom);
    return;
  }
  dossier.createFile(pj.copyBlob()).setName(nom);
  stats.sauves++;
  Logger.log('Sauvé : ' + etiquette + '  <-  ' + nom);
}
