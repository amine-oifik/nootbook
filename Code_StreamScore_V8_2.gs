/**
 * ════════════════════════════════════════════════════════════════════
 *  STREAMSCORE GUIDES — Code.gs V8.2 (conversion + confiance)
 * ════════════════════════════════════════════════════════════════════
 *  Publication (doPost), mémoire (Registry) et redirections : INCHANGÉES,
 *  le notebook continue de fonctionner tel quel.
 *
 *  V8.2 :
 *   • « How this guide was made » n'apparaît plus deux fois.
 *
 *  V8.1 (corrige la V8) :
 *   • titres affichés SANS empreinte « [4821] » (plus de clics depuis Google) ;
 *   • un seul tableau de prix par page (celui de l'article est retiré) ;
 *   • TOUS les liens vers l'offre marqués rel="sponsored", anciens articles
 *     compris (exigé par Google pour un site qui déclare une relation commerciale) ;
 *   • mention de transparence lisible, juste avant l'offre (exigence FTC) ;
 *   • pas de barre d'offre sur les pages About / Author / Contact / politique ;
 *   • balise <title> qui n'est plus coupée par Google, description coupée
 *     sur un mot entier.
 *
 *  V7–V8 :
 *   • nouvelle mise en page appliquée à TOUS les articles (anciens compris) :
 *     auteur, date, transparence, sommaire, encadré auteur, articles liés ;
 *   • données structurées propres (titre sans empreinte, vraie description,
 *     un seul nom de site) + FAQ conservée ;
 *   • pages E-E-A-T : ?p=about · ?p=author · ?p=editorial-policy · ?p=contact
 *   • pages de thème : ?hub=country · device · prices · uses · questions · rankings
 *   • plan du site XML : ?format=sitemap   ·   flux RSS : ?format=rss
 *
 *  INSTALLATION : colle ce fichier à la place de Code.gs, remplis les
 *  réglages ci-dessous, puis Déployer > Gérer les déploiements > Modifier >
 *  Nouvelle version (l'URL …/exec ne change pas).
 * ════════════════════════════════════════════════════════════════════
 */

// ─── Réglages ───────────────────────────────────────────────────────
const TOKEN        = 'Ayob@1221';                 // identique à APPSCRIPT_TOKEN dans le notebook
const FOLDER_NAME  = 'StreamScore Articles';
const SHEET_NAME   = 'Articles';
const SITE_NAME    = 'StreamScore Guides';
const SITE_TAGLINE = 'Honest IPTV guides based on real, hands-on testing';
const AUTHOR_NAME  = '';                          // ⚠️ À REMPLIR : ton prénom ou nom public (signal E-E-A-T important)
const AUTHOR_BIO   = 'I test IPTV services the way you would use them: every evening, on real devices, at peak time. ' +
                     'These guides are built on a six-week hands-on test (10 August – 20 September 2026) on a Fire TV Stick 4K Max, ' +
                     'a Xiaomi TV Box S, a Samsung QLED Series 7 and an iPhone 13.';
const CONTACT_EMAIL = '';                         // ⚠️ À REMPLIR : e-mail public de contact (sinon la page Contact est vide)
const DISCLOSURE   = 'Disclosure: this site has a commercial relationship with SellDigitalHub, the service ranked first. ' +
                     'Its place is based on our own six-week test; the other services are described from the information they publish.';
const MONEY_PAGES  = [                            // pages à faire ranker (liens depuis chaque article)
  ['Best IPTV subscription', 'https://selldigitalhub.com/best-iptv-subscription/'],
  ['IPTV subscription',      'https://selldigitalhub.com/iptv-subscription/']
];
// ─── Offre mise en avant (vrais prix ; remplace les liens par tes liens de commande) ───
const OFFER = {
  brand: 'SellDigitalHub',
  url: 'https://selldigitalhub.com/',            // page principale de l'offre
  trialUrl: 'https://selldigitalhub.com/',       // page de l'essai gratuit 24 h
  currency: '€',
  plans: [                                       // [durée, prix, lien de commande]
    ['1 month',   14.99, 'https://selldigitalhub.com/'],
    ['3 months',  29.99, 'https://selldigitalhub.com/'],
    ['6 months',  44.99, 'https://selldigitalhub.com/'],
    ['12 months', 69.99, 'https://selldigitalhub.com/']
  ],
  best: 3,                                       // index de la formule mise en avant (annuelle)
  utm: 'utm_source=streamscore&utm_medium=guide' // suivi des clics (laisse vide pour le retirer)
};
const RESERVED_    = ['about', 'author', 'editorial-policy', 'contact'];
const FP_RX_       = /\s*\[(?:[A-Z]{2,6}(?:-[A-Z0-9]{3,6})?|\d{3,4})\]\s*$/;

/* ─────────────────────────── Stockage ─────────────────────────── */

function getFolder_() {
  const props = PropertiesService.getScriptProperties();
  const id = props.getProperty('FOLDER_ID');
  if (id) {
    try { return DriveApp.getFolderById(id); } catch (e) { /* recréé plus bas */ }
  }
  const folder = DriveApp.createFolder(FOLDER_NAME);
  props.setProperty('FOLDER_ID', folder.getId());
  return folder;
}

function getSheet_() {
  const props = PropertiesService.getScriptProperties();
  let ss = null;
  const id = props.getProperty('SHEET_ID');
  if (id) {
    try { ss = SpreadsheetApp.openById(id); } catch (e) { ss = null; }
  }
  if (!ss) {
    ss = SpreadsheetApp.create(SITE_NAME + ' — index');
    props.setProperty('SHEET_ID', ss.getId());
  }
  let sh = ss.getSheetByName(SHEET_NAME);
  if (!sh) sh = ss.insertSheet(SHEET_NAME);
  if (sh.getLastRow() === 0) {
    sh.appendRow(['slug', 'title', 'keyword', 'description', 'fileId', 'created', 'updated', 'redirect']);
  } else if (!sh.getRange(1, 8).getValue()) {
    sh.getRange(1, 8).setValue('redirect');          // anciennes feuilles : colonne ajoutee
  }
  return sh;
}

/* Memoire des publications (tous sites) : Colab s'en sert pour ne jamais
   republier une requete ou un titre, meme apres un redemarrage. */
function getRegistry_() {
  const ss = getSheet_().getParent();
  let rg = ss.getSheetByName('Registry');
  if (!rg) rg = ss.insertSheet('Registry');
  if (rg.getLastRow() === 0) rg.appendRow(['date', 'site', 'keyword', 'title', 'url']);
  return rg;
}

function findRow_(sh, slug) {
  const last = sh.getLastRow();
  if (last < 2) return -1;
  const slugs = sh.getRange(2, 1, last - 1, 1).getValues();
  for (let i = 0; i < slugs.length; i++) {
    if (String(slugs[i][0]) === slug) return i + 2;
  }
  return -1;
}

function sanitizeSlug_(value) {
  return String(value || '')
    .toLowerCase()
    .normalize('NFD').replace(/[\u0300-\u036f]/g, '')
    .replace(/[^a-z0-9]+/g, '-')
    .replace(/^-+|-+$/g, '')
    .slice(0, 90) || ('article-' + Date.now());
}

function json_(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}

function escapeHtml_(s) {
  return String(s || '')
    .replace(/&/g, '&amp;').replace(/</g, '&lt;')
    .replace(/>/g, '&gt;').replace(/"/g, '&quot;');
}


/* ─────────────────────────── Réception ─────────────────────────── */

/**
 * Colab envoie : { token, slug, title, keyword, description, html }
 * Réponse      : { ok, url, slug, updated }
 */
function doPost(e) {
  const lock = LockService.getScriptLock();
  try {
    lock.waitLock(30000);

    if (!e || !e.postData || !e.postData.contents) {
      return json_({ ok: false, error: 'Corps de requête vide.' });
    }
    const data = JSON.parse(e.postData.contents);

    if (data.token !== TOKEN) {
      return json_({ ok: false, error: 'Token invalide : vérifie APPSCRIPT_TOKEN dans Colab.' });
    }

    // Memoire : { action: "log", entries: [{site, keyword, title, url}, …] }
    if (data.action === 'log') {
      const rg = getRegistry_();
      const now = new Date();
      const entries = [].concat(data.entries || []);
      entries.forEach(function (x) {
        rg.appendRow([now, String(x.site || ''), String(x.keyword || ''),
                      String(x.title || ''), String(x.url || '')]);
      });
      return json_({ ok: true, logged: entries.length });
    }

    // Redirection : { redirect_from: "slug" ou ["slug", …], redirect_to: "url" }
    if (data.redirect_from && data.redirect_to) {
      const shr = getSheet_();
      const froms = [].concat(data.redirect_from);
      let done = 0;
      froms.forEach(function (f) {
        const r = findRow_(shr, sanitizeSlug_(f));
        if (r > 0) { shr.getRange(r, 8).setValue(String(data.redirect_to)); done++; }
      });
      return json_({ ok: true, redirected: done });
    }

    if (!data.html || !data.title) {
      return json_({ ok: false, error: 'Champs "title" et "html" obligatoires.' });
    }

    const slug  = sanitizeSlug_(data.slug || data.title);
    const title = String(data.title).slice(0, 200);
    const now   = new Date();

    const folder = getFolder_();
    const sh     = getSheet_();
    const row    = findRow_(sh, slug);

    const file = folder.createFile(slug + '.html', String(data.html), MimeType.HTML);

    let updated = false;
    if (row > 0) {
      // Même slug : on remplace l'ancienne version.
      const oldId = sh.getRange(row, 5).getValue();
      try { DriveApp.getFileById(oldId).setTrashed(true); } catch (err) { /* déjà supprimé */ }
      sh.getRange(row, 2, 1, 6).setValues([[
        title, data.keyword || '', data.description || '', file.getId(),
        sh.getRange(row, 6).getValue() || now, now
      ]]);
      sh.getRange(row, 8).setValue('');
      updated = true;
    } else {
      sh.appendRow([slug, title, data.keyword || '', data.description || '',
                    file.getId(), now, now]);
    }

    const base = ScriptApp.getService().getUrl();
    return json_({ ok: true, slug: slug, updated: updated,
                   url: base + '?p=' + encodeURIComponent(slug) });
  } catch (err) {
    return json_({ ok: false, error: String(err) });
  } finally {
    try { lock.releaseLock(); } catch (e2) { /* rien */ }
  }
}




/* ─────────────────────────── Affichage V7 ─────────────────────────── */

function doGet(e) {
  const params = (e && e.parameter) || {};
  if (params.format === 'json') return listJson_();
  if (params.format === 'registry') return registryJson_();
  if (params.format === 'sitemap') return sitemap_();
  if (params.format === 'rss') return rss_();
  if (params.hub) return renderHub_(String(params.hub));
  if (params.p && RESERVED_.indexOf(String(params.p)) !== -1) return renderStatic_(String(params.p));
  if (params.p) return renderArticle_(params.p);
  return renderIndex_(params['new'] || '');
}

function registryJson_() {
  const rg = getRegistry_();
  const last = rg.getLastRow();
  const rows = last > 1 ? rg.getRange(2, 1, last - 1, 5).getValues() : [];
  return json_({ ok: true, count: rows.length, items: rows.map(function (r) {
    return { date: r[0] ? new Date(r[0]).toISOString() : '', site: String(r[1]),
             keyword: String(r[2]), title: String(r[3]), url: String(r[4]) };
  }) });
}

function listJson_() {
  const base = ScriptApp.getService().getUrl();
  const items = rows_(false).map(function (r) {
    return { slug: String(r[0]), title: String(r[1]), keyword: String(r[2]),
             url: base + '?p=' + encodeURIComponent(r[0]),
             updated: r[6] ? new Date(r[6]).toISOString() : '', redirect: String(r[7] || '') };
  });
  return json_({ ok: true, count: items.length, items: items });
}

/* ── Données ── */
function rows_(liveOnly) {
  const sh = getSheet_();
  const last = sh.getLastRow();
  let rows = last > 1 ? sh.getRange(2, 1, last - 1, 8).getValues() : [];
  if (liveOnly) rows = rows.filter(function (r) { return r[0] && !r[7]; });
  return rows;
}
const SHOW_FINGERPRINT = false;                  // true = afficher l'empreinte « [4821] » (déconseillé : moins de clics)
function cleanTitle_(t) { return SHOW_FINGERPRINT ? String(t || '').trim() : String(t || '').replace(FP_RX_, '').trim(); }
function decode_(s) {
  return String(s || '').replace(/&nbsp;/g, ' ').replace(/&quot;/g, '"').replace(/&#39;|&#x27;/g, "'")
    .replace(/&lt;/g, '<').replace(/&gt;/g, '>').replace(/&amp;/g, '&');
}
function stripTags_(s) { return decode_(String(s || '').replace(/<[^>]+>/g, ' ')).replace(/\s+/g, ' ').trim(); }
function fmt_(d) { return d ? Utilities.formatDate(new Date(d), 'GMT', 'd MMMM yyyy') : ''; }
function iso_(d) { return d ? new Date(d).toISOString() : new Date().toISOString(); }
function author_() { return AUTHOR_NAME || ('the ' + SITE_NAME + ' tester'); }

const HUBS_ = [
  ['rankings',  'Best IPTV Subscriptions', /best iptv (subscription|service|providers)|iptv providers|iptv subscription$|reviews|comparison/i],
  ['country',   'IPTV by Country',         /\b(usa|uk|canada|ireland|germany|france|australia|united states|united kingdom)\b/i],
  ['device',    'IPTV by Device',          /firestick|fire tv|smart tv|samsung|lg tv|android tv|apple tv|roku|shield|mag box|formuler|chromecast|iphone|\bpc\b/i],
  ['prices',    'Prices & Plans',          /price|cost|cheap|deal|trial|contract|month|buy/i],
  ['uses',      'IPTV by Use',             /sport|movie|famil|kid|news|4k|epg|catch|vod|ppv|connection|international|reliable/i],
  ['questions', 'Questions & Answers',     /\?|^(what|how|is|can|do|does|why|which)\b/i]
];
function hubOf_(keyword, title) {
  const s = String(keyword || '') + ' ' + String(title || '');
  if (/\?/.test(String(title || ''))) return HUBS_[5];
  for (let i = 1; i < HUBS_.length - 1; i++) if (HUBS_[i][2].test(s)) return HUBS_[i];
  return HUBS_[0];
}

/* ── Extraction du contenu d'un article déjà publié ── */
function extractArticle_(html) {
  const h1m = html.match(/<h1[^>]*>([\s\S]*?)<\/h1>/i);
  let main = '';
  let m = html.match(/<main[^>]*>([\s\S]*?)<\/main>/i);
  if (m) main = m[1];
  if (!main) { m = html.match(/<article[^>]*>([\s\S]*?)<\/article>/i); if (m) main = m[1]; }
  if (!main) { m = html.match(/<body[^>]*>([\s\S]*?)<\/body>/i); main = m ? m[1] : html; }
  const how = (html.match(/<section class="how">[\s\S]*?<\/section>/i) || [''])[0];
  const faq = [];
  (html.match(/<script type="application\/ld\+json">[\s\S]*?<\/script>/gi) || []).forEach(function (s) {
    if (/"FAQPage"/.test(s)) faq.push(s);
  });
  main = main.replace(/<script[\s\S]*?<\/script>/gi, '')
    // le template affiche déjà l'auteur et la transparence : on retire les doublons du texte
    .replace(/<p>\s*<(?:b|strong)>Disclosure:<\/(?:b|strong)>[\s\S]*?<\/p>/gi, '')
    .replace(/<p>\s*<em>By [\s\S]*?<\/em>\s*<\/p>/gi, '')
    .replace(/<p>\s*\*By [^<]*<\/p>/gi, '').replace(/<p>\s*\*\*Disclosure:\*\*[^<]*<\/p>/gi, '')
    // le gabarit affiche son propre tableau de prix : on retire celui de l'article (pas de doublon)
    .replace(/<div class="plans">[\s\S]*?<\/div>/gi, '')
    // l'encadré « How this guide was made » est ajouté une seule fois, à la fin
    .replace(/<section class="how">[\s\S]*?<\/section>/gi, '');
  main = sponsor_(main);
  return { h1: h1m ? stripTags_(h1m[1]) : '', main: main, how: how, faq: faq };
}

/* Liens vers l'offre : rel="sponsored" partout, anciens articles compris */
const OFFER_HOST_RX_ = /<a\b([^>]*?)href="(https?:\/\/(?:www\.)?selldigitalhub\.com[^"]*)"([^>]*)>/gi;
function sponsor_(html) {
  return String(html || '').replace(OFFER_HOST_RX_, function (all, a, href, b) {
    const attrs = (a + ' ' + b).replace(/\s*rel="[^"]*"/gi, '').replace(/\s+/g, ' ').trim();
    return '<a href="' + href + '" rel="sponsored noopener"' + (attrs ? ' ' + attrs : '') + '>';
  });
}

/* Coupe un texte sur un mot entier */
function clip_(s, n) {
  s = String(s || '').trim();
  if (s.length <= n) return s;
  const cut = s.slice(0, n + 1).replace(/\s+\S*$/, '');
  return cut.replace(/[,;:\-–—\s]+$/, '') + '…';
}

function withToc_(main) {
  const toc = [];
  let n = 0;
  const out = main.replace(/<h2([^>]*)>([\s\S]*?)<\/h2>/gi, function (all, attrs, inner) {
    n++;
    const id = 's' + n;
    toc.push('<li><a href="#' + id + '" target="_self">' + stripTags_(inner) + '</a></li>');
    return '<h2 id="' + id + '"' + attrs.replace(/\sid="[^"]*"/, '') + '>' + inner + '</h2>';
  });
  return { html: out, toc: toc.length > 2 ? '<nav class="toc" aria-label="Contents"><b>In this guide</b><ol>' + toc.join('') + '</ol></nav>' : '' };
}


/* ─────────────────────── Composants de conversion (V8) ─────────────────────── */

function planName_(label) { return String(label).replace(/(\d+)\s*months?/i, '$1-month') + ' plan'; }
function money_(n) { return OFFER.currency + Number(n).toFixed(2); }
function link_(url, placement) {
  if (!OFFER.utm) return url;
  return url + (url.indexOf('?') === -1 ? '?' : '&') + OFFER.utm + '&utm_campaign=' + encodeURIComponent(placement);
}
function planMath_(i) {
  const p = OFFER.plans[i], m = parseInt(p[0], 10) || 1, monthly = OFFER.plans[0][1];
  return { label: p[0], price: p[1], url: p[2], months: m, perMonth: p[1] / m, save: Math.max(0, monthly * m - p[1]) };
}
const ICON_ = {
  check: '<svg viewBox="0 0 24 24" width="18" height="18" aria-hidden="true"><path fill="currentColor" d="M9 16.2l-3.5-3.5L4 14.2l5 5 11-11-1.5-1.5z"/></svg>',
  shield: '<svg viewBox="0 0 24 24" width="18" height="18" aria-hidden="true"><path fill="currentColor" d="M12 2l8 3v6c0 5-3.4 9.7-8 11-4.6-1.3-8-6-8-11V5z"/></svg>',
  bolt: '<svg viewBox="0 0 24 24" width="18" height="18" aria-hidden="true"><path fill="currentColor" d="M13 2L4 14h7l-1 8 9-12h-7z"/></svg>'
};

/* Offre principale, en haut de page */
function offerCard_(placement) {
  const b = planMath_(OFFER.best);
  return '<section class="offer" aria-label="Our tested pick">' +
    '<div class="offer-l"><span class="badge">' + ICON_.shield + ' Our tested pick</span>' +
    '<p class="offer-h">' + OFFER.brand + ': the IPTV subscription we used for 6 weeks</p>' +
    '<ul class="ticks"><li>' + ICON_.check + 'Channel changes in about 1.4 s</li><li>' + ICON_.check + 'Stable through 3 big live games</li>' +
    '<li>' + ICON_.check + 'Support reply in about 8 minutes</li><li>' + ICON_.check + 'Catch-up up to 7 days, 4K on supported channels</li></ul></div>' +
    '<div class="offer-r"><div class="ribbon">Best value</div>' +
    '<div class="plan-name">' + planName_(b.label) + '</div>' +
    '<div class="price">' + money_(b.price) + '</div>' +
    '<div class="permonth">≈ ' + money_(b.perMonth) + ' / month' + (b.save ? ' · save ' + money_(b.save) + ' a year' : '') + '</div>' +
    '<a class="btn btn-xl" rel="sponsored noopener" href="' + link_(b.url, placement + '-annual') + '">Get the annual subscription →</a>' +
    '<a class="btn-ghost" rel="sponsored noopener" href="' + link_(OFFER.trialUrl, placement + '-trial') + '">Or try it free for 24 hours — no card</a>' +
    '<div class="micro">' + ICON_.check + '7-day refund · Card, PayPal, crypto · No contract</div></div></section>';
}

/* Bandeau de confiance */
function trustStrip_() {
  const items = ['Tested 6 weeks on 4 devices', '1.4 s channel switching', 'Support in ~8 min',
                 '24-h free trial, no card', '7-day refund', 'Card · PayPal · Crypto'];
  return '<ul class="trust">' + items.map(function (t) { return '<li>' + ICON_.check + t + '</li>'; }).join('') + '</ul>';
}

/* Tableau des 4 formules */
function pricing_(placement) {
  const cards = OFFER.plans.map(function (_p, i) {
    const m = planMath_(i), best = i === OFFER.best;
    return '<div class="pcard' + (best ? ' pbest' : '') + '">' + (best ? '<div class="ribbon">Best value</div>' : '') +
      '<div class="plan-name">' + m.label + '</div><div class="price">' + money_(m.price) + '</div>' +
      '<div class="permonth">' + money_(m.perMonth) + ' / month' + (m.save ? '<br><b>save ' + money_(m.save) + '</b>' : '<br>pay monthly') + '</div>' +
      '<a class="btn' + (best ? '' : ' btn-alt') + '" rel="sponsored noopener" href="' + link_(m.url, placement + '-' + m.months + 'm') + '">Choose ' + m.label + '</a></div>';
  }).join('');
  return '<section class="pricing" aria-label="Plans and prices"><h2>Plans and prices</h2>' +
    '<p class="muted">Same service on every plan. Start with the free 24-hour trial, then pick the length that suits you.</p>' +
    '<div class="pgrid">' + cards + '</div>' + trustStrip_() + '</section>';
}

/* Appel à l'action au milieu de l'article */
function midCta_(placement) {
  const b = planMath_(OFFER.best);
  return '<div class="midcta"><div><b>Ready to try it?</b> Test ' + OFFER.brand + ' free for 24 hours (no card), then save with the ' +
    planName_(b.label) + ' at ≈ ' + money_(b.perMonth) + '/month.</div>' +
    '<a class="btn" rel="sponsored noopener" href="' + link_(OFFER.trialUrl, placement + '-mid') + '">Start the free trial →</a></div>';
}

/* Barre d'offre fixe en bas d'écran (fermable) */
function stickyBar_(placement) {
  const b = planMath_(OFFER.best);
  return '<div class="sticky" id="sticky"><span><b>' + OFFER.brand + ' · ' + planName_(b.label) + '</b> · ' + money_(b.price) +
    ' <span class="hide-s">(≈ ' + money_(b.perMonth) + '/mo) · 24-h free trial</span></span>' +
    '<a class="btn" rel="sponsored noopener" href="' + link_(b.url, placement + '-sticky') + '">Get it →</a>' +
    '<button type="button" aria-label="Close" onclick="document.getElementById(\'sticky\').remove()">×</button></div>';
}

/* ── Gabarit commun ── */
const CSS7_ =
  ':root{--ink:#0b1220;--text:#243041;--muted:#5f6b7a;--line:#e3e8ef;--bg:#f4f6fa;--card:#fff;--brand:#0e7a5a;--brand2:#12a174;' +
  '--accent:#ffb000;--accent2:#ff8a00;--accent-ink:#1e1300;--navy:#0f1b33;--navy2:#16294d}' +
  '*{box-sizing:border-box}html{scroll-behavior:smooth}body{margin:0;background:var(--bg);color:var(--text);' +
  'font:18px/1.7 Inter,system-ui,-apple-system,"Segoe UI",Roboto,sans-serif;padding-bottom:76px}' +
  'a{color:var(--brand)}img{max-width:100%;height:auto}svg{vertical-align:-3px;margin-right:6px}' +
  '.top{background:var(--navy);color:#cfd8e6;box-shadow:0 1px 0 rgba(255,255,255,.06)}' +
  '.top .in{max-width:1120px;margin:0 auto;padding:12px 20px;display:flex;flex-wrap:wrap;gap:6px 22px;align-items:center}' +
  '.top a{color:#e8eef8;text-decoration:none;font-weight:600;font-size:.94rem}.top .brand{font-weight:900;font-size:1.15rem;color:#fff;margin-right:auto;letter-spacing:-.01em}' +
  '.top .brand span{color:var(--accent)}.top .navcta{background:var(--accent);color:var(--accent-ink);padding:7px 14px;border-radius:999px;font-weight:800}' +
  '.wrap{max-width:820px;margin:0 auto;padding:20px 20px 40px}.wide{max-width:1120px}' +
  '.crumbs{font-size:.86rem;color:var(--muted)}.crumbs a{color:var(--muted)}' +
  'h1{font-size:clamp(1.95rem,4.6vw,2.75rem);line-height:1.12;color:var(--ink);margin:.25em 0 .35em;letter-spacing:-.02em}' +
  'h2{font-size:1.5rem;line-height:1.25;color:var(--ink);margin:1.8em 0 .5em;scroll-margin-top:16px;letter-spacing:-.01em}h3{font-size:1.16rem;color:var(--ink);margin:1.3em 0 .4em}' +
  '.byline{display:flex;flex-wrap:wrap;gap:4px 16px;font-size:.92rem;color:var(--muted)}.byline b{color:var(--ink)}' +
  '.offer{display:grid;grid-template-columns:1.25fr 1fr;gap:0;background:linear-gradient(135deg,var(--navy),var(--navy2));color:#e8eef8;' +
  'border-radius:20px;overflow:hidden;box-shadow:0 18px 40px rgba(15,27,51,.28);margin:20px 0 10px}' +
  '.offer-l{padding:22px 24px}.offer-r{position:relative;background:#fff;color:var(--text);padding:22px 22px 18px;text-align:center}' +
  '.badge{display:inline-flex;align-items:center;background:rgba(18,161,116,.18);color:#7ff0c6;border:1px solid rgba(127,240,198,.35);' +
  'border-radius:999px;padding:4px 12px;font-size:.82rem;font-weight:800;text-transform:uppercase;letter-spacing:.04em}' +
  '.offer-h{color:#fff;font-size:1.35rem;font-weight:800;line-height:1.25;margin:.6em 0 .5em}.ticks{list-style:none;margin:0;padding:0;font-size:.97rem}.ticks li{margin:6px 0;color:#d7e1f0}.ticks svg{color:#33d69f}' +
  '.ribbon{position:absolute;top:14px;right:-36px;transform:rotate(35deg);background:var(--accent);color:var(--accent-ink);font-size:.72rem;font-weight:900;' +
  'padding:4px 40px;text-transform:uppercase;letter-spacing:.05em}' +
  '.plan-name{font-weight:800;color:var(--muted);text-transform:uppercase;letter-spacing:.06em;font-size:.82rem}' +
  '.price{font-size:2.6rem;font-weight:900;color:var(--ink);line-height:1.05;margin:4px 0}.permonth{color:var(--muted);font-size:.93rem;margin-bottom:12px}.permonth b{color:var(--brand)}' +
  '.btn{display:inline-block;background:linear-gradient(180deg,var(--accent),var(--accent2));color:var(--accent-ink);font-weight:900;text-decoration:none;' +
  'padding:13px 22px;border-radius:12px;box-shadow:0 6px 16px rgba(255,138,0,.35);transition:transform .08s}.btn:hover{transform:translateY(-1px)}' +
  '.btn-xl{display:block;font-size:1.08rem;padding:15px 18px}.btn-alt{background:var(--card);color:var(--ink);box-shadow:none;border:2px solid var(--line)}' +
  '.btn-ghost{display:block;margin:10px 0 6px;font-weight:700;color:var(--brand);font-size:.95rem}' +
  '.micro{font-size:.82rem;color:var(--muted)}.micro svg{color:var(--brand2)}' +
  '.disc{font-size:.9rem;color:var(--text);background:#fff8e6;border:1px solid #f3d38a;border-radius:12px;padding:9px 14px;margin:14px 0 0}' +
  '.trust{list-style:none;display:flex;flex-wrap:wrap;gap:8px;padding:0;margin:14px 0}.trust li{background:var(--card);border:1px solid var(--line);' +
  'border-radius:999px;padding:6px 12px;font-size:.88rem;font-weight:600;color:var(--ink)}.trust svg{color:var(--brand2)}' +
  '.toc{background:var(--card);border:1px solid var(--line);border-radius:16px;padding:12px 18px;margin:18px 0;font-size:.97rem}.toc ol{margin:.4em 0 0;padding-left:1.2em}' +
  '.answer{background:#e9f8f1;border-left:5px solid var(--brand2);border-radius:12px;padding:14px 16px;margin:16px 0}.answer b{display:block;color:var(--ink)}' +
  '.midcta{display:flex;flex-wrap:wrap;gap:12px;align-items:center;justify-content:space-between;background:#fff7e6;border:2px solid #ffd27a;' +
  'border-radius:16px;padding:16px 18px;margin:22px 0}.midcta div{flex:1 1 320px}' +
  '.tbl{overflow-x:auto;margin:14px 0}table{border-collapse:collapse;width:100%;font-size:.95rem;background:var(--card)}' +
  'th,td{border:1px solid var(--line);padding:9px 11px;text-align:left;vertical-align:top}th{background:var(--navy);color:#fff}tr:nth-child(even) td{background:#f8fafc}' +
  '.cta{background:linear-gradient(135deg,var(--navy),var(--navy2));color:#e8eef8;border-radius:16px;padding:18px 20px;margin:22px 0}.cta p{margin:0 0 10px}' +
  '.cta a{display:inline-block;background:linear-gradient(180deg,var(--accent),var(--accent2));color:var(--accent-ink);font-weight:900;text-decoration:none;padding:12px 20px;border-radius:12px}' +
  'details{background:var(--card);border:1px solid var(--line);border-radius:12px;padding:10px 14px;margin:8px 0}summary{cursor:pointer;font-weight:700;color:var(--ink)}' +
  '.how{background:var(--card);border:1px solid var(--line);border-radius:16px;padding:6px 18px 12px;margin:24px 0}' +
  '.pricing{margin:30px 0}.muted{color:var(--muted)}.pgrid{display:grid;grid-template-columns:repeat(4,1fr);gap:12px}' +
  '.pcard{position:relative;overflow:hidden;background:var(--card);border:2px solid var(--line);border-radius:18px;padding:18px 14px;text-align:center}' +
  '.pcard .price{font-size:1.9rem}.pcard .btn{display:block;padding:11px 10px;font-size:.95rem}' +
  '.pbest{border-color:var(--accent);box-shadow:0 14px 30px rgba(255,138,0,.22);transform:translateY(-4px)}' +
  '.author{display:flex;gap:14px;align-items:flex-start;background:var(--card);border:1px solid var(--line);border-radius:16px;padding:16px;margin:28px 0}' +
  '.avatar{flex:0 0 52px;height:52px;border-radius:50%;background:linear-gradient(135deg,var(--brand),var(--brand2));color:#fff;display:flex;align-items:center;justify-content:center;font-weight:900;font-size:1.3rem}' +
  '.related{background:var(--card);border:1px solid var(--line);border-radius:16px;padding:12px 18px;margin:24px 0}.related ul{margin:.3em 0 0;padding-left:1.1em}' +
  '.grid{display:grid;grid-template-columns:repeat(auto-fill,minmax(290px,1fr));gap:12px}' +
  '.card{display:block;background:var(--card);border:1px solid var(--line);border-radius:16px;padding:16px 18px;text-decoration:none;color:var(--text);transition:box-shadow .15s}' +
  '.card:hover{box-shadow:0 10px 24px rgba(15,27,51,.12)}.card b{display:block;color:var(--ink);line-height:1.3;margin-bottom:6px}.card span{display:block;color:var(--muted);font-size:.9rem}' +
  '.hero{background:radial-gradient(1200px 400px at 10% -10%,#1f3b75 0,transparent 60%),linear-gradient(135deg,var(--navy),#0b2a4a);color:#e8eef8;padding:40px 20px 30px}' +
  '.hero .in{max-width:1120px;margin:0 auto}.hero h1{color:#fff;margin-top:0}.hero p{color:#c9d5e8;max-width:62ch}' +
  '.hubs{display:flex;flex-wrap:wrap;gap:8px;margin:14px 0 0}.hubs a{background:rgba(255,255,255,.1);color:#fff;text-decoration:none;border-radius:999px;padding:6px 14px;font-size:.92rem}' +
  '.sticky{position:fixed;left:0;right:0;bottom:0;z-index:50;display:flex;gap:12px;align-items:center;justify-content:center;flex-wrap:wrap;' +
  'background:rgba(11,18,32,.96);color:#e8eef8;padding:10px 46px 10px 14px;font-size:.95rem;box-shadow:0 -8px 24px rgba(0,0,0,.25)}' +
  '.sticky .btn{padding:9px 16px}.sticky button{position:absolute;right:10px;top:50%;transform:translateY(-50%);background:none;border:0;color:#9fb0c8;font-size:1.5rem;cursor:pointer}' +
  'footer{background:var(--navy);color:#b9c6da;font-size:.9rem}footer .in{max-width:1120px;margin:0 auto;padding:22px 20px;display:flex;flex-wrap:wrap;gap:8px 20px}footer a{color:#e8eef8}' +
  '@media(max-width:760px){.top a:not(.brand):not(.navcta){display:none}.offer{grid-template-columns:1fr}.pgrid{grid-template-columns:1fr 1fr}.hide-s{display:none}.price{font-size:2.2rem}}';

function layout_(o) {
  const base = ScriptApp.getService().getUrl();
  const ld = (o.ld || []).map(function (x) {
    return typeof x === 'string' ? x : '<script type="application/ld+json">' + JSON.stringify(x).replace(/</g, '\\u003c') + '</script>';
  }).join('');
  const nav = '<a class="brand" href="' + base + '">' + SITE_NAME.replace(/(\S+)$/, '<span>$1</span>') + '</a>' +
    '<a href="' + base + '?hub=rankings">Rankings</a><a href="' + base + '?hub=device">By device</a>' +
    '<a href="' + base + '?hub=country">By country</a><a href="' + base + '?hub=prices">Prices</a>' +
    '<a href="' + base + '?p=editorial-policy">How we test</a>' +
    '<a class="navcta" rel="sponsored noopener" href="' + link_(OFFER.url, 'nav') + '">Get ' + OFFER.brand + '</a>';
  const foot = '<span>© ' + new Date().getFullYear() + ' ' + SITE_NAME + ' — ' + SITE_TAGLINE + '</span>' +
    '<a href="' + base + '?p=about">About</a><a href="' + base + '?p=author">Author</a>' +
    '<a href="' + base + '?p=editorial-policy">Editorial policy</a><a href="' + base + '?p=contact">Contact</a>' +
    '<a href="' + base + '?format=sitemap">Sitemap</a>';
  return '<!doctype html><html lang="en"><head><meta charset="utf-8"><base target="_top">' +
    '<link rel="preconnect" href="https://fonts.googleapis.com"><link href="https://fonts.googleapis.com/css2?family=Inter:wght@400;600;800;900&display=swap" rel="stylesheet">' +
    '<meta name="description" content="' + escapeHtml_(o.description || SITE_TAGLINE) + '">' +
    '<style>' + CSS7_ + '</style>' + ld + '</head><body>' +
    '<div class="top"><div class="in">' + nav + '</div></div>' + (o.hero || '') +
    '<div class="wrap' + (o.wide ? ' wide' : '') + '">' + o.body + '</div>' +
    '<footer><div class="in">' + foot + '</div></footer>' + (o.sticky === false ? '' : stickyBar_(o.placement || 'page')) + '</body></html>';
}

function out_(html, title) {
  return HtmlService.createHtmlOutput(html).setTitle(String(title))
    .addMetaTag('viewport', 'width=device-width, initial-scale=1');
}

function orgLd_() {
  const base = ScriptApp.getService().getUrl();
  const o = { '@context': 'https://schema.org', '@type': 'Organization', name: SITE_NAME, url: base };
  if (CONTACT_EMAIL) o.email = CONTACT_EMAIL;
  return o;
}


/* CTA du milieu : après la réponse courte, sinon après la 2e section */
function injectMid_(html, block) {
  const m = html.match(/<div class="answer">[\s\S]*?<\/div>/i);
  if (m) return html.replace(m[0], m[0] + block);
  let n = 0;
  return html.replace(/<h2[\s>]/gi, function (h) { n++; return n === 3 ? block + h : h; });
}

/* ── Article ── */
function renderArticle_(slug) {
  const sh = getSheet_();
  const row = findRow_(sh, sanitizeSlug_(slug));
  if (row < 0) return renderNotFound_();
  const v = sh.getRange(row, 1, 1, 8).getValues()[0];
  const title = cleanTitle_(v[1]);
  if (v[7]) return renderMoved_(String(v[7]), title);
  let stored;
  try { stored = DriveApp.getFileById(v[4]).getBlob().getDataAsString('UTF-8'); } catch (err) { return renderNotFound_(); }

  const base = ScriptApp.getService().getUrl();
  const url = base + '?p=' + encodeURIComponent(v[0]);
  const art = extractArticle_(stored);
  const hub = hubOf_(v[2], v[1]);
  const t = withToc_(art.main);
  const h1 = cleanTitle_(art.h1 || title);
  let desc = String(v[3] || '');
  if (!desc || /^\*|^By /.test(desc)) {
    const a = art.main.match(/<div class="answer">([\s\S]*?)<\/div>/i) || art.main.match(/<p>([\s\S]*?)<\/p>/i);
    desc = a ? stripTags_(a[1]).replace(/^The short answer\s*/i, '') : SITE_TAGLINE;
  }
  desc = clip_(desc, 155);

  const rel = rows_(true).filter(function (r) {
    return String(r[0]) !== String(v[0]) && hubOf_(r[2], r[1])[0] === hub[0];
  }).slice(-6).reverse();
  const related = '<aside class="related"><b>Related guides</b><ul>' +
    MONEY_PAGES.map(function (m) { return '<li><a href="' + m[1] + '" rel="sponsored">' + m[0] + ' — full guide</a></li>'; }).join('') +
    rel.map(function (r) { return '<li><a href="' + base + '?p=' + encodeURIComponent(r[0]) + '">' + escapeHtml_(cleanTitle_(r[1])) + '</a></li>'; }).join('') +
    '</ul></aside>';

  const body =
    '<div class="crumbs"><a href="' + base + '">Home</a> › <a href="' + base + '?hub=' + hub[0] + '">' + hub[1] + '</a></div>' +
    '<h1>' + escapeHtml_(h1) + '</h1>' +
    '<div class="byline"><span>By <a href="' + base + '?p=author">' + escapeHtml_(author_()) + '</a></span>' +
    '<span>Updated <b>' + fmt_(v[6]) + '</b></span><span>Based on a 6-week hands-on test</span></div>' +
    '<p class="disc">' + escapeHtml_(DISCLOSURE) + ' <a href="' + base + '?p=editorial-policy">How we test</a>.</p>' +
    offerCard_('article-top') +
    t.toc + injectMid_(t.html, midCta_('article')) + pricing_('article') + art.how +
    '<section class="author"><div class="avatar">' + escapeHtml_(author_().charAt(0).toUpperCase()) + '</div><div>' +
    '<b>About the author — ' + escapeHtml_(author_()) + '</b><p style="margin:.3em 0 0">' + escapeHtml_(AUTHOR_BIO) +
    ' <a href="' + base + '?p=editorial-policy">Our testing method</a>.</p></div></section>' + related;

  const ld = [
    { '@context': 'https://schema.org', '@type': 'Article', headline: h1.slice(0, 110), description: desc,
      author: { '@type': AUTHOR_NAME ? 'Person' : 'Organization', name: author_(), url: base + '?p=author' },
      publisher: orgLd_(), datePublished: iso_(v[5]), dateModified: iso_(v[6]),
      mainEntityOfPage: { '@type': 'WebPage', '@id': url }, inLanguage: 'en' },
    { '@context': 'https://schema.org', '@type': 'BreadcrumbList', itemListElement: [
      { '@type': 'ListItem', position: 1, name: SITE_NAME, item: base },
      { '@type': 'ListItem', position: 2, name: hub[1], item: base + '?hub=' + hub[0] },
      { '@type': 'ListItem', position: 3, name: h1, item: url }] }
  ].concat(art.faq);
  // au-delà d'environ 60 caractères Google coupe le titre : on n'ajoute le nom du site que s'il tient
  const tabTitle = (h1 + ' | ' + SITE_NAME).length <= 62 ? h1 + ' | ' + SITE_NAME : h1;
  return out_(layout_({ body: body, description: desc, ld: ld, placement: 'article' }), tabTitle);
}

/* ── Accueil ── */
function card_(base, r) {
  return '<a class="card" href="' + base + '?p=' + encodeURIComponent(r[0]) + '"><b>' + escapeHtml_(cleanTitle_(r[1])) + '</b>' +
         '<span>Updated ' + fmt_(r[6]) + '</span></a>';
}

function renderIndex_(onlyDay) {
  const base = ScriptApp.getService().getUrl();
  let rows = rows_(true);
  if (onlyDay) rows = rows.filter(function (r) { return r[6] && Utilities.formatDate(new Date(r[6]), 'GMT', 'yyMMdd') === String(onlyDay); });
  rows.sort(function (a, b) { return new Date(b[6]) - new Date(a[6]); });
  const hero = '<header class="hero"><div class="in"><h1>Honest IPTV Subscription Guides</h1>' +
    '<p>' + escapeHtml_(SITE_TAGLINE) + '. Rankings, prices, setup by device and buyer checklists, with our top pick measured over six weeks on real devices.</p>' +
    '<div class="hubs">' + HUBS_.map(function (h) { return '<a href="' + base + '?hub=' + h[0] + '">' + h[1] + '</a>'; }).join('') + '</div></div></header>';
  const pillars = '<h2>Start here</h2><div class="grid">' + MONEY_PAGES.map(function (m) {
    return '<a class="card" href="' + m[1] + '" rel="sponsored"><b>' + m[0] + ' — the complete guide</b><span>Our main ranking and buyer guide</span></a>';
  }).join('') + '</div>';
  const latest = '<h2>' + (onlyDay ? 'New and updated guides' : 'Latest guides') + '</h2>' +
    (rows.length ? '<div class="grid">' + rows.slice(0, 90).map(function (r) { return card_(base, r); }).join('') + '</div>'
                 : '<p>No guides published yet.</p>');
  const ld = [orgLd_(), { '@context': 'https://schema.org', '@type': 'ItemList', name: SITE_NAME + ' guides',
    itemListElement: rows.slice(0, 50).map(function (r, i) {
      return { '@type': 'ListItem', position: i + 1, name: cleanTitle_(r[1]), url: base + '?p=' + encodeURIComponent(r[0]) }; }) }];
  return out_(layout_({ hero: hero, body: offerCard_('home-top') + trustStrip_() + pillars + pricing_('home') + latest, wide: true, ld: ld, placement: 'home',
    description: 'Honest IPTV subscription guides: rankings, prices, setup by device and buyer checklists, based on real hands-on testing.' }),
    'Honest IPTV Subscription Guides & Rankings | ' + SITE_NAME);
}

/* ── Pages de thème ── */
function renderHub_(key) {
  const base = ScriptApp.getService().getUrl();
  const hub = HUBS_.filter(function (h) { return h[0] === key; })[0] || HUBS_[0];
  const rows = rows_(true).filter(function (r) { return hubOf_(r[2], r[1])[0] === hub[0]; })
    .sort(function (a, b) { return new Date(b[6]) - new Date(a[6]); });
  const body = '<div class="crumbs"><a href="' + base + '">Home</a> › ' + hub[1] + '</div><h1>' + hub[1] + '</h1>' +
    '<p>Every ' + hub[1].toLowerCase() + ' guide on ' + SITE_NAME + ', newest first. Each one is based on our six-week hands-on test.</p>' +
    (rows.length ? '<div class="grid">' + rows.map(function (r) { return card_(base, r); }).join('') + '</div>' : '<p>No guides yet.</p>');
  return out_(layout_({ body: body, wide: true, description: hub[1] + ' — honest IPTV guides from ' + SITE_NAME + '.',
    ld: [{ '@context': 'https://schema.org', '@type': 'CollectionPage', name: hub[1], url: base + '?hub=' + hub[0] }] }),
    hub[1] + ' | ' + SITE_NAME);
}

/* ── Pages E-E-A-T ── */
function renderStatic_(key) {
  const base = ScriptApp.getService().getUrl();
  const pages = {
    'about': ['About ' + SITE_NAME,
      '<p>' + SITE_NAME + ' publishes practical, honest guides to IPTV subscriptions: rankings, prices, setup by device and fixes for common problems.</p>' +
      '<h2>How these guides are made</h2><p>' + escapeHtml_(AUTHOR_BIO) + '</p>' +
      '<h2>How the site is funded</h2><p>' + escapeHtml_(DISCLOSURE) + '</p>' +
      '<p>Read our <a href="' + base + '?p=editorial-policy">editorial policy</a> or <a href="' + base + '?p=contact">contact us</a>.</p>'],
    'author': [author_(),
      '<p>' + escapeHtml_(AUTHOR_BIO) + '</p><h2>What I measured</h2><ul><li>Channel changes: about 1.4 seconds on average</li>' +
      '<li>Stability: 2 short drops, both under 5 seconds, in 24 hours of evening TV</li><li>Live events: Premier League (17 Aug), Champions League (27 Aug), NFL (14 Sep)</li>' +
      '<li>Catch-up: up to 7 days</li><li>Support: first WhatsApp reply in about 8 minutes</li></ul>'],
    'editorial-policy': ['How We Test and Review',
      '<h2>Hands-on testing</h2><p>Our top pick was used every evening for six weeks on four devices. We time channel changes, log every interruption, watch live events at peak time and time support replies.</p>' +
      '<h2>Other services</h2><p>Services we have not tested ourselves are described from the information they publish, and we say so on every page.</p>' +
      '<h2>Independence and disclosure</h2><p>' + escapeHtml_(DISCLOSURE) + '</p>' +
      '<h2>Updates and corrections</h2><p>Prices and services change. Guides are reviewed when they do, and the date at the top of each guide shows the latest update. If you spot an error, <a href="' + base + '?p=contact">tell us</a>.</p>'],
    'contact': ['Contact',
      CONTACT_EMAIL ? '<p>Questions, corrections or feedback: <a href="mailto:' + escapeHtml_(CONTACT_EMAIL) + '">' + escapeHtml_(CONTACT_EMAIL) + '</a>.</p>'
                    : '<p>Questions, corrections or feedback: contact details will be published here shortly.</p>']
  };
  const p = pages[key];
  const body = '<div class="crumbs"><a href="' + base + '">Home</a> › ' + escapeHtml_(p[0]) + '</div><h1>' + escapeHtml_(p[0]) + '</h1>' + p[1];
  const type = key === 'about' ? 'AboutPage' : key === 'contact' ? 'ContactPage' : key === 'author' ? 'ProfilePage' : 'WebPage';
  return out_(layout_({ body: body, sticky: false, description: clip_(stripTags_(p[1]), 155),
    ld: [orgLd_(), { '@context': 'https://schema.org', '@type': type, name: p[0], url: base + '?p=' + key }] }), p[0] + ' | ' + SITE_NAME);
}

/* ── Plan du site et flux ── */
function sitemap_() {
  const base = ScriptApp.getService().getUrl();
  const esc = function (s) { return String(s).replace(/&/g, '&amp;'); };
  const urls = [[base, new Date()]].concat(
    RESERVED_.map(function (k) { return [base + '?p=' + k, new Date()]; }),
    HUBS_.map(function (h) { return [base + '?hub=' + h[0], new Date()]; }),
    rows_(true).map(function (r) { return [base + '?p=' + encodeURIComponent(r[0]), r[6] || new Date()]; }));
  const xml = '<?xml version="1.0" encoding="UTF-8"?><urlset xmlns="http://www.sitemaps.org/schemas/sitemap/0.9">' +
    urls.map(function (u) { return '<url><loc>' + esc(u[0]) + '</loc><lastmod>' + iso_(u[1]).slice(0, 10) + '</lastmod></url>'; }).join('') + '</urlset>';
  return ContentService.createTextOutput(xml).setMimeType(ContentService.MimeType.XML);
}

function rss_() {
  const base = ScriptApp.getService().getUrl();
  const esc = function (s) { return String(s).replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;'); };
  const rows = rows_(true).sort(function (a, b) { return new Date(b[6]) - new Date(a[6]); }).slice(0, 50);
  const xml = '<?xml version="1.0" encoding="UTF-8"?><rss version="2.0"><channel><title>' + esc(SITE_NAME) + '</title><link>' + esc(base) + '</link>' +
    '<description>' + esc(SITE_TAGLINE) + '</description>' +
    rows.map(function (r) {
      const u = base + '?p=' + encodeURIComponent(r[0]);
      return '<item><title>' + esc(cleanTitle_(r[1])) + '</title><link>' + esc(u) + '</link><guid>' + esc(u) + '</guid>' +
             '<pubDate>' + new Date(r[6] || new Date()).toUTCString() + '</pubDate>' +
             (r[3] && !/^\*/.test(String(r[3])) ? '<description>' + esc(String(r[3]).slice(0, 300)) + '</description>' : '') + '</item>';
    }).join('') + '</channel></rss>';
  return ContentService.createTextOutput(xml).setMimeType(ContentService.MimeType.RSS);
}

/* ── Déplacé / introuvable ── */
function renderMoved_(target, title) {
  const safe = escapeHtml_(target);
  const page = '<!doctype html><html lang="en"><head><meta charset="utf-8"><meta http-equiv="refresh" content="0;url=' + safe + '">' +
    '<style>' + CSS7_ + '</style></head><body><div class="wrap" style="text-align:center;padding-top:80px"><h1>This guide has moved</h1>' +
    '<p>An updated version is available.</p><a class="btn" target="_top" href="' + safe + '">Open the updated guide</a></div>' +
    '<script>try{window.top.location.replace(' + JSON.stringify(target) + ')}catch(e){}</script></body></html>';
  return out_(page, cleanTitle_(title) + ' (moved)');
}

function renderNotFound_() {
  const base = ScriptApp.getService().getUrl();
  const body = '<h1>Guide not found</h1><p>This page was moved or removed. All our guides are listed on the home page.</p>' +
    '<a class="btn" href="' + base + '">See all guides</a>';
  return out_(layout_({ body: body, sticky: false, description: 'Guide not found' }), 'Guide not found | ' + SITE_NAME);
}

/* ── Test manuel ── */
function testSetup() {
  Logger.log('Dossier Drive : ' + getFolder_().getUrl());
  Logger.log('Feuille index : ' + getSheet_().getParent().getUrl());
  Logger.log('URL web app   : ' + ScriptApp.getService().getUrl());
}
