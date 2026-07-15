/**
 * Media Dezign Studio — réception des formulaires du site (devis.html + contact index.html)
 * Web App Apps Script unique : reçoit le POST du site, dispatche selon data.formType,
 * envoie un email au studio et un accusé de réception au client.
 *
 * Déploiement : voir README.md dans ce dossier.
 */

var CONFIG = {
  emailAdmin: 'mdz.studio60@gmail.com',
  expediteurNom: 'Media Dezign Studio',
  champsRequisDevis: ['nom', 'email'],
  champsRequisContact: ['nom', 'email', 'message'],
};

/* ---------- helpers ---------- */

function escapeHtml(s) {
  return String(s == null ? '' : s).replace(/[&<>"']/g, function (c) {
    return ({ '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;', "'": '&#39;' })[c];
  });
}

function estEmailValide(e) {
  return /^[^\s@]+@[^\s@]+\.[^\s@]+$/.test(String(e || ''));
}

function fmt(n) {
  return Math.round(Number(n) || 0).toLocaleString('fr-FR');
}

function fmtRange(mn, mx) {
  return fmt(mn) + ' – ' + fmt(mx) + ' €';
}

function notifierErreur(err, contexte) {
  try {
    GmailApp.sendEmail(CONFIG.emailAdmin, '[ERREUR script site] ' + contexte,
      'Une erreur est survenue :\n\n' + (err && err.stack ? err.stack : err));
  } catch (e) { Logger.log(err); }
}

function parseBody(e) {
  try {
    if (e && e.postData && e.postData.contents) return JSON.parse(e.postData.contents);
  } catch (err) { /* pas du JSON : on retombe sur les paramètres classiques */ }
  return (e && e.parameter) ? e.parameter : {};
}

function json(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj)).setMimeType(ContentService.MimeType.JSON);
}

/* ---------- point d'entrée Web App ---------- */

function doPost(e) {
  try {
    var data = parseBody(e);

    // Anti-spam : champ caché "website" qui doit rester vide (rempli par un bot)
    if (data.website) return json({ ok: true });

    if (data.formType === 'contact') return handleContact(data);
    return handleDevis(data);
  } catch (err) {
    notifierErreur(err, 'doPost');
    return json({ ok: false, error: 'Erreur serveur' });
  }
}

/* ---------- gabarit email commun (bandeau + logotype + footer, charte du site) ---------- */

function emailShell(innerHtml) {
  var primaire = '#E8005A';
  return '' +
  '<!DOCTYPE html><html lang="fr"><head><meta charset="utf-8">' +
  '<meta name="viewport" content="width=device-width,initial-scale=1"></head>' +
  '<body style="margin:0;padding:0;background:#f4f4f5;font-family:Arial,Helvetica,sans-serif;">' +
  '<table role="presentation" width="100%" cellpadding="0" cellspacing="0" style="background:#f4f4f5;padding:24px 0;">' +
  '<tr><td align="center">' +
  '<table role="presentation" width="600" cellpadding="0" cellspacing="0" style="width:600px;max-width:600px;background:#ffffff;border-radius:12px;overflow:hidden;">' +
  '<tr><td style="background:' + primaire + ';height:6px;line-height:6px;font-size:0;">&nbsp;</td></tr>' +
  '<tr><td style="padding:24px 32px 8px;">' +
  '<div style="font-family:Arial Black,Arial,sans-serif;font-size:20px;letter-spacing:.04em;text-transform:uppercase;color:#141414;">STUDIO<span style="color:' + primaire + ';">.AI</span></div>' +
  '<div style="font-size:10.5px;letter-spacing:.18em;text-transform:uppercase;color:#8a8a8e;margin-top:2px;">Media Dezign Studio</div>' +
  '</td></tr>' +
  innerHtml +
  '<tr><td style="padding:18px 32px;background:#fafafa;color:#9ca3af;font-size:12px;line-height:1.6;" align="center">' +
  'Media Dezign Studio — Studio créatif augmenté par l\'IA<br>' + CONFIG.emailAdmin +
  '</td></tr>' +
  '</table></td></tr></table></body></html>';
}

/* =========================================================
   DEVIS (configurateur devis.html)
   ========================================================= */

function handleDevis(data) {
  var manquants = CONFIG.champsRequisDevis.filter(function (c) { return !data[c] || !String(data[c]).trim(); });
  if (manquants.length) return json({ ok: false, error: 'Champs manquants : ' + manquants.join(', ') });
  if (!estEmailValide(data.email)) return json({ ok: false, error: 'Email invalide' });
  if (!data.groups || !data.groups.length) return json({ ok: false, error: 'Devis vide' });

  var pdfBlob = null;
  if (data.pdfBase64) {
    pdfBlob = Utilities.newBlob(Utilities.base64Decode(data.pdfBase64), 'application/pdf', data.pdfFilename || 'Devis.pdf');
  }

  envoyerEmailAdminDevis(data, pdfBlob);
  envoyerEmailClientDevis(data, pdfBlob);
  return json({ ok: true });
}

function buildRecapTexteDevis(d) {
  var lignes = [];
  (d.groups || []).forEach(function (g) {
    lignes.push('');
    lignes.push('[' + g.title + ']');
    (g.items || []).forEach(function (it) {
      lignes.push('• ' + it.label + ' : ' + fmtRange(it.min, it.max));
      if (it.note) lignes.push('   (' + it.note + ')');
    });
  });
  (d.adjustments || []).forEach(function (a) {
    lignes.push('• ' + a.label + ' : ' + fmt(a.min) + ' – ' + fmt(a.max) + ' €');
  });
  lignes.push('');
  lignes.push('ESTIMATION TOTALE : ' + fmtRange(d.totalMin, d.totalMax) + ' HT');
  return lignes.join('\n');
}

function envoyerEmailAdminDevis(d, pdfBlob) {
  var nomComplet = ((d.prenom || '') + ' ' + (d.nom || '')).trim();
  var corps = 'Nouvelle demande de devis — ' + (d.devisRef || '') + '\n\n' +
    '--- CLIENT ---\n' +
    'Nom : ' + nomComplet + '\n' +
    (d.societe ? 'Société : ' + d.societe + '\n' : '') +
    'Email : ' + d.email + '\n' +
    (d.tel ? 'Téléphone : ' + d.tel + '\n' : '') +
    (d.secteur ? 'Secteur : ' + d.secteur + '\n' : '') +
    'Délai souhaité : ' + (d.timelineLabel || '-') + '\n' +
    '\n--- PROJET ---' + buildRecapTexteDevis(d) + '\n' +
    (d.notes ? '\n--- NOTES DU CLIENT ---\n' + d.notes + '\n' : '') +
    '\nReçu le ' + new Date().toLocaleString('fr-FR');

  GmailApp.sendEmail(CONFIG.emailAdmin, 'Nouveau devis — ' + nomComplet + ' (' + (d.devisRef || '') + ')', corps, {
    replyTo: d.email,
    name: CONFIG.expediteurNom,
    attachments: pdfBlob ? [pdfBlob] : [],
  });
}

function buildClientEmailHtmlDevis(d) {
  var nomComplet = escapeHtml(((d.prenom || '') + ' ' + (d.nom || '')).trim() || 'vous');
  var primaire = '#E8005A';

  var lignesItems = '';
  (d.groups || []).forEach(function (g) {
    lignesItems += '<tr><td style="padding:10px 0 2px;font-size:11px;letter-spacing:.1em;text-transform:uppercase;color:' + primaire + ';font-weight:bold;">' + escapeHtml(g.title) + '</td></tr>';
    (g.items || []).forEach(function (it) {
      lignesItems += '<tr><td style="padding:4px 0;font-size:13.5px;color:#3a3a3e;border-bottom:1px dashed #e5e5e7;">' +
        '<table role="presentation" width="100%" cellpadding="0" cellspacing="0"><tr>' +
        '<td style="color:#3a3a3e;font-size:13.5px;">' + escapeHtml(it.label) + '</td>' +
        '<td align="right" style="color:#141414;font-weight:bold;font-size:13.5px;white-space:nowrap;">' + fmtRange(it.min, it.max) + '</td>' +
        '</tr></table></td></tr>';
    });
  });

  var corps =
    '<tr><td style="padding:16px 32px 8px;color:#141414;font-size:15px;line-height:1.6;">' +
    '<h1 style="margin:0 0 14px;font-size:21px;color:#141414;">Merci ' + nomComplet + ', votre devis est bien reçu !</h1>' +
    '<p style="margin:0 0 16px;color:#3a3a3e;">Vous trouverez le récapitulatif ci-dessous ainsi que le PDF détaillé en pièce jointe. ' +
    'Je reviens vers vous sous 24 à 48h avec une proposition affinée.</p>' +
    '</td></tr>' +
    '<tr><td style="padding:0 32px;">' +
    '<table role="presentation" width="100%" cellpadding="0" cellspacing="0" style="background:#f8f9fa;border-radius:8px;">' +
    '<tr><td style="padding:16px 18px;">' +
    '<table role="presentation" width="100%" cellpadding="0" cellspacing="0">' + lignesItems + '</table>' +
    '</td></tr>' +
    '</table>' +
    '</td></tr>' +
    '<tr><td style="padding:16px 32px 0;">' +
    '<table role="presentation" width="100%" cellpadding="0" cellspacing="0" style="background:' + primaire + ';border-radius:10px;">' +
    '<tr>' +
    '<td style="padding:14px 18px;color:#ffffff;font-size:12px;letter-spacing:.1em;text-transform:uppercase;font-weight:bold;">Estimation totale HT</td>' +
    '<td align="right" style="padding:14px 18px;color:#ffffff;font-size:18px;font-weight:bold;white-space:nowrap;">' + fmtRange(d.totalMin, d.totalMax) + '</td>' +
    '</tr>' +
    '</table>' +
    '</td></tr>' +
    '<tr><td style="padding:16px 32px 28px;font-size:11.5px;color:#8a8a8e;">Réf. ' + escapeHtml(d.devisRef || '') + ' · Estimation indicative HT, devis final confirmé après échange.</td></tr>';

  return emailShell(corps);
}

function envoyerEmailClientDevis(d, pdfBlob) {
  var nomComplet = ((d.prenom || '') + ' ' + (d.nom || '')).trim();
  var texte = 'Merci ' + nomComplet + ', votre devis (' + (d.devisRef || '') + ') a bien été reçu.\n' +
    'Le PDF détaillé est en pièce jointe. Je reviens vers vous sous 24 à 48h.\n' +
    buildRecapTexteDevis(d) + '\n\n' + CONFIG.emailAdmin;

  GmailApp.sendEmail(d.email, 'Votre devis ' + (d.devisRef || '') + ' — Media Dezign Studio', texte, {
    htmlBody: buildClientEmailHtmlDevis(d),
    name: CONFIG.expediteurNom,
    replyTo: CONFIG.emailAdmin,
    attachments: pdfBlob ? [pdfBlob] : [],
  });
}

/* =========================================================
   CONTACT (formulaire simple, section #contact de index.html)
   ========================================================= */

function handleContact(data) {
  var manquants = CONFIG.champsRequisContact.filter(function (c) { return !data[c] || !String(data[c]).trim(); });
  if (manquants.length) return json({ ok: false, error: 'Champs manquants : ' + manquants.join(', ') });
  if (!estEmailValide(data.email)) return json({ ok: false, error: 'Email invalide' });

  envoyerEmailAdminContact(data);
  envoyerEmailClientContact(data);
  return json({ ok: true });
}

function envoyerEmailAdminContact(d) {
  var corps = 'Nouveau message de contact\n\n' +
    'Nom : ' + d.nom + '\n' +
    'Email : ' + d.email + '\n' +
    (d.type ? 'Type de projet : ' + d.type + '\n' : '') +
    (d.budget ? 'Budget indicatif : ' + d.budget + '\n' : '') +
    '\nMessage :\n' + d.message + '\n' +
    '\nReçu le ' + new Date().toLocaleString('fr-FR');

  GmailApp.sendEmail(CONFIG.emailAdmin, 'Nouveau contact — ' + d.nom, corps, {
    replyTo: d.email,
    name: CONFIG.expediteurNom,
  });
}

function buildClientEmailHtmlContact(d) {
  var nomComplet = escapeHtml(d.nom || 'vous');
  var corps =
    '<tr><td style="padding:16px 32px 8px;color:#141414;font-size:15px;line-height:1.6;">' +
    '<h1 style="margin:0 0 14px;font-size:21px;color:#141414;">Merci ' + nomComplet + ', votre message est bien reçu !</h1>' +
    '<p style="margin:0 0 16px;color:#3a3a3e;">Je reviens vers vous rapidement avec une première piste et une estimation.</p>' +
    '</td></tr>' +
    '<tr><td style="padding:0 32px 28px;">' +
    '<table role="presentation" width="100%" cellpadding="0" cellspacing="0" style="background:#f8f9fa;border-radius:8px;">' +
    '<tr><td style="padding:16px 18px;font-size:13.5px;color:#3a3a3e;line-height:1.7;">' +
    (d.type ? '<strong style="color:#141414;">Type de projet :</strong> ' + escapeHtml(d.type) + '<br>' : '') +
    (d.budget ? '<strong style="color:#141414;">Budget indicatif :</strong> ' + escapeHtml(d.budget) + '<br>' : '') +
    '<strong style="color:#141414;">Message :</strong><br>' + escapeHtml(d.message || '').replace(/\n/g, '<br>') +
    '</td></tr></table></td></tr>';

  return emailShell(corps);
}

function envoyerEmailClientContact(d) {
  var texte = 'Merci ' + (d.nom || '') + ', votre message a bien été reçu. Je reviens vers vous rapidement.\n\n' + (d.message || '');

  GmailApp.sendEmail(d.email, 'Votre message — Media Dezign Studio', texte, {
    htmlBody: buildClientEmailHtmlContact(d),
    name: CONFIG.expediteurNom,
    replyTo: CONFIG.emailAdmin,
  });
}
