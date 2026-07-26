// Générateur du deck "Recherche masters 2027" — consomme deck/data/phase-{1..4}.json
const fs = require('fs');
const path = require('path');
const pptxgen = require('pptxgenjs');

const DATA = {};
for (const n of [1, 2, 3, 4]) {
  DATA[n] = JSON.parse(fs.readFileSync(path.join(__dirname, 'data', `phase-${n}.json`), 'utf-8'));
}
const DOMAINS = [];
for (const n of [1, 2, 3, 4]) for (const d of DATA[n].domaines) DOMAINS.push({ ...d, phase: n });
const TR = DATA[4].transversal;

// ---------- palette ----------
const C = {
  dark: '2A2350',      // fond des diapositives sombres
  darkSoft: '3A3168',
  accent: '6D5BD0',
  ink: '26262E',
  muted: '6B6B75',
  bg: 'FFFFFF',
  tint: 'F4F2FB',
  tintRow: 'FAF9FD',
  line: 'D9D5EA',
  white: 'FFFFFF',
  green: '1E8E3E', amber: 'C98A00', red: 'D93025', black: '3B3B41',
};
const F = { head: 'Cambria', body: 'Calibri' };

const pres = new pptxgen();
pres.layout = 'LAYOUT_WIDE'; // 13.33 x 7.5
pres.defineSlideMaster({
  title: 'LIGHT',
  background: { color: C.bg },
  slideNumber: { x: 12.55, y: 7.08, w: 0.6, h: 0.3, fontSize: 9, color: C.muted, fontFace: F.body },
});
pres.defineSlideMaster({
  title: 'DARK',
  background: { color: C.dark },
  slideNumber: { x: 12.55, y: 7.08, w: 0.6, h: 0.3, fontSize: 9, color: 'B9B3D6', fontFace: F.body },
});

// ---------- helpers ----------
const clip = (s, n) => (s && s.length > n ? s.slice(0, n - 1).trimEnd() + '…' : (s || ''));

function slideLight(title, kicker) {
  const s = pres.addSlide({ masterName: 'LIGHT' });
  if (kicker) s.addText(kicker.toUpperCase(), { x: 0.55, y: 0.28, w: 12.2, h: 0.32, fontSize: 11, color: C.accent, fontFace: F.body, bold: true, charSpacing: 2, margin: 0 });
  s.addText(title, { x: 0.55, y: 0.58, w: 12.2, h: 0.85, fontSize: 30, bold: true, color: C.ink, fontFace: F.head, margin: 0 });
  return s;
}
function slideDark(title, subtitle, kicker) {
  const s = pres.addSlide({ masterName: 'DARK' });
  if (kicker) s.addText(kicker.toUpperCase(), { x: 0.9, y: 1.7, w: 11.5, h: 0.4, fontSize: 13, color: 'B9B3D6', bold: true, charSpacing: 3, fontFace: F.body, margin: 0 });
  s.addText(title, { x: 0.9, y: 2.15, w: 11.5, h: 1.6, fontSize: 40, bold: true, color: C.white, fontFace: F.head, margin: 0 });
  if (subtitle) s.addText(subtitle, { x: 0.9, y: 3.9, w: 11.0, h: 1.9, fontSize: 16, color: 'D9D5EA', fontFace: F.body, margin: 0, fit: 'shrink' });
  return s;
}
function bullets(s, items, opt = {}) {
  const o = { x: 0.55, y: 1.6, w: 12.2, h: 5.4, size: 14, color: C.ink, ...opt };
  const arr = items.map((t, i) => ({
    text: t, options: { bullet: { code: '2022', indent: 12 }, breakLine: i < items.length - 1, paraSpaceAfter: 8 },
  }));
  s.addText(arr, { x: o.x, y: o.y, w: o.w, h: o.h, fontSize: o.size, color: o.color, fontFace: F.body, valign: 'top', margin: 0, fit: 'shrink', lineSpacingMultiple: 1.05 });
}
function panel(s, x, y, w, h, title, items, opt = {}) {
  s.addShape('roundRect', { x, y, w, h, rectRadius: 0.06, fill: { color: opt.fill || C.tint }, line: { color: C.line, width: 0.75 } });
  s.addText(title, { x: x + 0.22, y: y + 0.14, w: w - 0.44, h: 0.35, fontSize: 13.5, bold: true, color: opt.titleColor || C.accent, fontFace: F.body, margin: 0 });
  const arr = items.map((t, i) => ({ text: t, options: { bullet: { code: '2022', indent: 10 }, breakLine: i < items.length - 1, paraSpaceAfter: 6 } }));
  s.addText(arr, { x: x + 0.22, y: y + 0.55, w: w - 0.44, h: h - 0.72, fontSize: opt.size || 12.5, color: C.ink, fontFace: F.body, valign: 'top', margin: 0, fit: 'shrink', lineSpacingMultiple: 1.03 });
}
function table(s, header, rows, colW, opt = {}) {
  const trows = [header.map(h => ({ text: h, options: { bold: true, color: C.white, fill: { color: C.dark }, fontSize: opt.hSize || 10.5, valign: 'middle' } }))];
  rows.forEach((r, i) => {
    trows.push(r.map(c => ({ text: c, options: { fill: { color: i % 2 ? C.tintRow : C.bg }, color: C.ink, fontSize: opt.size || 10, valign: 'middle' } })));
  });
  s.addTable(trows, {
    x: opt.x ?? 0.55, y: opt.y ?? 1.6, w: opt.w ?? 12.2, colW,
    border: { type: 'solid', color: C.line, pt: 0.5 }, fontFace: F.body, margin: 0.04, autoPage: false,
  });
}
const chunk = (a, n) => { const r = []; for (let i = 0; i < a.length; i += n) r.push(a.slice(i, i + n)); return r; };

// ============================================================
// OUVERTURE
// ============================================================
{
  const s = pres.addSlide({ masterName: 'DARK' });
  s.addText('RECHERCHE CONSOLIDÉE · 4 PHASES × 3 LANGUES (FR · EN · ES) · 12 PASSES + 4 ARBITRAGES', { x: 0.9, y: 1.5, w: 11.5, h: 0.4, fontSize: 12.5, color: 'B9B3D6', bold: true, charSpacing: 2, fontFace: F.body, margin: 0 });
  s.addText('Étudier à l’étranger : cartographie des masters\net du financement — rentrée 2027', { x: 0.9, y: 2.0, w: 11.6, h: 2.0, fontSize: 38, bold: true, color: C.white, fontFace: F.head, margin: 0 });
  s.addText([
    { text: 'Exécuté le 26 juillet 2026 · cible : rentrée septembre 2027', options: { breakLine: true, paraSpaceAfter: 6 } },
    { text: '10 domaines · ~110 programmes examinés · ~35 dispositifs de financement · 633 sources', options: { breakLine: true, paraSpaceAfter: 6 } },
    { text: 'Profil : ingénieure des systèmes (Bogotá) — IA/agents · 3D/Blender/mocap · danse/biomécanique · recherche accessibilité des parcs (Kennedy)', options: {} },
  ], { x: 0.9, y: 4.35, w: 11.4, h: 1.6, fontSize: 15, color: 'D9D5EA', fontFace: F.body, margin: 0 });
  s.addText('Chaque donnée porte un badge de confiance — aucune décision sans relecture de l’URL source (voir diapositive 2).', { x: 0.9, y: 6.35, w: 11.4, h: 0.6, fontSize: 12, italic: true, color: '9C94C6', fontFace: F.body, margin: 0 });
}
{
  const s = slideLight('Comment lire ce deck — niveaux de confiance', 'Légende');
  const rows = [
    ['🟢', 'Confirmé', 'Trouvé dans ≥2 langues avec URL officielle concordante', 'Utilisable tel quel'],
    ['🟡', 'Partiel', 'Une seule langue et/ou source secondaire', 'À CONFIRMER sur l’URL avant décision'],
    ['🔴', 'Contradictoire', 'Les langues donnent des valeurs différentes', 'Les deux valeurs sont affichées — jamais tranché en silence'],
    ['⚫', 'Absent', 'Introuvable dans les trois passes', '« non trouvé — vérifier sur <URL> ». Jamais d’estimation'],
  ];
  table(s, ['Badge', 'Niveau', 'Critère', 'Traitement'], rows, [0.8, 1.7, 5.4, 4.3], { size: 12, hSize: 12, y: 1.75 });
  bullets(s, [
    'Contexte méthodologique important : les pages officielles étaient largement bloquées (erreurs 403) pendant la recherche — la plupart des faits proviennent de résumés de moteur de recherche citant des URL officielles. Le deck est donc massivement 🟡.',
    'Presque aucune date du cycle 2027 n’était publiée en juillet 2026 : « proj. » = projection du patron 2026, à re-vérifier dès l’automne 2026.',
    'Chaque montant et deadline porte son année de validité. Les témoignages de forums sont étiquetés « anecdotique » et ne portent jamais un 🟢.',
  ], { y: 4.35, size: 13, h: 2.6 });
}
{
  const s = slideLight('Synthèse exécutive — la carte de décision', 'Vue d’ensemble');
  const rows = DOMAINS.map(d => {
    const avg = d.alignement.reduce((a, b) => a + b.score, 0) / d.alignement.length;
    const dots = '●'.repeat(Math.round(avg)) + '○'.repeat(3 - Math.round(avg));
    const top = d.programmes.find(p => p.retenu);
    return [clip(d.nom, 44), dots, top ? clip(`${top.univ} — ${top.intitule}`, 52) : '—', clip(d.risque[0] || '', 60)];
  });
  table(s, ['Domaine', 'Alignement (4 priorités)', 'Cible n°1', 'Risque principal'], rows, [3.6, 1.5, 3.9, 3.2], { size: 9.5, y: 1.55 });
  s.addText('Alignement = moyenne des scores 0-3 sur ses 4 priorités (impact · création visuelle · technique · cerveau/monde), calculée depuis les dossiers de phase. Détail par domaine dans les blocs suivants.', { x: 0.55, y: 6.85, w: 12.2, h: 0.5, fontSize: 10, italic: true, color: C.muted, fontFace: F.body, margin: 0 });
}
{
  const s = slideLight('Les 3 recommandations principales', 'Synthèse exécutive');
  panel(s, 0.55, 1.55, 12.2, 1.75, '1 · Portefeuille Erasmus Mundus — la seule voie de financement intégral sans condition d’expérience ni de retour', [
    'IMLEX (imagerie/XR, semestre au Japon) en tête + EMAI (IA), IMIM (imagerie médicale), SMACCs (smart cities), DEAI, RE_PLAY (jeux) — bourse ≈1 400 €/mois + scolarité + voyage 🟡.',
    'Deadlines groupées nov. 2026 → mars 2027 (IMLEX ~21 janv. 2027 🟡 proj.) ; candidater à 4-6 consortiums mutualise essais et lettres. Aucun frais de dossier.',
  ], { size: 12 });
  panel(s, 0.55, 3.45, 12.2, 1.75, '2 · France publique via bourse Eiffel — le meilleur alignement académique (HCI + visualisation + accessibilité)', [
    'Paris-Saclay HCI (labos AVIZ/ILDA) 🟡, MoSIG Grenoble 🟡, CNN Paris-Saclay (neuro) 🟡 — frais 3 941 €/an 🟡 (2026-27) effacés par le statut boursier Eiffel (~1 200 €/mois 🟡).',
    'Candidature UNIQUEMENT via l’université : deadlines internes oct.–nov. 2026 🟡 (proj.) — le jalon le plus tôt du calendrier. Une seule tentative à vie. Atouts : français + semestre INSA Lyon.',
  ], { size: 12 });
  panel(s, 0.55, 5.35, 12.2, 1.75, '3 · Barcelone / Espagne via Fundación Carolina — la fusion maximale du profil au coût le plus bas', [
    'UPC Neuroingeniería y Rehabilitación (Institut Guttmann) ≈9 496 € 🟡 (2025-26) : serious games + réhabilitation + biomécanique + accessibilité en un seul programme. Alternatives : UPF CSIM, URJC, UPM Neurotecnología (au catalogue Carolina 🟡).',
    'Espagne = visa gratuit, travail 30 h/sem 🟡, sans homologation du titre 🟡. Carolina (~12 janv.–2 mars 2027 🟡 proj.) exige de croiser son catalogue dès l’ouverture ; préférence d’expérience 🔴 à vérifier.',
  ], { size: 12 });
}

// ============================================================
// BLOCS DOMAINES (10 × 7 diapositives)
// ============================================================
const PHASE_NAMES = { 1: 'Phase 1 · Champs-ponts', 2: 'Phase 2 · Science + code + impact', 3: 'Phase 3 · Corps + créatif', 4: 'Phase 4 · Le plus autonome' };
DOMAINS.forEach((d, di) => {
  const kicker = `${PHASE_NAMES[d.phase]} — Domaine ${di + 1}/10`;
  // 1. divider + définition
  slideDark(d.nom, d.definition.join('  ·  '), kicker);
  // 2. alignement
  {
    const s = slideLight('Alignement avec ses 4 priorités', `${kicker} — ${clip(d.nom, 60)}`);
    d.alignement.forEach((a, i) => {
      const y = 1.75 + i * 1.28;
      s.addShape('roundRect', { x: 0.55, y, w: 12.2, h: 1.1, rectRadius: 0.05, fill: { color: C.tint }, line: { color: C.line, width: 0.75 } });
      s.addText(a.priorite, { x: 0.8, y: y + 0.12, w: 3.4, h: 0.45, fontSize: 14, bold: true, color: C.ink, fontFace: F.body, margin: 0 });
      const filled = Math.max(0, Math.min(3, a.score));
      for (let k = 0; k < 3; k++) {
        s.addShape('ellipse', { x: 0.82 + k * 0.42, y: y + 0.62, w: 0.3, h: 0.3, fill: { color: k < filled ? C.accent : 'DDD9EE' }, line: { color: C.accent, width: 0.5 } });
      }
      s.addText(a.note, { x: 4.4, y: y + 0.1, w: 8.1, h: 0.92, fontSize: 12, color: C.ink, fontFace: F.body, margin: 0, valign: 'middle', fit: 'shrink' });
    });
  }
  // 3. quotidien
  {
    const s = slideLight('Le quotidien — étudiante, puis professionnelle', `${kicker} — ${clip(d.nom, 60)}`);
    panel(s, 0.55, 1.6, 5.95, 3.1, 'En master', d.quotidienMaster, {});
    panel(s, 6.8, 1.6, 5.95, 3.1, 'À 3 ans d’expérience', d.quotidienPro, {});
    panel(s, 0.55, 4.9, 12.2, 2.15, 'État du marché 2026', d.marche2026, {});
  }
  // 4. biblio + 20h
  {
    const s = slideLight('Bibliographie canonique & parcours d’entrée en 20 h', `${kicker} — ${clip(d.nom, 60)}`);
    panel(s, 0.55, 1.6, 6.75, 5.45, 'Les références qui comptent', d.biblio, { size: 12 });
    panel(s, 7.55, 1.6, 5.2, 5.45, 'Tester le domaine en 20 heures', d.parcours20h, { size: 12 });
  }
  // 5. programmes (tableaux, ≤7 lignes)
  {
    const hasPf = d.programmes.some(p => p.portfolio !== undefined);
    const header = hasPf
      ? ['Université — programme', 'Pays', 'Frais non-UE', 'Deadline 2027', 'Portfolio', 'Financement clé']
      : ['Université — programme', 'Pays', 'Frais non-UE', 'Deadline 2027', 'Financement clé'];
    const colW = hasPf ? [3.5, 1.0, 2.2, 2.2, 1.1, 2.2] : [3.7, 1.05, 2.35, 2.45, 2.65];
    const rows = d.programmes.map(p => {
      const base = [
        clip(`${p.univ} — ${p.intitule}`, 72) + (p.retenu ? '  ★' : ''),
        clip(p.pays, 14),
        clip(`${p.fraisBadge} ${p.frais}`, 46),
        clip(`${p.deadlineBadge} ${p.deadline}`, 46),
      ];
      if (hasPf) base.push(clip(p.portfolio || '⚫', 20));
      base.push(clip(`${p.finBadge} ${p.financement}`, 48));
      return base;
    });
    chunk(rows, 7).forEach((part, pi, all) => {
      const s = slideLight(`Programmes de master${all.length > 1 ? ` (${pi + 1}/${all.length})` : ''}`, `${kicker} — ${clip(d.nom, 60)}`);
      table(s, header, part, colW, { size: 9.5, y: 1.6 });
      s.addText('★ = programme retenu (fiche détaillée en annexe). Tous les faits : voir badge + année de validité ; URL dans les dossiers de phase et en annexe.', { x: 0.55, y: 6.9, w: 12.2, h: 0.45, fontSize: 9.5, italic: true, color: C.muted, fontFace: F.body, margin: 0 });
    });
  }
  // 6. fusion
  {
    const s = slideLight('Fusion avec son profil CS / IA', `${kicker} — ${clip(d.nom, 60)}`);
    bullets(s, d.fusion, { y: 1.7, size: 14.5 });
  }
  // 7. risque
  {
    const s = slideLight('Risque — le scénario où elle regretterait ce choix', `${kicker} — ${clip(d.nom, 60)}`);
    s.addShape('roundRect', { x: 0.55, y: 1.7, w: 12.2, h: 4.9, rectRadius: 0.06, fill: { color: 'FBF1F0' }, line: { color: 'E5B9B5', width: 0.75 } });
    const arr = d.risque.map((t, i) => ({ text: t, options: { bullet: { code: '2022', indent: 12 }, breakLine: i < d.risque.length - 1, paraSpaceAfter: 10 } }));
    const sl = s;
    sl.addText(arr, { x: 0.9, y: 2.0, w: 11.5, h: 4.3, fontSize: 14.5, color: '7A2E27', fontFace: F.body, valign: 'top', margin: 0, fit: 'shrink', lineSpacingMultiple: 1.06 });
  }
});

// ============================================================
// TRANSVERSAL
// ============================================================
slideDark('Transversal', 'Comparatif global · programmes hybrides · financement · rétroplanning · risques · ce qui reste à vérifier', 'Consolidation des 4 phases');

// comparatif global des retenus
{
  const rows = [];
  DOMAINS.forEach(d => d.programmes.filter(p => p.retenu).forEach(p => rows.push([
    clip(`${p.univ} — ${p.intitule}`, 66), clip(p.pays, 14), clip(d.nom.split('/')[0].split('(')[0], 26),
    clip(`${p.fraisBadge} ${p.frais}`, 40), clip(`${p.deadlineBadge} ${p.deadline}`, 40), clip(`${p.finBadge} ${p.financement}`, 42),
  ])));
  chunk(rows, 7).forEach((part, pi, all) => {
    const s = slideLight(`Tableau comparatif global — programmes retenus (${pi + 1}/${all.length})`, 'Transversal');
    table(s, ['Université — programme', 'Pays', 'Domaine', 'Frais non-UE', 'Deadline 2027', 'Financement'], part, [3.35, 0.95, 1.9, 2.0, 2.0, 2.0], { size: 9, y: 1.55 });
  });
}
// hybrides
{
  const s = slideLight('Programmes hybrides — sa meilleure zone', 'Transversal');
  s.addText('Les programmes qui croisent plusieurs de ses domaines sont ceux où son profil atypique (code + 3D + mouvement + accessibilité) cesse d’être « dispersé » pour devenir l’argument central.', { x: 0.55, y: 1.5, w: 12.2, h: 0.65, fontSize: 12.5, italic: true, color: C.muted, fontFace: F.body, margin: 0 });
  const rows = TR.hybrides.map(h => [clip(h.nom, 44), clip(h.ou, 22), clip(`${h.badge || ''} ${h.pourquoi}`, 88), clip(h.url, 46)]);
  table(s, ['Programme', 'Où', 'Pourquoi c’est sa zone', 'URL'], rows.slice(0, 8), [3.1, 1.6, 4.9, 2.6], { size: 9.5, y: 2.25 });
}
// EMJM
{
  const rows = TR.emjm.map(e => [clip(e.nom, 22), clip(e.domaine, 24), clip(e.consortium, 30), clip(e.montant, 30), clip(e.deadline, 34), clip(e.fit, 44)]);
  chunk(rows, 7).forEach((part, pi, all) => {
    const s = slideLight(`Erasmus Mundus — les EMJM pertinents${all.length > 1 ? ` (${pi + 1}/${all.length})` : ''}`, 'Transversal · financement intégral, sans condition d’expérience ni de retour');
    table(s, ['EMJM', 'Domaine', 'Consortium', 'Bourse', 'Deadline 2027', 'Adéquation profil'], part, [1.5, 1.7, 2.3, 2.1, 2.2, 2.4], { size: 9, y: 1.6 });
    if (pi === all.length - 1) s.addText(clip('Écartés : ' + (DATA[4].transversal.emjmEcartes || []).join(' · '), 260), { x: 0.55, y: 6.55, w: 12.2, h: 0.8, fontSize: 9.5, italic: true, color: C.muted, fontFace: F.body, margin: 0, fit: 'shrink' });
  });
}
// Chevening
{
  const s = slideLight('Chevening (UK) — le cas limite à trancher en août 2026', 'Transversal');
  bullets(s, TR.chevening, { y: 1.7, size: 13.5 });
}
// financement — deux diapositives : colombien/hispano puis reste
{
  const fin = TR.financement || [];
  const isCol = f => /colfuturo|icetex|carolina|minciencias|cancillería|cancilleria|aecid/i.test(f.dispositif);
  const colRows = fin.filter(isCol);
  const others = fin.filter(f => !isCol(f));
  const mk = f => [clip(f.dispositif, 30), clip(f.couverture, 44), clip(f.eligibilite, 52), clip(`${f.badge || ''} ${f.deadline}`, 34), clip(f.verdict, 16)];
  chunk(colRows.map(mk), 7).forEach((part, pi, all) => {
    const s = slideLight(`Financement colombien & hispanophone${all.length > 1 ? ` (${pi + 1}/${all.length})` : ''}`, 'Transversal · ce que seul l’espagnol a révélé');
    table(s, ['Dispositif', 'Couverture', 'Éligibilité pour elle', 'Deadline 2027', 'Verdict'], part, [1.9, 3.0, 3.4, 2.3, 1.6], { size: 9.5, y: 1.6 });
  });
  chunk(others.map(mk), 7).forEach((part, pi, all) => {
    const s = slideLight(`Financement — autres dispositifs (${pi + 1}/${all.length})`, 'Transversal · Eiffel, EM, bourses internes, Asie, wildcards');
    table(s, ['Dispositif', 'Couverture', 'Éligibilité pour elle', 'Deadline 2027', 'Verdict'], part, [1.9, 3.0, 3.4, 2.3, 1.6], { size: 9.5, y: 1.6 });
  });
}
// options non envisagées
{
  const s = slideLight('Options qu’elle n’avait pas envisagées', 'Transversal');
  const items = TR.optionsNonEnvisagees.map(o => `${o.nom} — ${o.description}`);
  bullets(s, items, { y: 1.7, size: 13.5 });
}
// rétroplanning
{
  chunk(TR.retroplanning, 4).forEach((part, pi, all) => {
    const s = slideLight(`Rétroplanning juillet 2026 → rentrée 2027 (${pi + 1}/${all.length})`, 'Transversal');
    part.forEach((per, i) => {
      const y = 1.6 + i * 1.36;
      s.addShape('roundRect', { x: 0.55, y, w: 2.5, h: 1.2, rectRadius: 0.05, fill: { color: C.dark } });
      s.addText(per.periode, { x: 0.7, y: y + 0.05, w: 2.2, h: 1.1, fontSize: 13, bold: true, color: C.white, fontFace: F.body, valign: 'middle', margin: 0, fit: 'shrink' });
      const arr = per.jalons.map((t, k) => ({ text: t, options: { bullet: { code: '2022', indent: 10 }, breakLine: k < per.jalons.length - 1, paraSpaceAfter: 4 } }));
      s.addText(arr, { x: 3.35, y: y + 0.02, w: 9.4, h: 1.28, fontSize: 10.5, color: C.ink, fontFace: F.body, valign: 'middle', margin: 0, fit: 'shrink', lineSpacingMultiple: 1.0 });
    });
  });
}
// logistique
{
  const s = slideLight('Logistique documentaire & visas depuis Bogotá', 'Transversal');
  bullets(s, TR.logistique, { y: 1.7, size: 13 });
}
// analyse de risque par domaine
{
  const rows = DOMAINS.map(d => [clip(d.nom, 46), clip(d.risque[0] || '', 105)]);
  chunk(rows, 5).forEach((part, pi, all) => {
    const s = slideLight(`Analyse de risque par domaine, sans complaisance (${pi + 1}/${all.length})`, 'Transversal');
    table(s, ['Domaine', 'Le scénario de regret'], part, [4.0, 8.2], { size: 11, y: 1.6 });
  });
}
// contradictions 🔴
{
  const allContr = [];
  for (const n of [1, 2, 3, 4]) (DATA[n].contradictions || []).forEach(c => allContr.push([`P${n}`, clip(c.sujet, 34), clip(c.valeurA, 44), clip(c.valeurB, 44), clip(c.lecture, 52)]));
  chunk(allContr, 6).forEach((part, pi, all) => {
    const s = slideLight(`🔴 Registre des contradictions (${pi + 1}/${all.length})`, 'Ce que nous n’avons pas pu trancher — les deux valeurs sont affichées');
    table(s, ['Phase', 'Sujet', 'Valeur A (source)', 'Valeur B (source)', 'Lecture'], part, [0.7, 2.3, 3.0, 3.0, 3.2], { size: 8.5, y: 1.6 });
  });
}
// non vérifiés ⚫
{
  const allNv = [];
  for (const n of [1, 2, 3, 4]) (DATA[n].nonVerifies || []).forEach(x => allNv.push(`[P${n}] ${x}`));
  chunk(allNv, 13).forEach((part, pi, all) => {
    const s = slideLight(`⚫ Ce que nous n’avons pas pu vérifier (${pi + 1}/${all.length})`, 'À consulter soi-même avant toute décision');
    bullets(s, part.map(t => clip(t, 150)), { y: 1.55, size: 10.5 });
  });
}

// ============================================================
// ANNEXES
// ============================================================
slideDark('Annexes', 'Fiches des programmes retenus · bibliographies intégrales · index des sources', '');

DOMAINS.forEach(d => {
  (d.annexes || []).forEach(a => {
    const s = slideLight(clip(a.titre, 78), `Annexe · ${clip(d.nom, 64)}`);
    bullets(s, a.lignes.map(l => clip(l, 190)), { y: 1.6, size: 12.5 });
  });
});
DOMAINS.forEach(d => {
  const items = (d.biblioComplete || []).map(t => clip(t, 140));
  chunk(items, 14).forEach((part, pi, all) => {
    const s = slideLight(`Bibliographie intégrale${all.length > 1 ? ` (${pi + 1}/${all.length})` : ''}`, `Annexe · ${clip(d.nom, 64)}`);
    bullets(s, part, { y: 1.55, size: 10.5 });
  });
});
{
  const s = slideLight('Index des sources', 'Annexe');
  bullets(s, [
    '633 URLs consultées le 2026-07-26, listées intégralement dans masters-research/sources.md (dépôt git), groupées par passe de recherche.',
    'Dossiers consolidés avec sources par fait : phase-1-dossier.md · phase-2-dossier.md · phase-3-dossier.md · phase-4-dossier.md.',
    'Rapports bruts par langue : masters-research/raw/phase-{1..4}-{FR,EN,ES}.md.',
    'Rappel : aucune passe n’a pu lire directement la plupart des pages officielles (403) — les URL citées sont le point de départ obligatoire de toute vérification.',
  ], { y: 1.8, size: 14 });
}

pres.writeFile({ fileName: path.join(__dirname, '..', 'masters-2027-deck.pptx') }).then(() => {
  console.log('deck écrit : masters-research/masters-2027-deck.pptx');
});
