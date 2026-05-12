/**
 * ══════════════════════════════════════════════════════════
 *  PILOTAGE DES CONTRÔLES FISCAUX — app.js
 *  Architecture : Data Engine · Charts (D3) · UI · Filters
 * ══════════════════════════════════════════════════════════
 */

'use strict';

/* ─────────────────────────────────────────────────────────
   1. ÉTAT GLOBAL
───────────────────────────────────────────────────────── */
const STATE = {
  rawData:      [],   // données brutes importées
  filteredData: [],   // données après filtres
  page:         1,
  pageSize:     15,
  searchTerm:   '',
};

/* ─────────────────────────────────────────────────────────
   2. UTILITAIRES
───────────────────────────────────────────────────────── */

/** Convertit un numéro de série Excel en Date JS */
function excelDateToJS(serial) {
  if (!serial || isNaN(+serial)) return null;
  const n = +serial;
  // Excel epoch : 1er janv 1900 (avec le bug du 29 fév 1900)
  const d = new Date(Date.UTC(1899, 11, 30) + n * 86400000);
  return isNaN(d.getTime()) ? null : d;
}

/** Parse une valeur de date (numéro Excel ou string) */
function parseDate(val) {
  if (!val) return null;
  if (val instanceof Date) return isNaN(val.getTime()) ? null : val;
  if (typeof val === 'number') return excelDateToJS(val);
  if (typeof val === 'string') {
    const d = new Date(val);
    return isNaN(d.getTime()) ? null : d;
  }
  return null;
}

/** Formate une date en jj/mm/aaaa */
function fmtDate(d) {
  if (!d) return '–';
  return d.toLocaleDateString('fr-FR');
}

/** Différence en jours entre deux dates */
function daysDiff(a, b) {
  if (!a || !b) return null;
  return Math.round((b - a) / 86400000);
}

/** Formate un nombre avec séparateur de milliers */
function fmtNum(n, dec = 0) {
  if (n === null || n === undefined || isNaN(n)) return '–';
  return n.toLocaleString('fr-FR', { minimumFractionDigits: dec, maximumFractionDigits: dec });
}

/** Formate un pourcentage */
function fmtPct(n) {
  if (n === null || n === undefined || isNaN(n)) return '–';
  return (n >= 0 ? '+' : '') + n.toFixed(1) + '%';
}

/** Toast notification */
function toast(msg, dur = 3000) {
  const el = document.getElementById('toast');
  el.textContent = msg;
  el.classList.remove('hidden');
  clearTimeout(toast._t);
  toast._t = setTimeout(() => el.classList.add('hidden'), dur);
}

/** Couleurs palette D3 */
const PALETTE = ['#1a56db','#059669','#d97706','#dc2626','#7c3aed','#0891b2','#db2777','#65a30d'];

/* ─────────────────────────────────────────────────────────
   3. PARSER XLSX (via JSZip + DOMParser)
───────────────────────────────────────────────────────── */

/**
 * Lit un fichier .xlsx et retourne un tableau de lignes en JSON.
 * Utilise JSZip pour décompresser et DOMParser pour lire le XML.
 */
async function parseXLSX(file) {
  const buf = await file.arrayBuffer();
  const zip = await JSZip.loadAsync(buf);

  // 1. Lire les chaînes partagées
  const sst = await zip.file('xl/sharedStrings.xml')?.async('text');
  const strings = [];
  if (sst) {
    const doc = new DOMParser().parseFromString(sst, 'application/xml');
    doc.querySelectorAll('si').forEach(si => {
      const ts = si.querySelectorAll('t');
      strings.push([...ts].map(t => t.textContent).join(''));
    });
  }

  // 2. Lire la 1ère feuille
  const sheetFiles = Object.keys(zip.files).filter(f => f.match(/xl\/worksheets\/sheet\d+\.xml/));
  if (!sheetFiles.length) throw new Error('Aucune feuille trouvée');
  const sheetXml = await zip.file(sheetFiles[0]).async('text');
  const sheet = new DOMParser().parseFromString(sheetXml, 'application/xml');

  // 3. Parser les cellules
  const rows = [];
  sheet.querySelectorAll('row').forEach(row => {
    const cells = {};
    row.querySelectorAll('c').forEach(c => {
      const ref  = c.getAttribute('r');         // ex: A3
      const type = c.getAttribute('t');
      const v    = c.querySelector('v')?.textContent;
      const col  = ref.replace(/[0-9]/g, '');   // ex: A, BC

      let val = null;
      if (type === 's') val = strings[+v] ?? '';
      else if (type === 'b') val = v === '1';
      else if (v !== undefined && v !== null) val = isNaN(+v) ? v : +v;

      cells[col] = val;
    });
    rows.push(cells);
  });

  if (!rows.length) return [];

  // 4. Convertir en tableau objet avec en-têtes
  const header = rows[0];
  // Map colonnes lettre → nom d'en-tête
  const colMap = {};
  Object.entries(header).forEach(([col, val]) => { colMap[col] = String(val ?? col); });

  return rows.slice(1).map(row => {
    const obj = {};
    Object.entries(colMap).forEach(([col, name]) => {
      obj[name] = row[col] ?? null;
    });
    return obj;
  }).filter(r => Object.values(r).some(v => v !== null));
}

/** Parser CSV */
function parseCSV(text) {
  const lines = text.split('\n').map(l => l.trim()).filter(Boolean);
  if (!lines.length) return [];
  const sep = lines[0].includes(';') ? ';' : ',';
  const headers = lines[0].split(sep).map(h => h.replace(/^"|"$/g, '').trim());
  return lines.slice(1).map(line => {
    const vals = line.split(sep).map(v => v.replace(/^"|"$/g, '').trim());
    const obj = {};
    headers.forEach((h, i) => {
      const v = vals[i] ?? '';
      obj[h] = v === '' ? null : (isNaN(+v) ? v : +v);
    });
    return obj;
  });
}

/* ─────────────────────────────────────────────────────────
   4. ENRICHISSEMENT DES DONNÉES
───────────────────────────────────────────────────────── */
function enrichData(rawRows) {
  return rawRows.map((row, idx) => {
    // Normaliser les noms de colonnes (trim + suppression BOM)
    const r = {};
    Object.entries(row).forEach(([k, v]) => { r[k.trim().replace(/^\uFEFF/, '')] = v; });

    // Dates
    const dateCree      = parseDate(r['DATE_CREE'] || r['DATE CREE']);
    const dateNP_chef   = parseDate(r['DATE VALID NP CHEF CENTRE']);
    const dateNP_central= parseDate(r['DATE VALID NP CENTRAL']);
    const dateND_chef   = parseDate(r['DATE VALID ND CHEF CENTRE']);
    const dateND_central= parseDate(r['DATE VALID ND CENTRAL']);
    const dateTO_chef   = parseDate(r['DATE VALID TO CHEF CENTRE']);
    const dateTO_central= parseDate(r['DATE VALID TO CENTRAL']);

    // Date NP = la plus récente entre chef et central
    const dateNP = dateNP_central || dateNP_chef;
    // Date ND = la plus récente entre chef et central
    const dateND = dateND_central || dateND_chef;
    // Date TO
    const dateTO = dateTO_central || dateTO_chef;

    // Statut
    const hasCloture = !!(dateND || dateTO);
    const statut = hasCloture ? 'Clôturé' : 'En cours';
    const hasTO  = !!(dateTO_chef || dateTO_central || r['VALID TO CENTRAL']);

    // Montants
    const mntPpalNP = +r['MNT PPAL Not Prim'] || 0;
    const mntAmdNP  = +r['MNT AMD Not Prim']  || 0;
    const mntPpalND = +r['MNT PPAL Not Def']  || 0;
    const mntAmdND  = +r['MNT AMD Not Def']   || 0;
    const montantNP = mntPpalNP + mntAmdNP;
    const montantND = mntPpalND + mntAmdND;
    const ecartPct  = montantNP > 0 ? ((montantND - montantNP) / montantNP * 100) : null;

    // Délais
    const delaiNPND    = daysDiff(dateNP, dateND);
    const delaiTotal   = daysDiff(dateCree, dateND);

    // Aging
    let aging = null;
    if (delaiNPND !== null) {
      if (delaiNPND <= 15)      aging = '0-15 j (Rapide)';
      else if (delaiNPND <= 30) aging = '16-30 j (Normal)';
      else if (delaiNPND <= 60) aging = '31-60 j (À surveiller)';
      else                      aging = '60+ j (En retard)';
    }

    // Centre / Vérificateur avec valeurs par défaut
    const centre = r['CENTRE GESTIONNAIRE'] || r['CENTRE GESTIONNAIRE '] || 'Non renseigné';
    const verif  = r['VERIFICATEUR'] || r['VERIFICATEUR '] || 'Non renseigné';
    const type   = r['TYPE CONTRÔLE'] || r['TYPE CONTROLE'] || 'Non renseigné';
    const ref    = r['REFERENCE'] || `Ref ${String(idx + 1).padStart(3, '0')}`;

    // Extraire l'année de création
    const anneeCreation = dateCree ? dateCree.getFullYear() : null;

    return {
      ...r,
      _ref:          ref,
      _type:         type,
      _centre:       centre,
      _verif:        verif,
      _statut:       statut,
      _hasTO:        hasTO,
      _dateCree:     dateCree,
      _dateNP:       dateNP,
      _dateND:       dateND,
      _dateTO:       dateTO,
      _montantNP:    montantNP,
      _montantND:    montantND,
      _mntPpalNP:    mntPpalNP,
      _mntAmdNP:     mntAmdNP,
      _mntPpalND:    mntPpalND,
      _mntAmdND:     mntAmdND,
      _ecartPct:     ecartPct,
      _delaiNPND:    delaiNPND,
      _delaiTotal:   delaiTotal,
      _aging:        aging,
      _anneeCreation:anneeCreation,
    };
  });
}

/* ─────────────────────────────────────────────────────────
   5. FILTRES
───────────────────────────────────────────────────────── */
function applyFilters() {
  const types   = getSelectValues('f-type');
  const centres = getSelectValues('f-centre');
  const verifs  = getSelectValues('f-verif');
  const annee   = document.getElementById('f-annee').value;

  const statutChecked = getCheckValues('f-statut');
  const toChecked     = getCheckValues('f-to');

  STATE.filteredData = STATE.rawData.filter(d => {
    if (types.length   && !types.includes(d._type))    return false;
    if (centres.length && !centres.includes(d._centre)) return false;
    if (verifs.length  && !verifs.includes(d._verif))  return false;
    if (annee && String(d._anneeCreation) !== annee)   return false;
    if (!statutChecked.includes(d._statut))            return false;
    if (d._hasTO  && !toChecked.includes('oui'))       return false;
    if (!d._hasTO && !toChecked.includes('non'))       return false;
    return true;
  });

  STATE.page = 1;
  document.getElementById('sidebar-count').textContent = STATE.filteredData.length;
  refreshAll();
}

function getSelectValues(id) {
  const sel = document.getElementById(id);
  return [...sel.selectedOptions].map(o => o.value).filter(v => v !== '');
}

function getCheckValues(id) {
  return [...document.getElementById(id).querySelectorAll('input:checked')].map(i => i.value);
}

/* ─────────────────────────────────────────────────────────
   6. REMPLISSAGE DES FILTRES
───────────────────────────────────────────────────────── */
function populateFilters() {
  const unique = key => [...new Set(STATE.rawData.map(d => d[key]).filter(Boolean))].sort();

  fillSelect('f-type',   unique('_type'));
  fillSelect('f-centre', unique('_centre'));
  fillSelect('f-verif',  unique('_verif'));

  const annees = [...new Set(STATE.rawData.map(d => d._anneeCreation).filter(Boolean))].sort();
  fillSelect('f-annee', annees, false);
}

function fillSelect(id, values, multi = true) {
  const sel = document.getElementById(id);
  const empty = multi ? '<option value="">Tous</option>' : '<option value="">Toutes</option>';
  sel.innerHTML = empty + values.map(v => `<option value="${v}">${v}</option>`).join('');
}

/* ─────────────────────────────────────────────────────────
   7. CALCUL DES KPI
───────────────────────────────────────────────────────── */
function computeKPI(data) {
  const total     = data.length;
  const clotures  = data.filter(d => d._statut === 'Clôturé').length;
  const encours   = total - clotures;
  const tauxCloture = total ? (clotures / total * 100) : 0;
  const withTO    = data.filter(d => d._hasTO).length;
  const tauxTO    = total ? (withTO / total * 100) : 0;

  const totalNP   = data.reduce((s, d) => s + (d._montantNP || 0), 0);
  const totalND   = data.reduce((s, d) => s + (d._montantND || 0), 0);
  const ecarts    = data.map(d => d._ecartPct).filter(v => v !== null);
  const ecartMoy  = ecarts.length ? ecarts.reduce((a, b) => a + b, 0) / ecarts.length : null;

  const totalPpalNP = data.reduce((s, d) => s + (d._mntPpalNP || 0), 0);
  const totalAmdNP  = data.reduce((s, d) => s + (d._mntAmdNP  || 0), 0);
  const totalPpalND = data.reduce((s, d) => s + (d._mntPpalND || 0), 0);
  const totalAmdND  = data.reduce((s, d) => s + (d._mntAmdND  || 0), 0);

  const delais    = data.map(d => d._delaiNPND).filter(v => v !== null && v >= 0);
  const delaiMoy  = delais.length ? delais.reduce((a, b) => a + b, 0) / delais.length : null;
  const delaiMin  = delais.length ? Math.min(...delais) : null;
  const delaiMax  = delais.length ? Math.max(...delais) : null;
  const enRetard  = delais.filter(d => d > 60).length;

  const centresTO = [...new Set(data.filter(d => d._hasTO).map(d => d._centre))].length;
  const verifsTO  = [...new Set(data.filter(d => d._hasTO).map(d => d._verif))].length;

  return {
    total, clotures, encours, tauxCloture,
    withTO, tauxTO,
    totalNP, totalND, ecartMoy,
    totalPpalNP, totalAmdNP, totalPpalND, totalAmdND,
    delaiMoy, delaiMin, delaiMax, enRetard,
    centresTO, verifsTO,
  };
}

/* ─────────────────────────────────────────────────────────
   8. MISE À JOUR DES KPI CARDS
───────────────────────────────────────────────────────── */
function updateKPICards(kpi) {
  setText('kpi-total',         fmtNum(kpi.total));
  setText('kpi-clotures',      fmtNum(kpi.clotures));
  setText('kpi-encours',       fmtNum(kpi.encours));
  setText('kpi-taux-cloture',  kpi.tauxCloture.toFixed(1) + '%');
  setText('kpi-to',            fmtNum(kpi.withTO));
  setText('kpi-taux-to',       kpi.tauxTO.toFixed(1) + '%');
  setText('kpi-np',            fmtNum(kpi.totalNP));
  setText('kpi-nd',            fmtNum(kpi.totalND));
  setText('kpi-ecart',         fmtPct(kpi.ecartMoy));

  // Finance tab
  setText('fin-np-ppal', fmtNum(kpi.totalPpalNP));
  setText('fin-np-amd',  fmtNum(kpi.totalAmdNP));
  setText('fin-nd-ppal', fmtNum(kpi.totalPpalND));
  setText('fin-nd-amd',  fmtNum(kpi.totalAmdND));

  // Délais tab
  setText('del-moy',    kpi.delaiMoy !== null ? Math.round(kpi.delaiMoy) + ' j' : '–');
  setText('del-min',    kpi.delaiMin !== null ? kpi.delaiMin + ' j' : '–');
  setText('del-max',    kpi.delaiMax !== null ? kpi.delaiMax + ' j' : '–');
  setText('del-retard', fmtNum(kpi.enRetard));

  // TO tab
  setText('to-count',   fmtNum(kpi.withTO));
  setText('to-taux',    kpi.tauxTO.toFixed(1) + '%');
  setText('to-centres', fmtNum(kpi.centresTO));
  setText('to-verifs',  fmtNum(kpi.verifsTO));
}

function setText(id, val) {
  const el = document.getElementById(id);
  if (el) el.textContent = val;
}

/* ─────────────────────────────────────────────────────────
   9. GRAPHIQUES D3
───────────────────────────────────────────────────────── */

/** PIE CHART générique */
function drawPie(selector, data, labelFn, valueFn) {
  const svg = d3.select(selector);
  svg.selectAll('*').remove();

  const W = svg.node().parentElement.clientWidth || 320;
  const H = Math.min(W, 260);
  const R = Math.min(W, H) / 2 - 30;

  svg.attr('viewBox', `0 0 ${W} ${H}`);

  const g = svg.append('g').attr('transform', `translate(${W * 0.38},${H / 2})`);

  const pie  = d3.pie().value(d => valueFn(d)).sort(null);
  const arc  = d3.arc().innerRadius(R * 0.5).outerRadius(R);
  const arcH = d3.arc().innerRadius(R * 0.5).outerRadius(R + 6);
  const color = d3.scaleOrdinal(PALETTE);

  const arcs = g.selectAll('.arc')
    .data(pie(data))
    .join('g').attr('class', 'arc');

  arcs.append('path')
    .attr('d', arc)
    .attr('fill', (d, i) => color(i))
    .attr('stroke', '#fff')
    .attr('stroke-width', 2)
    .style('cursor', 'pointer')
    .on('mouseover', function(event, d) {
      d3.select(this).attr('d', arcH);
    })
    .on('mouseout', function(event, d) {
      d3.select(this).attr('d', arc);
    });

  // Légende à droite
  const legendX = W * 0.38 + R + 20;
  const legendStartY = H / 2 - (data.length * 18) / 2;
  data.forEach((d, i) => {
    svg.append('rect')
      .attr('x', legendX).attr('y', legendStartY + i * 22 - 8)
      .attr('width', 12).attr('height', 12)
      .attr('rx', 2).attr('fill', color(i));
    svg.append('text')
      .attr('x', legendX + 16).attr('y', legendStartY + i * 22)
      .attr('dominant-baseline', 'middle')
      .attr('font-size', '11px').attr('fill', '#334155')
      .text(`${labelFn(d)} (${valueFn(d)})`);
  });
}

/** BAR CHART horizontal */
function drawBarH(selector, data, labelFn, valueFn, color = '#1a56db') {
  const svg = d3.select(selector);
  svg.selectAll('*').remove();

  const W = svg.node().parentElement.clientWidth || 400;
  const m = { top: 10, right: 50, bottom: 10, left: 120 };
  const H = Math.max(data.length * 36 + m.top + m.bottom, 100);

  svg.attr('viewBox', `0 0 ${W} ${H}`);

  const x = d3.scaleLinear()
    .domain([0, d3.max(data, valueFn) * 1.1 || 1])
    .range([0, W - m.left - m.right]);

  const y = d3.scaleBand()
    .domain(data.map(labelFn))
    .range([0, H - m.top - m.bottom])
    .padding(0.2);

  const g = svg.append('g').attr('transform', `translate(${m.left},${m.top})`);

  g.selectAll('rect')
    .data(data)
    .join('rect')
    .attr('y', d => y(labelFn(d)))
    .attr('height', y.bandwidth())
    .attr('x', 0)
    .attr('width', d => x(valueFn(d)))
    .attr('rx', 4)
    .attr('fill', (d, i) => typeof color === 'function' ? color(d, i) : color)
    .style('cursor', 'default');

  // Labels valeurs
  g.selectAll('.val-label')
    .data(data)
    .join('text')
    .attr('class', 'val-label')
    .attr('x', d => x(valueFn(d)) + 5)
    .attr('y', d => y(labelFn(d)) + y.bandwidth() / 2)
    .attr('dominant-baseline', 'middle')
    .attr('font-size', '11px')
    .attr('fill', '#334155')
    .text(d => fmtNum(valueFn(d)));

  // Labels axes
  g.selectAll('.axis-label')
    .data(data)
    .join('text')
    .attr('class', 'axis-label')
    .attr('x', -5)
    .attr('y', d => y(labelFn(d)) + y.bandwidth() / 2)
    .attr('dominant-baseline', 'middle')
    .attr('text-anchor', 'end')
    .attr('font-size', '11px')
    .attr('fill', '#64748b')
    .text(d => {
      const lbl = String(labelFn(d));
      return lbl.length > 16 ? lbl.slice(0, 15) + '…' : lbl;
    });
}

/** BAR CHART vertical */
function drawBarV(selector, data, labelFn, valueFn, color = '#1a56db', fmt = fmtNum) {
  const svg = d3.select(selector);
  svg.selectAll('*').remove();

  const W = svg.node().parentElement.clientWidth || 400;
  const m = { top: 30, right: 10, bottom: 60, left: 55 };
  const H = 260;

  svg.attr('viewBox', `0 0 ${W} ${H}`);

  const allVals = data.map(valueFn);
  const minV = d3.min(allVals) || 0;
  const maxV = d3.max(allVals) || 1;
  const yMin = Math.min(0, minV * 1.1);
  const yMax = maxV * 1.1;

  const x = d3.scaleBand()
    .domain(data.map(labelFn))
    .range([0, W - m.left - m.right])
    .padding(0.2);

  const y = d3.scaleLinear()
    .domain([yMin, yMax])
    .range([H - m.bottom, m.top]);

  const g = svg.append('g').attr('transform', `translate(${m.left},0)`);

  // Axe Y
  g.append('g')
    .call(d3.axisLeft(y).ticks(4).tickFormat(d => fmt(d)))
    .call(g => g.selectAll('.domain, .tick line').attr('stroke', '#e2e8f0'))
    .call(g => g.selectAll('.tick text').attr('font-size', '10px').attr('fill', '#64748b'));

  // Ligne zéro
  if (yMin < 0) {
    g.append('line')
      .attr('x1', 0).attr('x2', W - m.left - m.right)
      .attr('y1', y(0)).attr('y2', y(0))
      .attr('stroke', '#94a3b8').attr('stroke-width', 1);
  }

  // Barres
  g.selectAll('rect')
    .data(data)
    .join('rect')
    .attr('x', d => x(labelFn(d)))
    .attr('width', x.bandwidth())
    .attr('y', d => valueFn(d) >= 0 ? y(valueFn(d)) : y(0))
    .attr('height', d => Math.abs(y(valueFn(d)) - y(0)))
    .attr('rx', 3)
    .attr('fill', (d, i) => typeof color === 'function' ? color(d, i) : color);

  // Labels X
  g.selectAll('.x-label')
    .data(data)
    .join('text')
    .attr('class', 'x-label')
    .attr('x', d => x(labelFn(d)) + x.bandwidth() / 2)
    .attr('y', H - m.bottom + 12)
    .attr('text-anchor', 'middle')
    .attr('font-size', '10px')
    .attr('fill', '#64748b')
    .text(d => {
      const lbl = String(labelFn(d));
      return lbl.length > 8 ? lbl.slice(0, 7) + '…' : lbl;
    });

  // Labels valeurs au-dessus
  g.selectAll('.val-label')
    .data(data)
    .join('text')
    .attr('class', 'val-label')
    .attr('x', d => x(labelFn(d)) + x.bandwidth() / 2)
    .attr('y', d => y(Math.max(0, valueFn(d))) - 4)
    .attr('text-anchor', 'middle')
    .attr('font-size', '9px')
    .attr('fill', '#334155')
    .text(d => fmt(valueFn(d)));
}

/** BAR CHART NP vs ND groupé */
function drawGroupedBarNPND(selector, data) {
  const svg = d3.select(selector);
  svg.selectAll('*').remove();

  const W = svg.node().parentElement.clientWidth || 500;
  const m = { top: 20, right: 20, bottom: 70, left: 60 };
  const H = 280;

  svg.attr('viewBox', `0 0 ${W} ${H}`);

  const inner = W - m.left - m.right;
  const keys = ['_montantNP', '_montantND'];
  const colorsMap = { '_montantNP': '#1a56db', '_montantND': '#059669' };

  const x0 = d3.scaleBand().domain(data.map(d => d._ref)).range([0, inner]).padding(0.2);
  const x1 = d3.scaleBand().domain(keys).range([0, x0.bandwidth()]).padding(0.05);
  const y  = d3.scaleLinear()
    .domain([0, d3.max(data, d => Math.max(d._montantNP, d._montantND)) * 1.15 || 1])
    .range([H - m.bottom, m.top]);

  const g = svg.append('g').attr('transform', `translate(${m.left},0)`);

  g.append('g')
    .call(d3.axisLeft(y).ticks(4).tickFormat(fmtNum))
    .call(ax => ax.selectAll('.domain,.tick line').attr('stroke', '#e2e8f0'))
    .call(ax => ax.selectAll('.tick text').attr('font-size', '10px').attr('fill', '#64748b'));

  const barGroup = g.selectAll('.bar-group')
    .data(data).join('g').attr('class', 'bar-group')
    .attr('transform', d => `translate(${x0(d._ref)},0)`);

  barGroup.selectAll('rect')
    .data(d => keys.map(k => ({ key: k, val: d[k] || 0 })))
    .join('rect')
    .attr('x', d => x1(d.key))
    .attr('width', x1.bandwidth())
    .attr('y', d => y(d.val))
    .attr('height', d => H - m.bottom - y(d.val))
    .attr('rx', 3)
    .attr('fill', d => colorsMap[d.key]);

  g.selectAll('.x-label')
    .data(data).join('text').attr('class', 'x-label')
    .attr('x', d => x0(d._ref) + x0.bandwidth() / 2)
    .attr('y', H - m.bottom + 14)
    .attr('text-anchor', 'middle')
    .attr('font-size', '10px').attr('fill', '#64748b')
    .text(d => d._ref);

  // Légende
  const legendData = [{ label: 'NP', color: '#1a56db' }, { label: 'ND', color: '#059669' }];
  const legG = svg.append('g').attr('transform', `translate(${W / 2 - 50}, 5)`);
  legendData.forEach((ld, i) => {
    legG.append('rect').attr('x', i * 70).attr('width', 12).attr('height', 12).attr('rx', 2).attr('fill', ld.color);
    legG.append('text').attr('x', i * 70 + 16).attr('y', 9).attr('font-size', '11px').attr('fill', '#334155').text(ld.label);
  });
}

/* ─────────────────────────────────────────────────────────
   10. REFRESH DE TOUS LES GRAPHIQUES
───────────────────────────────────────────────────────── */
function refreshCharts(data, kpi) {
  const counts = d => d.length;

  // — Vue d'ensemble —
  // PIE: type de contrôle
  const byType = groupCount(data, '_type');
  drawPie('#chart-type-pie', byType, d => d.key, d => d.val);

  // PIE: statut
  const byStatut = groupCount(data, '_statut');
  drawPie('#chart-statut-pie', byStatut, d => d.key, d => d.val);

  // BAR H: vérificateurs
  const byVerif = groupCount(data, '_verif').sort((a, b) => b.val - a.val);
  drawBarH('#chart-verif-bar', byVerif, d => d.key, d => d.val);

  // — Finance —
  if (data.length) {
    const topNP = data.slice().sort((a, b) => b._montantNP - a._montantNP).slice(0, 10);
    drawGroupedBarNPND('#chart-np-nd-bar', topNP);

    // PIE: Principal vs Amende NP
    const npSplit = [
      { key: 'Principal', val: kpi.totalPpalNP },
      { key: 'Amende',    val: kpi.totalAmdNP  },
    ];
    drawPie('#chart-np-split', npSplit, d => d.key, d => d.val);

    // Ecart % bar
    const withEcart = data.filter(d => d._ecartPct !== null);
    drawBarV('#chart-ecart-bar', withEcart, d => d._ref, d => d._ecartPct,
      (d) => d._ecartPct >= 0 ? '#059669' : '#dc2626',
      n => fmtNum(n, 1) + '%'
    );

    // PIE: Principal vs Amende ND
    const ndSplit = [
      { key: 'Principal', val: kpi.totalPpalND },
      { key: 'Amende',    val: kpi.totalAmdND  },
    ];
    drawPie('#chart-nd-split', ndSplit, d => d.key, d => d.val);
  }

  // — Délais —
  // PIE aging
  const agingCounts = groupCount(data.filter(d => d._aging), '_aging');
  const agingOrder = ['0-15 j (Rapide)', '16-30 j (Normal)', '31-60 j (À surveiller)', '60+ j (En retard)'];
  const agingColMap = {
    '0-15 j (Rapide)': '#059669',
    '16-30 j (Normal)': '#1a56db',
    '31-60 j (À surveiller)': '#d97706',
    '60+ j (En retard)': '#dc2626',
  };
  const agingSorted = agingOrder
    .map(k => agingCounts.find(a => a.key === k))
    .filter(Boolean);
  drawPie('#chart-aging', agingSorted, d => d.key, d => d.val);

  // BAR: délai NP→ND par dossier
  const withDelai = data.filter(d => d._delaiNPND !== null);
  drawBarV('#chart-delai-bar', withDelai, d => d._ref, d => d._delaiNPND,
    (d) => (d._delaiNPND > 60 ? '#dc2626' : d._delaiNPND > 30 ? '#d97706' : '#059669')
  );

  // BAR: délai total
  const withDelaiTotal = data.filter(d => d._delaiTotal !== null);
  drawBarV('#chart-delai-total-bar', withDelaiTotal, d => d._ref, d => d._delaiTotal, '#7c3aed');

  // BAR H: délai moyen par vérificateur
  const verifGroups = {};
  withDelai.forEach(d => {
    if (!verifGroups[d._verif]) verifGroups[d._verif] = [];
    verifGroups[d._verif].push(d._delaiNPND);
  });
  const delaiVerif = Object.entries(verifGroups).map(([k, v]) => ({
    key: k,
    val: Math.round(v.reduce((a, b) => a + b, 0) / v.length)
  })).sort((a, b) => b.val - a.val);
  drawBarH('#chart-delai-verif', delaiVerif, d => d.key, d => d.val);

  // — TO —
  const toData = data.filter(d => d._hasTO);
  const toPie = [
    { key: 'Avec TO',   val: toData.length },
    { key: 'Sans TO',   val: data.length - toData.length },
  ];
  drawPie('#chart-to-pie', toPie, d => d.key, d => d.val);

  const toByType = groupCount(toData, '_type');
  drawBarH('#chart-to-type', toByType, d => d.key, d => d.val, '#dc2626');
}

/** Groupe et compte par clé */
function groupCount(data, key) {
  const map = {};
  data.forEach(d => {
    const k = d[key] || 'Non renseigné';
    map[k] = (map[k] || 0) + 1;
  });
  return Object.entries(map).map(([key, val]) => ({ key, val })).sort((a, b) => b.val - a.val);
}

/* ─────────────────────────────────────────────────────────
   11. MATRICE KPI
───────────────────────────────────────────────────────── */
function renderMatrix(data) {
  const container = document.getElementById('matrix-table');
  const types = [...new Set(data.map(d => d._type))].sort();

  const metrics = [
    { label: 'Nombre de dossiers',    fn: arr => arr.length },
    { label: 'Dossiers clôturés',     fn: arr => arr.filter(d => d._statut === 'Clôturé').length },
    { label: 'En cours',              fn: arr => arr.filter(d => d._statut === 'En cours').length },
    { label: 'Taux de clôture',       fn: arr => arr.length ? (arr.filter(d => d._statut === 'Clôturé').length / arr.length * 100).toFixed(1) + '%' : '–' },
    { label: 'Taxations d\'Office',   fn: arr => arr.filter(d => d._hasTO).length },
    { label: 'Total NP',              fn: arr => fmtNum(arr.reduce((s, d) => s + d._montantNP, 0)) },
    { label: 'Total ND',              fn: arr => fmtNum(arr.reduce((s, d) => s + d._montantND, 0)) },
    { label: 'Délai moyen (j)',       fn: arr => { const d = arr.map(x => x._delaiNPND).filter(v => v !== null); return d.length ? Math.round(d.reduce((a, b) => a + b, 0) / d.length) + ' j' : '–'; } },
  ];

  const total = data;
  let html = '<table><thead><tr><th>Indicateur</th>';
  types.forEach(t => { html += `<th>${t}</th>`; });
  html += '<th>TOTAL</th></tr></thead><tbody>';

  metrics.forEach(m => {
    html += `<tr><td><strong>${m.label}</strong></td>`;
    types.forEach(t => {
      const sub = data.filter(d => d._type === t);
      html += `<td>${m.fn(sub)}</td>`;
    });
    html += `<td><strong>${m.fn(total)}</strong></td></tr>`;
  });

  html += '</tbody></table>';
  container.innerHTML = html;
}

/* ─────────────────────────────────────────────────────────
   12. TABLE TO
───────────────────────────────────────────────────────── */
function renderTOTable(data) {
  const toData = data.filter(d => d._hasTO);
  const wrap = document.getElementById('to-table-wrap');
  if (!toData.length) { wrap.innerHTML = '<p style="color:#64748b;padding:1rem">Aucun dossier avec Taxation d\'Office</p>'; return; }

  let html = `<table><thead><tr>
    <th>Référence</th><th>Type</th><th>Vérificateur</th><th>Centre</th>
    <th>Date TO Chef</th><th>Date TO Central</th><th>Montant NP</th><th>Montant ND</th>
  </tr></thead><tbody>`;
  toData.forEach(d => {
    html += `<tr>
      <td>${d._ref}</td>
      <td><span class="badge badge-red">${d._type}</span></td>
      <td>${d._verif}</td>
      <td>${d._centre}</td>
      <td>${fmtDate(d._dateTO)}</td>
      <td>${fmtDate(d._dateTO)}</td>
      <td>${fmtNum(d._montantNP)}</td>
      <td>${fmtNum(d._montantND)}</td>
    </tr>`;
  });
  html += '</tbody></table>';
  wrap.innerHTML = html;
}

/* ─────────────────────────────────────────────────────────
   13. TABLEAU DÉTAILLÉ PAGINÉ
───────────────────────────────────────────────────────── */
function renderDetailTable() {
  const term = STATE.searchTerm.toLowerCase();
  const filtered = term
    ? STATE.filteredData.filter(d =>
        Object.values(d).some(v => String(v).toLowerCase().includes(term))
      )
    : STATE.filteredData;

  const total = filtered.length;
  const pages = Math.max(1, Math.ceil(total / STATE.pageSize));
  STATE.page = Math.min(STATE.page, pages);
  const start = (STATE.page - 1) * STATE.pageSize;
  const slice = filtered.slice(start, start + STATE.pageSize);

  const wrap = document.getElementById('detail-table-wrap');
  let html = `<table><thead><tr>
    <th>Référence</th><th>Raison Sociale</th><th>NIF</th>
    <th>Type</th><th>Vérificateur</th><th>Centre</th>
    <th>Statut</th><th>TO</th>
    <th>Montant NP</th><th>Montant ND</th>
    <th>Écart %</th><th>Délai NP→ND</th>
    <th>Date Créée</th><th>Date ND</th><th>Étape courante</th>
  </tr></thead><tbody>`;

  slice.forEach(d => {
    const statutBadge = d._statut === 'Clôturé'
      ? `<span class="badge badge-green">Clôturé</span>`
      : `<span class="badge badge-orange">En cours</span>`;
    const toBadge = d._hasTO
      ? `<span class="badge badge-red">OUI</span>`
      : `<span class="badge badge-grey">Non</span>`;
    const ecartColor = d._ecartPct === null ? '' : d._ecartPct >= 0 ? 'color:#059669' : 'color:#dc2626';

    html += `<tr>
      <td><strong>${d._ref}</strong></td>
      <td>${d['RAISON SOCIALE'] || '–'}</td>
      <td>${d['NIF'] || '–'}</td>
      <td><span class="badge badge-blue">${d._type}</span></td>
      <td>${d._verif}</td>
      <td>${d._centre}</td>
      <td>${statutBadge}</td>
      <td>${toBadge}</td>
      <td>${fmtNum(d._montantNP)}</td>
      <td>${fmtNum(d._montantND)}</td>
      <td style="${ecartColor}">${fmtPct(d._ecartPct)}</td>
      <td>${d._delaiNPND !== null ? d._delaiNPND + ' j' : '–'}</td>
      <td>${fmtDate(d._dateCree)}</td>
      <td>${fmtDate(d._dateND)}</td>
      <td>${d['ETAPE COURANTE'] || '–'}</td>
    </tr>`;
  });
  html += '</tbody></table>';
  wrap.innerHTML = html;

  // Pagination
  renderPagination(pages, total, start, slice.length);
}

function renderPagination(pages, total, start, sliceLen) {
  const pag = document.getElementById('pagination');
  let html = `<span class="page-info">${start + 1}–${start + sliceLen} sur ${total}</span>`;

  const btn = (p, label, active = false) =>
    `<button class="page-btn ${active ? 'active' : ''}" data-p="${p}">${label}</button>`;

  html += btn(1, '«');
  html += btn(Math.max(1, STATE.page - 1), '‹');

  const range = 2;
  for (let p = Math.max(1, STATE.page - range); p <= Math.min(pages, STATE.page + range); p++) {
    html += btn(p, p, p === STATE.page);
  }

  html += btn(Math.min(pages, STATE.page + 1), '›');
  html += btn(pages, '»');

  pag.innerHTML = html;
  pag.querySelectorAll('.page-btn').forEach(b => {
    b.addEventListener('click', () => {
      STATE.page = +b.dataset.p;
      renderDetailTable();
    });
  });
}

/* ─────────────────────────────────────────────────────────
   14. EXPORT CSV
───────────────────────────────────────────────────────── */
function exportCSV() {
  const cols = ['_ref','RAISON SOCIALE','NIF','_type','_verif','_centre','_statut','_hasTO',
                '_montantNP','_montantND','_ecartPct','_delaiNPND','_dateCree','_dateND'];
  const headers = ['Référence','Raison Sociale','NIF','Type','Vérificateur','Centre','Statut','TO',
                   'Montant NP','Montant ND','Écart %','Délai NP→ND','Date Créée','Date ND'];

  const rows = [headers.join(';')];
  STATE.filteredData.forEach(d => {
    rows.push(cols.map(c => {
      let v = d[c];
      if (v instanceof Date) return fmtDate(v);
      if (v === null || v === undefined) return '';
      return String(v).replace(/;/g, ',');
    }).join(';'));
  });

  const blob = new Blob(['\uFEFF' + rows.join('\n')], { type: 'text/csv;charset=utf-8;' });
  const url  = URL.createObjectURL(blob);
  const a    = Object.assign(document.createElement('a'), { href: url, download: 'controles_fiscaux.csv' });
  document.body.appendChild(a);
  a.click();
  document.body.removeChild(a);
  URL.revokeObjectURL(url);
}

/* ─────────────────────────────────────────────────────────
   15. REFRESH GLOBAL
───────────────────────────────────────────────────────── */
function refreshAll() {
  const data = STATE.filteredData;
  const kpi  = computeKPI(data);
  updateKPICards(kpi);
  refreshCharts(data, kpi);
  renderMatrix(data);
  renderTOTable(data);
  renderDetailTable();
}

/* ─────────────────────────────────────────────────────────
   16. GESTION SESSIONSTORAGE
───────────────────────────────────────────────────────── */
function saveDataToSession(fileName) {
  try {
    const sessionData = {
      rawData: STATE.rawData,
      fileName: fileName,
      timestamp: new Date().toISOString(),
    };
    // Replacer pour convertir les Dates en format ISO
    sessionStorage.setItem('fiscal_dashboard_data', JSON.stringify(sessionData, (key, value) => {
      if (value instanceof Date) {
        return { __isDate: true, value: value.toISOString() };
      }
      return value;
    }));
  } catch (e) {
    console.warn('Impossible de sauvegarder en sessionStorage:', e);
  }
}

function loadDataFromSession() {
  try {
    const data = sessionStorage.getItem('fiscal_dashboard_data');
    if (!data) return null;
    // Reviver pour reconvertir les Dates
    return JSON.parse(data, (key, value) => {
      if (value && value.__isDate) {
        return new Date(value.value);
      }
      return value;
    });
  } catch (e) {
    console.warn('Impossible de charger depuis sessionStorage:', e);
    return null;
  }
}

function clearSessionData() {
  try {
    sessionStorage.removeItem('fiscal_dashboard_data');
  } catch (e) {
    console.warn('Impossible de nettoyer sessionStorage:', e);
  }
}

function restoreFromSession() {
  const sessionData = loadDataFromSession();
  if (!sessionData) return false;

  try {
    STATE.rawData = sessionData.rawData;
    STATE.filteredData = [...STATE.rawData];

    populateFilters();
    document.getElementById('sidebar-count').textContent = STATE.rawData.length;
    document.getElementById('header-file-info').textContent =
      `${sessionData.fileName} · ${STATE.rawData.length} dossier${STATE.rawData.length > 1 ? 's' : ''}`;

    showDashboard();
    refreshAll();
    return true;
  } catch (e) {
    console.error('Erreur lors de la restauration:', e);
    clearSessionData();
    return false;
  }
}

/* ─────────────────────────────────────────────────────────
   17. CHARGEMENT DU FICHIER
───────────────────────────────────────────────────────── */
async function loadFile(file) {
  if (!file) return;
  toast('Chargement en cours…');

  try {
    let raw;
    if (file.name.endsWith('.csv')) {
      const text = await file.text();
      raw = parseCSV(text);
    } else {
      raw = await parseXLSX(file);
    }

    if (!raw.length) throw new Error('Aucune donnée trouvée dans le fichier.');

    STATE.rawData = enrichData(raw);
    STATE.filteredData = [...STATE.rawData];

    // ✅ Sauvegarder en sessionStorage
    saveDataToSession(file.name);

    populateFilters();
    document.getElementById('sidebar-count').textContent = STATE.rawData.length;
    document.getElementById('header-file-info').textContent =
      `${file.name} · ${STATE.rawData.length} dossier${STATE.rawData.length > 1 ? 's' : ''}`;

    showDashboard();
    refreshAll();
    toast(`✅ ${STATE.rawData.length} dossiers chargés avec succès`);
  } catch (e) {
    toast('❌ Erreur : ' + e.message);
    console.error(e);
  }
}

/* ─────────────────────────────────────────────────────────
   17. GESTION DE L'UI
───────────────────────────────────────────────────────── */
function showDashboard() {
  document.getElementById('welcome-screen').classList.add('hidden');
  document.getElementById('dashboard').classList.remove('hidden');
}

function showWelcome() {
  document.getElementById('welcome-screen').classList.remove('hidden');
  document.getElementById('dashboard').classList.add('hidden');
}

function resetFilters() {
  document.getElementById('f-type').selectedIndex   = -1;
  document.getElementById('f-centre').selectedIndex = -1;
  document.getElementById('f-verif').selectedIndex  = -1;
  document.getElementById('f-annee').value          = '';
  document.querySelectorAll('#f-statut input, #f-to input').forEach(i => i.checked = true);
  applyFilters();
}

/* ─────────────────────────────────────────────────────────
   18. INITIALISATION
───────────────────────────────────────────────────────── */
document.addEventListener('DOMContentLoaded', () => {
  // Date du jour dans le header
  document.getElementById('header-date').textContent =
    new Date().toLocaleDateString('fr-FR', { weekday: 'long', year: 'numeric', month: 'long', day: 'numeric' });

  // 🔄 Vérifier si des données sont en sessionStorage et les restaurer
  if (!restoreFromSession()) {
    // Si aucune donnée n'a pu être restaurée, afficher l'écran d'accueil
    showWelcome();
  }

  /* ── File input ── */
  const fileInput = document.getElementById('file-input');
  fileInput.addEventListener('change', e => loadFile(e.target.files[0]));

  /* ── Drag & Drop ── */
  const zone = document.getElementById('upload-zone');
  zone.addEventListener('click', (e) => {
    // Éviter le double click si l'utilisateur a cliqué sur le label ou le file-input
    if (e.target.tagName === 'LABEL' || e.target.tagName === 'INPUT' || e.target.closest('label[for="file-input"]')) {
      return;
    }
    fileInput.click();
  });
  zone.addEventListener('dragover', e => { e.preventDefault(); zone.classList.add('drag-over'); });
  zone.addEventListener('dragleave', () => zone.classList.remove('drag-over'));
  zone.addEventListener('drop', e => {
    e.preventDefault();
    zone.classList.remove('drag-over');
    loadFile(e.dataTransfer.files[0]);
  });

  /* ── Réimport ── */
  document.getElementById('btn-reimport').addEventListener('click', () => {
    fileInput.value = '';
    fileInput.click();
  });

  /* ── Onglets ── */
  document.querySelectorAll('.tab-btn').forEach(btn => {
    btn.addEventListener('click', () => {
      document.querySelectorAll('.tab-btn').forEach(b => b.classList.remove('active'));
      document.querySelectorAll('.tab-content').forEach(c => c.classList.add('hidden'));
      btn.classList.add('active');
      document.getElementById('tab-' + btn.dataset.tab).classList.remove('hidden');
      // Redessiner les graphiques de l'onglet actif (pour résoudre les problèmes de taille)
      if (STATE.filteredData.length) {
        setTimeout(() => refreshCharts(STATE.filteredData, computeKPI(STATE.filteredData)), 50);
      }
    });
  });

  /* ── Filtres ── */
  ['f-type', 'f-centre', 'f-verif', 'f-annee'].forEach(id => {
    document.getElementById(id).addEventListener('change', applyFilters);
  });
  document.querySelectorAll('#f-statut input, #f-to input').forEach(i => {
    i.addEventListener('change', applyFilters);
  });
  document.getElementById('btn-reset-filters').addEventListener('click', resetFilters);

  /* ── Recherche ── */
  document.getElementById('search-input').addEventListener('input', e => {
    STATE.searchTerm = e.target.value;
    STATE.page = 1;
    renderDetailTable();
  });

  /* ── Export CSV ── */
  document.getElementById('btn-export-csv').addEventListener('click', exportCSV);

  /* ── Resize → redessiner les graphiques ── */
  let resizeTimer;
  window.addEventListener('resize', () => {
    clearTimeout(resizeTimer);
    resizeTimer = setTimeout(() => {
      if (STATE.filteredData.length) {
        refreshCharts(STATE.filteredData, computeKPI(STATE.filteredData));
      }
    }, 200);
  });
});
