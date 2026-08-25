/* 
   DASHBOARD CONCRESUPER — script.js

   - Carregamento automático tenta vários nomes de arquivo.
   - Ranking por % da Meta 100%.
   - Filtro de dias abre no dia mais recente com dados.
   - Gráfico com 2 séries: "Total Produzido" e "Meta 100%".
   - "Total Produzido" = valor registrado no dia selecionado no filtro.
   - Reparo de nomes: remove "undefined" e corrige nomes corrompidos
     comparando com a lista conhecida de vendedores e filiais.
    */

/* ---------- 1. CONFIGURAÇÃO ---------- */
const MONTHS_PT = ['Janeiro','Fevereiro','Março','Abril','Maio','Junho','Julho','Agosto','Setembro','Outubro','Novembro','Dezembro'];

const COLORS = {
  yellow: '#F5C542',
  gold:   '#FFC83D',
  muted:  '#AAB4C3',
  green:  '#00D084',
  border: '#F5B400',
  red:    '#E5533D'
};

const PROPOSALS_DAILY_META = 5;

/* Arquivos de dados que o dashboard tenta carregar automaticamente */
const CANDIDATE_FILES = [
  'dados.xlsx', 'Dados.xlsx', 'DADOS.xlsx', 'dados.xls',
  'data.xlsx', 'Data.xlsx', 'dados.csv', 'Dados.csv',
  'planilha.xlsx', 'base.xlsx'
];

/* Listas conhecidas de vendedores e filiais (usadas para reparar nomes corrompidos) */
const KNOWN_VENDORS = ['Alex','Eduardo','Cesar','Marcelo','Michel','Diego','Diogo','Emerson','Junior','Lucas'];
const KNOWN_BRANCHES = ['Filial Cascavel','Filial Toledo','Filial Rondon','Filial Palotina','Filial Guaira','Filial Matelandia','Filial CDC'];

const nf  = new Intl.NumberFormat('pt-BR', { maximumFractionDigits: 1 });
const nf0 = new Intl.NumberFormat('pt-BR', { maximumFractionDigits: 0 });

/* ---------- 2. ESTADO GLOBAL ---------- */
const state = {
  daily: [], goals: [], proposals: [],
  months: [], branches: [], vendors: [],
  isCumulative: true,
  prod: { month: 'all', day: 'all', branch: 'all', vendor: 'all' },
  prop: { month: 'all', vendor: 'all' }
};

const charts = { daily: null, dayCompare: null, vendorCharts: {} };

/* ---------- 3. UTILITÁRIOS ---------- */
function norm(s) {
  return String(s ?? '').toLowerCase().normalize('NFD').replace(/[\u0300-\u036f]/g, '').trim();
}
function str(v) { return (v === null || v === undefined) ? '' : String(v).trim(); }

/* Limpa nomes: remove "undefined", caracteres invisíveis e espaços duplos */
function cleanName(s) {
  return String(s ?? '')
    .replace(/undefined/gi, '')
    .replace(/[\uFEFF\u200B-\u200D\u2060]/g, '')
    .replace(/\s+/g, ' ')
    .trim();
}

/* Distância de edição (Levenshtein) para comparar nomes parecidos */
function editDist(a, b) {
  const m = a.length, n = b.length;
  const dp = Array.from({ length: m + 1 }, () => Array(n + 1).fill(0));
  for (let i = 0; i <= m; i++) dp[i][0] = i;
  for (let j = 0; j <= n; j++) dp[0][j] = j;
  for (let i = 1; i <= m; i++) {
    for (let j = 1; j <= n; j++) {
      dp[i][j] = Math.min(
        dp[i - 1][j] + 1,
        dp[i][j - 1] + 1,
        dp[i - 1][j - 1] + (a[i - 1] === b[j - 1] ? 0 : 1)
      );
    }
  }
  return dp[m][n];
}

/* Repara um nome corrompido comparando com a lista conhecida */
function repairName(raw, knownList) {
  const c = cleanName(raw);
  if (!c) return '';
  const nc = norm(c);
  for (const k of knownList) if (norm(k) === nc) return k;   // exato
  for (const k of knownList) {
    const nk = norm(k);
    if (nc.includes(nk) || nk.includes(nc)) return k;        // um contém o outro
  }
  let best = null, bestD = Infinity;
  for (const k of knownList) {
    const d = editDist(nc, norm(k));
    if (d < bestD) { bestD = d; best = k; }
  }
  if (bestD <= 2) return best;                               // quase igual
  return c;                                                  // sem match: devolve limpo
}

/* Busca o valor de uma coluna aceitando vários nomes possíveis na planilha */
function findCol(row, names) {
  for (const n of names) {
    if (row[n] !== undefined && row[n] !== null) return row[n];
  }
  return null;
}

function safeNumber(v) {
  if (v === null || v === undefined || v === '') return null;
  if (typeof v === 'number') return Number.isFinite(v) ? v : null;
  let s = String(v).trim().replace(/\s/g, '');
  if (s === '' || s === '-') return null;
  if (/^-?\d{1,3}(\.\d{3})+(,\d+)?$/.test(s)) s = s.replace(/\./g, '').replace(',', '.');
  else s = s.replace(',', '.');
  const n = parseFloat(s);
  return Number.isFinite(n) ? n : null;
}

function parseDateCell(v) {
  if (v === null || v === undefined || v === '') return null;
  if (v instanceof Date && !isNaN(v)) return v;
  if (typeof v === 'number') {
    const d = XLSX.SSF.parse_date_code(v);
    return d ? new Date(d.y, d.m - 1, d.d) : null;
  }
  if (typeof v === 'string') {
    const s = v.trim();
    let m = s.match(/(\d{4})-(\d{1,2})-(\d{1,2})/);
    if (m) return new Date(+m[1], +m[2] - 1, +m[3]);
    m = s.match(/(\d{1,2})\/(\d{1,2})\/(\d{4})/);
    if (m) return new Date(+m[3], +m[2] - 1, +m[1]);
    m = s.match(/(\d{1,2})-(\d{1,2})-(\d{4})/);
    if (m) return new Date(+m[3], +m[2] - 1, +m[1]);
  }
  return null;
}

function dateKey(d) {
  return d.getFullYear() + '-' + String(d.getMonth() + 1).padStart(2, '0') + '-' + String(d.getDate()).padStart(2, '0');
}
function fmtDate(d) {
  return d ? String(d.getDate()).padStart(2, '0') + '/' + String(d.getMonth() + 1).padStart(2, '0') + '/' + d.getFullYear() : '—';
}
function fmtNum(v) {
  if (v === null || v === undefined || !isFinite(v)) return '—';
  return nf.format(v);
}
function escapeHtml(s) {
  return String(s ?? '').replace(/[&<>"']/g, c => ({ '&':'&amp;', '<':'&lt;', '>':'&gt;', '"':'&quot;', "'":'&#39;' }[c]));
}
function pick(row, ...keys) {
  for (const k of keys) if (row[k] !== undefined && row[k] !== null) return row[k];
  return null;
}
function progressMarkers() {
  return [25, 50, 75, 100].map(p => '<span class="progress-marker" style="left:' + p + '%"></span>').join('');
}

/* ---------- 4. LEITURA DA PLANILHA ---------- */
async function loadFromFetch() {
  for (const fileName of CANDIDATE_FILES) {
    try {
      const res = await fetch(fileName + '?ts=' + Date.now());
      if (!res.ok) continue;
      const buf = await res.arrayBuffer();
      processWorkbook(buf, fileName);
      return;
    } catch (e) { /* tenta o próximo nome */ }
  }
  setStatus('Nenhum arquivo carregado — publique o ' + CANDIDATE_FILES[0] + ' na pasta do site ou use "🔄 ATUALIZAR DADOS".');
}
function onFileSelected(file) {
  const reader = new FileReader();
  reader.onload = ev => processWorkbook(ev.target.result, file.name);
  reader.readAsArrayBuffer(file);
}
function processWorkbook(data, name) {
  const wb = XLSX.read(data, { type: 'array', cellDates: true });
  try { state.daily = parseDaily(wb); } catch (e) { state.daily = []; }
  try { state.goals = parseGoals(wb); } catch (e) { state.goals = []; }
  try { state.proposals = parseProposals(wb); } catch (e) { state.proposals = []; }
  normalizeNames();
  state.isCumulative = detectCumulative();
  deriveDimensions();
  if (state.prod.day === 'all') state.prod.day = latestDayKey();
  renderFilters();
  renderAll();
  const modo = state.isCumulative ? 'acumulado' : 'diário';
  setStatus(name + ' · ' + state.daily.length + ' registros diários · ' + state.goals.length + ' metas · ' + state.proposals.length + ' propostas · vendedores: ' + state.vendors.length + ' · Produzido: ' + modo);
}

/* Unifica a grafia dos nomes e repara nomes corrompidos em todas as abas */
function normalizeNames() {
  const vMap = new Map();
  const fix = (raw, knownList) => {
    const c = cleanName(raw);
    if (!c) return '';
    const k = norm(c);
    if (!vMap.has(k)) vMap.set(k, repairName(raw, knownList) || c);
    return vMap.get(k);
  };
  state.daily.forEach(r => {
    r.vendedor = fix(r.vendedor, KNOWN_VENDORS);
    r.filial = fix(r.filial, KNOWN_BRANCHES);
  });
  state.goals.forEach(g => { g.vendedor = fix(g.vendedor, KNOWN_VENDORS); });
  state.proposals.forEach(p => { p.vendedor = fix(p.vendedor, KNOWN_VENDORS); });
}

/* ---------- 5. PARSING DAS ABAS ---------- */
function getSheet(wb, namePattern) {
  for (const n of wb.SheetNames) if (norm(n).includes(norm(namePattern))) return wb.Sheets[n];
  return null;
}
function parseDaily(wb) {
  const ws = getSheet(wb, 'Dados Diários');
  if (!ws) return [];
  const rows = XLSX.utils.sheet_to_json(ws, { defval: null });
  const out = [];
  for (const r of rows) {
    const vendedor = cleanName(str(findCol(r, ['Vendedor', 'vendedor', 'VENDEDOR'])));
    const data = parseDateCell(findCol(r, ['Data', 'data', 'DATA']));
    if (!vendedor || !data) continue;
    out.push({
      cidade: cleanName(str(findCol(r, ['Cidade', 'cidade', 'CIDADE']))),
      filial: cleanName(str(findCol(r, ['Filial', 'filial', 'FILIAL', 'Filial/Cidade', 'Loja']))),
      vendedor,
      produzido: safeNumber(findCol(r, ['Produzido', 'produzido', 'PRODUZIDO'])) ?? 0,
      metaDia: safeNumber(findCol(r, ['Meta', 'meta', 'META', 'Meta do Dia'])),
      data
    });
  }
  const seen = new Map();
  for (const rec of out) {
    const k = norm(rec.vendedor) + '|' + dateKey(rec.data);
    seen.set(k, rec);
  }
  return [...seen.values()];
}
function parseGoals(wb) {
  const ws = getSheet(wb, 'Metas 100%');
  if (!ws) return [];
  const rows = XLSX.utils.sheet_to_json(ws, { defval: null });
  const out = [];
  for (const r of rows) {
    const vendedor = cleanName(str(findCol(r, ['Vendedor', 'vendedor', 'VENDEDOR'])));
    if (!vendedor) continue;
    const mes = str(findCol(r, ['Mês', 'Mes', 'mês', 'mes', 'MES']));
    const cidade = cleanName(str(findCol(r, ['Cidade', 'cidade', 'CIDADE'])));
    const meta = safeNumber(findCol(r, ['Meta 100%', 'Meta100', 'Meta 100', 'meta100', 'Meta']));
    out.push({ mes, vendedor, cidade, meta, metaKey: norm(mes) + '|' + norm(vendedor) + '|' + norm(cidade) });
  }
  return out;
}
function parseProposals(wb) {
  const ws = getSheet(wb, 'Propostas');
  if (!ws) return [];
  const rows = XLSX.utils.sheet_to_json(ws, { header: 1, defval: null, raw: true });
  if (!rows.length) return [];

  const header = rows[0] || [];
  const mesIdx = header.findIndex(h => norm(h) === 'mes');
  const vendIdx = header.findIndex(h => norm(h) === 'vendedor');
  if (mesIdx < 0 || vendIdx < 0) return [];

  const dayCols = [];
  header.forEach((h, i) => {
    const m = String(h ?? '').match(/^dia\s*(\d+)$/i);
    if (m) dayCols.push({ idx: i, dia: +m[1] });
  });

  const records = [];
  let currentMonth = null;
  let currentDates = null;

  for (const r of rows) {
    if (!r || !r.length) continue;
    const vname = cleanName(str(r[vendIdx]));

    if (norm(vname) === 'data') {
      const dates = {};
      for (const c of dayCols) {
        const d = parseDateCell(r[c.idx]);
        if (d) dates[c.dia] = d;
      }
      currentDates = dates;
      continue;
    }

    const mesCell = str(r[mesIdx]);
    if (mesCell) currentMonth = mesCell;

    if (!vname) continue;
    if (!currentMonth) continue;

    for (const c of dayCols) {
      const val = safeNumber(r[c.idx]);
      if (val === null) continue;
      records.push({
        mes: currentMonth,
        vendedor: vname,
        dia: c.dia,
        data: currentDates && currentDates[c.dia] ? currentDates[c.dia] : null,
        valor: val
      });
    }
  }
  return records;
}

/* ---------- 6. DETECÇÃO: ACUMULADO OU DIÁRIO ---------- */
function detectCumulative() {
  const groups = {};
  for (const rec of state.daily) {
    const k = norm(rec.vendedor) + '|' + rec.data.getFullYear() + '-' + rec.data.getMonth();
    if (!groups[k]) groups[k] = [];
    groups[k].push(rec);
  }
  let ascending = 0, total = 0;
  for (const k in groups) {
    const list = groups[k].sort((a, b) => a.data - b.data);
    for (let i = 1; i < list.length; i++) {
      total++;
      if ((list[i].produzido || 0) >= (list[i - 1].produzido || 0)) ascending++;
    }
  }
  if (total === 0) return true;
  return ascending / total >= 0.6;
}

/* ---------- 7. DIMENSÕES ---------- */
function deriveDimensions() {
  const map = new Map();
  const addFromDate = d => {
    if (!d) return;
    const key = d.getFullYear() + '-' + d.getMonth();
    if (!map.has(key)) map.set(key, { key, name: MONTHS_PT[d.getMonth()], year: d.getFullYear(), m: d.getMonth() });
  };
  for (const r of state.daily) addFromDate(r.data);
  for (const p of state.proposals) if (p.data) addFromDate(p.data);
  for (const g of state.goals) {
    if (g.meta === null) continue;
    const m = MONTHS_PT.indexOf(g.mes);
    if (m < 0) continue;
    let year = null;
    for (const v of map.values()) if (v.m === m) { year = v.year; break; }
    const key = (year ?? 'x') + '-' + m;
    if (!map.has(key)) map.set(key, { key, name: g.mes, year, m });
  }
  state.months = [...map.values()].sort((a, b) => (a.year ?? 9999) - (b.year ?? 9999) || a.m - b.m);

  const br = new Set();
  for (const r of state.daily) if (r.filial) br.add(r.filial);
  state.branches = [...br].sort((a, b) => a.localeCompare(b, 'pt-BR'));

  const vd = new Set();
  for (const r of state.daily) vd.add(r.vendedor);
  for (const g of state.goals) vd.add(g.vendedor);
  for (const p of state.proposals) vd.add(p.vendedor);
  state.vendors = [...vd].sort((a, b) => a.localeCompare(b, 'pt-BR'));
}

/* ---------- 8. CÁLCULOS ---------- */
function getSelectedMonths(which) {
  const sel = which === 'prod' ? state.prod.month : state.prop.month;
  if (sel === 'all') return state.months;
  const m = state.months.find(x => x.key === sel);
  return m ? [m] : [];
}
function findGoal(monthName, vendedor, cidade) {
  const key = norm(monthName) + '|' + norm(vendedor) + '|' + norm(cidade);
  return state.goals.find(g => g.metaKey === key)
      || state.goals.find(g => norm(g.mes) === norm(monthName) && norm(g.vendedor) === norm(vendedor));
}
function getCutoff() {
  if (state.prod.day === 'all') return null;
  const d = new Date(state.prod.day + 'T00:00:00');
  return isNaN(d) ? null : d;
}

/* Último dia com dados (produzido > 0) */
function latestDayKey() {
  let last = null;
  for (const r of state.daily) {
    if ((r.produzido || 0) > 0 && (!last || r.data > last)) last = r.data;
  }
  return last ? dateKey(last) : 'all';
}

function computeVendorStats() {
  const months = getSelectedMonths('prod');
  const cutoff = getCutoff();
  const daySelected = cutoff !== null;

  const filtered = state.daily.filter(rec => {
    if (!months.some(mm => mm.m === rec.data.getMonth() && mm.year === rec.data.getFullYear())) return false;
    if (state.prod.branch !== 'all' && norm(rec.filial) !== norm(state.prod.branch)) return false;
    if (state.prod.vendor !== 'all' && norm(rec.vendedor) !== norm(state.prod.vendor)) return false;
    if (cutoff && rec.data > cutoff) return false;
    return true;
  });

  const totals = {};
  for (const rec of filtered) {
    const key = norm(rec.vendedor);
    if (!totals[key]) totals[key] = { vendedor: rec.vendedor, filial: rec.filial, cidade: rec.cidade, produced: 0, _last: {} };
    const t = totals[key];

    if (daySelected) {
      if (dateKey(rec.data) === state.prod.day) t.produced += rec.produzido || 0;
    } else if (state.isCumulative) {
      const mk = rec.data.getFullYear() + '-' + rec.data.getMonth();
      if (!t._last[mk] || rec.data > t._last[mk].data) t._last[mk] = rec;
    } else {
      t.produced += rec.produzido || 0;
    }
  }

  if (!daySelected && state.isCumulative) {
    for (const key of Object.keys(totals)) {
      const t = totals[key];
      for (const mk in t._last) t.produced += t._last[mk].produzido || 0;
      delete t._last;
    }
  }

  for (const key of Object.keys(totals)) {
    const t = totals[key];
    let m100 = 0, has = false;
    for (const mm of months) {
      const goal = findGoal(mm.name, t.vendedor, t.cidade);
      if (goal && goal.meta !== null) { m100 += goal.meta; has = true; }
    }
    t.meta100 = has ? m100 : null;
  }

  const arr = Object.values(totals).map(t => {
    const meta = t.meta100;
    const pct = meta ? (t.produced / meta) * 100 : null;
    return {
      vendedor: t.vendedor, filial: t.filial, cidade: t.cidade,
      meta, produced: t.produced, pct, falta: meta !== null ? meta - t.produced : null,
      meta25: meta !== null ? meta * 0.25 : null,
      meta50: meta !== null ? meta * 0.50 : null,
      meta75: meta !== null ? meta * 0.75 : null
    };
  });
  arr.sort((a, b) => (b.pct ?? -Infinity) - (a.pct ?? -Infinity));
  return arr;
}
function computeTotals(vendorStats) {
  let produced = 0, meta = 0;
  for (const v of vendorStats) { produced += v.produced || 0; if (v.meta !== null) meta += v.meta; }
  const pct = meta ? (produced / meta) * 100 : null;
  return { produced, meta, pct, falta: meta ? meta - produced : null };
}
function computeBranchStats(vendorStats) {
  const map = {};
  for (const v of vendorStats) {
    const key = norm(v.filial || v.cidade || 'Sem filial');
    const name = v.filial || v.cidade || 'Sem filial';
    if (!map[key]) map[key] = { filial: name, produced: 0, metaSum: 0, hasMeta: false };
    map[key].produced += v.produced || 0;
    if (v.meta !== null) { map[key].metaSum += v.meta; map[key].hasMeta = true; }
  }
  const arr = Object.values(map).map(b => {
    const meta = b.hasMeta ? b.metaSum : null;
    const pct = meta ? (b.produced / meta) * 100 : null;
    return { filial: b.filial, meta, produced: b.produced, pct, falta: meta !== null ? meta - b.produced : null };
  });
  arr.sort((a, b) => (b.pct ?? -Infinity) - (a.pct ?? -Infinity));
  return arr;
}

function computeDailyEvolution() {
  const months = getSelectedMonths('prod');
  const cutoff = getCutoff();
  const filtered = state.daily.filter(rec => {
    if (!months.some(mm => mm.m === rec.data.getMonth() && mm.year === rec.data.getFullYear())) return false;
    if (state.prod.branch !== 'all' && norm(rec.filial) !== norm(state.prod.branch)) return false;
    if (state.prod.vendor !== 'all' && norm(rec.vendedor) !== norm(state.prod.vendor)) return false;
    if (cutoff && rec.data > cutoff) return false;
    return true;
  });
  const byDate = {};
  for (const rec of filtered) {
    const k = dateKey(rec.data);
    if (!byDate[k]) byDate[k] = { date: rec.data, produced: 0 };
    byDate[k].produced = rec.produzido || 0;
  }
  return Object.values(byDate).sort((a, b) => a.date - b.date);
}

/* Gráfico: "Total Produzido" = valor registrado no dia selecionado (sem subtrair o dia anterior) */
function computeDayCompare() {
  const months = getSelectedMonths('prod');
  if (!months.length) return [];

  let targetKey = null;
  if (state.prod.day !== 'all') {
    targetKey = state.prod.day;
  } else {
    let lastDate = null;
    for (const rec of state.daily) {
      if (!months.some(mm => mm.m === rec.data.getMonth() && mm.year === rec.data.getFullYear())) continue;
      if (state.prod.branch !== 'all' && norm(rec.filial) !== norm(state.prod.branch)) continue;
      if (state.prod.vendor !== 'all' && norm(rec.vendedor) !== norm(state.prod.vendor)) continue;
      if (!lastDate || rec.data > lastDate) lastDate = rec.data;
    }
    if (!lastDate) return [];
    targetKey = dateKey(lastDate);
  }

  const map = {};

  for (const rec of state.daily) {
    if (!months.some(mm => mm.m === rec.data.getMonth() && mm.year === rec.data.getFullYear())) continue;
    if (state.prod.branch !== 'all' && norm(rec.filial) !== norm(state.prod.branch)) continue;
    if (state.prod.vendor !== 'all' && norm(rec.vendedor) !== norm(state.prod.vendor)) continue;
    if (dateKey(rec.data) !== targetKey) continue;
    const key = norm(rec.vendedor);
    if (!map[key]) map[key] = { vendedor: rec.vendedor, cidade: rec.cidade, produced: 0 };
    map[key].produced += rec.produzido || 0;
  }

  const arr = Object.values(map).map(v => {
    const goal = findGoal(months[0] ? months[0].name : '', v.vendedor, v.cidade);
    v.meta100 = goal && goal.meta !== null ? goal.meta : null;
    return v;
  });
  arr.sort((a, b) => b.produced - a.produced);
  return arr;
}

/* ---------- PROPOSTAS ---------- */
function computeProposals() {
  const months = getSelectedMonths('prop');
  let recs = state.proposals.filter(p => {
    const inMonth = p.data
      ? months.some(mm => mm.year === p.data.getFullYear() && mm.m === p.data.getMonth())
      : months.some(mm => norm(mm.name) === norm(p.mes));
    return inMonth;
  });
  if (state.prop.vendor !== 'all') recs = recs.filter(p => norm(p.vendedor) === norm(state.prop.vendor));

  const byVendor = {};
  const daysSet = new Set();
  let total = 0;
  for (const p of recs) {
    total += p.valor;
    if (p.data) daysSet.add(dateKey(p.data));
    const key = norm(p.vendedor);
    if (!byVendor[key]) byVendor[key] = { vendedor: p.vendedor, total: 0, dias: new Set() };
    byVendor[key].total += p.valor;
    if (p.data) byVendor[key].dias.add(dateKey(p.data));
  }
  const diasDistintos = daysSet.size;
  const media = diasDistintos ? total / diasDistintos : 0;
  const vendors = Object.values(byVendor).map(v => ({
    vendedor: v.vendedor,
    total: v.total,
    mediaDia: v.dias.size ? v.total / v.dias.size : 0,
    participacao: total ? (v.total / total) * 100 : 0
  }));
  vendors.sort((a, b) => b.total - a.total);

  const series = buildVendorProposalSeries(recs, months);

  return { total, media, diasDistintos, vendors, leader: vendors.length ? vendors[0] : null, series };
}

function buildVendorProposalSeries(recs, months) {
  if (!recs.length) return [];
  let maxDate = null;
  for (const p of state.proposals) {
    if (!p.data) continue;
    if (!months.some(mm => mm.year === p.data.getFullYear() && mm.m === p.data.getMonth())) continue;
    if (!maxDate || p.data > maxDate) maxDate = p.data;
  }
  if (!maxDate) return [];

  const byVendor = {};
  for (const p of recs) {
    if (!p.data) continue;
    const key = norm(p.vendedor);
    if (!byVendor[key]) byVendor[key] = { vendedor: p.vendedor, map: {} };
    const dk = dateKey(p.data);
    byVendor[key].map[dk] = (byVendor[key].map[dk] || 0) + p.valor;
  }

  const days = [];
  const cursor = new Date(maxDate.getFullYear(), maxDate.getMonth(), 1);
  while (cursor <= maxDate) {
    days.push(new Date(cursor));
    cursor.setDate(cursor.getDate() + 1);
  }

  const out = [];
  for (const key in byVendor) {
    const v = byVendor[key];
    const points = [];
    let acc = 0;
    for (const d of days) {
      const dk = dateKey(d);
      const val = v.map[dk] || 0;
      acc += val;
      points.push({ date: d, value: val, acc });
    }
    out.push({ vendedor: v.vendedor, points });
  }
  out.sort((a, b) => a.vendedor.localeCompare(b.vendedor, 'pt-BR'));
  return out;
}

/* ---------- 9. RENDERIZAÇÃO ---------- */
function setStatus(msg) { document.getElementById('fileStatus').textContent = msg; }

function renderFilters() {
  const monthsP = getSelectedMonths('prod');

  document.getElementById('filterMonthP').innerHTML =
    '<option value="all">Todos os meses</option>' +
    state.months.map(m => '<option value="' + m.key + '" ' + (state.prod.month === m.key ? 'selected' : '') + '>' + m.name + (m.year ? ' ' + m.year : '') + '</option>').join('');

  const days = new Set();
  for (const r of state.daily) {
    if (monthsP.some(mm => mm.m === r.data.getMonth() && mm.year === r.data.getFullYear()) && (r.produzido || 0) > 0) {
      days.add(dateKey(r.data));
    }
  }
  const sortedDays = [...days].sort();
  document.getElementById('filterDayP').innerHTML =
    '<option value="all">Todos os dias</option>' +
    sortedDays.map(d => '<option value="' + d + '" ' + (state.prod.day === d ? 'selected' : '') + '>' + fmtDate(new Date(d + 'T00:00:00')) + '</option>').join('');

  document.getElementById('filterBranchP').innerHTML =
    '<option value="all">Todas as filiais</option>' +
    state.branches.map(b => '<option value="' + escapeHtml(b) + '" ' + (state.prod.branch === b ? 'selected' : '') + '>' + escapeHtml(b) + '</option>').join('');

  document.getElementById('filterVendorP').innerHTML =
    '<option value="all">Todos os vendedores</option>' +
    state.vendors.map(v => '<option value="' + escapeHtml(v) + '" ' + (state.prod.vendor === v ? 'selected' : '') + '>' + escapeHtml(v) + '</option>').join('');

  document.getElementById('filterMonthS').innerHTML =
    '<option value="all">Todos os meses</option>' +
    state.months.map(m => '<option value="' + m.key + '" ' + (state.prop.month === m.key ? 'selected' : '') + '>' + m.name + (m.year ? ' ' + m.year : '') + '</option>').join('');

  document.getElementById('filterVendorS').innerHTML =
    '<option value="all">Todos os vendedores</option>' +
    state.vendors.map(v => '<option value="' + escapeHtml(v) + '" ' + (state.prop.vendor === v ? 'selected' : '') + '>' + escapeHtml(v) + '</option>').join('');
}

function pctColor(pct) {
  if (pct === null) return COLORS.muted;
  if (pct >= 100) return COLORS.green;
  if (pct >= 80) return COLORS.yellow;
  return COLORS.red;
}
function renderProducaoCards(totals) {
  const color = pctColor(totals.pct);
  document.getElementById('cardsProducao').innerHTML =
    '<div class="kpi"><span class="kpi-label">🎯 Meta 100%</span><span class="kpi-value">' + fmtNum(totals.meta) + '</span></div>' +
    '<div class="kpi"><span class="kpi-label">📦 Total Produzido</span><span class="kpi-value">' + fmtNum(totals.produced) + '</span></div>' +
    '<div class="kpi kpi-highlight"><span class="kpi-label">% da Meta</span><span class="kpi-value" style="color:' + color + '">' + (totals.pct === null ? '—' : nf.format(totals.pct) + '%') + '</span><div class="progress">' + progressMarkers() + '<div class="progress-fill" style="width:' + Math.min(100, totals.pct ?? 0) + '%;background:' + color + '"></div></div></div>' +
    '<div class="kpi"><span class="kpi-label">Falta para Meta</span><span class="kpi-value">' + fmtNum(totals.falta) + '</span></div>';
}

/* Ranking por % da Meta 100% (exibe % em vez de valores absolutos) */
function renderRankingProducao(list) {
  const el = document.getElementById('rankingVendors');
  const ranked = list.filter(v => v.meta !== null).sort((a, b) => (b.pct ?? -Infinity) - (a.pct ?? -Infinity));
  if (!ranked.length) { el.innerHTML = '<p class="empty">Sem metas cadastradas para o período.</p>'; return; }
  const medals = ['🥇', '🥈', '🥉'];
  el.innerHTML = ranked.map((v, i) => {
    const medal = i < 3 ? medals[i] : '<span class="pos">' + (i + 1) + 'º</span>';
    const color = pctColor(v.pct);
    const bar = v.pct === null ? '' : progressMarkers() + '<div class="progress-fill" style="width:' + Math.min(100, v.pct) + '%;background:' + color + '"></div>';
    return '<div class="rank-row"><div class="rank-medal">' + medal + '</div><div class="rank-info"><div class="rank-name">' + escapeHtml(v.vendedor) + '</div><div class="progress">' + bar + '</div></div><div class="rank-pct" style="color:' + color + '">' + (v.pct === null ? '—' : nf.format(v.pct) + '%') + '</div></div>';
  }).join('');
}

/* Tabela com a Meta 100% e as metas fragmentadas (25/50/75) com as datas */
function renderTableVendors(list) {
  const el = document.getElementById('tableVendors');
  if (!list.length) { el.innerHTML = '<p class="empty">Sem dados para os filtros selecionados.</p>'; return; }

  const note = '<div class="table-note">' +
    '<div class="note-title">📌 Meta de Ritmo</div>' +
    '<p>Para mantermos um ritmo constante ao longo do mês, nosso objetivo é:</p>' +
    '<div class="note-milestones">' +
    '<div class="milestone"><span class="m-pct">25%</span> até o dia 10</div>' +
    '<div class="milestone"><span class="m-pct">50%</span> até o dia 17</div>' +
    '<div class="milestone"><span class="m-pct">75%</span> até o dia 24</div>' +
    '<div class="milestone"><span class="m-pct">100%</span> até o dia 31</div>' +
    '</div></div>';

  const rows = list.map((v, i) => {
    const cls = i === 0 ? 'row-best' : i < 3 ? 'row-top' : '';
    const metaCell = v.meta === null ? '<span class="tag-warn">Sem meta</span>' : fmtNum(v.meta);
    const m25 = v.meta25 === null ? '—' : fmtNum(v.meta25);
    const m50 = v.meta50 === null ? '—' : fmtNum(v.meta50);
    const m75 = v.meta75 === null ? '—' : fmtNum(v.meta75);
    const color = pctColor(v.pct);
    return '<tr class="' + cls + '">' +
      '<td class="pos-cell">' + (i + 1) + 'º</td>' +
      '<td class="name-cell"><strong>' + escapeHtml(v.vendedor) + '</strong></td>' +
      '<td>' + metaCell + '</td>' +
      '<td>' + m25 + '</td>' +
      '<td>' + m50 + '</td>' +
      '<td>' + m75 + '</td>' +
      '<td>' + fmtNum(v.produced) + '</td>' +
      '<td class="pct-cell" style="color:' + color + '">' + (v.pct === null ? '—' : nf.format(v.pct) + '%') + '</td>' +
      '<td>' + fmtNum(v.falta) + '</td>' +
      '</tr>';
  }).join('');
  el.innerHTML = note + '<table class="tbl"><thead><tr>' +
    '<th>Posição</th><th>Vendedor</th><th>Meta 100%</th>' +
    '<th>Meta 25% (10)</th><th>Meta 50% (17)</th><th>Meta 75% (24)</th>' +
    '<th>Produzido</th><th>% da Meta</th><th>Falta</th>' +
    '</tr></thead><tbody>' + rows + '</tbody></table>';
}
function renderTableBranches(list) {
  const el = document.getElementById('tableBranches');
  if (!list.length) { el.innerHTML = '<p class="empty">Sem dados para os filtros selecionados.</p>'; return; }
  const rows = list.map((b, i) => {
    const cls = i === 0 ? 'row-best' : i < 3 ? 'row-top' : '';
    const metaCell = b.meta === null ? '<span class="tag-warn">Sem meta cadastrada</span>' : fmtNum(b.meta);
    const color = pctColor(b.pct);
    return '<tr class="' + cls + '"><td class="pos-cell">' + (i + 1) + 'º</td><td class="name-cell"><strong>' + escapeHtml(b.filial) + '</strong></td><td>' + metaCell + '</td><td>' + fmtNum(b.produced) + '</td><td class="pct-cell" style="color:' + color + '">' + (b.pct === null ? '—' : nf.format(b.pct) + '%') + '</td><td>' + fmtNum(b.falta) + '</td></tr>';
  }).join('');
  el.innerHTML = '<table class="tbl"><thead><tr><th>Posição</th><th>Filial</th><th>Meta</th><th>Produzido</th><th>% Atingido</th><th>Falta</th></tr></thead><tbody>' + rows + '</tbody></table>';
}

/* ---------- 10. GRÁFICOS ---------- */
const baseOptions = {
  responsive: true,
  maintainAspectRatio: false,
  interaction: { mode: 'index', intersect: false },
  plugins: {
    legend: { labels: { color: COLORS.muted, usePointStyle: true } },
    tooltip: { backgroundColor: '#252B3D', titleColor: '#fff', bodyColor: '#fff', borderColor: COLORS.border, borderWidth: 1 }
  },
  scales: {
    x: { ticks: { color: COLORS.muted, maxRotation: 45 }, grid: { color: 'rgba(255,255,255,0.05)' } },
    y: { ticks: { color: COLORS.muted, callback: v => nf0.format(v) }, grid: { color: 'rgba(255,255,255,0.05)' } }
  }
};
function renderDailyChart(evo, metaRef) {
  const canvas = document.getElementById('chartDaily');
  const empty = document.getElementById('emptyDaily');
  if (charts.daily) { charts.daily.destroy(); charts.daily = null; }
  if (!evo.length) { canvas.classList.add('hidden'); empty.classList.remove('hidden'); return; }
  canvas.classList.remove('hidden'); empty.classList.add('hidden');
  const labels = evo.map(e => fmtDate(e.date).slice(0, 5));
  const datasets = [
    { label: 'Produzido', data: evo.map(e => e.produced), borderColor: COLORS.green, backgroundColor: 'rgba(0,208,132,0.15)', fill: true, tension: 0.25, pointRadius: 3, borderWidth: 2, yAxisID: 'y' },
    { label: 'Meta 100%', data: evo.map(() => metaRef), borderColor: COLORS.gold, borderDash: [6, 6], pointRadius: 0, borderWidth: 2, yAxisID: 'y1' }
  ];
  charts.daily = new Chart(canvas, {
    type: 'line',
    data: { labels, datasets },
    options: {
      responsive: true,
      maintainAspectRatio: false,
      interaction: { mode: 'index', intersect: false },
      plugins: {
        legend: { labels: { color: COLORS.muted, usePointStyle: true } },
        tooltip: { backgroundColor: '#252B3D', titleColor: '#fff', bodyColor: '#fff', borderColor: COLORS.border, borderWidth: 1 }
      },
      scales: {
        x: { ticks: { color: COLORS.muted, maxRotation: 45 }, grid: { color: 'rgba(255,255,255,0.05)' } },
        y: {
          type: 'linear', position: 'left',
          ticks: { color: COLORS.green, callback: v => nf0.format(v) },
          grid: { color: 'rgba(255,255,255,0.05)' },
          title: { display: true, text: 'Produzido', color: COLORS.green }
        },
        y1: {
          type: 'linear', position: 'right',
          ticks: { color: COLORS.gold, callback: v => nf0.format(v) },
          grid: { drawOnChartArea: false },
          title: { display: true, text: 'Meta 100%', color: COLORS.gold }
        }
      }
    }
  });
}

/* Gráfico com apenas 2 séries: Total Produzido e Meta 100% */
function renderDayCompareChart(list, dayLabel) {
  const canvas = document.getElementById('chartDayCompare');
  const empty = document.getElementById('emptyDayCompare');
  if (charts.dayCompare) { charts.dayCompare.destroy(); charts.dayCompare = null; }
  if (!list.length) { canvas.classList.add('hidden'); empty.classList.remove('hidden'); return; }
  canvas.classList.remove('hidden'); empty.classList.add('hidden');
  const labels = list.map(v => v.vendedor);
  const datasets = [
    { label: 'Total Produzido', data: list.map(v => v.produced), backgroundColor: 'rgba(0,208,132,0.85)', borderRadius: 4 },
    { label: 'Meta 100%', data: list.map(v => v.meta100 ?? 0), backgroundColor: 'rgba(245,197,66,1)', borderRadius: 4 }
  ];
  charts.dayCompare = new Chart(canvas, {
    type: 'bar',
    data: { labels, datasets },
    options: { ...baseOptions, plugins: { ...baseOptions.plugins, title: { display: true, text: 'Total Produzido × Meta 100% — ' + dayLabel, color: COLORS.white, font: { size: 14, weight: '700' } } } }
  });
}

function renderVendorProposalCharts(series) {
  const container = document.getElementById('vendorProposalCharts');
  if (charts.vendorCharts) { Object.values(charts.vendorCharts).forEach(c => c.destroy()); }
  charts.vendorCharts = {};
  container.innerHTML = '';
  if (!series.length) {
    container.innerHTML = '<p class="empty">Sem propostas registradas para o período. Os dados aparecerão conforme a planilha for alimentada.</p>';
    return;
  }
  series.forEach((s, idx) => {
    const wrap = document.createElement('div');
    wrap.className = 'card';
    wrap.innerHTML = '<h2>' + escapeHtml(s.vendedor) + '</h2><div class="chart-box-sm"><canvas id="vpc-' + idx + '"></canvas></div>';
    container.appendChild(wrap);
    const labels = s.points.map(p => fmtDate(p.date).slice(0, 5));
    const datasets = [
      { type: 'bar', label: 'Propostas (dia)', data: s.points.map(p => p.value), backgroundColor: 'rgba(245,197,66,0.75)', borderRadius: 4 },
      { type: 'line', label: 'Acumulado no mês', data: s.points.map(p => p.acc), borderColor: COLORS.green, backgroundColor: 'rgba(0,208,132,0.10)', fill: true, tension: 0.25, pointRadius: 2, borderWidth: 2 },
      { type: 'line', label: 'Meta diária (' + PROPOSALS_DAILY_META + ')', data: s.points.map(() => PROPOSALS_DAILY_META), borderColor: COLORS.gold, borderDash: [6, 6], pointRadius: 0, borderWidth: 2 }
    ];
    charts.vendorCharts[idx] = new Chart(document.getElementById('vpc-' + idx), {
      type: 'bar',
      data: { labels, datasets },
      options: { ...baseOptions }
    });
  });
}

/* ---------- 11. ABA PROPOSTAS ---------- */
function renderPropostasCards(p) {
  const leaderName = p.leader ? escapeHtml(p.leader.vendedor) : '—';
  const leaderVal = p.leader ? fmtNum(p.leader.total) : '';
  document.getElementById('cardsPropostas').innerHTML =
    '<div class="kpi"><span class="kpi-label">📋 Total de Propostas</span><span class="kpi-value">' + fmtNum(p.total) + '</span></div>' +
    '<div class="kpi"><span class="kpi-label">Média Diária</span><span class="kpi-value">' + fmtNum(p.media) + '</span></div>' +
    '<div class="kpi kpi-highlight"><span class="kpi-label">🎯 Meta Diária</span><span class="kpi-value kpi-small">' + PROPOSALS_DAILY_META + ' propostas</span><span style="color:' + COLORS.green + ';font-size:13px;font-weight:700">por vendedor, por dia</span></div>' +
    '<div class="kpi"><span class="kpi-label">Vendedor Líder</span><span class="kpi-value kpi-small">' + leaderName + '</span><span style="color:' + COLORS.green + ';font-size:13px;font-weight:700">' + leaderVal + '</span></div>';
}
function renderRankingPropostas(list) {
  const el = document.getElementById('rankingProposals');
  if (!list.length) { el.innerHTML = '<p class="empty">Sem propostas registradas para o período.</p>'; return; }
  const medals = ['🥇', '🥈', '🥉'];
  el.innerHTML = list.map((v, i) => {
    const medal = i < 3 ? medals[i] : '<span class="pos">' + (i + 1) + 'º</span>';
    return '<div class="rank-row"><div class="rank-medal">' + medal + '</div><div class="rank-info"><div class="rank-name">' + escapeHtml(v.vendedor) + '</div><div class="progress"><div class="progress-fill" style="width:' + Math.min(100, v.participacao) + '%;background:' + COLORS.yellow + '"></div></div></div><div class="rank-pct" style="color:' + COLORS.yellow + '">' + fmtNum(v.total) + '</div></div>';
  }).join('');
}
function renderTablePropostas(list) {
  const el = document.getElementById('tableProposals');
  if (!list.length) { el.innerHTML = '<p class="empty">Sem propostas registradas para o período.</p>'; return; }
  const rows = list.map((v, i) => {
    const cls = i === 0 ? 'row-best' : i < 3 ? 'row-top' : '';
    return '<tr class="' + cls + '"><td class="pos-cell">' + (i + 1) + 'º</td><td class="name-cell"><strong>' + escapeHtml(v.vendedor) + '</strong></td><td>' + fmtNum(v.total) + '</td><td>' + fmtNum(v.mediaDia) + '</td><td class="pct-cell">' + nf.format(v.participacao) + '%</td></tr>';
  }).join('');
  el.innerHTML = '<table class="tbl"><thead><tr><th>Posição</th><th>Vendedor</th><th>Propostas</th><th>Média/Dia</th><th>Participação</th></tr></thead><tbody>' + rows + '</tbody></table>';
}

/* ---------- 12. ORQUESTRAÇÃO ---------- */
function renderAll() {
  if (!state.daily.length && !state.proposals.length) return;

  const vendorStats = computeVendorStats();
  const totals = computeTotals(vendorStats);
  renderProducaoCards(totals);
  renderRankingProducao(vendorStats);
  renderTableVendors(vendorStats);
  renderTableBranches(computeBranchStats(vendorStats));
  renderDailyChart(computeDailyEvolution(), totals.meta ? totals.meta : 0);

  const dayCompare = computeDayCompare();
  const dayLabel = state.prod.day !== 'all'
    ? fmtDate(new Date(state.prod.day + 'T00:00:00'))
    : 'Último dia alimentado';
  renderDayCompareChart(dayCompare, dayLabel);

  const prop = computeProposals();
  renderPropostasCards(prop);
  renderRankingPropostas(prop.vendors);
  renderTablePropostas(prop.vendors);
  renderVendorProposalCharts(prop.series);
}

function bindEvents() {
  document.getElementById('btnUpdate').addEventListener('click', () => document.getElementById('fileInput').click());
  document.getElementById('fileInput').addEventListener('change', e => {
    const f = e.target.files[0];
    if (f) onFileSelected(f);
    e.target.value = '';
  });

  document.getElementById('filterMonthP').addEventListener('change', e => { state.prod.month = e.target.value; state.prod.day = latestDayKey(); renderFilters(); renderAll(); });
  document.getElementById('filterDayP').addEventListener('change', e => { state.prod.day = e.target.value; renderAll(); });
  document.getElementById('filterBranchP').addEventListener('change', e => { state.prod.branch = e.target.value; renderAll(); });
  document.getElementById('filterVendorP').addEventListener('change', e => { state.prod.vendor = e.target.value; renderAll(); });
  document.getElementById('btnClearP').addEventListener('click', () => {
    state.prod = { month: 'all', day: latestDayKey(), branch: 'all', vendor: 'all' };
    renderFilters(); renderAll();
  });

  document.getElementById('filterMonthS').addEventListener('change', e => { state.prop.month = e.target.value; renderAll(); });
  document.getElementById('filterVendorS').addEventListener('change', e => { state.prop.vendor = e.target.value; renderAll(); });
  document.getElementById('btnClearS').addEventListener('click', () => {
    state.prop = { month: 'all', vendor: 'all' };
    renderFilters(); renderAll();
  });

  document.querySelectorAll('.tab').forEach(t => t.addEventListener('click', () => {
    document.querySelectorAll('.tab').forEach(x => x.classList.remove('active'));
    document.querySelectorAll('.tab-panel').forEach(x => x.classList.remove('active'));
    t.classList.add('active');
    document.getElementById('tab-' + t.dataset.tab).classList.add('active');
    if (charts.daily) charts.daily.resize();
    if (charts.dayCompare) charts.dayCompare.resize();
    Object.values(charts.vendorCharts).forEach(c => c.resize());
  }));
}

document.addEventListener('DOMContentLoaded', () => { bindEvents(); loadFromFetch(); });