// src/pages/RetailReportPage.js
// Cashier & Betshop Report + Deposits & Withdraws by Market
// Source: shift_report_detailed_*.csv
import React, { useState, useRef, useCallback, useMemo } from 'react';
import Papa from 'papaparse';
import * as XLSX from 'xlsx';
import { saveAs } from 'file-saver';
import jsPDF from 'jspdf';
import autoTable from 'jspdf-autotable';
import { loadRecapSaved, loadListe } from '../lib/savedFiles';

const PAGE = 50;

// ── Helpers ───────────────────────────────────────────────────────────────────
function nv(row, col) {
  const v = parseFloat(row[col]);
  return isNaN(v) ? 0 : v;
}
function fmt(v) {
  if (v == null) return '—';
  if (Math.abs(v) >= 1000000) return (v/1000000).toFixed(1)+'M';
  if (Math.abs(v) >= 1000)    return (v/1000).toFixed(0)+'K';
  return Math.round(v).toLocaleString('fr-FR');
}
function fmtEUR(v) {
  return '€'+(v||0).toFixed(2);
}
function pctBadge(prev, curr) {
  if (!prev || prev === 0) return '';
  const p = ((curr - prev) / prev) * 100;
  const color = p >= 0 ? 'var(--deposit)' : 'var(--danger)';
  return `<span style="font-size:10px;padding:1px 5px;border-radius:4px;
    background:${p>=0?'var(--deposit-bg)':'var(--danger-bg)'};color:${color};
    margin-left:4px;">${p>=0?'+':''}${p.toFixed(1)}%</span>`;
}

function extractParts(betshopName) {
  const parts = (betshopName || '').split('.');
  return {
    country: parts[0] || 'Unknown',
    market:  parts[1] || 'Unknown',
    shop:    parts[parts.length-1] || betshopName,
  };
}

function parseShiftDate(raw) {
  // Format: "22.06.26 20:20" → { month: "YYYY-MM", day: "YYYY-MM-DD", time: "HH:MM" }
  if (!raw) return null;
  const m = raw.match(/(\d{2})\.(\d{2})\.(\d{2})(?:\s+(\d{2}:\d{2}))?/);
  if (!m) return null;
  const [, dd, mm, yy, time] = m;
  return {
    month: `20${yy}-${mm}`,
    day:   `20${yy}-${mm}-${dd}`,
    time:  time || '',
  };
}

// ── Merge deux stores agrégés ─────────────────────────────────────────────────
function mergeStores(storeA, storeB) {
  const out = JSON.parse(JSON.stringify(storeA)); // deep copy
  Object.entries(storeB).forEach(([country, byMonth]) => {
    if (!out[country]) out[country] = {};
    Object.entries(byMonth).forEach(([monthKey, monthData]) => {
      if (!out[country][monthKey]) {
        out[country][monthKey] = monthData;
        return;
      }
      // Fusionner betshop
      Object.entries(monthData.betshop||{}).forEach(([shop, b]) => {
        if (!out[country][monthKey].betshop[shop]) {
          out[country][monthKey].betshop[shop] = b;
        } else {
          const t = out[country][monthKey].betshop[shop];
          t.payin_count += b.payin_count; t.payin_money += b.payin_money;
          t.paid_out    += b.paid_out;    t.balance     += b.balance;
          t.dep_eur     += b.dep_eur;     t.wd_eur      += b.wd_eur;
          t.dep_local   += b.dep_local;   t.wd_local    += b.wd_local;
          t.shifts      += b.shifts;      t.cashierCount += b.cashierCount;
        }
      });
      // Fusionner cashier
      Object.entries(monthData.cashier||{}).forEach(([name, c]) => {
        if (!out[country][monthKey].cashier[name]) {
          out[country][monthKey].cashier[name] = c;
        } else {
          const t = out[country][monthKey].cashier[name];
          t.payin_count += c.payin_count; t.payin_money += c.payin_money;
          t.paid_out    += c.paid_out;    t.balance     += c.balance;
          t.dep_eur     += c.dep_eur;     t.wd_eur      += c.wd_eur;
          t.shifts      += c.shifts;
        }
      });
      // Fusionner market
      Object.entries(monthData.market||{}).forEach(([mkt, mk]) => {
        if (!out[country][monthKey].market[mkt]) {
          out[country][monthKey].market[mkt] = mk;
        } else {
          const t = out[country][monthKey].market[mkt];
          t.dep_eur     += mk.dep_eur;   t.wd_eur      += mk.wd_eur;
          t.dep_local   += mk.dep_local; t.wd_local    += mk.wd_local;
          t.payin_count += mk.payin_count; t.payin_money += mk.payin_money;
          t.shopCount   += mk.shopCount;
        }
      });
    });
  });
  return out;
}

// ── Main aggregation ──────────────────────────────────────────────────────────
function aggregateShifts(rows) {
  // Structure: { byCountry: { country: { byMonth: { YYYY-MM: { betshop, cashier, market } } } } }
  const out = {};
  rows.forEach(row => {
    const { country, market, shop } = extractParts(row['Betshop name']);
    const username = (row['Username'] || '').trim();
    const parsed = parseShiftDate(row['Shift start']);
    if (!username || !parsed) return;
    const { month: monthKey, day: dayKey, time: shiftTime } = parsed;

    if (!out[country]) out[country] = {};
    if (!out[country][monthKey]) out[country][monthKey] = {
      betshop: {}, cashier: {}, market: {}
    };
    const m = out[country][monthKey];

    const payin_count = nv(row, 'Payin count');
    const payin_money = nv(row, 'Payin money');
    const paid_out    = nv(row, 'Paid out money');
    const balance     = nv(row, 'Shift balance');
    const dep_eur     = nv(row, 'Total online deposits EUR');
    const wd_eur      = nv(row, 'Total online withdraws EUR');
    const dep_local    = nv(row, 'Total online deposits');
    const wd_local     = nv(row, 'Total online withdraws');
    const kill_count      = nv(row, 'Kill count');
    const kill_money      = nv(row, 'Kill money');
    const kill_money_eur  = nv(row, 'Kill money EUR');
    const revoke_count    = nv(row, 'Revoke count');
    const revoke_money    = nv(row, 'Revoke money');
    const revoke_money_eur= nv(row, 'Revoke money EUR');
    const paid_out_count  = nv(row, 'Paid out count');
    const payin_money_eur = nv(row, 'Payin money EUR');
    const paid_out_eur    = nv(row, 'Paid out money EUR');
    const winner_count    = nv(row, 'Winner count');
    const winner_money    = nv(row, 'Winner money');
    const winner_money_eur= nv(row, 'Winner money EUR');
    const cash_dep        = nv(row, 'Total cash deposit');
    const cash_dep_eur    = nv(row, 'Total cash deposit EUR');
    const cash_wd         = nv(row, 'Total cash withdraws');
    const cash_wd_eur     = nv(row, 'Total cash withdraws EUR');
    const payin_tax       = nv(row, 'Payin money tax');
    const payin_tax_eur   = nv(row, 'Payin money tax EUR');

    // By betshop
    if (!m.betshop[shop]) m.betshop[shop] = {
      name:shop, market, country,
      payin_count:0, payin_money:0, paid_out:0, balance:0,
      dep_eur:0, wd_eur:0, dep_local:0, wd_local:0,
      kill_count:0, kill_money:0, kill_money_eur:0,
      revoke_count:0, revoke_money:0, revoke_money_eur:0,
      paid_out_count:0, payin_money_eur:0, paid_out_eur:0,
      winner_count:0, winner_money:0, winner_money_eur:0,
      cash_dep:0, cash_dep_eur:0, cash_wd:0, cash_wd_eur:0,
      payin_tax:0, payin_tax_eur:0,
      shifts:0, cashiers: new Set()
    };
    const b = m.betshop[shop];
    b.payin_count += payin_count; b.payin_money += payin_money;
    b.paid_out    += paid_out;    b.balance     += balance;
    b.dep_eur     += dep_eur;     b.wd_eur      += wd_eur;
    b.dep_local    += dep_local;    b.wd_local     += wd_local;
    b.kill_count       += kill_count;       b.kill_money       += kill_money;       b.kill_money_eur   += kill_money_eur;
    b.revoke_count     += revoke_count;     b.revoke_money     += revoke_money;     b.revoke_money_eur += revoke_money_eur;
    b.paid_out_count   += paid_out_count;   b.payin_money_eur  += payin_money_eur;  b.paid_out_eur     += paid_out_eur;
    b.winner_count     += winner_count;     b.winner_money     += winner_money;     b.winner_money_eur += winner_money_eur;
    b.cash_dep         += cash_dep;         b.cash_dep_eur     += cash_dep_eur;
    b.cash_wd          += cash_wd;          b.cash_wd_eur      += cash_wd_eur;
    b.payin_tax        += payin_tax;        b.payin_tax_eur    += payin_tax_eur;
    b.shifts      += 1;           b.cashiers.add(username);
    // byDay
    if (!b.byDay) b.byDay = {};
    if (!b.byDay[dayKey]) b.byDay[dayKey] = { payin_count:0, payin_money:0, paid_out:0, balance:0, shifts:0, kill_count:0, kill_money:0, revoke_count:0, revoke_money:0 };
    b.byDay[dayKey].payin_count   += payin_count;   b.byDay[dayKey].payin_money   += payin_money;
    b.byDay[dayKey].paid_out      += paid_out;       b.byDay[dayKey].balance       += balance;
    b.byDay[dayKey].kill_count    += kill_count;     b.byDay[dayKey].kill_money    += kill_money;
    b.byDay[dayKey].revoke_count  += revoke_count;   b.byDay[dayKey].revoke_money  += revoke_money;
    b.byDay[dayKey].shifts        += 1;

    // By cashier
    const cKey = username;
    if (!m.cashier[cKey]) m.cashier[cKey] = {
      name:username, betshop:shop, market, country,
      payin_count:0, payin_money:0, paid_out:0, balance:0,
      dep_eur:0, wd_eur:0, shifts:0,
      kill_count:0, kill_money:0, kill_money_eur:0,
      revoke_count:0, revoke_money:0, revoke_money_eur:0,
      paid_out_count:0, payin_money_eur:0, paid_out_eur:0,
      winner_count:0, winner_money:0, winner_money_eur:0,
      cash_dep:0, cash_dep_eur:0, cash_wd:0, cash_wd_eur:0,
      payin_tax:0, payin_tax_eur:0
    };
    const c = m.cashier[cKey];
    c.payin_count += payin_count; c.payin_money += payin_money;
    c.paid_out    += paid_out;    c.balance     += balance;
    c.dep_eur      += dep_eur;      c.wd_eur       += wd_eur;
    c.kill_count       += kill_count;       c.kill_money       += kill_money;       c.kill_money_eur   += kill_money_eur;
    c.revoke_count     += revoke_count;     c.revoke_money     += revoke_money;     c.revoke_money_eur += revoke_money_eur;
    c.paid_out_count   += paid_out_count;   c.payin_money_eur  += payin_money_eur;  c.paid_out_eur     += paid_out_eur;
    c.winner_count     += winner_count;     c.winner_money     += winner_money;     c.winner_money_eur += winner_money_eur;
    c.cash_dep         += cash_dep;         c.cash_dep_eur     += cash_dep_eur;
    c.cash_wd          += cash_wd;          c.cash_wd_eur      += cash_wd_eur;
    c.payin_tax        += payin_tax;        c.payin_tax_eur    += payin_tax_eur;
    c.shifts      += 1;
    // byDay
    if (!c.byDay) c.byDay = {};
    if (!c.byDay[dayKey]) c.byDay[dayKey] = { payin_count:0, payin_money:0, paid_out:0, balance:0, shifts:0, kill_count:0, kill_money:0, revoke_count:0, revoke_money:0, shiftList:[] };
    c.byDay[dayKey].payin_count   += payin_count;   c.byDay[dayKey].payin_money   += payin_money;
    c.byDay[dayKey].paid_out      += paid_out;       c.byDay[dayKey].balance       += balance;
    c.byDay[dayKey].kill_count    += kill_count;     c.byDay[dayKey].kill_money    += kill_money;
    c.byDay[dayKey].revoke_count  += revoke_count;   c.byDay[dayKey].revoke_money  += revoke_money;
    c.byDay[dayKey].shifts        += 1;
    c.byDay[dayKey].shiftList.push({ time: shiftTime, payin_count, payin_money, paid_out, balance });

    // By market
    if (!m.market[market]) m.market[market] = {
      name:market, country,
      dep_eur:0, wd_eur:0, dep_local:0, wd_local:0,
      payin_count:0, payin_money:0, shopCount:0, shops: new Set()
    };
    const mk = m.market[market];
    mk.dep_eur     += dep_eur;  mk.wd_eur     += wd_eur;
    mk.dep_local   += dep_local; mk.wd_local  += wd_local;
    mk.payin_count += payin_count; mk.payin_money += payin_money;
    mk.shops.add(shop);
    // byDay
    if (!mk.byDay) mk.byDay = {};
    if (!mk.byDay[dayKey]) mk.byDay[dayKey] = { dep_eur:0, wd_eur:0, dep_local:0, wd_local:0, payin_count:0, payin_money:0, shops: new Set() };
    mk.byDay[dayKey].dep_eur     += dep_eur;   mk.byDay[dayKey].wd_eur      += wd_eur;
    mk.byDay[dayKey].dep_local   += dep_local; mk.byDay[dayKey].wd_local    += wd_local;
    mk.byDay[dayKey].payin_count += payin_count; mk.byDay[dayKey].payin_money += payin_money;
    mk.byDay[dayKey].shops.add(shop);
  });
  // Convert Sets to counts
  Object.values(out).forEach(byMonth =>
    Object.values(byMonth).forEach(monthData => {
      Object.values(monthData.betshop).forEach(b => { b.cashierCount = b.cashiers.size; delete b.cashiers; });
      Object.values(monthData.market).forEach(mk => {
        mk.shopCount = mk.shops.size; delete mk.shops;
        if (mk.byDay) Object.values(mk.byDay).forEach(d => { d.shopCount = d.shops.size; delete d.shops; });
      });
    })
  );
  return out;
}

// ── Upload slot multi-fichiers ────────────────────────────────────────────────
function UploadSlot({ loadedFiles, onFiles, onAddFiles }) {
  const ref = useRef();
  const addRef = useRef();
  const hasFiles = loadedFiles && loadedFiles.length > 0;
  return (
    <div>
      <div
        onClick={() => ref.current.click()}
        onDragOver={e => e.preventDefault()}
        onDrop={e => {
          e.preventDefault();
          const files = Array.from(e.dataTransfer.files).filter(f => f.name.endsWith('.csv'));
          if (files.length) onFiles(files);
        }}
        className={`drop-zone ${hasFiles ? 'has-file' : ''}`} style={{ cursor:'pointer' }}>
        <input ref={ref} type="file" accept=".csv" hidden multiple
          onChange={e => onFiles(Array.from(e.target.files))} />
        <div className="drop-icon">{hasFiles ? '✅' : '📋'}</div>
        <div className="drop-label">
          {hasFiles
            ? `${loadedFiles.length} fichier(s) chargé(s) — cliquer pour remplacer`
            : 'shift_report_detailed_*.csv — glisser plusieurs fichiers'}
        </div>
        <div className="drop-sub">Username · Betshop name · Payin count · Payin money · Deposits EUR · Withdraws EUR</div>
        {hasFiles && (
          <div style={{ marginTop:8, display:'flex', flexWrap:'wrap', gap:4, justifyContent:'center' }}>
            {loadedFiles.map((f,i) => (
              <span key={i} className="file-chip">✓ {f}</span>
            ))}
          </div>
        )}
      </div>
      {hasFiles && (
        <div style={{ marginTop:8, textAlign:'center' }}>
          <input ref={addRef} type="file" accept=".csv" hidden multiple
            onChange={e => onAddFiles(Array.from(e.target.files))} />
          <button className="btn sm" onClick={() => addRef.current.click()}
            style={{ fontSize:12, padding:'6px 14px' }}>
            ➕ Ajouter un mois
          </button>
        </div>
      )}
    </div>
  );
}

// ── KPI card ──────────────────────────────────────────────────────────────────
function KpiCard({ label, value, color, sub }) {
  return (
    <div style={{ background:'var(--bg2)', border:'1px solid var(--border)',
      borderRadius:10, padding:'12px 16px', flex:1 }}>
      <div style={{ fontSize:22, fontWeight:700, color: color||'var(--text)' }}>{value}</div>
      <div style={{ fontSize:11, color:'var(--muted)', marginTop:3, fontWeight:600,
        letterSpacing:'.05em', textTransform:'uppercase' }}>{label}</div>
      {sub && <div style={{ fontSize:11, color:'var(--muted)', marginTop:2 }}>{sub}</div>}
    </div>
  );
}

// ── Sortable table header ─────────────────────────────────────────────────────
function Th({ label, k, sort, onSort, right }) {
  const active = sort.key === k;
  return (
    <th onClick={() => onSort(k)} style={{ cursor:'pointer', textAlign: right?'right':'left',
      color: active ? 'var(--accent)' : 'var(--muted)', userSelect:'none', whiteSpace:'nowrap' }}>
      {label} {active ? (sort.dir==='desc'?'↓':'↑') : ''}
    </th>
  );
}

// ── Pagination ────────────────────────────────────────────────────────────────
function Pagination({ page, total, totalPages, onPage }) {
  if (totalPages <= 1) return null;
  const start = (page-1)*PAGE+1, end = Math.min(page*PAGE, total);
  const pages = [];
  for (let i=1;i<=totalPages;i++) {
    if (i===1||i===totalPages||Math.abs(i-page)<=2) pages.push(i);
    else if (pages[pages.length-1]!=='…') pages.push('…');
  }
  return (
    <div style={{ display:'flex', alignItems:'center', justifyContent:'space-between',
      padding:'10px 14px', borderTop:'1px solid var(--border)', fontSize:12, color:'var(--muted)' }}>
      <span>{start}–{end} of {total.toLocaleString()}</span>
      <div style={{ display:'flex', gap:3 }}>
        <button className="btn sm" onClick={() => onPage(page-1)} disabled={page===1}>‹</button>
        {pages.map((p,i) => typeof p==='number'
          ? <button key={i} className="btn sm" onClick={() => onPage(p)}
              style={{ background:p===page?'var(--accent)':'', color:p===page?'#fff':'',
                borderColor:p===page?'var(--accent)':'' }}>{p}</button>
          : <span key={i} style={{ padding:'0 4px', lineHeight:'28px' }}>…</span>
        )}
        <button className="btn sm" onClick={() => onPage(page+1)} disabled={page===totalPages}>›</button>
      </div>
    </div>
  );
}

// ── Main Component ────────────────────────────────────────────────────────────
export default function RetailReportPage() {
  const [loadedFiles, setLoadedFiles] = useState([]); // noms des fichiers chargés
  const [store,       setStore]       = useState(null); // aggregated data
  const [loading,     setLoading]     = useState(false);
  const [country, setCountry] = useState('');
  const [month,   setMonth]   = useState('');
  const [selectedMkt, setSelectedMkt] = useState(null);
  const [mainTab, setMainTab] = useState('performance'); // 'performance' | 'deposits'
  const [subTab,  setSubTab]  = useState('betshop');     // 'betshop' | 'cashier'
  const [search,  setSearch]  = useState('');
  // ── RECAP zone/super + Liste propriétaires ──
  const [recapLookup, setRecapLookup] = useState({}); // betshopName → {zone, super}
  const [propLookup,  setPropLookup]  = useState({}); // betshopName → proprietaire

  const [filterZone,  setFilterZone]  = useState(''); // filtre actif zone
  const [filterSuper, setFilterSuper] = useState(''); // filtre actif super
  const [filterProp,  setFilterProp]  = useState(''); // filtre actif propriétaire
  const [filterCity,  setFilterCity]  = useState(''); // filtre actif ville
  const [cityLookup,  setCityLookup]  = useState({}); // betshopName → city
  const cityRef = useRef();
  const [showPerfColCfg, setShowPerfColCfg] = useState(false);
  const [visiblePerfCols, setVisiblePerfCols] = useState([
    'shifts','payin_count','payin_money','paid_out','balance','dep_eur','wd_eur'
  ]);
  const [draftPerfCols, setDraftPerfCols] = useState([
    'shifts','payin_count','payin_money','paid_out','balance','dep_eur','wd_eur'
  ]);
  const [page,    setPage]    = useState(1);
  const [sort,    setSort]    = useState({ key:'payin_count', dir:'desc' });
  const [mktSort,      setMktSort]      = useState({ key:'dep_eur', dir:'desc' });
  const [selectedRow,  setSelectedRow]  = useState(null); // { type:'betshop'|'cashier', data }
  // ── Comparaison de périodes ──
  const [cmpMode,   setCmpMode]   = useState('2months'); // '2months' | 'range'
  const [cmpMonthA, setCmpMonthA] = useState('');
  const [cmpMonthB, setCmpMonthB] = useState('');
  const [cmpFrom,   setCmpFrom]   = useState('');
  const [cmpTo,     setCmpTo]     = useState('');
  const [cmpEntity, setCmpEntity] = useState('market');   // 'market' | 'betshop' | 'cashier'
  const [cmpKpi,    setCmpKpi]    = useState('tickets');  // 'tickets' | 'payin' | 'dep' | 'wd'
  const [cmpTop,    setCmpTop]    = useState(10);         // Nb entités affichées
  const [cmpSearch, setCmpSearch] = useState('');         // Filtre recherche entité

  // ── Charger cityMap depuis fichier LISTE XLSX local ──
  const loadCityMapLocal = (file) => {
    const reader = new FileReader();
    reader.onload = e => {
      try {
        const wb2  = XLSX.read(new Uint8Array(e.target.result), { type:'array' });
        const ws2  = wb2.Sheets[wb2.SheetNames[0]];
        const rows2 = XLSX.utils.sheet_to_json(ws2, { header:1, defval:'' });
        const nrm2 = s => String(s||'').trim().toLowerCase().replace(/\s+/g,' ');
        const clk  = {};
        rows2.slice(1).forEach(r => {
          const shop = String(r[2]||'').trim(); // col C = Shop
          const city = String(r[3]||'').trim(); // col D = City
          if (shop && city) clk[nrm2(shop)] = city;
        });
        setCityLookup(clk);
      } catch(err) { console.warn('cityMap load error', err); }
    };
    reader.readAsArrayBuffer(file);
  };

  // ── Charger RECAP (zone/super) + LISTE (propriétaires) depuis Firebase ──
  React.useEffect(() => {
    const nrm = s => String(s||'').trim().toLowerCase().replace(/\s+/g,' ');
    const run = async () => {
      // RECAP : zoneData[feuille] = [[zone, super, salle, nomTech], ...]
      try {
        const res = await loadRecapSaved();
        if (res?.zoneData) {
          const lk = {};
          Object.values(res.zoneData).forEach(rows => {
            rows.forEach(r => {
              const zone = Array.isArray(r) ? r[0] : (r.z||r.ZONES||'');
              const sup  = Array.isArray(r) ? r[1] : (r.sup||r.SUPER||'');
              const sd   = Array.isArray(r) ? r[2] : (r.sd||r.SALLES||'');
              const st   = Array.isArray(r) ? r[3] : (r.st||r.NOMS_TECH||'');
              const zs = String(zone||'').trim(), ss = String(sup||'').trim();
              if (st) lk[nrm(st)] = { zone:zs, super:ss };
              if (sd) lk[nrm(sd)] = { zone:zs, super:ss };
            });
          });
          setRecapLookup(lk);
        }
      } catch(e) { console.warn('RECAP load error', e); }
      // LISTE : propMap[proprietaire] = [shop1, shop2, ...]
      try {
        const res2 = await loadListe();
        if (res2?.propMap) {
          const plk = {};
          Object.entries(res2.propMap).forEach(([prop, shops]) => {
            (Array.isArray(shops)?shops:[]).forEach(s => { if(s) plk[nrm(s)] = prop; });
          });
          setPropLookup(plk);
        }
        // Charger aussi le cityLookup si disponible dans Firebase
        if (res2?.cityMap && Object.keys(res2.cityMap).length > 0) {
          const clk = {};
          Object.entries(res2.cityMap).forEach(([shop, city]) => {
            if (shop && city) clk[nrm(shop)] = city;
          });
          setCityLookup(clk);
        }
      } catch(e) { console.warn('LISTE load error', e); }
    };
    run();
  }, []); // eslint-disable-line

  // Parse un tableau de fichiers CSV et fusionne dans le store existant
  const parseAndMerge = useCallback((files, existingStore, existingNames) => {
    setLoading(true);
    let remaining = files.length;
    const allRows = [];
    files.forEach(f => {
      Papa.parse(f, {
        header: true, skipEmptyLines: true, dynamicTyping: false,
        complete: ({ data }) => {
          allRows.push(...data);
          remaining -= 1;
          if (remaining === 0) {
            const newAgg = aggregateShifts(allRows);
            // Fusionner avec le store existant
            const merged = existingStore ? mergeStores(existingStore, newAgg) : newAgg;
            setStore(merged);
            const countries = Object.keys(merged).sort();
            const c = countries[0] || '';
            setCountry(c);
            const months = c ? Object.keys(merged[c]).sort().reverse() : [];
            setMonth(months[0] || '');
            setPage(1);
            setLoadedFiles([...(existingNames||[]), ...files.map(f2 => f2.name)]);
            setLoading(false);
          }
        },
        error: () => { remaining -= 1; if (remaining === 0) setLoading(false); },
      });
    });
  }, []);

  const loadFiles  = useCallback((files) => parseAndMerge(files, null, []),  [parseAndMerge]);
  const addFiles   = useCallback((files) => parseAndMerge(files, store, loadedFiles), [parseAndMerge, store, loadedFiles]);

  // Available options
  const countries = useMemo(() => store ? Object.keys(store).sort() : [], [store]);
  const months    = useMemo(() =>
    (store && country && store[country]) ? Object.keys(store[country]).sort().reverse() : [],
  [store, country]);

  // Current month data
  const monthData = useMemo(() =>
    (store && country && month && store[country]?.[month]) ? store[country][month] : null,
  [store, country, month]);

  // Betshop rows enrichies — toujours les betshops (pour KPIs filtrés)
  const betshopRows = useMemo(() => {
    if (!monthData) return [];
    const nrm = s => String(s||'').trim().toLowerCase().replace(/\s+/g,' ');
    return Object.values(monthData.betshop).map(r => {
      const matchKey = nrm(r.name);
      const ri = recapLookup[matchKey] || {};
      return {
        ...r,
        zone:  ri.zone  || '',
        super: ri.super || '',
        prop:  propLookup[matchKey] || '',
        city:  cityLookup[matchKey] || '',
      };
    });
  }, [monthData, recapLookup, propLookup, cityLookup]);

  // Performance rows — enrichies dynamiquement avec zone/super/prop depuis recapLookup
  const perfRows = useMemo(() => {
    if (!monthData) return [];
    const src = subTab === 'betshop' ? monthData.betshop : monthData.cashier;
    const nrm = s => String(s||'').trim().toLowerCase().replace(/\s+/g,' ');
    return Object.values(src).map(r => {
      // Pour betshop : match sur r.name ; pour cashier : match sur r.betshop
      const matchKey = subTab === 'betshop' ? nrm(r.name) : nrm(r.betshop);
      const ri = recapLookup[matchKey] || {};
      return {
        ...r,
        zone:  ri.zone  || '',
        super: ri.super || '',
        prop:  propLookup[matchKey] || '',
        city:  cityLookup[matchKey] || '',
      };
    });
  }, [monthData, subTab, recapLookup, propLookup, cityLookup]);

  // Listes distinctes pour les selects — calculées depuis perfRows enrichies
  const zoneList  = useMemo(() => {
    if (!monthData) return [];
    const nrm = s => String(s||'').trim().toLowerCase().replace(/\s+/g,' ');
    return [...new Set(Object.values(monthData.betshop)
      .filter(b => {
        if (!filterSuper) return true;
        const ri = recapLookup[nrm(b.name)] || {};
        return ri.super === filterSuper;
      })
      .map(b => {
        const ri = recapLookup[nrm(b.name)] || {};
        return ri.zone || '';
      }).filter(Boolean))].sort();
  }, [monthData, recapLookup, filterSuper]);
  const superList = useMemo(() => {
    if (!monthData) return [];
    const nrm = s => String(s||'').trim().toLowerCase().replace(/\s+/g,' ');
    return [...new Set(Object.values(monthData.betshop)
      .filter(b => {
        if (!filterZone) return true;
        const ri = recapLookup[nrm(b.name)] || {};
        return ri.zone === filterZone;
      })
      .map(b => {
        const ri = recapLookup[nrm(b.name)] || {};
        return ri.super || '';
      }).filter(Boolean))].sort();
  }, [monthData, recapLookup, filterZone]);
  const propList  = useMemo(() => {
    if (!monthData) return [];
    const nrm = s => String(s||'').trim().toLowerCase().replace(/\s+/g,' ');
    return [...new Set(Object.values(monthData.betshop)
      .filter(b => {
        const ri = recapLookup[nrm(b.name)] || {};
        if (filterZone  && ri.zone  !== filterZone)  return false;
        if (filterSuper && ri.super !== filterSuper) return false;
        return true;
      })
      .map(b => propLookup[nrm(b.name)] || '')
      .filter(Boolean))].sort();
  }, [monthData, propLookup, recapLookup, filterZone, filterSuper]);
  // cityList = vraies villes (depuis cityLookup Firebase) filtrées selon filtres actifs
  const cityList  = useMemo(() => {
    const nrm = s => String(s||'').trim().toLowerCase().replace(/\s+/g,' ');
    return [...new Set(betshopRows
      .filter(r =>
        (!filterZone  || r.zone  === filterZone)  &&
        (!filterSuper || r.super === filterSuper) &&
        (!filterProp  || r.prop  === filterProp)
      )
      .map(r => cityLookup[nrm(r.name)] || '')
      .filter(Boolean))].sort();
  }, [betshopRows, cityLookup, filterZone, filterSuper, filterProp]);

  const filtered = useMemo(() => {
    let rows = perfRows;
    // ── Filtres Zone / Super / Propriétaire ──
    if (filterZone)  rows = rows.filter(r => r.zone  === filterZone);
    if (filterSuper) rows = rows.filter(r => r.super === filterSuper);
    if (filterProp)  rows = rows.filter(r => r.prop  === filterProp);
    if (filterCity)  rows = rows.filter(r => r.city === filterCity);
    // ── Filtre texte ──
    if (search.trim()) {
      const q = search.toLowerCase();
      rows = rows.filter(r =>
        r.name.toLowerCase().includes(q) ||
        (r.market||'').toLowerCase().includes(q) ||
        (r.betshop||'').toLowerCase().includes(q)
      );
    }
    return [...rows].sort((a,b) => {
      const va = a[sort.key] ?? 0, vb = b[sort.key] ?? 0;
      return sort.dir === 'desc' ? vb - va : va - vb;
    });
  }, [perfRows, search, sort, filterZone, filterSuper, filterProp, filterCity]);

  const totalPages = Math.max(1, Math.ceil(filtered.length / PAGE));
  const pageRows   = filtered.slice((page-1)*PAGE, page*PAGE);

  // Market rows
  const mktRows = useMemo(() => {
    if (!monthData) return [];
    return [...Object.values(monthData.market)].sort((a,b) => {
      const va = a[mktSort.key] ?? 0, vb = b[mktSort.key] ?? 0;
      return mktSort.dir === 'desc' ? vb - va : va - vb;
    });
  }, [monthData, mktSort]);

  // KPIs — recalculés depuis perfRows filtrées si un filtre est actif
  const kpis = useMemo(() => {
    if (!monthData) return null;
    const hasFilter = filterZone || filterSuper || filterProp || filterCity;
    if (hasFilter) {
      // Toujours calculer depuis betshopRows (indépendant du subTab actif)
      const bsFiltered = betshopRows.filter(r =>
        (!filterZone  || r.zone   === filterZone)  &&
        (!filterSuper || r.super  === filterSuper) &&
        (!filterProp  || r.prop   === filterProp)  &&
        (!filterCity  || r.city === filterCity)
      );
      const bsNames = new Set(bsFiltered.map(b => b.name));
      const csFiltered = Object.values(monthData.cashier).filter(r => bsNames.has(r.betshop));
      return {
        betshops: bsFiltered.length,
        cashiers: csFiltered.length,
        shifts:   csFiltered.reduce((s,r) => s+r.shifts, 0),
        tickets:  bsFiltered.reduce((s,r) => s+r.payin_count, 0),
        payin:    bsFiltered.reduce((s,r) => s+r.payin_money, 0),
        dep_eur:  bsFiltered.reduce((s,r) => s+(r.dep_eur||0), 0),
        wd_eur:   bsFiltered.reduce((s,r) => s+(r.wd_eur||0), 0),
        filtered: true,
      };
    }
    const bs = Object.values(monthData.betshop);
    const cs = Object.values(monthData.cashier);
    return {
      betshops:   bs.length,
      cashiers:   cs.length,
      shifts:     cs.reduce((s,r) => s+r.shifts, 0),
      tickets:    bs.reduce((s,r) => s+r.payin_count, 0),
      payin:      bs.reduce((s,r) => s+r.payin_money, 0),
      dep_eur:    Object.values(monthData.market).reduce((s,r) => s+r.dep_eur, 0),
      wd_eur:     Object.values(monthData.market).reduce((s,r) => s+r.wd_eur, 0),
    };
  }, [monthData, betshopRows, filterZone, filterSuper, filterProp, filterCity]);

  // Calcule les KPIs agrégés pour une liste de monthKeys
  function kpisForMonths(monthKeys) {
    if (!store || !country) return null;
    const countryStore = store[country];
    if (!countryStore) return null;
    const valid = monthKeys.filter(m => countryStore[m]);
    if (!valid.length) return null;
    const betshops  = new Set();
    const cashiers  = new Set();
    let shifts=0, tickets=0, payin=0, dep=0, wd=0;
    valid.forEach(m => {
      const md = countryStore[m];
      Object.keys(md.betshop).forEach(k => betshops.add(k));
      Object.keys(md.cashier).forEach(k => cashiers.add(k));
      Object.values(md.cashier).forEach(r => { shifts += r.shifts; });
      Object.values(md.betshop).forEach(r => { tickets += r.payin_count; payin += r.payin_money; });
      Object.values(md.market).forEach(r => { dep += r.dep_eur; wd += r.wd_eur; });
    });
    return { betshops:betshops.size, cashiers:cashiers.size, shifts, tickets, payin, dep, wd, months:valid };
  }

  // Agrège les KPIs par entité (market/betshop/cashier) pour une liste de mois
  // Champs internes : payin_count, payin_money, dep_eur, wd_eur, shifts
  // Champs exposés  : tickets,     payin,        dep,     wd,     shifts
  function detailForMonths(monthKeys, entityType) {
    if (!store || !country) return {};
    const countryStore = store[country];
    const valid = monthKeys.filter(m => countryStore?.[m]);
    const result = {};
    valid.forEach(m => {
      const md = countryStore[m];
      const src = md[entityType] || {};
      Object.entries(src).forEach(([key, r]) => {
        const label = r.name || key;
        if (!result[label]) result[label] = { name:label, tickets:0, payin:0, dep:0, wd:0, shifts:0 };
        result[label].tickets += r.payin_count  || 0;
        result[label].payin   += r.payin_money  || 0;
        result[label].dep     += r.dep_eur      || 0;
        result[label].wd      += r.wd_eur       || 0;
        result[label].shifts  += r.shifts       || 0;
      });
    });
    return result;
  }

  function onSort(k) {
    setSort(s => ({ key:k, dir: s.key===k && s.dir==='desc' ? 'asc' : 'desc' }));
    setPage(1);
  }
  function onMktSort(k) {
    setMktSort(s => ({ key:k, dir: s.key===k && s.dir==='desc' ? 'asc' : 'desc' }));
  }

  // ── Export helpers ──────────────────────────────────────────────────────────
  const label = `${country}_${month}`;

  function makeBlob(wb) {
    return new Blob([XLSX.write(wb,{bookType:'xlsx',type:'array'})],
      {type:'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'});
  }

  function getBetshopRows() {
    if (!monthData) return [];
    return Object.values(monthData.betshop)
      .sort((a,b) => b.payin_count - a.payin_count);
  }
  function getCashierRows() {
    if (!monthData) return [];
    return Object.values(monthData.cashier)
      .sort((a,b) => b.payin_count - a.payin_count);
  }
  function getMktRows() {
    if (!monthData) return [];
    return Object.values(monthData.market)
      .sort((a,b) => b.dep_eur - a.dep_eur);
  }

  // 1. Top 10 Betshops Excel
  function exportTop10BetshopsXLSX() {
    const rows = getBetshopRows().slice(0,10).map((r,i) => ({
      Rang: i+1, Betshop: r.name, Marché: r.market,
      Caissiers: r.cashierCount, Shifts: r.shifts,
      Tickets: r.payin_count, 'Payin (XAF)': Math.round(r.payin_money),
      'Paid out (XAF)': Math.round(r.paid_out), 'Balance (XAF)': Math.round(r.balance),
    }));
    const wb = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(wb, XLSX.utils.json_to_sheet(rows), 'Top 10 Betshops');
    saveAs(makeBlob(wb), `Top10_Betshops_${label}.xlsx`);
  }

  // 2. Top 10 Betshops PDF — jsPDF portrait
  function exportTop10BetshopsPDF() {
    const rows = getBetshopRows().slice(0,10);
    const doc = new jsPDF({ orientation:'portrait', unit:'mm', format:'a4' });
    const now = new Date().toLocaleDateString('fr-FR');
    // En-tête
    doc.setFillColor(31, 56, 100);
    doc.rect(0, 0, 210, 22, 'F');
    doc.setTextColor(255,255,255);
    doc.setFontSize(14); doc.setFont('helvetica','bold');
    doc.text('Pipeline CM — Top 10 Betshops', 14, 10);
    doc.setFontSize(9); doc.setFont('helvetica','normal');
    doc.text(`${country} · ${month} · Généré le ${now}`, 14, 17);
    // Sous-titre
    doc.setTextColor(100,100,100);
    doc.setFontSize(9);
    doc.text('Classement par nombre de tickets (Payin count)', 14, 30);
    // Tableau
    autoTable(doc, {
      startY: 35,
      head: [['#','Betshop','Marché','Caissiers','Shifts','Tickets','Payin (XAF)','Balance (XAF)']],
      body: rows.map((r,i) => [
        i+1, r.name, r.market||'—', r.cashierCount, r.shifts,
        r.payin_count.toLocaleString('fr-FR'),
        Math.round(r.payin_money).toLocaleString('fr-FR'),
        Math.round(r.balance).toLocaleString('fr-FR'),
      ]),
      styles: { fontSize:9, cellPadding:3 },
      headStyles: { fillColor:[31,56,100], textColor:255, fontStyle:'bold' },
      alternateRowStyles: { fillColor:[245,248,255] },
      columnStyles: {
        0: { halign:'center', cellWidth:8 },
        5: { halign:'right' }, 6: { halign:'right' }, 7: { halign:'right' },
      },
      didParseCell: (data) => {
        if (data.section === 'body' && data.column.index === 7) {
          const val = rows[data.row.index]?.balance || 0;
          data.cell.styles.textColor = val >= 0 ? [15,110,86] : [197,90,17];
        }
      },
    });
    // Pied de page
    const pageH = doc.internal.pageSize.height;
    doc.setFontSize(8); doc.setTextColor(160,160,160);
    doc.text('Pipeline CM — SUPERGOOAL / CAMPRESJ', 14, pageH - 8);
    doc.text(`Page 1`, 196, pageH - 8, { align:'right' });
    doc.save(`Top10_Betshops_${label}.pdf`);
  }

  // 3. Top 10 Caissiers Excel
  function exportTop10CashiersXLSX() {
    const rows = getCashierRows().slice(0,10).map((r,i) => ({
      Rang: i+1, Caissier: r.name, Betshop: r.betshop, Marché: r.market,
      Shifts: r.shifts, Tickets: r.payin_count,
      'Payin (XAF)': Math.round(r.payin_money),
    }));
    const wb = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(wb, XLSX.utils.json_to_sheet(rows), 'Top 10 Caissiers');
    saveAs(makeBlob(wb), `Top10_Caissiers_${label}.xlsx`);
  }

  // 4. Top 10 Caissiers PDF — jsPDF portrait
  function exportTop10CashiersPDF() {
    const rows = getCashierRows().slice(0,10);
    const doc = new jsPDF({ orientation:'portrait', unit:'mm', format:'a4' });
    const now = new Date().toLocaleDateString('fr-FR');
    // En-tête
    doc.setFillColor(31, 78, 121);
    doc.rect(0, 0, 210, 22, 'F');
    doc.setTextColor(255,255,255);
    doc.setFontSize(14); doc.setFont('helvetica','bold');
    doc.text('Pipeline CM — Top 10 Caissiers', 14, 10);
    doc.setFontSize(9); doc.setFont('helvetica','normal');
    doc.text(`${country} · ${month} · Généré le ${now}`, 14, 17);
    // Sous-titre
    doc.setTextColor(100,100,100);
    doc.setFontSize(9);
    doc.text('Classement par nombre de tickets (Payin count)', 14, 30);
    // Tableau
    autoTable(doc, {
      startY: 35,
      head: [['#','Caissier','Betshop','Marché','Shifts','Tickets','Payin (XAF)']],
      body: rows.map((r,i) => [
        i+1, r.name, r.betshop||'—', r.market||'—', r.shifts,
        r.payin_count.toLocaleString('fr-FR'),
        Math.round(r.payin_money).toLocaleString('fr-FR'),
      ]),
      styles: { fontSize:9, cellPadding:3 },
      headStyles: { fillColor:[31,78,121], textColor:255, fontStyle:'bold' },
      alternateRowStyles: { fillColor:[245,248,255] },
      columnStyles: {
        0: { halign:'center', cellWidth:8 },
        4: { halign:'right' }, 5: { halign:'right' }, 6: { halign:'right' },
      },
    });
    // Pied de page
    const pageH = doc.internal.pageSize.height;
    doc.setFontSize(8); doc.setTextColor(160,160,160);
    doc.text('Pipeline CM — SUPERGOOAL / CAMPRESJ', 14, pageH - 8);
    doc.text('Page 1', 196, pageH - 8, { align:'right' });
    doc.save(`Top10_Caissiers_${label}.pdf`);
  }

  // 5. Performance complète multi-onglets
  function exportPerformanceComplete() {
    const wb = XLSX.utils.book_new();
    // Onglet Betshops
    const bRows = getBetshopRows().map((r,i) => ({
      Rang: i+1, Betshop: r.name, Marché: r.market, Caissiers: r.cashierCount,
      Shifts: r.shifts, Tickets: r.payin_count,
      'Payin (XAF)': Math.round(r.payin_money), 'Paid out': Math.round(r.paid_out),
      'Balance': Math.round(r.balance), 'Dépôts EUR': r.dep_eur.toFixed(2),
      'Retraits EUR': r.wd_eur.toFixed(2),
    }));
    XLSX.utils.book_append_sheet(wb, XLSX.utils.json_to_sheet(bRows), 'Betshops');
    // Onglet Caissiers
    const cRows = getCashierRows().map((r,i) => ({
      Rang: i+1, Caissier: r.name, Betshop: r.betshop, Marché: r.market,
      Shifts: r.shifts, Tickets: r.payin_count, 'Payin (XAF)': Math.round(r.payin_money),
    }));
    XLSX.utils.book_append_sheet(wb, XLSX.utils.json_to_sheet(cRows), 'Caissiers');
    // Onglet Marchés
    const mRows = getMktRows().map((r,i) => ({
      Rang: i+1, Marché: r.name, Shops: r.shopCount, Tickets: r.payin_count,
      'Dépôts EUR': r.dep_eur.toFixed(2), 'Retraits EUR': r.wd_eur.toFixed(2),
      'Net EUR': (r.dep_eur - r.wd_eur).toFixed(2),
    }));
    XLSX.utils.book_append_sheet(wb, XLSX.utils.json_to_sheet(mRows), 'Marchés');
    // Onglet Résumé
    const tot = getBetshopRows();
    const sumRows = [{
      Pays: country, Mois: month,
      Betshops: tot.length,
      Caissiers: getCashierRows().length,
      'Total Tickets': tot.reduce((s,r)=>s+r.payin_count,0),
      'Total Payin (XAF)': Math.round(tot.reduce((s,r)=>s+r.payin_money,0)),
      'Total Dépôts EUR': getMktRows().reduce((s,r)=>s+r.dep_eur,0).toFixed(2),
      'Total Retraits EUR': getMktRows().reduce((s,r)=>s+r.wd_eur,0).toFixed(2),
    }];
    XLSX.utils.book_append_sheet(wb, XLSX.utils.json_to_sheet(sumRows), 'Résumé');
    saveAs(makeBlob(wb), `Performance_Complete_${label}.xlsx`);
  }

  // 6. Dépôts & Retraits par marché
  function exportDepositsMarche() {
    const rows = getMktRows().map((r,i) => ({
      Rang: i+1, Marché: r.name, Shops: r.shopCount, Tickets: r.payin_count,
      'Dépôts EUR': r.dep_eur.toFixed(2), 'Retraits EUR': r.wd_eur.toFixed(2),
      'Net EUR': (r.dep_eur - r.wd_eur).toFixed(2),
    }));
    const wb = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(wb, XLSX.utils.json_to_sheet(rows), 'Dépôts par Marché');
    saveAs(makeBlob(wb), `Deposits_Marche_${label}.xlsx`);
  }

  // 7. Comparatif mois précédent
  function exportComparatifMoisPrec() {
    if (!store || !country) return;
    const allMonths = Object.keys(store[country]||{}).sort().reverse();
    if (allMonths.length < 2) { alert('Au moins 2 mois nécessaires pour la comparaison'); return; }
    const mCurr = allMonths[0], mPrev = allMonths[1];
    const curr = store[country][mCurr]?.betshop || {};
    const prev = store[country][mPrev]?.betshop || {};
    const allShops = [...new Set([...Object.keys(curr), ...Object.keys(prev)])];
    const rows = allShops.map(shop => {
      const c = curr[shop] || {payin_count:0,payin_money:0};
      const p = prev[shop] || {payin_count:0,payin_money:0};
      const delta = c.payin_count - p.payin_count;
      const pct = p.payin_count > 0 ? ((delta/p.payin_count)*100).toFixed(1)+'%' : '—';
      return {
        Betshop: shop, Marché: c.market || p.market || '',
        [`Tickets ${mCurr}`]: c.payin_count,
        [`Tickets ${mPrev}`]: p.payin_count,
        'Δ Tickets': delta, '% Évolution': pct,
        [`Payin ${mCurr} (XAF)`]: Math.round(c.payin_money),
        [`Payin ${mPrev} (XAF)`]: Math.round(p.payin_money),
      };
    }).sort((a,b) => b[`Tickets ${mCurr}`] - a[`Tickets ${mCurr}`]);
    const wb = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(wb, XLSX.utils.json_to_sheet(rows), `${mCurr} vs ${mPrev}`);
    saveAs(makeBlob(wb), `Comparatif_${mCurr}_vs_${mPrev}_${country}.xlsx`);
  }

  // 8. Comparatif multi-mois
  function exportComparatifMultiMois() {
    if (!store || !country) return;
    const allMonths = Object.keys(store[country]||{}).sort();
    if (allMonths.length < 2) { alert('Au moins 2 mois nécessaires'); return; }
    const allShops = [...new Set(allMonths.flatMap(m =>
      Object.keys(store[country][m]?.betshop || {})))];
    const wb = XLSX.utils.book_new();
    // Un onglet par mois
    allMonths.forEach(m => {
      const bs = store[country][m]?.betshop || {};
      const rows = allShops.map(shop => {
        const r = bs[shop] || {payin_count:0,payin_money:0,paid_out:0,balance:0};
        return { Betshop: shop, Marché: r.market||'', Tickets: r.payin_count,
          'Payin (XAF)': Math.round(r.payin_money), 'Balance': Math.round(r.balance) };
      }).filter(r => r.Tickets > 0).sort((a,b) => b.Tickets - a.Tickets);
      XLSX.utils.book_append_sheet(wb, XLSX.utils.json_to_sheet(rows), m.slice(0,31));
    });
    // Onglet synthèse évolution
    const synthRows = allShops.map(shop => {
      const row = { Betshop: shop };
      allMonths.forEach(m => {
        row[m] = store[country][m]?.betshop?.[shop]?.payin_count || 0;
      });
      return row;
    }).filter(r => allMonths.some(m => r[m] > 0))
      .sort((a,b) => b[allMonths[allMonths.length-1]] - a[allMonths[allMonths.length-1]]);
    XLSX.utils.book_append_sheet(wb, XLSX.utils.json_to_sheet(synthRows), 'Synthèse');
    saveAs(makeBlob(wb), `Comparatif_MultiMois_${country}.xlsx`);
  }

  // 9. Tout en ZIP
  async function exportAllZip() {
    const JSZip = (await import('jszip')).default;
    const zip = new JSZip();
    const add = (wb, name) => {
      const buf = XLSX.write(wb,{bookType:'xlsx',type:'array'});
      zip.file(name, buf);
    };

    // Top 10 Betshops
    const wb1 = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(wb1, XLSX.utils.json_to_sheet(
      getBetshopRows().slice(0,10).map((r,i) => ({
        Rang:i+1,Betshop:r.name,Marché:r.market,Tickets:r.payin_count,
        'Payin (XAF)':Math.round(r.payin_money),'Balance':Math.round(r.balance)
      }))), 'Top10');
    add(wb1, `Top10_Betshops_${label}.xlsx`);

    // Top 10 Caissiers
    const wb2 = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(wb2, XLSX.utils.json_to_sheet(
      getCashierRows().slice(0,10).map((r,i) => ({
        Rang:i+1,Caissier:r.name,Betshop:r.betshop,Tickets:r.payin_count,
        'Payin (XAF)':Math.round(r.payin_money)
      }))), 'Top10');
    add(wb2, `Top10_Caissiers_${label}.xlsx`);

    // Performance complète (réutilise la fonction)
    const wb3 = XLSX.utils.book_new();
    const bRows = getBetshopRows().map((r,i) => ({
      Rang:i+1,Betshop:r.name,Marché:r.market,Caissiers:r.cashierCount,
      Shifts:r.shifts,Tickets:r.payin_count,'Payin (XAF)':Math.round(r.payin_money),
      'Paid out':Math.round(r.paid_out),'Balance':Math.round(r.balance),
      'Dépôts EUR':r.dep_eur.toFixed(2),'Retraits EUR':r.wd_eur.toFixed(2),
    }));
    const cRows = getCashierRows().map((r,i) => ({
      Rang:i+1,Caissier:r.name,Betshop:r.betshop,Marché:r.market,
      Shifts:r.shifts,Tickets:r.payin_count,'Payin (XAF)':Math.round(r.payin_money),
    }));
    const mRows2 = getMktRows().map((r,i) => ({
      Rang:i+1,Marché:r.name,Shops:r.shopCount,Tickets:r.payin_count,
      'Dépôts EUR':r.dep_eur.toFixed(2),'Retraits EUR':r.wd_eur.toFixed(2),
      'Net EUR':(r.dep_eur-r.wd_eur).toFixed(2),
    }));
    XLSX.utils.book_append_sheet(wb3, XLSX.utils.json_to_sheet(bRows), 'Betshops');
    XLSX.utils.book_append_sheet(wb3, XLSX.utils.json_to_sheet(cRows), 'Caissiers');
    XLSX.utils.book_append_sheet(wb3, XLSX.utils.json_to_sheet(mRows2), 'Marchés');
    add(wb3, `Performance_Complete_${label}.xlsx`);

    const blob = await zip.generateAsync({type:'blob'});
    saveAs(blob, `Retail_Reports_${label}.zip`);
  }

  // ── DrillDown Panel ─────────────────────────────────────────────────────────
  function DrillDown({ row, type, onClose }) {
    // ── Hooks AVANT tout early return (règle React) ──
    const [visibleCols, setVisibleCols] = React.useState(['shifts','tickets','payin','paid_out','balance']);
    const [showColCfg,  setShowColCfg]  = React.useState(false);
    const [draftCols,   setDraftCols]   = React.useState(['shifts','tickets','payin','paid_out','balance']);

    if (!row || !row.byDay) return null;
    const days = Object.keys(row.byDay).sort();
    const maxTickets = Math.max(...days.map(d => row.byDay[d].payin_count), 1);
    const barW = Math.max(18, Math.min(48, Math.floor(520 / days.length)));

    // Définition de toutes les colonnes disponibles
    const ALL_COLS = [
      { k:'shifts',      label:'Shifts',       render: dv => <td key="shifts" style={{ padding:'5px 10px', textAlign:'right', color:'var(--muted)' }}>{dv.shifts}</td>,                                                                                total: () => row.shifts },
      { k:'tickets',     label:'Tickets',      render: dv => <td key="tickets" style={{ padding:'5px 10px', textAlign:'right', fontWeight:700, color:'#22C55E' }}>{dv.payin_count.toLocaleString('fr-FR')}</td>,                                       total: () => row.payin_count.toLocaleString('fr-FR') },
      { k:'payin',       label:'Payin (XAF)',  render: dv => <td key="payin" style={{ padding:'5px 10px', textAlign:'right', fontFamily:'monospace' }}>{fmt(dv.payin_money)}</td>,                                                                     total: () => fmt(row.payin_money) },
      { k:'paid_out',    label:'Paid out',     render: dv => <td key="paid_out" style={{ padding:'5px 10px', textAlign:'right', fontFamily:'monospace', color:'var(--muted)' }}>{fmt(dv.paid_out)}</td>,                                               total: () => fmt(row.paid_out) },
      { k:'balance',      label:'Balance',       render: dv => <td key="balance"      style={{ padding:'5px 10px', textAlign:'right', fontFamily:'monospace', fontWeight:600, color:(dv.balance||0)>=0?'var(--deposit)':'var(--danger)' }}>{fmt(dv.balance)}</td>,      total: () => <span style={{color:(row.balance||0)>=0?'var(--deposit)':'var(--danger)'}}>{fmt(row.balance)}</span> },
      { k:'kill_count',   label:'Kill count',    render: dv => <td key="kill_count"   style={{ padding:'5px 10px', textAlign:'right', fontFamily:'monospace', color:'#F87171' }}>{(dv.kill_count||0).toLocaleString('fr-FR')}</td>,   total: () => (row.kill_count||0).toLocaleString('fr-FR') },
      { k:'kill_money',   label:'Kill money',    render: dv => <td key="kill_money"   style={{ padding:'5px 10px', textAlign:'right', fontFamily:'monospace', color:'#F87171' }}>{fmt(dv.kill_money||0)}</td>,                         total: () => fmt(row.kill_money||0) },
      { k:'revoke_count', label:'Revoke count',  render: dv => <td key="revoke_count" style={{ padding:'5px 10px', textAlign:'right', fontFamily:'monospace', color:'#FBBF24' }}>{(dv.revoke_count||0).toLocaleString('fr-FR')}</td>, total: () => (row.revoke_count||0).toLocaleString('fr-FR') },
      { k:'revoke_money', label:'Revoke money',  render: dv => <td key="revoke_money" style={{ padding:'5px 10px', textAlign:'right', fontFamily:'monospace', color:'#FBBF24' }}>{fmt(dv.revoke_money||0)}</td>,                     total: () => fmt(row.revoke_money||0) },
    ];

    const activeCols = ALL_COLS.filter(c2 => visibleCols.includes(c2.k));

    return (
      <div style={{
        margin:'12px 0 16px', background:'var(--bg2)',
        border:'1px solid var(--accent)', borderRadius:12, padding:'16px 18px',
        position:'relative',
      }}>

        {/* ── Modal Column Configuration ── */}
        {showColCfg && (
          <div style={{
            position:'fixed', inset:0, background:'rgba(0,0,0,0.45)', zIndex:1000,
            display:'flex', alignItems:'center', justifyContent:'center',
          }}>
            <div style={{
              background:'var(--bg2)', borderRadius:14, padding:'24px', width:340,
              boxShadow:'0 20px 60px rgba(0,0,0,0.4)', maxHeight:'80vh', overflowY:'auto',
            }}>
              <div style={{ display:'flex', justifyContent:'space-between', alignItems:'center', marginBottom:18 }}>
                <span style={{ fontSize:13, fontWeight:700, color:'var(--text)', textTransform:'uppercase', letterSpacing:'.06em' }}>
                  Colonnes visibles
                </span>
                <button onClick={() => setShowColCfg(false)} style={{
                  background:'none', border:'none', fontSize:18, cursor:'pointer', color:'var(--muted)',
                }}>✕</button>
              </div>
              {ALL_COLS.map(col => (
                <div key={col.k} style={{
                  display:'flex', alignItems:'center', gap:12, padding:'10px 0',
                  borderBottom:'1px solid var(--border)',
                }}>
                  <div
                    onClick={() => setDraftCols(prev =>
                      prev.includes(col.k) ? prev.filter(x => x !== col.k) : [...prev, col.k]
                    )}
                    style={{
                      width:22, height:22, borderRadius:6, cursor:'pointer', flexShrink:0,
                      border: draftCols.includes(col.k) ? 'none' : '2px solid var(--border)',
                      background: draftCols.includes(col.k) ? '#16A34A' : 'transparent',
                      display:'flex', alignItems:'center', justifyContent:'center',
                    }}
                  >
                    {draftCols.includes(col.k) && <span style={{ color:'#fff', fontSize:13, fontWeight:700 }}>✓</span>}
                  </div>
                  <span style={{ fontSize:13, color:'var(--text)' }}>{col.label}</span>
                </div>
              ))}
              <div style={{ display:'flex', gap:10, marginTop:20, justifyContent:'flex-end' }}>
                <button onClick={() => { setDraftCols(visibleCols); setShowColCfg(false); }} style={{
                  padding:'7px 18px', borderRadius:8, border:'1px solid var(--border)',
                  background:'var(--bg3)', color:'var(--muted)', cursor:'pointer', fontSize:12, fontWeight:600,
                }}>Annuler</button>
                <button onClick={() => { setVisibleCols(draftCols); setShowColCfg(false); }} style={{
                  padding:'7px 18px', borderRadius:8, border:'none',
                  background:'#16A34A', color:'#fff', cursor:'pointer', fontSize:12, fontWeight:700,
                }}>Enregistrer</button>
              </div>
            </div>
          </div>
        )}

        {/* Header */}
        <div style={{ display:'flex', alignItems:'center', justifyContent:'space-between', marginBottom:14 }}>
          <div>
            <span style={{ fontSize:13, fontWeight:700, color:'var(--text)' }}>
              {type === 'betshop' ? '🏪' : '👤'} {row.name}
            </span>
            <span style={{ fontSize:11, color:'var(--muted)', marginLeft:8 }}>
              {row.market} · {days.length} jours actifs · {row.shifts} shifts · {row.payin_count.toLocaleString('fr-FR')} tickets
            </span>
          </div>
          <div style={{ display:'flex', gap:6, alignItems:'center' }}>
            <button onClick={() => {
              const XLSX = window.XLSX;
              if (!XLSX) { alert('XLSX non chargé'); return; }
              const days2 = Object.keys(row.byDay).sort();
              const header = ['Date','Shifts','Tickets','Payin (XAF)','Paid out','Balance'];
              const dataRows = days2.map(d => {
                const dv = row.byDay[d];
                return [d.slice(5), dv.shifts, dv.payin_count, dv.payin_money, dv.paid_out, dv.balance||0];
              });
              dataRows.push(['TOTAL', row.shifts, row.payin_count, row.payin_money, row.paid_out, row.balance||0]);
              const ws = XLSX.utils.aoa_to_sheet([header, ...dataRows]);
              const wb = XLSX.utils.book_new();
              XLSX.utils.book_append_sheet(wb, ws, row.name);
              XLSX.writeFile(wb, `drilldown_${row.name}.xlsx`);
            }} style={{
              fontSize:11, padding:'4px 10px', borderRadius:6, border:'none', cursor:'pointer',
              background:'#16A34A', color:'#fff', fontWeight:700,
            }}>📥 Excel</button>
            <button onClick={() => {
              const days2 = Object.keys(row.byDay).sort();
              const rows2 = days2.map(d => {
                const dv = row.byDay[d];
                return `<tr><td>${d.slice(5)}</td><td>${dv.shifts}</td><td>${dv.payin_count.toLocaleString('fr-FR')}</td><td>${dv.payin_money.toLocaleString('fr-FR')}</td><td>${dv.paid_out.toLocaleString('fr-FR')}</td><td style="color:${(dv.balance||0)>=0?'green':'red'}">${(dv.balance||0).toLocaleString('fr-FR')}</td></tr>`;
              }).join('');
              const win = window.open('', '_blank');
              win.document.write(`<!DOCTYPE html><html><head><title>${row.name}</title>
                <style>body{font-family:sans-serif;padding:20px;font-size:12px;}
                h2{margin-bottom:4px;}p{color:#64748b;margin:0 0 12px;}
                table{border-collapse:collapse;width:100%;}
                th,td{border:1px solid #e2e8f0;padding:6px 10px;text-align:right;}
                th{background:#f1f5f9;text-align:left;}td:first-child{text-align:left;}
                tfoot td{font-weight:700;background:#f8fafc;}</style></head>
                <body>
                <h2>👤 ${row.name}</h2>
                <p>${row.market} · ${days2.length} jours · ${row.shifts} shifts · ${row.payin_count.toLocaleString('fr-FR')} tickets</p>
                <table><thead><tr><th>Date</th><th>Shifts</th><th>Tickets</th><th>Payin (XAF)</th><th>Paid out</th><th>Balance</th></tr></thead>
                <tbody>${rows2}</tbody>
                <tfoot><tr><td>TOTAL</td><td>${row.shifts}</td><td>${row.payin_count.toLocaleString('fr-FR')}</td><td>${row.payin_money.toLocaleString('fr-FR')}</td><td>${row.paid_out.toLocaleString('fr-FR')}</td><td>${(row.balance||0).toLocaleString('fr-FR')}</td></tr></tfoot>
                </table>
                <script>window.onload=()=>window.print();<\/script></body></html>`);
              win.document.close();
            }} style={{
              fontSize:11, padding:'4px 10px', borderRadius:6, border:'none', cursor:'pointer',
              background:'#DC2626', color:'#fff', fontWeight:700,
            }}>📄 PDF</button>
            <button onClick={() => { setDraftCols(visibleCols); setShowColCfg(true); }} style={{
              fontSize:11, padding:'4px 10px', borderRadius:6, border:'1px solid var(--border)',
              background:'var(--bg3)', color:'var(--text)', cursor:'pointer', fontWeight:600,
            }}>⚙️ Colonnes</button>
            <button onClick={onClose} style={{
              background:'none', border:'none', fontSize:16, cursor:'pointer',
              color:'var(--muted)', lineHeight:1, marginLeft:4,
            }}>✕</button>
          </div>
        </div>

        {/* Graphique barres */}
        <div style={{ marginBottom:14 }}>
          <div style={{ display:'flex', alignItems:'center', gap:16, marginBottom:8 }}>
            <div style={{ fontSize:11, color:'var(--muted)', fontWeight:600,
              textTransform:'uppercase', letterSpacing:'.05em' }}>
              Tickets / jour
            </div>
            {(() => {
              const vals = days.map(d => row.byDay[d].payin_count);
              const avg  = vals.length ? Math.round(vals.reduce((a,b)=>a+b,0) / vals.length) : 0;
              const med  = vals.length ? [...vals].sort((a,b)=>a-b)[Math.floor(vals.length/2)] : 0;
              // Meilleur jour (date)
              const bestDay = days.reduce((a,b) => row.byDay[a].payin_count >= row.byDay[b].payin_count ? a : b, days[0]);
              // % jours au-dessus de la moyenne
              const aboveAvg = vals.filter(v => v > avg).length;
              const pctAbove = vals.length ? Math.round(aboveAvg / vals.length * 100) : 0;
              return (
                <div style={{ display:'flex', gap:8, flexWrap:'wrap' }}>
                  <div style={{ fontSize:11, background:'rgba(99,102,241,.12)', borderRadius:5,
                    padding:'3px 9px', color:'var(--accent)', fontWeight:600 }}>
                    Moy. <span style={{ fontFamily:'monospace' }}>{avg.toLocaleString('fr-FR')}</span>
                  </div>
                  <div style={{ fontSize:11, background:'rgba(52,211,153,.1)', borderRadius:5,
                    padding:'3px 9px', color:'#34D399', fontWeight:600 }}>
                    Méd. <span style={{ fontFamily:'monospace' }}>{med.toLocaleString('fr-FR')}</span>
                  </div>
                  <div style={{ fontSize:11, background:'rgba(251,191,36,.1)', borderRadius:5,
                    padding:'3px 9px', color:'#FBBF24', fontWeight:600 }}>
                    Max <span style={{ fontFamily:'monospace' }}>{maxTickets.toLocaleString('fr-FR')}</span>
                  </div>
                  <div style={{ fontSize:11, background:'rgba(96,165,250,.1)', borderRadius:5,
                    padding:'3px 9px', color:'#60A5FA', fontWeight:600 }}>
                    📅 <span style={{ fontFamily:'monospace' }}>{bestDay.slice(5)}</span>
                  </div>
                  <div style={{ fontSize:11, background:'rgba(167,139,250,.1)', borderRadius:5,
                    padding:'3px 9px', color:'#A78BFA', fontWeight:600 }}>
                    {days.length} jours actifs
                  </div>
                  <div style={{ fontSize:11, background:'rgba(251,113,133,.1)', borderRadius:5,
                    padding:'3px 9px', color:'#FB7185', fontWeight:600 }}>
                    {pctAbove}% &gt; moy.
                  </div>
                </div>
              );
            })()}
          </div>
          <div style={{ overflowX:'auto' }}>
            <div style={{ display:'flex', alignItems:'flex-end', gap:3,
              minWidth: days.length * (barW + 3), height:130, paddingBottom:22, position:'relative' }}>
              {days.map(d => {
                const val = row.byDay[d].payin_count;
                const h = Math.max(4, Math.round((val / maxTickets) * 88));
                const dayLabel = d.slice(8); // "DD"
                const isMax = val === maxTickets;
                // Valeur abrégée au-dessus de chaque barre
                const label = val >= 1000
                  ? (val >= 1000000 ? (val/1000000).toFixed(1)+'M' : Math.round(val/1000)+'K')
                  : val.toString();
                return (
                  <div key={d} style={{ display:'flex', flexDirection:'column',
                    alignItems:'center', flex:'0 0 auto', width:barW, position:'relative' }}>
                    {/* Valeur au-dessus — toujours visible */}
                    <div style={{
                      fontSize: isMax ? 9 : 8,
                      color: isMax ? 'var(--accent)' : 'var(--muted)',
                      fontWeight: isMax ? 700 : 500,
                      marginBottom:2, whiteSpace:'nowrap',
                      opacity: val === 0 ? 0 : 1,
                    }}>
                      {label}
                    </div>
                    <div title={`${d} : ${val.toLocaleString('fr-FR')} tickets`}
                      style={{
                        width:'100%', height:h, borderRadius:'3px 3px 0 0',
                        background: isMax ? 'var(--accent)' : 'rgba(99,102,241,.45)',
                        transition:'height .2s',
                      }}/>
                    <div style={{ fontSize:9, color:'var(--muted)', marginTop:3,
                      transform:'rotate(-40deg)', transformOrigin:'top center',
                      whiteSpace:'nowrap', marginLeft:-4 }}>
                      {dayLabel}
                    </div>
                  </div>
                );
              })}
            </div>
          </div>
        </div>

        {/* Tableau jour par jour */}
        <div style={{ fontSize:11, color:'var(--muted)', fontWeight:600,
          textTransform:'uppercase', letterSpacing:'.05em', marginBottom:6 }}>
          Détail jour par jour
        </div>
        <div style={{ overflowX:'auto', maxHeight:220, overflowY:'auto',
          border:'0.5px solid var(--border)', borderRadius:8 }}>
          <table style={{ width:'100%', borderCollapse:'collapse', fontSize:11 }}>
            <thead>
              <tr style={{ background:'var(--bg3)', position:'sticky', top:0 }}>
                <th style={{ padding:'6px 10px', textAlign:'left', color:'var(--muted)', fontWeight:600 }}>Date</th>
                {activeCols.map(col => (
                  <th key={col.k} style={{ padding:'6px 10px', textAlign:'right', color:'var(--muted)', fontWeight:600 }}>
                    {col.label}
                  </th>
                ))}
              </tr>
            </thead>
            <tbody>
              {days.map((d, i) => {
                const dv = row.byDay[d];
                return (
                  <tr key={d} style={{ borderTop:'0.5px solid var(--border)',
                    background: i%2===0?'transparent':'rgba(255,255,255,.015)' }}>
                    <td style={{ padding:'5px 10px', fontWeight:600, color:'var(--text)' }}>{d.slice(5)}</td>
                    {activeCols.map(col => col.render(dv))}
                  </tr>
                );
              })}
              <tr style={{ borderTop:'2px solid var(--border)', background:'var(--bg3)', fontWeight:700 }}>
                <td style={{ padding:'6px 10px' }}>TOTAL</td>
                {activeCols.map(col => (
                  <td key={col.k} style={{ padding:'6px 10px', textAlign:'right', fontFamily:'monospace' }}>
                    {col.total()}
                  </td>
                ))}
              </tr>
            </tbody>
          </table>
        </div>
      </div>
    );
  }

  // ── DrillDown Marché ────────────────────────────────────────────────────────
  function DrillDownMarket({ row, onClose }) {
    if (!row || !row.byDay) return null;
    const days = Object.keys(row.byDay).sort();
    const maxDep = Math.max(...days.map(d => row.byDay[d].dep_eur), 1);

    return (
      <div style={{
        margin:'12px 0 16px', background:'var(--bg2)',
        border:'1px solid #34D399', borderRadius:12, padding:'16px 18px',
        position:'relative',
      }}>
        {/* Header */}
        <div style={{ display:'flex', alignItems:'center', justifyContent:'space-between', marginBottom:14 }}>
          <div>
            <span style={{ fontSize:13, fontWeight:700, color:'var(--text)' }}>
              💶 {row.name}
            </span>
            <span style={{ fontSize:11, color:'var(--muted)', marginLeft:8 }}>
              {days.length} jours actifs · {row.shopCount} shops · {row.payin_count.toLocaleString('fr-FR')} tickets
            </span>
          </div>
          <div style={{ display:'flex', gap:6, alignItems:'center' }}>
            <button onClick={() => {
              const XLSX = window.XLSX;
              if (!XLSX) { alert('XLSX non chargé'); return; }
              const days2 = Object.keys(row.byDay).sort();
              const header = ['Date','Shops','Tickets','Dépôts EUR','Retraits EUR','Net EUR'];
              const dataRows = days2.map(d => {
                const dv = row.byDay[d];
                const net = dv.dep_eur - dv.wd_eur;
                return [d.slice(5), dv.shopCount||0, dv.payin_count, dv.dep_eur, dv.wd_eur, net];
              });
              const totNet = row.dep_eur - row.wd_eur;
              dataRows.push(['TOTAL', row.shopCount, row.payin_count, row.dep_eur, row.wd_eur, totNet]);
              const ws = XLSX.utils.aoa_to_sheet([header, ...dataRows]);
              const wb = XLSX.utils.book_new();
              XLSX.utils.book_append_sheet(wb, ws, row.name);
              XLSX.writeFile(wb, `marche_${row.name}.xlsx`);
            }} style={{
              fontSize:11, padding:'4px 10px', borderRadius:6, border:'none', cursor:'pointer',
              background:'#16A34A', color:'#fff', fontWeight:700,
            }}>📥 Excel</button>
            <button onClick={() => {
              const days2 = Object.keys(row.byDay).sort();
              const rows2 = days2.map(d => {
                const dv = row.byDay[d];
                const net = dv.dep_eur - dv.wd_eur;
                return `<tr><td>${d.slice(5)}</td><td>${dv.shopCount||'—'}</td><td>${dv.payin_count.toLocaleString('fr-FR')}</td><td style="color:green">€${dv.dep_eur.toFixed(2)}</td><td style="color:#dc2626">€${dv.wd_eur.toFixed(2)}</td><td style="color:${net>=0?'green':'#dc2626'}">${net>=0?'+':''}€${net.toFixed(2)}</td></tr>`;
              }).join('');
              const totNet = row.dep_eur - row.wd_eur;
              const win = window.open('', '_blank');
              win.document.write(`<!DOCTYPE html><html><head><title>${row.name}</title>
                <style>body{font-family:sans-serif;padding:20px;font-size:12px;}
                h2{margin-bottom:4px;}p{color:#64748b;margin:0 0 12px;}
                table{border-collapse:collapse;width:100%;}
                th,td{border:1px solid #e2e8f0;padding:6px 10px;text-align:right;}
                th{background:#f1f5f9;text-align:left;}td:first-child{text-align:left;}
                tfoot td{font-weight:700;background:#f8fafc;}</style></head>
                <body>
                <h2>💶 ${row.name}</h2>
                <p>${days2.length} jours · ${row.shopCount} shops · ${row.payin_count.toLocaleString('fr-FR')} tickets</p>
                <table><thead><tr><th>Date</th><th>Shops</th><th>Tickets</th><th>Dépôts EUR</th><th>Retraits EUR</th><th>Net EUR</th></tr></thead>
                <tbody>${rows2}</tbody>
                <tfoot><tr><td>TOTAL</td><td>${row.shopCount}</td><td>${row.payin_count.toLocaleString('fr-FR')}</td><td style="color:green">€${row.dep_eur.toFixed(2)}</td><td style="color:#dc2626">€${row.wd_eur.toFixed(2)}</td><td style="color:${totNet>=0?'green':'#dc2626'}">${totNet>=0?'+':''}€${totNet.toFixed(2)}</td></tr></tfoot>
                </table>
                <script>window.onload=()=>window.print();<\/script></body></html>`);
              win.document.close();
            }} style={{
              fontSize:11, padding:'4px 10px', borderRadius:6, border:'none', cursor:'pointer',
              background:'#DC2626', color:'#fff', fontWeight:700,
            }}>📄 PDF</button>
            <button onClick={onClose} style={{
              background:'none', border:'none', fontSize:16, cursor:'pointer',
              color:'var(--muted)', lineHeight:1, marginLeft:4,
            }}>✕</button>
          </div>
        </div>

        {/* Graphique double barres Dépôts / Retraits */}
        <div style={{ marginBottom:14 }}>
          <div style={{ display:'flex', gap:16, marginBottom:8 }}>
            <span style={{ fontSize:11, color:'var(--muted)', fontWeight:600,
              textTransform:'uppercase', letterSpacing:'.05em' }}>Dépôts & Retraits EUR / jour</span>
            <span style={{ fontSize:11 }}>
              <span style={{ display:'inline-block', width:10, height:10,
                background:'#34D399', borderRadius:2, marginRight:4 }}/>Dépôts
              <span style={{ display:'inline-block', width:10, height:10,
                background:'#F87171', borderRadius:2, margin:'0 4px 0 12px' }}/>Retraits
            </span>
          </div>
          <div style={{ overflowX:'auto' }}>
            {(() => {
              const barW = Math.max(20, Math.min(52, Math.floor(520 / days.length)));
              return (
                <div style={{ display:'flex', alignItems:'flex-end', gap:4,
                  minWidth: days.length * (barW + 4), height:120, paddingBottom:24, position:'relative' }}>
                  {days.map(d => {
                    const dv = row.byDay[d];
                    const hDep = Math.max(3, Math.round((dv.dep_eur / maxDep) * 96));
                    const hWd  = Math.max(0, Math.round((dv.wd_eur  / maxDep) * 96));
                    const net  = dv.dep_eur - dv.wd_eur;
                    const dayLabel = d.slice(8);
                    return (
                      <div key={d} title={`${d}
Dépôts: €${dv.dep_eur.toFixed(2)}
Retraits: €${dv.wd_eur.toFixed(2)}
Net: €${net.toFixed(2)}`}
                        style={{ display:'flex', flexDirection:'column', alignItems:'center',
                          flex:'0 0 auto', width:barW, gap:1 }}>
                        <div style={{ display:'flex', alignItems:'flex-end', gap:2, height:96 }}>
                          <div style={{ width:(barW/2)-1, height:hDep, background:'#34D399',
                            borderRadius:'3px 3px 0 0' }}/>
                          <div style={{ width:(barW/2)-1, height:hWd,  background:'#F87171',
                            borderRadius:'3px 3px 0 0' }}/>
                        </div>
                        <div style={{ fontSize:9, color:'var(--muted)', marginTop:3,
                          transform:'rotate(-40deg)', transformOrigin:'top center',
                          whiteSpace:'nowrap', marginLeft:-4 }}>
                          {dayLabel}
                        </div>
                      </div>
                    );
                  })}
                </div>
              );
            })()}
          </div>
        </div>

        {/* Tableau jour par jour */}
        <div style={{ fontSize:11, color:'var(--muted)', fontWeight:600,
          textTransform:'uppercase', letterSpacing:'.05em', marginBottom:6 }}>
          Détail jour par jour
        </div>
        <div style={{ overflowX:'auto', maxHeight:220, overflowY:'auto',
          border:'0.5px solid var(--border)', borderRadius:8 }}>
          <table style={{ width:'100%', borderCollapse:'collapse', fontSize:11 }}>
            <thead>
              <tr style={{ background:'var(--bg3)', position:'sticky', top:0 }}>
                <th style={{ padding:'6px 10px', textAlign:'left',  color:'var(--muted)', fontWeight:600 }}>Date</th>
                <th style={{ padding:'6px 10px', textAlign:'right', color:'var(--muted)', fontWeight:600 }}>Shops</th>
                <th style={{ padding:'6px 10px', textAlign:'right', color:'var(--muted)', fontWeight:600 }}>Tickets</th>
                <th style={{ padding:'6px 10px', textAlign:'right', color:'#34D399',      fontWeight:600 }}>Dépôts EUR</th>
                <th style={{ padding:'6px 10px', textAlign:'right', color:'#F87171',      fontWeight:600 }}>Retraits EUR</th>
                <th style={{ padding:'6px 10px', textAlign:'right', color:'var(--muted)', fontWeight:600 }}>Net EUR</th>
              </tr>
            </thead>
            <tbody>
              {days.map((d, i) => {
                const dv  = row.byDay[d];
                const net = dv.dep_eur - dv.wd_eur;
                return (
                  <tr key={d} style={{ borderTop:'0.5px solid var(--border)',
                    background: i%2===0?'transparent':'rgba(255,255,255,.015)' }}>
                    <td style={{ padding:'5px 10px', fontWeight:600 }}>{d.slice(5)}</td>
                    <td style={{ padding:'5px 10px', textAlign:'right', color:'var(--muted)' }}>{dv.shopCount||'—'}</td>
                    <td style={{ padding:'5px 10px', textAlign:'right', fontWeight:700, color:'#22C55E' }}>
                      {dv.payin_count.toLocaleString('fr-FR')}
                    </td>
                    <td style={{ padding:'5px 10px', textAlign:'right', color:'#34D399', fontWeight:600 }}>
                      €{dv.dep_eur.toFixed(2)}
                    </td>
                    <td style={{ padding:'5px 10px', textAlign:'right', color:'#F87171' }}>
                      €{dv.wd_eur.toFixed(2)}
                    </td>
                    <td style={{ padding:'5px 10px', textAlign:'right', fontWeight:600,
                      color: net >= 0 ? 'var(--deposit)' : 'var(--danger)' }}>
                      {net >= 0 ? '+' : ''}€{net.toFixed(2)}
                    </td>
                  </tr>
                );
              })}
              {/* Total */}
              <tr style={{ borderTop:'2px solid var(--border)', background:'var(--bg3)', fontWeight:700 }}>
                <td style={{ padding:'6px 10px' }}>TOTAL</td>
                <td style={{ padding:'6px 10px', textAlign:'right' }}>{row.shopCount}</td>
                <td style={{ padding:'6px 10px', textAlign:'right', color:'#22C55E' }}>
                  {row.payin_count.toLocaleString('fr-FR')}
                </td>
                <td style={{ padding:'6px 10px', textAlign:'right', color:'#34D399' }}>€{row.dep_eur.toFixed(2)}</td>
                <td style={{ padding:'6px 10px', textAlign:'right', color:'#F87171' }}>€{row.wd_eur.toFixed(2)}</td>
                <td style={{ padding:'6px 10px', textAlign:'right', fontWeight:600,
                  color: (row.dep_eur - row.wd_eur) >= 0 ? 'var(--deposit)' : 'var(--danger)' }}>
                  {(row.dep_eur - row.wd_eur) >= 0 ? '+' : ''}€{(row.dep_eur - row.wd_eur).toFixed(2)}
                </td>
              </tr>
            </tbody>
          </table>
        </div>
      </div>
    );
  }

  // ── Export Panel UI ───────────────────────────────────────────────────────
  function ExportPanel() {
    const [open, setOpen] = React.useState(false);
    const [zipping, setZipping] = React.useState(false);
    const btnBase = {
      fontSize:11, padding:'4px 10px', borderRadius:5, cursor:'pointer',
      border:'0.5px solid var(--border)', fontWeight:500,
    };
    const xlsBtn = { ...btnBase, background:'#E6F1FB', color:'#185FA5', borderColor:'#B5D4F4' };
    const pdfBtn = { ...btnBase, background:'#FAECE7', color:'#993C1D', borderColor:'#F5C4B3' };
    const zipBtn = { ...btnBase, background:'var(--bg3)', color:'var(--muted)', fontSize:12, padding:'6px 14px' };

    return (
      <div style={{ marginBottom:14 }}>
        <button onClick={() => setOpen(o => !o)} style={{
          fontSize:12, padding:'7px 14px', borderRadius:7,
          border:'0.5px solid var(--border)', background:'var(--bg3)',
          color:'var(--text)', cursor:'pointer', fontWeight:500,
        }}>
          <i className="ti ti-download" style={{ fontSize:13, marginRight:5 }} aria-hidden="true"/>
          Exports {open ? '▲' : '▼'}
        </button>

        {open && (
          <div style={{
            marginTop:8, background:'var(--bg2)', border:'0.5px solid var(--border)',
            borderRadius:10, padding:'14px 16px',
          }}>
            <div style={{ display:'grid', gridTemplateColumns:'1fr 1fr', gap:8, marginBottom:12 }}>

              <div style={{ display:'flex', alignItems:'center', justifyContent:'space-between',
                padding:'8px 12px', background:'var(--bg3)', borderRadius:7 }}>
                <span style={{ fontSize:12, color:'var(--text)' }}>🏪 Top 10 Betshops</span>
                <div style={{ display:'flex', gap:5 }}>
                  <button style={xlsBtn} onClick={exportTop10BetshopsXLSX}>📊 Excel</button>
                  <button style={pdfBtn} onClick={exportTop10BetshopsPDF}>📄 PDF</button>
                </div>
              </div>

              <div style={{ display:'flex', alignItems:'center', justifyContent:'space-between',
                padding:'8px 12px', background:'var(--bg3)', borderRadius:7 }}>
                <span style={{ fontSize:12, color:'var(--text)' }}>👤 Top 10 Caissiers</span>
                <div style={{ display:'flex', gap:5 }}>
                  <button style={xlsBtn} onClick={exportTop10CashiersXLSX}>📊 Excel</button>
                  <button style={pdfBtn} onClick={exportTop10CashiersPDF}>📄 PDF</button>
                </div>
              </div>

              <div style={{ display:'flex', alignItems:'center', justifyContent:'space-between',
                padding:'8px 12px', background:'var(--bg3)', borderRadius:7 }}>
                <span style={{ fontSize:12, color:'var(--text)' }}>📋 Performance complète</span>
                <button style={xlsBtn} onClick={exportPerformanceComplete}>📊 Excel (4 onglets)</button>
              </div>

              <div style={{ display:'flex', alignItems:'center', justifyContent:'space-between',
                padding:'8px 12px', background:'var(--bg3)', borderRadius:7 }}>
                <span style={{ fontSize:12, color:'var(--text)' }}>💶 Dépôts & Retraits</span>
                <button style={xlsBtn} onClick={exportDepositsMarche}>📊 Excel</button>
              </div>

              <div style={{ display:'flex', alignItems:'center', justifyContent:'space-between',
                padding:'8px 12px', background:'var(--bg3)', borderRadius:7 }}>
                <span style={{ fontSize:12, color:'var(--text)' }}>🔄 Comparatif mois préc.</span>
                <button style={xlsBtn} onClick={exportComparatifMoisPrec}>📊 Excel</button>
              </div>

              <div style={{ display:'flex', alignItems:'center', justifyContent:'space-between',
                padding:'8px 12px', background:'var(--bg3)', borderRadius:7 }}>
                <span style={{ fontSize:12, color:'var(--text)' }}>📅 Comparatif multi-mois</span>
                <button style={xlsBtn} onClick={exportComparatifMultiMois}>📊 Excel (1 onglet/mois)</button>
              </div>

            </div>

            <button style={{ ...zipBtn, width:'100%' }}
              disabled={zipping}
              onClick={async () => { setZipping(true); await exportAllZip(); setZipping(false); }}>
              {zipping ? '⟳ Compression...' : '📦 Tout télécharger en .zip (Top10 + Performance + Dépôts)'}
            </button>
          </div>
        )}
      </div>
    );
  }

  // ── Render ────────────────────────────────────────────────────────────────
  if (!store) {
    return (
      <div>
        <p className="section-label">Upload shift report(s)</p>
        <div style={{ fontSize:13, color:'var(--muted)', marginBottom:'1rem' }}>
          Rapport de shifts niveau caissier — agrégation automatique par betshop, caissier et marché.
          <br/>Vous pouvez glisser <strong>plusieurs fichiers CSV</strong> d'un coup (multi-mois).
        </div>
        <UploadSlot
          loadedFiles={loadedFiles}
          onFiles={loadFiles}
          onAddFiles={addFiles}
        />
        {loading && <p style={{ color:'var(--muted)', fontSize:13, marginTop:12 }}>⟳ Chargement...</p>}
      </div>
    );
  }

  return (
    <div>
      {/* ── Controls ── */}
      <div style={{ display:'flex', gap:10, alignItems:'center', marginBottom:16, flexWrap:'wrap' }}>
        <select value={country} onChange={e => {
          setCountry(e.target.value);
          const ms = store[e.target.value] ? Object.keys(store[e.target.value]).sort().reverse() : [];
          setMonth(ms[0]||''); setPage(1);
        }} className="field-input" style={{ minWidth:130 }}>
          {countries.map(c => <option key={c} value={c}>{c}</option>)}
        </select>
        <select value={month} onChange={e => { setMonth(e.target.value); setPage(1); }}
          className="field-input" style={{ minWidth:110 }}>
          {months.map(m => <option key={m} value={m}>{m}</option>)}
        </select>
        <div style={{ flex:1 }}/>
        <UploadSlot
          loadedFiles={loadedFiles}
          onFiles={loadFiles}
          onAddFiles={addFiles}
        />
        <button className="btn sm" onClick={() => { setStore(null); setLoadedFiles([]); }}
          style={{ whiteSpace:'nowrap' }}>
          🗑 Réinitialiser
        </button>
      </div>

      {/* ── KPIs ── */}
      {kpis && (
        <div>
          {kpis.filtered && (
            <div style={{ fontSize:11, fontWeight:600, color:'#60A5FA',
              background:'rgba(96,165,250,.12)', padding:'4px 10px', borderRadius:6,
              marginBottom:8, display:'inline-flex', alignItems:'center', gap:6 }}>
              📊 KPIs filtrés ·&nbsp;
              <span style={{ color:'var(--text)' }}>
                {[filterZone, filterSuper, filterProp, filterCity].filter(Boolean).join(' · ')}
              </span>
            </div>
          )}
          <div style={{ display:'flex', gap:10, marginBottom:16, flexWrap:'wrap' }}>
          <KpiCard label="Betshops"  value={kpis.betshops} color="var(--accent)"/>
          <KpiCard label="Cashiers"  value={kpis.cashiers} color="#A78BFA"/>
          <KpiCard label="Shifts"    value={kpis.shifts.toLocaleString()}/>
          <KpiCard label="Tickets"   value={fmt(kpis.tickets)} color="#22C55E"/>
          <KpiCard label="Payin"     value={fmt(kpis.payin)} color="#60A5FA"
            sub={country === 'Cameroon' ? 'XAF' : ''}/>
          <KpiCard label="Dépôts"    value={fmtEUR(kpis.dep_eur)} color="#34D399"/>
          <KpiCard label="Retraits"  value={fmtEUR(kpis.wd_eur)} color="#F87171"/>
          </div>
        </div>
      )}

      {/* ── Export Panel ── */}
      <ExportPanel/>

      {/* ── Main tabs ── */}
      <div style={{ display:'flex', gap:6, marginBottom:14 }}>
        {[
          { k:'performance', label:'🎫 Performance (Betshop / Cashier)' },
          { k:'deposits',    label:'💶 Dépôts & Retraits par Marché' },
          { k:'compare',     label:'📊 Comparaison de Périodes' },
        ].map(t => (
          <button key={t.k} onClick={() => { setMainTab(t.k); setPage(1); setSelectedMkt(null); setSelectedRow(null); }} style={{
            fontSize:12, padding:'7px 14px', borderRadius:7, border:'none',
            cursor:'pointer', fontWeight:600,
            background: mainTab===t.k ? 'var(--accent)' : 'var(--bg3)',
            color: mainTab===t.k ? '#fff' : 'var(--muted)',
          }}>{t.label}</button>
        ))}
      </div>

      {/* ── Performance tab ── */}
      {mainTab === 'performance' && (
        <>
          <div style={{ display:'flex', gap:6, marginBottom:10 }}>
            {[
              { k:'betshop', label:`🏪 Par Betshop (${monthData ? Object.keys(monthData.betshop).length : 0})` },
              { k:'cashier', label:`👤 Par Caissier (${monthData ? Object.keys(monthData.cashier).length : 0})` },
            ].map(t => (
              <button key={t.k} onClick={() => { setSubTab(t.k); setPage(1); onSort('payin_count'); setSelectedRow(null); }} style={{
                fontSize:12, padding:'6px 12px', borderRadius:6, border:'none',
                cursor:'pointer', fontWeight:600,
                background: subTab===t.k ? 'var(--bg)' : 'var(--bg3)',
                color: subTab===t.k ? 'var(--text)' : 'var(--muted)',
                boxShadow: subTab===t.k ? '0 0 0 1px var(--border)' : 'none',
              }}>{t.label}</button>
            ))}
            <div style={{ flex:1 }}/>
            {/* Filtres Zone / Super / Propriétaire */}
            {zoneList.length > 0 && (
              <select value={filterZone} onChange={e => {
                const newZone = e.target.value;
                setFilterZone(newZone);
                setFilterCity('');
                setPage(1);
                // Auto-sélectionner le super si la zone n'en a qu'un seul
                if (newZone && monthData) {
                  const nrm2 = s => String(s||'').trim().toLowerCase().replace(/\s+/g,' ');
                  const supers = [...new Set(Object.values(monthData.betshop)
                    .filter(b => { const ri = recapLookup[nrm2(b.name)]||{}; return ri.zone === newZone; })
                    .map(b => { const ri = recapLookup[nrm2(b.name)]||{}; return ri.super||''; })
                    .filter(Boolean))];
                  setFilterSuper(supers.length === 1 ? supers[0] : '');
                } else {
                  setFilterSuper('');
                }
              }}
                className="field-input" style={{ fontSize:11, minWidth:100 }}>
                <option value="">🌍 Zone</option>
                {zoneList.map(z => <option key={z} value={z}>{z}</option>)}
              </select>
            )}
            {superList.length > 0 && (
              <select value={filterSuper} onChange={e => { setFilterSuper(e.target.value); setFilterZone(''); setFilterProp(''); setPage(1); }}
                className="field-input" style={{ fontSize:11, minWidth:120 }}>
                <option value="">👤 Superviseur</option>
                {superList.filter(s => !filterZone || (store?.[country]?.[month]?.betshop && Object.values(store[country][month].betshop).some(b=>b.zone===filterZone&&b.super===s))).map(s => <option key={s} value={s}>{s}</option>)}
              </select>
            )}
            {propList.length > 0 && (
              <select value={filterProp} onChange={e => { setFilterProp(e.target.value); setFilterCity(''); setPage(1); }}
                className="field-input" style={{ fontSize:11, minWidth:120 }}>
                <option value="">🏠 Propriétaire</option>
                {propList.map(p => <option key={p} value={p}>{p}</option>)}
              </select>
            )}
            {/* Bouton chargement liste salles pour cityMap */}
            <input ref={cityRef} type="file" accept=".xlsx" hidden
              onChange={e => { if(e.target.files[0]) { loadCityMapLocal(e.target.files[0]); } e.target.value=''; }}/>
            {cityList.length === 0 && Object.keys(cityLookup).length === 0 && (
              <button onClick={() => cityRef.current.click()} title="Charger la liste des salles pour activer le filtre Ville"
                style={{ fontSize:10, padding:'4px 8px', borderRadius:5, border:'1px dashed var(--border)',
                  background:'transparent', color:'var(--muted)', cursor:'pointer', whiteSpace:'nowrap' }}>
                🏙️ + Liste salles
              </button>
            )}
            {cityList.length > 0 && (
              <select value={filterCity} onChange={e => { setFilterCity(e.target.value); setPage(1); }}
                className="field-input" style={{ fontSize:11, minWidth:110 }}>
                <option value="">🏙️ Ville</option>
                {cityList.map(ci => <option key={ci} value={ci}>{ci}</option>)}
              </select>
            )}
            {(filterZone||filterSuper||filterProp||filterCity) && (
              <button onClick={() => { setFilterZone(''); setFilterSuper(''); setFilterProp(''); setFilterCity(''); setPage(1); }} style={{
                fontSize:11, padding:'4px 8px', borderRadius:5, border:'none', cursor:'pointer',
                background:'var(--danger, #F87171)', color:'#fff', fontWeight:700,
              }}>✕ Reset</button>
            )}
            {(filterZone||filterSuper||filterProp||filterCity) && (() => {
              // Export tableau filtré
              const exportFilteredXLSX = () => {
                const cols = ['name','market','zone','super','prop','betshop','shifts','payin_count','payin_money','paid_out','balance','dep_eur','wd_eur'];
                const labels = ['Nom','Marché','Zone','Superviseur','Propriétaire','Betshop','Shifts','Tickets','Payin XAF','Paid out','Balance','Dépôts EUR','Retraits EUR'];
                const dataRows = filtered.map(r => {
                  const obj = {};
                  cols.forEach((k,i) => { obj[labels[i]] = r[k] ?? ''; });
                  return obj;
                });
                const ws = XLSX.utils.json_to_sheet(dataRows);
                const wb = XLSX.utils.book_new();
                XLSX.utils.book_append_sheet(wb, ws, 'Filtré');
                const tag = [filterZone,filterSuper,filterProp,filterCity].filter(Boolean).join('_');
                XLSX.writeFile(wb, `Retail_${subTab}_${tag}.xlsx`);
              };
              const exportFilteredPDF = () => {
                const tag = [filterZone,filterSuper,filterProp,filterCity].filter(Boolean).join(' · ');
                const doc = new jsPDF({ orientation:'landscape', unit:'mm', format:'a4' });
                doc.setFontSize(12);
                doc.text(`Retail — ${subTab === 'betshop' ? 'Betshops' : 'Caissiers'} · ${tag}`, 14, 14);
                doc.setFontSize(9);
                doc.text(`${filtered.length} lignes · Exporté le ${new Date().toLocaleDateString('fr-FR')}`, 14, 20);
                const head = subTab === 'betshop'
                  ? [['Betshop','Marché','Zone','Super','Prop','Shifts','Tickets','Payin XAF','Balance']]
                  : [['Caissier','Marché','Zone','Super','Prop','Betshop','Shifts','Tickets','Payin XAF','Balance']];
                const body = filtered.map(r => subTab === 'betshop'
                  ? [r.name,r.market,r.zone||'—',r.super||'—',r.prop||'—',r.shifts,r.payin_count.toLocaleString('fr-FR'),fmt(r.payin_money),fmt(r.balance)]
                  : [r.name,r.market,r.zone||'—',r.super||'—',r.prop||'—',r.betshop||'—',r.shifts,r.payin_count.toLocaleString('fr-FR'),fmt(r.payin_money),fmt(r.balance)]
                );
                autoTable(doc, { head, body, startY:26, styles:{ fontSize:7 }, headStyles:{ fillColor:[79,70,229] } });
                doc.save(`Retail_${subTab}_${tag.replace(/ · /g,'_')}.pdf`);
              };
              return (
                <>
                  <button onClick={exportFilteredXLSX} style={{
                    fontSize:11, padding:'4px 10px', borderRadius:6, border:'none', cursor:'pointer',
                    background:'#16A34A', color:'#fff', fontWeight:700,
                  }}>📥 Excel</button>
                  <button onClick={exportFilteredPDF} style={{
                    fontSize:11, padding:'4px 10px', borderRadius:6, border:'none', cursor:'pointer',
                    background:'#DC2626', color:'#fff', fontWeight:700,
                  }}>📄 PDF</button>
                </>
              );
            })()}
            <button onClick={() => { setDraftPerfCols(visiblePerfCols); setShowPerfColCfg(true); }} style={{
              fontSize:11, padding:'5px 11px', borderRadius:6, border:'1px solid var(--border)',
              background:'var(--bg3)', color:'var(--text)', cursor:'pointer', fontWeight:600,
            }}>⚙️ Colonnes</button>
            <input value={search} onChange={e => { setSearch(e.target.value); setPage(1); }}
              placeholder="Rechercher..." className="field-input" style={{ width:180, fontSize:12 }}/>
            <span style={{ fontSize:12, color:'var(--muted)', alignSelf:'center', whiteSpace:'nowrap' }}>
              {filtered.length.toLocaleString()} lignes
            </span>
          </div>

          {/* ── Modal Colonnes Performance ── */}
          {showPerfColCfg && (() => {
            const PERF_COLS = [
              { k:'shifts',      label:'Shifts',                  desc:'Nombre de shifts travaillés' },
              { k:'payin_count', label:'Tickets',                 desc:'Nombre de tickets (Payin count)' },
              { k:'payin_money', label:'Payin (XAF)',             desc:'Montant total encaissé' },
              { k:'paid_out',    label:'Paid out',                desc:'Montant total payé' },
              { k:'balance',     label:'Balance',                 desc:'Payin − Paid out' },
              { k:'dep_eur',     label:'Dépôts EUR',              desc:'Total online deposits EUR' },
              { k:'wd_eur',      label:'Retraits EUR',            desc:'Total online withdraws EUR' },
              { k:'dep_local',   label:'Dépôts (local)',          desc:'Total online deposits (monnaie locale)' },
              { k:'wd_local',    label:'Retraits (local)',        desc:'Total online withdraws (monnaie locale)' },
              { k:'kill_count',  label:'Kill count',              desc:'Nombre de kills' },
              { k:'kill_money',  label:'Kill money',              desc:'Montant total des kills' },
              { k:'revoke_count',label:'Revoke count',            desc:'Nombre de révocations' },
              { k:'revoke_money',    label:'Revoke money',        desc:'Montant total des révocations (XAF)' },
              { k:'revoke_money_eur',label:'Revoke money EUR',    desc:'Montant total des révocations (EUR)' },
              { k:'kill_money_eur',  label:'Kill money EUR',      desc:'Montant total des kills (EUR)' },
              { k:'paid_out_count',  label:'Paid out count',      desc:'Nombre de paiements effectués' },
              { k:'payin_money_eur', label:'Payin EUR',           desc:'Montant encaissé en EUR' },
              { k:'paid_out_eur',    label:'Paid out EUR',        desc:'Montant payé en EUR' },
              { k:'winner_count',    label:'Winner count',        desc:'Nombre de tickets gagnants' },
              { k:'winner_money',    label:'Winner money',        desc:'Montant total des gains (XAF)' },
              { k:'winner_money_eur',label:'Winner money EUR',    desc:'Montant total des gains (EUR)' },
              { k:'cash_dep',        label:'Cash dépôts',         desc:'Total cash deposit (XAF)' },
              { k:'cash_dep_eur',    label:'Cash dépôts EUR',     desc:'Total cash deposit (EUR)' },
              { k:'cash_wd',         label:'Cash retraits',       desc:'Total cash withdraws (XAF)' },
              { k:'cash_wd_eur',     label:'Cash retraits EUR',   desc:'Total cash withdraws (EUR)' },
              { k:'payin_tax',       label:'Payin tax',           desc:'Taxe sur payin (XAF)' },
              { k:'payin_tax_eur',   label:'Payin tax EUR',       desc:'Taxe sur payin (EUR)' },
              { k:'zone',  label:'Zone',         desc:'Zone géographique (depuis RECAP)' },
              { k:'super', label:'Superviseur',  desc:'Superviseur en charge (depuis RECAP)' },
              { k:'prop',  label:'Propriétaire', desc:'Propriétaire de la salle (depuis Liste)' },
              ...(subTab==='betshop' ? [{ k:'cashierCount', label:'Nb Caissiers', desc:'Nombre de caissiers actifs' }] : []),
              ...(subTab==='cashier' ? [{ k:'betshop',      label:'Betshop',      desc:"Salle d'affectation" }] : []),
            ];
            return (
              <div style={{
                position:'fixed', inset:0, background:'rgba(0,0,0,0.45)', zIndex:1000,
                display:'flex', alignItems:'center', justifyContent:'center',
              }}>
                <div style={{
                  background:'var(--bg2)', borderRadius:14, padding:'24px', width:380,
                  boxShadow:'0 20px 60px rgba(0,0,0,0.4)', maxHeight:'80vh', display:'flex', flexDirection:'column',
                }}>
                  <div style={{ display:'flex', justifyContent:'space-between', alignItems:'center', marginBottom:6 }}>
                    <span style={{ fontSize:13, fontWeight:700, color:'var(--text)', textTransform:'uppercase', letterSpacing:'.06em' }}>
                      Configuration des colonnes
                    </span>
                    <button onClick={() => setShowPerfColCfg(false)} style={{
                      background:'none', border:'none', fontSize:18, cursor:'pointer', color:'var(--muted)',
                    }}>✕</button>
                  </div>
                  <p style={{ fontSize:11, color:'var(--muted)', marginBottom:14 }}>
                    Sélectionnez les colonnes à afficher dans le tableau.
                  </p>
                  <div style={{ overflowY:'auto', flex:1 }}>
                    {PERF_COLS.map(col => (
                      <div key={col.k}
                        onClick={() => setDraftPerfCols(prev =>
                          prev.includes(col.k) ? prev.filter(x => x !== col.k) : [...prev, col.k]
                        )}
                        style={{
                          display:'flex', alignItems:'center', gap:12, padding:'10px 8px',
                          borderBottom:'1px solid var(--border)', cursor:'pointer',
                          background: draftPerfCols.includes(col.k) ? 'rgba(99,102,241,.06)' : 'transparent',
                          borderRadius:6, marginBottom:2,
                        }}
                      >
                        <div style={{
                          width:22, height:22, borderRadius:6, flexShrink:0,
                          border: draftPerfCols.includes(col.k) ? 'none' : '2px solid var(--border)',
                          background: draftPerfCols.includes(col.k) ? '#16A34A' : 'transparent',
                          display:'flex', alignItems:'center', justifyContent:'center',
                        }}>
                          {draftPerfCols.includes(col.k) && <span style={{ color:'#fff', fontSize:13, fontWeight:700 }}>✓</span>}
                        </div>
                        <div>
                          <div style={{ fontSize:13, fontWeight:600, color:'var(--text)' }}>{col.label}</div>
                          <div style={{ fontSize:10, color:'var(--muted)' }}>{col.desc}</div>
                        </div>
                      </div>
                    ))}
                  </div>
                  <div style={{ display:'flex', gap:10, marginTop:16, justifyContent:'space-between', alignItems:'center' }}>
                    <button onClick={() => setDraftPerfCols(['shifts','payin_count','payin_money','paid_out','balance','dep_eur','wd_eur'])} style={{
                      fontSize:11, padding:'6px 12px', borderRadius:6, border:'1px solid var(--border)',
                      background:'var(--bg3)', color:'var(--muted)', cursor:'pointer',
                    }}>Réinitialiser</button>
                    <div style={{ display:'flex', gap:8 }}>
                      <button onClick={() => { setDraftPerfCols(visiblePerfCols); setShowPerfColCfg(false); }} style={{
                        padding:'7px 16px', borderRadius:8, border:'1px solid var(--border)',
                        background:'var(--bg3)', color:'var(--muted)', cursor:'pointer', fontSize:12, fontWeight:600,
                      }}>Annuler</button>
                      <button onClick={() => { setVisiblePerfCols(draftPerfCols); setShowPerfColCfg(false); }} style={{
                        padding:'7px 16px', borderRadius:8, border:'none',
                        background:'#16A34A', color:'#fff', cursor:'pointer', fontSize:12, fontWeight:700,
                      }}>Enregistrer</button>
                    </div>
                  </div>
                </div>
              </div>
            );
          })()}

          <div style={{ background:'var(--bg2)', border:'1px solid var(--border)', borderRadius:10, overflow:'hidden' }}>
            <div style={{ overflowX:'auto', overflowY:'auto', maxHeight:520 }}>
              <table style={{ width:'100%', borderCollapse:'collapse', fontSize:12 }}>
                <thead>
                  <tr style={{ borderBottom:'1px solid var(--border)' }}>
                    <Th label={subTab==='betshop'?'Betshop':'Caissier'} k="name" sort={sort} onSort={onSort}/>
                    <Th label="Marché" k="market" sort={sort} onSort={onSort}/>
                    {visiblePerfCols.includes('zone')  && <Th label="Zone"          k="zone"  sort={sort} onSort={onSort}/>}
                    {visiblePerfCols.includes('super') && <Th label="Superviseur"    k="super" sort={sort} onSort={onSort}/>}
                    {visiblePerfCols.includes('prop')  && <Th label="Propriétaire"   k="prop"  sort={sort} onSort={onSort}/>}
                    {subTab==='betshop' && visiblePerfCols.includes('cashierCount') && <Th label="Caissiers" k="cashierCount" sort={sort} onSort={onSort} right/>}
                    {subTab==='cashier' && visiblePerfCols.includes('betshop')      && <Th label="Betshop"   k="betshop"     sort={sort} onSort={onSort}/>}
                    {visiblePerfCols.includes('shifts')      && <Th label="Shifts"       k="shifts"      sort={sort} onSort={onSort} right/>}
                    {visiblePerfCols.includes('payin_count') && <Th label="Tickets"      k="payin_count" sort={sort} onSort={onSort} right/>}
                    {visiblePerfCols.includes('payin_money') && <Th label="Payin"        k="payin_money" sort={sort} onSort={onSort} right/>}
                    {visiblePerfCols.includes('paid_out')    && <Th label="Paid out"     k="paid_out"    sort={sort} onSort={onSort} right/>}
                    {visiblePerfCols.includes('balance')     && <Th label="Balance"      k="balance"     sort={sort} onSort={onSort} right/>}
                    {visiblePerfCols.includes('dep_eur')     && <Th label="Dépôts EUR"   k="dep_eur"     sort={sort} onSort={onSort} right/>}
                    {visiblePerfCols.includes('wd_eur')      && <Th label="Retraits EUR" k="wd_eur"      sort={sort} onSort={onSort} right/>}
                    {visiblePerfCols.includes('dep_local')   && <Th label="Dépôts local" k="dep_local"   sort={sort} onSort={onSort} right/>}
                    {visiblePerfCols.includes('wd_local')    && <Th label="Ret. local"    k="wd_local"     sort={sort} onSort={onSort} right/>}
                    {visiblePerfCols.includes('kill_count')  && <Th label="Kill count"   k="kill_count"   sort={sort} onSort={onSort} right/>}
                    {visiblePerfCols.includes('kill_money')  && <Th label="Kill money"   k="kill_money"   sort={sort} onSort={onSort} right/>}
                    {visiblePerfCols.includes('revoke_count')&& <Th label="Revoke count" k="revoke_count" sort={sort} onSort={onSort} right/>}
                    {visiblePerfCols.includes('revoke_money')    && <Th label="Revoke money"    k="revoke_money"     sort={sort} onSort={onSort} right/>}
                    {visiblePerfCols.includes('revoke_money_eur') && <Th label="Revoke EUR"     k="revoke_money_eur" sort={sort} onSort={onSort} right/>}
                    {visiblePerfCols.includes('kill_money_eur')   && <Th label="Kill EUR"        k="kill_money_eur"   sort={sort} onSort={onSort} right/>}
                    {visiblePerfCols.includes('paid_out_count')   && <Th label="Paid out cnt"   k="paid_out_count"   sort={sort} onSort={onSort} right/>}
                    {visiblePerfCols.includes('payin_money_eur')  && <Th label="Payin EUR"       k="payin_money_eur"  sort={sort} onSort={onSort} right/>}
                    {visiblePerfCols.includes('paid_out_eur')     && <Th label="Paid out EUR"    k="paid_out_eur"     sort={sort} onSort={onSort} right/>}
                    {visiblePerfCols.includes('winner_count')     && <Th label="Winner cnt"      k="winner_count"     sort={sort} onSort={onSort} right/>}
                    {visiblePerfCols.includes('winner_money')     && <Th label="Winner money"    k="winner_money"     sort={sort} onSort={onSort} right/>}
                    {visiblePerfCols.includes('winner_money_eur') && <Th label="Winner EUR"      k="winner_money_eur" sort={sort} onSort={onSort} right/>}
                    {visiblePerfCols.includes('cash_dep')         && <Th label="Cash dep"        k="cash_dep"         sort={sort} onSort={onSort} right/>}
                    {visiblePerfCols.includes('cash_dep_eur')     && <Th label="Cash dep EUR"    k="cash_dep_eur"     sort={sort} onSort={onSort} right/>}
                    {visiblePerfCols.includes('cash_wd')          && <Th label="Cash ret"        k="cash_wd"          sort={sort} onSort={onSort} right/>}
                    {visiblePerfCols.includes('cash_wd_eur')      && <Th label="Cash ret EUR"    k="cash_wd_eur"      sort={sort} onSort={onSort} right/>}
                    {visiblePerfCols.includes('payin_tax')        && <Th label="Payin tax"       k="payin_tax"        sort={sort} onSort={onSort} right/>}
                    {visiblePerfCols.includes('payin_tax_eur')    && <Th label="Payin tax EUR"   k="payin_tax_eur"    sort={sort} onSort={onSort} right/>}
                  </tr>
                </thead>
                <tbody>
                  {pageRows.map((r, i) => {
                    const isSelected = selectedRow?.data?.name === r.name;
                    return (
                    <tr key={i}
                      onClick={() => setSelectedRow(isSelected ? null : { type: subTab, data: r })}
                      style={{ borderBottom:'0.5px solid var(--border)', cursor:'pointer',
                        background: isSelected
                          ? 'rgba(99,102,241,.12)'
                          : i%2===0?'transparent':'rgba(255,255,255,.02)',
                        outline: isSelected ? '1px solid var(--accent)' : 'none',
                      }}>
                      <td style={{ padding:'6px 10px', fontWeight:600, fontFamily:'monospace', fontSize:11 }}>{r.name}</td>
                      <td style={{ padding:'6px 10px', fontSize:11, color:'var(--muted)' }}>{r.market}</td>
                      {visiblePerfCols.includes('zone')  && <td style={{ padding:'6px 10px', fontSize:11, color:'#60A5FA' }}>{r.zone  || '—'}</td>}
                      {visiblePerfCols.includes('super') && <td style={{ padding:'6px 10px', fontSize:11, color:'#A78BFA' }}>{r.super || '—'}</td>}
                      {visiblePerfCols.includes('prop')  && <td style={{ padding:'6px 10px', fontSize:11, color:'#34D399' }}>{r.prop  || '—'}</td>}
                      {subTab==='betshop' && visiblePerfCols.includes('cashierCount') && <td style={{ padding:'6px 10px', textAlign:'right' }}>{r.cashierCount}</td>}
                      {subTab==='cashier' && visiblePerfCols.includes('betshop')      && <td style={{ padding:'6px 10px', fontSize:11, color:'var(--muted)' }}>{r.betshop}</td>}
                      {visiblePerfCols.includes('shifts')      && <td style={{ padding:'6px 10px', textAlign:'right', color:'var(--muted)' }}>{r.shifts}</td>}
                      {visiblePerfCols.includes('payin_count') && <td style={{ padding:'6px 10px', textAlign:'right', fontWeight:600, color:'#22C55E' }}>{r.payin_count.toLocaleString('fr-FR')}</td>}
                      {visiblePerfCols.includes('payin_money') && <td style={{ padding:'6px 10px', textAlign:'right', fontFamily:'monospace', fontSize:11 }}>{fmt(r.payin_money)}</td>}
                      {visiblePerfCols.includes('paid_out')    && <td style={{ padding:'6px 10px', textAlign:'right', fontFamily:'monospace', fontSize:11, color:'var(--muted)' }}>{fmt(r.paid_out)}</td>}
                      {visiblePerfCols.includes('balance')     && <td style={{ padding:'6px 10px', textAlign:'right', fontFamily:'monospace', fontSize:11, color:(r.balance||0)>=0?'var(--deposit)':'var(--danger)' }}>{fmt(r.balance)}</td>}
                      {visiblePerfCols.includes('dep_eur')     && <td style={{ padding:'6px 10px', textAlign:'right', fontFamily:'monospace', fontSize:11, color:'#34D399' }}>{fmtEUR(r.dep_eur||0)}</td>}
                      {visiblePerfCols.includes('wd_eur')      && <td style={{ padding:'6px 10px', textAlign:'right', fontFamily:'monospace', fontSize:11, color:'#F87171' }}>{fmtEUR(r.wd_eur||0)}</td>}
                      {visiblePerfCols.includes('dep_local')   && <td style={{ padding:'6px 10px', textAlign:'right', fontFamily:'monospace', fontSize:11 }}>{fmt(r.dep_local||0)}</td>}
                      {visiblePerfCols.includes('wd_local')     && <td style={{ padding:'6px 10px', textAlign:'right', fontFamily:'monospace', fontSize:11 }}>{fmt(r.wd_local||0)}</td>}
                      {visiblePerfCols.includes('kill_count')  && <td style={{ padding:'6px 10px', textAlign:'right', fontFamily:'monospace', fontSize:11, color:'#F87171' }}>{(r.kill_count||0).toLocaleString('fr-FR')}</td>}
                      {visiblePerfCols.includes('kill_money')  && <td style={{ padding:'6px 10px', textAlign:'right', fontFamily:'monospace', fontSize:11, color:'#F87171' }}>{fmt(r.kill_money||0)}</td>}
                      {visiblePerfCols.includes('revoke_count')&& <td style={{ padding:'6px 10px', textAlign:'right', fontFamily:'monospace', fontSize:11, color:'#FBBF24' }}>{(r.revoke_count||0).toLocaleString('fr-FR')}</td>}
                      {visiblePerfCols.includes('revoke_money')     && <td style={{ padding:'6px 10px', textAlign:'right', fontFamily:'monospace', fontSize:11, color:'#FBBF24' }}>{fmt(r.revoke_money||0)}</td>}
                      {visiblePerfCols.includes('revoke_money_eur') && <td style={{ padding:'6px 10px', textAlign:'right', fontFamily:'monospace', fontSize:11, color:'#FBBF24' }}>{fmtEUR(r.revoke_money_eur||0)}</td>}
                      {visiblePerfCols.includes('kill_money_eur')   && <td style={{ padding:'6px 10px', textAlign:'right', fontFamily:'monospace', fontSize:11, color:'#F87171' }}>{fmtEUR(r.kill_money_eur||0)}</td>}
                      {visiblePerfCols.includes('paid_out_count')   && <td style={{ padding:'6px 10px', textAlign:'right', fontFamily:'monospace', fontSize:11 }}>{(r.paid_out_count||0).toLocaleString('fr-FR')}</td>}
                      {visiblePerfCols.includes('payin_money_eur')  && <td style={{ padding:'6px 10px', textAlign:'right', fontFamily:'monospace', fontSize:11 }}>{fmtEUR(r.payin_money_eur||0)}</td>}
                      {visiblePerfCols.includes('paid_out_eur')     && <td style={{ padding:'6px 10px', textAlign:'right', fontFamily:'monospace', fontSize:11 }}>{fmtEUR(r.paid_out_eur||0)}</td>}
                      {visiblePerfCols.includes('winner_count')     && <td style={{ padding:'6px 10px', textAlign:'right', fontFamily:'monospace', fontSize:11, color:'#34D399' }}>{(r.winner_count||0).toLocaleString('fr-FR')}</td>}
                      {visiblePerfCols.includes('winner_money')     && <td style={{ padding:'6px 10px', textAlign:'right', fontFamily:'monospace', fontSize:11, color:'#34D399' }}>{fmt(r.winner_money||0)}</td>}
                      {visiblePerfCols.includes('winner_money_eur') && <td style={{ padding:'6px 10px', textAlign:'right', fontFamily:'monospace', fontSize:11, color:'#34D399' }}>{fmtEUR(r.winner_money_eur||0)}</td>}
                      {visiblePerfCols.includes('cash_dep')         && <td style={{ padding:'6px 10px', textAlign:'right', fontFamily:'monospace', fontSize:11 }}>{fmt(r.cash_dep||0)}</td>}
                      {visiblePerfCols.includes('cash_dep_eur')     && <td style={{ padding:'6px 10px', textAlign:'right', fontFamily:'monospace', fontSize:11 }}>{fmtEUR(r.cash_dep_eur||0)}</td>}
                      {visiblePerfCols.includes('cash_wd')          && <td style={{ padding:'6px 10px', textAlign:'right', fontFamily:'monospace', fontSize:11 }}>{fmt(r.cash_wd||0)}</td>}
                      {visiblePerfCols.includes('cash_wd_eur')      && <td style={{ padding:'6px 10px', textAlign:'right', fontFamily:'monospace', fontSize:11 }}>{fmtEUR(r.cash_wd_eur||0)}</td>}
                      {visiblePerfCols.includes('payin_tax')        && <td style={{ padding:'6px 10px', textAlign:'right', fontFamily:'monospace', fontSize:11 }}>{fmt(r.payin_tax||0)}</td>}
                      {visiblePerfCols.includes('payin_tax_eur')    && <td style={{ padding:'6px 10px', textAlign:'right', fontFamily:'monospace', fontSize:11 }}>{fmtEUR(r.payin_tax_eur||0)}</td>}
                    </tr>
                  );})}
                </tbody>
              </table>
            </div>
            <Pagination page={page} total={filtered.length} totalPages={totalPages} onPage={setPage}/>
          </div>
          {selectedRow && (
            <DrillDown
              row={selectedRow.data}
              type={selectedRow.type}
              onClose={() => setSelectedRow(null)}
            />
          )}
        </>
      )}

      {/* ── Deposits tab ── */}
      {mainTab === 'deposits' && (
        <div style={{ background:'var(--bg2)', border:'1px solid var(--border)', borderRadius:10, overflow:'hidden' }}>
          <div style={{ overflowX:'auto', overflowY:'auto', maxHeight:560 }}>
            <table style={{ width:'100%', borderCollapse:'collapse', fontSize:12 }}>
              <thead>
                <tr style={{ borderBottom:'1px solid var(--border)' }}>
                  <Th label="Marché"        k="name"        sort={mktSort} onSort={onMktSort}/>
                  <Th label="Shops"         k="shopCount"   sort={mktSort} onSort={onMktSort} right/>
                  <Th label="Tickets"       k="payin_count" sort={mktSort} onSort={onMktSort} right/>
                  <Th label="Dépôts EUR"    k="dep_eur"     sort={mktSort} onSort={onMktSort} right/>
                  <Th label="Retraits EUR"  k="wd_eur"      sort={mktSort} onSort={onMktSort} right/>
                  <Th label="Net EUR"       k="net_eur"     sort={mktSort} onSort={onMktSort} right/>
                </tr>
              </thead>
              <tbody>
                {mktRows.map((r, i) => {
                  const net = r.dep_eur - r.wd_eur;
                  const isMktSelected = selectedMkt?.name === r.name;
                  return (
                    <tr key={i}
                      onClick={() => setSelectedMkt(isMktSelected ? null : r)}
                      style={{ borderBottom:'0.5px solid var(--border)', cursor:'pointer',
                        background: isMktSelected
                          ? 'rgba(52,211,153,.1)'
                          : i%2===0?'transparent':'rgba(255,255,255,.02)',
                        outline: isMktSelected ? '1px solid #34D399' : 'none',
                      }}>
                      <td style={{ padding:'7px 10px', fontWeight:600 }}>{r.name}</td>
                      <td style={{ padding:'7px 10px', textAlign:'right', color:'var(--muted)' }}>{r.shopCount}</td>
                      <td style={{ padding:'7px 10px', textAlign:'right', color:'#22C55E', fontWeight:600 }}>{r.payin_count.toLocaleString('fr-FR')}</td>
                      <td style={{ padding:'7px 10px', textAlign:'right', color:'#34D399', fontWeight:600 }}>{fmtEUR(r.dep_eur)}</td>
                      <td style={{ padding:'7px 10px', textAlign:'right', color:'#F87171' }}>{fmtEUR(r.wd_eur)}</td>
                      <td style={{ padding:'7px 10px', textAlign:'right', fontWeight:600,
                        color: net >= 0 ? 'var(--deposit)' : 'var(--danger)' }}>
                        {net >= 0 ? '+' : ''}{fmtEUR(net)}
                      </td>
                    </tr>
                  );
                })}
                {/* Total row */}
                {mktRows.length > 0 && (() => {
                  const tot = mktRows.reduce((acc,r) => ({
                    dep:acc.dep+r.dep_eur, wd:acc.wd+r.wd_eur, shops:acc.shops+r.shopCount,
                    tickets:acc.tickets+r.payin_count
                  }), {dep:0,wd:0,shops:0,tickets:0});
                  const net = tot.dep - tot.wd;
                  return (
                    <tr style={{ borderTop:'2px solid var(--border)', background:'var(--bg3)', fontWeight:700 }}>
                      <td style={{ padding:'7px 10px' }}>TOTAL</td>
                      <td style={{ padding:'7px 10px', textAlign:'right' }}>{tot.shops}</td>
                      <td style={{ padding:'7px 10px', textAlign:'right', color:'#22C55E' }}>{tot.tickets.toLocaleString('fr-FR')}</td>
                      <td style={{ padding:'7px 10px', textAlign:'right', color:'#34D399' }}>{fmtEUR(tot.dep)}</td>
                      <td style={{ padding:'7px 10px', textAlign:'right', color:'#F87171' }}>{fmtEUR(tot.wd)}</td>
                      <td style={{ padding:'7px 10px', textAlign:'right', color: net>=0?'var(--deposit)':'var(--danger)' }}>
                        {net>=0?'+':''}{fmtEUR(net)}
                      </td>
                    </tr>
                  );
                })()}
              </tbody>
            </table>
          </div>
        {selectedMkt && mainTab === 'deposits' && (
          <DrillDownMarket
            row={selectedMkt}
            onClose={() => setSelectedMkt(null)}
          />
        )}
        </div>
      )}

      {/* ── Compare tab ── */}
      {mainTab === 'compare' && (() => {
        // Déterminer les monthKeys de chaque période
        let keysA = [], keysB = [];
        if (cmpMode === '2months') {
          keysA = cmpMonthA ? [cmpMonthA] : (months[1] ? [months[1]] : []);
          keysB = cmpMonthB ? [cmpMonthB] : (months[0] ? [months[0]] : []);
        } else {
          // Plage : tous les mois de cmpFrom à cmpTo inclus
          const sorted = [...months].sort();
          keysA = sorted.filter(m => m >= cmpFrom && m <= cmpTo);
          // Plage B : même durée immédiatement avant
          if (keysA.length && cmpFrom) {
            const dur = keysA.length;
            const all = [...months].sort();
            const startIdx = all.indexOf(keysA[0]);
            keysB = all.slice(Math.max(0, startIdx - dur), startIdx);
          }
        }
        const kA = kpisForMonths(keysA);
        const kB = kpisForMonths(keysB);

        const rows = [
          { label:'🏪 Betshops actifs',   ka: kA?.betshops, kb: kB?.betshops, fmt: v => v?.toLocaleString('fr-FR') ?? '—', higherIsBetter: true },
          { label:'👤 Caissiers actifs',  ka: kA?.cashiers, kb: kB?.cashiers, fmt: v => v?.toLocaleString('fr-FR') ?? '—', higherIsBetter: true },
          { label:'🔄 Shifts',            ka: kA?.shifts,   kb: kB?.shifts,   fmt: v => v?.toLocaleString('fr-FR') ?? '—', higherIsBetter: true },
          { label:'🎫 Tickets',           ka: kA?.tickets,  kb: kB?.tickets,  fmt: v => fmt(v), higherIsBetter: true },
          { label:'💰 Payin (XAF)',       ka: kA?.payin,    kb: kB?.payin,    fmt: v => fmt(v), higherIsBetter: true },
          { label:'💶 Dépôts EUR',        ka: kA?.dep,      kb: kB?.dep,      fmt: v => fmtEUR(v ?? 0), higherIsBetter: true },
          { label:'💸 Retraits EUR',      ka: kA?.wd,       kb: kB?.wd,       fmt: v => fmtEUR(v ?? 0), higherIsBetter: false },
        ];

        const maxVal = (row) => Math.max(row.ka ?? 0, row.kb ?? 0, 1);

        return (
          <div>
            {/* Sélecteur de mode */}
            <div style={{ display:'flex', gap:8, marginBottom:16, alignItems:'center', flexWrap:'wrap' }}>
              <div style={{ display:'flex', gap:4 }}>
                {[{ k:'2months', l:'2 Mois' }, { k:'range', l:'Plage → Plage' }].map(m => (
                  <button key={m.k} onClick={() => setCmpMode(m.k)} style={{
                    fontSize:12, padding:'6px 12px', borderRadius:6, border:'none', cursor:'pointer', fontWeight:600,
                    background: cmpMode===m.k ? 'var(--accent)' : 'var(--bg3)',
                    color: cmpMode===m.k ? '#fff' : 'var(--muted)',
                  }}>{m.l}</button>
                ))}
              </div>

              {cmpMode === '2months' ? (
                <>
                  <select value={cmpMonthA || months[1] || ''} onChange={e => setCmpMonthA(e.target.value)}
                    className="field-input" style={{ minWidth:110, fontSize:12 }}>
                    {months.map(m => <option key={m} value={m}>{m}</option>)}
                  </select>
                  <span style={{ color:'var(--muted)', fontSize:13 }}>vs</span>
                  <select value={cmpMonthB || months[0] || ''} onChange={e => setCmpMonthB(e.target.value)}
                    className="field-input" style={{ minWidth:110, fontSize:12 }}>
                    {months.map(m => <option key={m} value={m}>{m}</option>)}
                  </select>
                </>
              ) : (
                <>
                  <span style={{ fontSize:12, color:'var(--muted)' }}>Période A :</span>
                  <select value={cmpFrom} onChange={e => setCmpFrom(e.target.value)}
                    className="field-input" style={{ minWidth:110, fontSize:12 }}>
                    <option value="">Début</option>
                    {[...months].sort().map(m => <option key={m} value={m}>{m}</option>)}
                  </select>
                  <span style={{ fontSize:12, color:'var(--muted)' }}>→</span>
                  <select value={cmpTo} onChange={e => setCmpTo(e.target.value)}
                    className="field-input" style={{ minWidth:110, fontSize:12 }}>
                    <option value="">Fin</option>
                    {[...months].sort().map(m => <option key={m} value={m}>{m}</option>)}
                  </select>
                  {keysA.length > 0 && <span style={{ fontSize:11, color:'var(--muted)' }}>({keysA.length} mois) — Période B auto : {keysB.length} mois précédents</span>}
                </>
              )}
            </div>

            {/* Libellés des périodes */}
            <div style={{ display:'grid', gridTemplateColumns:'180px 1fr 1fr 90px', gap:1, marginBottom:8 }}>
              {['', keysA.join(', ') || '—', keysB.join(', ') || '—', 'Δ%'].map((h, i) => (
                <div key={i} style={{ padding:'6px 10px', background:'var(--bg3)', fontSize:11, fontWeight:700,
                  color:'var(--muted)', textAlign: i > 0 ? 'right' : 'left', borderRadius: i===0?'8px 0 0 0':i===3?'0 8px 0 0':'0' }}>
                  {h}
                </div>
              ))}
            </div>

            {/* Tableau de comparaison */}
            <div style={{ background:'var(--bg2)', border:'1px solid var(--border)', borderRadius:10, overflow:'hidden' }}>
              {rows.map((row, i) => {
                const delta = row.ka != null && row.kb != null && row.kb !== 0
                  ? ((row.ka - row.kb) / row.kb) * 100 : null;
                const good = delta == null ? null : row.higherIsBetter ? delta >= 0 : delta <= 0;
                const barA = (row.ka ?? 0) / maxVal(row);
                const barB = (row.kb ?? 0) / maxVal(row);
                return (
                  <div key={i} style={{
                    display:'grid', gridTemplateColumns:'180px 1fr 1fr 90px',
                    borderBottom: i < rows.length-1 ? '1px solid var(--border)' : 'none',
                  }}>
                    {/* Label */}
                    <div style={{ padding:'10px 12px', fontSize:12, fontWeight:600, color:'var(--text)',
                      background: i%2===0 ? 'var(--bg2)' : 'var(--bg3)', display:'flex', alignItems:'center' }}>
                      {row.label}
                    </div>
                    {/* Valeur A + barre */}
                    {[{val: row.ka, bar: barA, color:'#60A5FA'}, {val: row.kb, bar: barB, color:'#A78BFA'}].map((side, si) => (
                      <div key={si} style={{ padding:'8px 12px', background: i%2===0 ? 'var(--bg2)' : 'var(--bg3)', textAlign:'right' }}>
                        <div style={{ fontSize:13, fontWeight:700, color: si===0 ? '#60A5FA' : '#A78BFA', marginBottom:4 }}>
                          {row.fmt(side.val)}
                        </div>
                        <div style={{ height:4, background:'var(--border)', borderRadius:2, overflow:'hidden' }}>
                          <div style={{ height:'100%', width:`${side.bar*100}%`, background:side.color, borderRadius:2, transition:'width .3s' }}/>
                        </div>
                      </div>
                    ))}
                    {/* Delta */}
                    <div style={{ padding:'10px 12px', textAlign:'right', background: i%2===0 ? 'var(--bg2)' : 'var(--bg3)',
                      display:'flex', alignItems:'center', justifyContent:'flex-end' }}>
                      {delta != null ? (
                        <span style={{ fontSize:12, fontWeight:700, padding:'2px 7px', borderRadius:5,
                          background: good ? 'var(--deposit-bg, rgba(52,211,153,0.15))' : 'var(--danger-bg, rgba(248,113,113,0.15))',
                          color: good ? '#34D399' : '#F87171' }}>
                          {delta >= 0 ? '+' : ''}{delta.toFixed(1)}%
                        </span>
                      ) : <span style={{ color:'var(--muted)', fontSize:11 }}>—</span>}
                    </div>
                  </div>
                );
              })}
            </div>

            {(!kA && !kB) && (
              <p style={{ color:'var(--muted)', fontSize:13, textAlign:'center', marginTop:20 }}>
                Sélectionnez des périodes pour lancer la comparaison.
              </p>
            )}

            {/* ── Section Diagramme par entité ── */}
            {(kA || kB) && (() => {
              const dA = detailForMonths(keysA, cmpEntity);
              const dB = detailForMonths(keysB, cmpEntity);

              // Union des noms d'entités présents dans A ou B
              const allNames = Array.from(new Set([...Object.keys(dA), ...Object.keys(dB)]));

              // KPI config
              const kpiCfg = {
                tickets: { label:'Tickets',      fmt: v => fmt(v),        colorA:'#60A5FA', colorB:'#93C5FD' },
                payin:   { label:'Payin (XAF)',  fmt: v => fmt(v),        colorA:'#34D399', colorB:'#6EE7B7' },
                dep:     { label:'Dépôts EUR',   fmt: v => fmtEUR(v||0),  colorA:'#A78BFA', colorB:'#C4B5FD' },
                wd:      { label:'Retraits EUR',  fmt: v => fmtEUR(v||0), colorA:'#F87171', colorB:'#FCA5A5' },
              };
              const kCfg = kpiCfg[cmpKpi];

              // Filtrer par recherche puis trier par valeur A desc, prendre les top N
              const filteredNames = cmpSearch.trim()
                ? allNames.filter(n => n.toLowerCase().includes(cmpSearch.trim().toLowerCase()))
                : allNames;
              const sorted = filteredNames
                .map(n => ({ name:n, vA: dA[n]?.[cmpKpi] ?? 0, vB: dB[n]?.[cmpKpi] ?? 0 }))
                .sort((a,b) => b.vA - a.vA)
                .slice(0, cmpTop);

              const maxV = Math.max(...sorted.map(r => Math.max(r.vA, r.vB)), 1);

              // SVG chart — barres groupées horizontales
              const BAR_H = 18;
              const GAP   = 6;
              const LABEL_W = 140;
              const CHART_W = 520;
              const ROW_H   = BAR_H * 2 + GAP + 14;
              const SVG_H   = sorted.length * ROW_H + 30;

              return (
                <div style={{ marginTop:28 }}>
                  {/* Sous-onglets entité + KPI */}
                  <div style={{ display:'flex', gap:8, alignItems:'center', flexWrap:'wrap', marginBottom:14 }}>
                    <span style={{ fontSize:12, color:'var(--muted)', fontWeight:600 }}>Détail par :</span>
                    {[
                      { k:'market',  l:'🌍 Marché' },
                      { k:'betshop', l:'🏪 Betshop' },
                      { k:'cashier', l:'👤 Caissier' },
                    ].map(e => (
                      <button key={e.k} onClick={() => { setCmpEntity(e.k); setCmpSearch(''); }} style={{
                        fontSize:12, padding:'5px 11px', borderRadius:6, border:'none', cursor:'pointer', fontWeight:600,
                        background: cmpEntity===e.k ? 'var(--bg)' : 'var(--bg3)',
                        color: cmpEntity===e.k ? 'var(--text)' : 'var(--muted)',
                        boxShadow: cmpEntity===e.k ? '0 0 0 1px var(--border)' : 'none',
                      }}>{e.l}</button>
                    ))}
                    <div style={{ width:1, height:20, background:'var(--border)', margin:'0 4px' }}/>
                    <span style={{ fontSize:12, color:'var(--muted)', fontWeight:600 }}>KPI :</span>
                    {[
                      { k:'tickets', l:'🎫 Tickets' },
                      { k:'payin',   l:'💰 Payin' },
                      { k:'dep',     l:'💶 Dépôts' },
                      { k:'wd',      l:'💸 Retraits' },
                    ].map(e => (
                      <button key={e.k} onClick={() => setCmpKpi(e.k)} style={{
                        fontSize:12, padding:'5px 11px', borderRadius:6, border:'none', cursor:'pointer', fontWeight:600,
                        background: cmpKpi===e.k ? 'var(--accent)' : 'var(--bg3)',
                        color: cmpKpi===e.k ? '#fff' : 'var(--muted)',
                      }}>{e.l}</button>
                    ))}
                    <div style={{ width:1, height:20, background:'var(--border)', margin:'0 4px' }}/>
                    <span style={{ fontSize:12, color:'var(--muted)' }}>Top</span>
                    {[5,10,20,50].map(n => (
                      <button key={n} onClick={() => setCmpTop(n)} style={{
                        fontSize:12, padding:'4px 9px', borderRadius:5, border:'none', cursor:'pointer',
                        background: cmpTop===n ? 'var(--accent)' : 'var(--bg3)',
                        color: cmpTop===n ? '#fff' : 'var(--muted)', fontWeight:600,
                      }}>{n}</button>
                    ))}
                  </div>

                  {/* Barre de recherche */}
                  <div style={{ display:'flex', alignItems:'center', gap:8, margin:'10px 0 6px' }}>
                    <div style={{ position:'relative', flexGrow:1, maxWidth:320 }}>
                      <span style={{ position:'absolute', left:9, top:'50%', transform:'translateY(-50%)',
                        fontSize:13, color:'var(--muted)', pointerEvents:'none' }}>🔍</span>
                      <input
                        type="text"
                        value={cmpSearch}
                        onChange={e => setCmpSearch(e.target.value)}
                        placeholder={`Rechercher ${cmpEntity === 'market' ? 'un marché' : cmpEntity === 'betshop' ? 'une salle' : 'un caissier'}…`}
                        style={{
                          width:'100%', paddingLeft:30, paddingRight:cmpSearch ? 28 : 10,
                          paddingTop:6, paddingBottom:6, fontSize:12, borderRadius:7,
                          border:'1px solid var(--border)', background:'var(--bg3)',
                          color:'var(--text)', outline:'none', boxSizing:'border-box',
                        }}
                      />
                      {cmpSearch && (
                        <button onClick={() => setCmpSearch('')} style={{
                          position:'absolute', right:6, top:'50%', transform:'translateY(-50%)',
                          background:'none', border:'none', cursor:'pointer',
                          color:'var(--muted)', fontSize:14, lineHeight:1, padding:0,
                        }}>✕</button>
                      )}
                    </div>
                    <span style={{ fontSize:11, color:'var(--muted)', whiteSpace:'nowrap' }}>
                      {cmpSearch
                        ? `${sorted.length} résultat${sorted.length>1?'s':''} sur ${filteredNames.length} / ${allNames.length}`
                        : `${allNames.length} ${cmpEntity === 'market' ? 'marchés' : cmpEntity === 'betshop' ? 'salles' : 'caissiers'}`
                      }
                    </span>
                  </div>

                  {/* Légende */}
                  <div style={{ display:'flex', gap:16, marginBottom:10 }}>
                    {[
                      { color: kCfg.colorA, label: `Période A  (${keysA.join(', ') || '—'})` },
                      { color: kCfg.colorB, label: `Période B  (${keysB.join(', ') || '—'})` },
                    ].map((leg,i) => (
                      <div key={i} style={{ display:'flex', alignItems:'center', gap:6, fontSize:11, color:'var(--muted)' }}>
                        <div style={{ width:14, height:10, background:leg.color, borderRadius:2 }}/>
                        {leg.label}
                      </div>
                    ))}
                  </div>

                  {/* ── VUE RECHERCHE : barres verticales par mois ── */}
                  {cmpSearch.trim() && sorted.length > 0 ? (() => {
                    const allMonths = [...months].sort();
                    const COLORS = ['#818CF8','#60A5FA','#34D399','#F87171','#FBBF24','#A78BFA','#FB7185','#38BDF8'];

                    // Valeurs par entité × mois
                    const series = sorted.map((r, idx) => {
                      const vals = allMonths.map(m => {
                        const d = detailForMonths([m], cmpEntity);
                        return d[r.name]?.[cmpKpi] ?? 0;
                      });
                      return { name: r.name, vals, color: COLORS[idx % COLORS.length] };
                    });

                    const nSeries  = series.length;
                    const nMonths  = allMonths.length;
                    const maxVal2  = Math.max(...series.flatMap(s => s.vals), 1);

                    // Dimensions
                    const M_L = 64, M_R = 24, M_T = 30, M_B = 44;
                    const CW  = Math.max(480, nMonths * (nSeries * 22 + 16));
                    const CH  = 200;
                    const W   = CW + M_L + M_R;
                    const H   = CH + M_T + M_B;

                    const grpW   = CW / nMonths;        // largeur d'un groupe de mois
                    const barW   = Math.min(22, (grpW - 12) / nSeries);
                    const grpOff = (grpW - barW * nSeries) / 2;
                    const bx     = (mi, si) => M_L + mi * grpW + grpOff + si * barW;
                    const by     = v => M_T + CH - (v / maxVal2) * CH;
                    const bh     = v => (v / maxVal2) * CH;

                    const yTicks = [0, 0.25, 0.5, 0.75, 1].map(f => maxVal2 * f);

                    // Export Excel
                    const exportXlsx = () => {
                      const XLSX = window.XLSX;
                      if (!XLSX) { alert('XLSX non chargé'); return; }
                      const header = ['Entité', ...allMonths];
                      const rows2  = series.map(s => [s.name, ...s.vals]);
                      const ws     = XLSX.utils.aoa_to_sheet([header, ...rows2]);
                      const wb     = XLSX.utils.book_new();
                      XLSX.utils.book_append_sheet(wb, ws, 'Comparaison');
                      XLSX.writeFile(wb, `comparaison_${cmpEntity}_${cmpKpi}.xlsx`);
                    };

                    // Export PDF via impression SVG dans une page
                    const exportPdf = () => {
                      const svgEl = document.getElementById('cmp-search-svg');
                      if (!svgEl) return;
                      const svgData = new XMLSerializer().serializeToString(svgEl);
                      const win = window.open('', '_blank');
                      win.document.write(`<!DOCTYPE html><html><head><title>Export</title>
                        <style>body{margin:0;background:#fff;font-family:sans-serif;}
                        h3{padding:12px 16px;margin:0;font-size:14px;}
                        svg{display:block;margin:0 auto;background:#fff;}</style></head>
                        <body>
                        <h3>${cmpEntity} · ${kCfg.label} · ${sorted.map(s=>s.name).join(', ')}</h3>
                        ${svgData.replace(/var\(--\w[\w-]*\)/g,'#64748b').replace(/var\(--border\)/g,'#e2e8f0').replace(/var\(--muted\)/g,'#94a3b8').replace(/var\(--bg2\)/g,'#fff').replace(/var\(--text\)/g,'#1e293b')}
                        <script>window.onload=()=>{window.print();}<\/script>
                        </body></html>`);
                      win.document.close();
                    };

                    return (
                      <div style={{ background:'var(--bg2)', border:'1px solid var(--border)', borderRadius:10, padding:'16px', overflowX:'auto' }}>
                        {/* En-tête : légende + boutons export */}
                        <div style={{ display:'flex', justifyContent:'space-between', alignItems:'center', flexWrap:'wrap', gap:8, marginBottom:12 }}>
                          <div style={{ display:'flex', gap:14, flexWrap:'wrap' }}>
                            {series.map(s => (
                              <div key={s.name} style={{ display:'flex', alignItems:'center', gap:5, fontSize:11, color:'var(--text)' }}>
                                <div style={{ width:12, height:12, background:s.color, borderRadius:3 }}/>
                                <span style={{ fontFamily:'monospace', fontWeight:600 }}>{s.name}</span>
                              </div>
                            ))}
                          </div>
                          <div style={{ display:'flex', gap:6 }}>
                            <button onClick={exportXlsx} style={{
                              fontSize:11, padding:'5px 11px', borderRadius:6, border:'none', cursor:'pointer',
                              background:'#16A34A', color:'#fff', fontWeight:700, display:'flex', alignItems:'center', gap:4,
                            }}>📥 Excel</button>
                            <button onClick={exportPdf} style={{
                              fontSize:11, padding:'5px 11px', borderRadius:6, border:'none', cursor:'pointer',
                              background:'#DC2626', color:'#fff', fontWeight:700, display:'flex', alignItems:'center', gap:4,
                            }}>📄 PDF</button>
                          </div>
                        </div>

                        {/* SVG barres verticales */}
                        <svg id="cmp-search-svg" width={W} height={H} style={{ display:'block', overflow:'visible', background:'var(--bg2)' }}>
                          {/* Grille Y */}
                          {yTicks.map((v, i) => (
                            <g key={i}>
                              <line x1={M_L} y1={by(v)} x2={M_L + CW} y2={by(v)}
                                stroke="var(--border)" strokeWidth={0.5} strokeDasharray="4,3"/>
                              <text x={M_L - 6} y={by(v) + 4} textAnchor="end" fontSize={9} fill="var(--muted)">
                                {kCfg.fmt(v)}
                              </text>
                            </g>
                          ))}

                          {/* Axe X */}
                          <line x1={M_L} y1={M_T + CH} x2={M_L + CW} y2={M_T + CH}
                            stroke="var(--border)" strokeWidth={1}/>

                          {/* Groupes de barres par mois */}
                          {allMonths.map((m, mi) => (
                            <g key={m}>
                              {/* Label mois */}
                              <text x={M_L + mi * grpW + grpW / 2} y={H - M_B + 16}
                                textAnchor="middle" fontSize={10} fill="var(--muted)" fontWeight="600">
                                {m}
                              </text>
                              {/* Barres par entité */}
                              {series.map((s, si) => {
                                const v   = s.vals[mi];
                                const x   = bx(mi, si);
                                const h   = Math.max(bh(v), 2);
                                const y   = by(v);
                                // Couleur plus sombre si c'est le mois le plus haut de cette série
                                const isMax = v === Math.max(...s.vals);
                                return (
                                  <g key={si}>
                                    <rect x={x} y={y} width={barW - 2} height={h}
                                      fill={s.color} opacity={isMax ? 1 : 0.65} rx={2}/>
                                    {/* Valeur au-dessus si la barre est assez haute */}
                                    {h > 18 && (
                                      <text x={x + (barW - 2) / 2} y={y - 3}
                                        textAnchor="middle" fontSize={8} fill={s.color} fontWeight="bold">
                                        {kCfg.fmt(v)}
                                      </text>
                                    )}
                                    {/* Valeur toujours visible en haut de barre petite */}
                                    {h <= 18 && v > 0 && (
                                      <text x={x + (barW - 2) / 2} y={y - 3}
                                        textAnchor="middle" fontSize={8} fill={s.color} fontWeight="bold">
                                        {kCfg.fmt(v)}
                                      </text>
                                    )}
                                  </g>
                                );
                              })}
                            </g>
                          ))}
                        </svg>

                        {/* Tableau récap sous le graphique */}
                        <div style={{ marginTop:16, overflowX:'auto' }}>
                          <table style={{ width:'100%', borderCollapse:'collapse', fontSize:11 }}>
                            <thead>
                              <tr>
                                <th style={{ textAlign:'left', padding:'5px 8px', color:'var(--muted)', borderBottom:'1px solid var(--border)' }}>Entité</th>
                                {allMonths.map(m => (
                                  <th key={m} style={{ textAlign:'right', padding:'5px 8px', color:'var(--muted)', borderBottom:'1px solid var(--border)' }}>{m}</th>
                                ))}
                                <th style={{ textAlign:'right', padding:'5px 8px', color:'var(--muted)', borderBottom:'1px solid var(--border)' }}>Total</th>
                              </tr>
                            </thead>
                            <tbody>
                              {series.map((s, si) => (
                                <tr key={s.name} style={{ background: si%2===0 ? 'var(--bg2)':'var(--bg3)' }}>
                                  <td style={{ padding:'5px 8px', fontFamily:'monospace', fontWeight:700, color:s.color }}>{s.name}</td>
                                  {s.vals.map((v, mi) => (
                                    <td key={mi} style={{ textAlign:'right', padding:'5px 8px', color:'var(--text)' }}>
                                      {kCfg.fmt(v)}
                                    </td>
                                  ))}
                                  <td style={{ textAlign:'right', padding:'5px 8px', fontWeight:700, color:s.color }}>
                                    {kCfg.fmt(s.vals.reduce((a,b)=>a+b,0))}
                                  </td>
                                </tr>
                              ))}
                            </tbody>
                          </table>
                        </div>
                      </div>
                    );
                  })()

                  : (
                  /* ── VUE GÉNÉRALE : barres groupées A vs B ── */
                  <div style={{ background:'var(--bg2)', border:'1px solid var(--border)', borderRadius:10, padding:'12px 16px', overflowX:'auto' }}>
                    {sorted.length === 0
                      ? <p style={{ color:'var(--muted)', fontSize:12 }}>Aucune donnée.</p>
                      : (
                        <svg width={LABEL_W + CHART_W + 80} height={SVG_H} style={{ display:'block' }}>
                          {sorted.map((r, i) => {
                            const y = 20 + i * ROW_H;
                            const wA = (r.vA / maxV) * CHART_W;
                            const wB = (r.vB / maxV) * CHART_W;
                            const delta = r.vB !== 0 ? ((r.vA - r.vB) / r.vB * 100) : null;
                            const dColor = delta == null ? '#64748B' : delta >= 0 ? '#34D399' : '#F87171';
                            return (
                              <g key={r.name}>
                                <text x={LABEL_W - 8} y={y + BAR_H - 2} textAnchor="end"
                                  fontSize={10} fill="var(--muted)" style={{ fontFamily:'monospace' }}>
                                  {r.name.length > 16 ? r.name.slice(0,15)+'…' : r.name}
                                </text>
                                <rect x={LABEL_W} y={y} width={Math.max(wA,2)} height={BAR_H} fill={kCfg.colorA} rx={3}/>
                                <text x={LABEL_W + wA + 5} y={y + BAR_H - 4} fontSize={9} fill={kCfg.colorA} fontWeight="bold">
                                  {kCfg.fmt(r.vA)}
                                </text>
                                <rect x={LABEL_W} y={y + BAR_H + GAP} width={Math.max(wB,2)} height={BAR_H} fill={kCfg.colorB} rx={3}/>
                                <text x={LABEL_W + wB + 5} y={y + BAR_H + GAP + BAR_H - 4} fontSize={9} fill={kCfg.colorB} fontWeight="bold">
                                  {kCfg.fmt(r.vB)}
                                </text>
                                {delta != null && (
                                  <text x={LABEL_W + CHART_W + 10} y={y + BAR_H} fontSize={10} fill={dColor} fontWeight="bold">
                                    {delta >= 0 ? '+' : ''}{delta.toFixed(1)}%
                                  </text>
                                )}
                                {i < sorted.length-1 && (
                                  <line x1={0} y1={y + ROW_H - 4} x2={LABEL_W + CHART_W + 70} y2={y + ROW_H - 4}
                                    stroke="var(--border)" strokeWidth={0.5} strokeDasharray="4,4"/>
                                )}
                              </g>
                            );
                          })}
                        </svg>
                      )
                    }
                  </div>
                  )}
                </div>
              );
            })()}
          </div>
        );
      })()}

    </div>
  );
}

// ── EXPORTS MODULE (appended) ─────────────────────────────────────────────────
