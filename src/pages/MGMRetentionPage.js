// src/pages/MGMRetentionPage.js
import React, { useState, useRef, useCallback } from 'react';
import Papa from 'papaparse';
import * as XLSX from 'xlsx';
import { saveAs } from 'file-saver';

const PAGE = 50;

function nv(row, col) { const v = parseFloat(row[col]); return isNaN(v) ? 0 : v; }
function fmt(v) {
  if (!v && v !== 0) return '—';
  if (Math.abs(v) >= 1e6) return (v / 1e6).toFixed(1) + 'M';
  if (Math.abs(v) >= 1000) return (v / 1000).toFixed(0) + 'K';
  return Math.round(v).toLocaleString();
}

function parseFile(file) {
  return new Promise((res) => {
    const ext = file.name.split('.').pop().toLowerCase();
    if (ext === 'csv') {
      Papa.parse(file, { header: true, skipEmptyLines: true, dynamicTyping: true, complete: r => res(r.data) });
    } else {
      const rd = new FileReader();
      rd.onload = e => {
        const wb = XLSX.read(e.target.result, { type: 'array' });
        res(XLSX.utils.sheet_to_json(wb.Sheets[wb.SheetNames[0]], { defval: 0 }));
      };
      rd.readAsArrayBuffer(file);
    }
  });
}

function runAnalysis(current, compare) {
  const cm = {};
  current.forEach(r => { cm[r['Account ID']] = r; });
  const sport = [], casino = [], deposit = [];
  compare.forEach(p => {
    const id = p['Account ID'], now = cm[id];
    const pt = nv(p, 'Standard Payin Tickets') + nv(p, 'Bonus Payin Tickets');
    const pc = nv(p, 'Casino Payin'), pd = nv(p, 'Deposit Money');
    const ct = now ? nv(now, 'Standard Payin Tickets') + nv(now, 'Bonus Payin Tickets') : 0;
    const cc = now ? nv(now, 'Casino Payin') : 0;
    const cd = now ? nv(now, 'Deposit Money') : 0;
    const b = {
      id, currency: p['Currency'] || 'XAF',
      name: ((p['First Name'] || '') + ' ' + (p['Last Name'] || '')).trim(),
      phone: p['Phone Number'] || '', email: p['Email'] || '',
      city: p['City'] || '', country: p['Country'] || '',
    };
    if (pt > 0 && ct === 0) sport.push({ ...b, pt, pg: nv(p, 'SportBetting Gross'), pp: nv(p, 'Standard Payin Money') });
    if (pc > 0 && cc === 0) casino.push({ ...b, pc, pcg: nv(p, 'Casino Gross') });
    if (pd > 0 && cd === 0) deposit.push({ ...b, pd });
  });
  sport.sort((a, b) => b.pp - a.pp);
  casino.sort((a, b) => b.pc - a.pc);
  deposit.sort((a, b) => b.pd - a.pd);
  return { sport, casino, deposit };
}

// ── Upload Slot ──────────────────────────────────
function UploadSlot({ tag, tagClass, icon, label, file, onFile, accept = '.csv,.xlsx,.xls' }) {
  const ref = useRef();
  const loaded = !!file;
  return (
    <div
      onClick={() => ref.current.click()}
      onDragOver={e => e.preventDefault()}
      onDrop={e => { e.preventDefault(); if (e.dataTransfer.files[0]) onFile(e.dataTransfer.files[0]); }}
      className={`drop-zone ${loaded ? 'has-file' : ''}`}
      style={{ cursor: 'pointer' }}
    >
      <input ref={ref} type="file" accept={accept} hidden onChange={e => e.target.files[0] && onFile(e.target.files[0])} />
      <div style={{ display: 'inline-block', fontSize: 10, fontWeight: 700, padding: '3px 10px', borderRadius: 6, marginBottom: 8, background: tagClass === 'curr' ? 'rgba(37,99,235,0.12)' : 'rgba(100,116,139,0.12)', color: tagClass === 'curr' ? 'var(--accent)' : 'var(--muted)' }}>{tag}</div>
      <div className="drop-icon">{loaded ? '✅' : icon}</div>
      <div className="drop-label">{loaded ? 'Loaded — click to replace' : label}</div>
      {loaded && <div className="file-chip">✓ {file.name} · {file.rows?.toLocaleString()} players</div>}
    </div>
  );
}

// ── Retention Table ──────────────────────────────
function RetTable({ rows, tab }) {
  if (!rows.length) return (
    <div style={{ textAlign: 'center', padding: '3rem', color: 'var(--muted)' }}>
      <div style={{ fontSize: 32, marginBottom: 8 }}>🎯</div>
      <div style={{ fontSize: 14 }}>No targets in this category</div>
    </div>
  );
  const cols = tab === 'sport'
    ? ['Account ID', 'Country', 'Phone', 'Tickets (prev)', 'Payin (prev)', 'Status']
    : tab === 'casino'
    ? ['Account ID', 'Country', 'Phone', 'Casino Payin (prev)', 'Status']
    : ['Account ID', 'Country', 'Phone', 'Deposit (prev)', 'Status'];

  return (
    <div style={{ overflowX: 'auto' }}>
      <table className="hist-table">
        <thead>
          <tr>{cols.map(c => <th key={c}>{c}</th>)}</tr>
        </thead>
        <tbody>
          {rows.map((r, i) => (
            <tr key={i}>
              <td style={{ fontWeight: 600, fontFamily: 'JetBrains Mono, monospace' }}>{r.id}</td>
              <td style={{ fontSize: 12 }}>{r.country} <span style={{ color: 'var(--muted)', fontSize: 11 }}>{r.city}</span></td>
              <td style={{ fontFamily: 'JetBrains Mono, monospace', fontSize: 12 }}>{r.phone}</td>
              {tab === 'sport' && <>
                <td style={{ fontFamily: 'JetBrains Mono, monospace' }}>{r.pt?.toLocaleString()}</td>
                <td style={{ fontFamily: 'JetBrains Mono, monospace' }}>{Math.round(r.pp).toLocaleString()} {r.currency}</td>
              </>}
              {tab === 'casino' && <td style={{ fontFamily: 'JetBrains Mono, monospace' }}>{Math.round(r.pc).toLocaleString()} {r.currency}</td>}
              {tab === 'deposit' && <td style={{ fontFamily: 'JetBrains Mono, monospace' }}>{Math.round(r.pd).toLocaleString()} {r.currency}</td>}
              <td><span style={{ display: 'inline-block', fontSize: 10, fontWeight: 700, padding: '2px 8px', borderRadius: 5, background: 'var(--danger-bg)', color: 'var(--danger)' }}>0 {tab.toUpperCase()}</span></td>
            </tr>
          ))}
        </tbody>
      </table>
    </div>
  );
}

// ── Main Component ───────────────────────────────
export default function MGMRetentionPage({ activeTab, onNavigate, results: extResults, onResults }) {
  // Retention state
  const [curr, setCurr]       = useState(null);
  const [comp, setComp]       = useState(null);
  const [results, setResults] = [extResults, onResults];
  const [_view, _setView]      = useState('upload');
  const view = activeTab ? 'results' : _view;
  const setView = (v) => { _setView(v); };
  const [_tab, _setTab]        = useState('sport');
  const tab = activeTab || _tab;
  const setTab = (t) => { _setTab(t); if (onNavigate) onNavigate('mgm-' + t); };
  const [search, setSearch]   = useState('');
  const [page, setPage]       = useState(1);

  // Bonus Validator state
  const [bTargets, setBTargets] = useState(null);
  const [bOffer, setBOffer]     = useState(null);
  const [bType, setBType]       = useState('sport');
  const [bMinD, setBMinD]       = useState(0);
  const [bMinP, setBMinP]       = useState(0);
  const [bLabel, setBLabel]     = useState('');
  const [bResults, setBResults] = useState(null);
  const [bSearch, setBSearch]   = useState('');
  const [bPage, setBPage]       = useState(1);

  const loadFile = useCallback(async (file, setter) => {
    const data = await parseFile(file);
    file.rows = data.length;
    setter({ file, data });
  }, []);

  const doRun = useCallback(() => {
    if (!curr || !comp) return;
    const r = runAnalysis(curr.data, comp.data);
    setResults(r);
    setTab('sport'); setPage(1); setSearch('');
    setView('results');
    if (onNavigate) onNavigate('mgm-sport');
  }, [curr, comp]);

  const doBonus = useCallback(() => {
    if (!bTargets || !bOffer) return;
    const isSport = bType === 'sport';
    const om = {};
    bOffer.data.forEach(r => { om[r['Account ID']] = r; });
    const valid = [], invalid = [];
    bTargets.data.forEach(tr => {
      const id = tr['Account ID'] || tr['id'] || tr['ID']; if (!id) return;
      const o = om[id];
      const dep = o ? nv(o, 'Deposit Money') : 0;
      const payin = o ? (isSport ? nv(o, 'Standard Payin Money') : nv(o, 'Casino Payin')) : 0;
      const cur2 = (o && o['Currency']) || tr['Currency'] || 'XAF';
      const nm = tr['Name'] || tr['name'] || ((tr['First Name'] || '') + ' ' + (tr['Last Name'] || '')).trim() || '';
      const pl = { id, name: nm, phone: tr['Phone'] || tr['Phone Number'] || '', city: tr['City'] || tr['city'] || '', country: tr['Country'] || tr['country'] || '', currency: cur2, dep, payin };
      const checks = [];
      if (bMinD > 0) checks.push(dep >= bMinD);
      if (bMinP > 0) checks.push(payin >= bMinP);
      if (checks.length === 0) return;
      if (checks.every(Boolean)) valid.push(pl); else invalid.push(pl);
    });
    valid.sort((a, b) => b.dep - a.dep);
    setBResults({ valid, invalid, isSport });
    setBPage(1); setBSearch('');
  }, [bTargets, bOffer, bType, bMinD, bMinP]);

  const exportRet = () => {
    if (!results) return;
    const wb = XLSX.utils.book_new();
    const sh = (rows, t) => XLSX.utils.json_to_sheet(rows.map(r => {
      const o = { 'Account ID': r.id, Name: r.name, Country: r.country, City: r.city, Phone: r.phone, Email: r.email, Currency: r.currency };
      if (t === 'sport') { o['Tickets(Prev)'] = r.pt; o['Payin(Prev)'] = r.pp; o['Gross(Prev)'] = r.pg; }
      else if (t === 'casino') { o['CasinoPayin(Prev)'] = r.pc; o['CasinoGross(Prev)'] = r.pcg; }
      else { o['Deposit(Prev)'] = r.pd; }
      return o;
    }));
    XLSX.utils.book_append_sheet(wb, sh(results.sport, 'sport'), 'Sport');
    XLSX.utils.book_append_sheet(wb, sh(results.casino, 'casino'), 'Casino');
    XLSX.utils.book_append_sheet(wb, sh(results.deposit, 'deposit'), 'Deposit');
    const blob = new Blob([XLSX.write(wb, { bookType: 'xlsx', type: 'array' })], { type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet' });
    saveAs(blob, `MGM_Retention_${new Date().toISOString().slice(0, 10)}.xlsx`);
  };

  const exportBonus = () => {
    if (!bResults) return;
    const wb = XLSX.utils.book_new();
    const rows = bResults.valid.map(r => {
      const o = { 'Account ID': r.id, Name: r.name, Country: r.country, City: r.city, Phone: r.phone, Currency: r.currency, Deposit: r.dep };
      o[bResults.isSport ? 'Sport Payin' : 'Casino Payin'] = r.payin;
      o['Bonus Status'] = 'VALID';
      return o;
    });
    XLSX.utils.book_append_sheet(wb, XLSX.utils.json_to_sheet(rows), 'Valid for Bonus');
    const blob = new Blob([XLSX.write(wb, { bookType: 'xlsx', type: 'array' })], { type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet' });
    saveAs(blob, `MGM_BonusValid_${new Date().toISOString().slice(0, 10)}.xlsx`);
  };

  // Filtered results
  const allRows = results ? results[tab] : [];
  const q = search.toLowerCase();
  const filtered = q ? allRows.filter(r => r.name.toLowerCase().includes(q) || String(r.id).includes(q) || r.city.toLowerCase().includes(q) || String(r.phone).includes(q)) : allRows;
  const totalPages = Math.max(1, Math.ceil(filtered.length / PAGE));
  const pageRows = filtered.slice((page - 1) * PAGE, page * PAGE);

  const bFiltered = bResults ? (bSearch ? bResults.valid.filter(r => String(r.id).includes(bSearch) || r.name.toLowerCase().includes(bSearch) || String(r.phone).includes(bSearch)) : bResults.valid) : [];
  const bTotalPages = Math.max(1, Math.ceil(bFiltered.length / PAGE));
  const bPageRows = bFiltered.slice((bPage - 1) * PAGE, bPage * PAGE);

  const uniqueTotal = results ? Object.keys(Object.fromEntries([...results.sport, ...results.casino, ...results.deposit].map(r => [r.id, 1]))).length : 0;

  const TABS = [
    { k: 'sport',   label: '⚽ Sport',   count: results?.sport.length || 0,   color: 'var(--sport)',   bg: 'var(--sport-bg)' },
    { k: 'casino',  label: '🎰 Casino',  count: results?.casino.length || 0,  color: '#D97706',        bg: 'var(--casino-bg)' },
    { k: 'deposit', label: '💳 Deposit', count: results?.deposit.length || 0, color: 'var(--deposit)', bg: 'var(--deposit-bg)' },
  ];

  return (
    <div>
      {view === 'upload' ? (
        <>
          {/* Retention Upload */}
          <p className="section-label">Retention Target Lists</p>
          <div style={{ fontSize: 13, color: 'var(--muted)', marginBottom: '1rem' }}>
            Upload two files: current month and comparison month. Finds players with zero activity in current month who were active before.
          </div>

          <div style={{ display: 'grid', gridTemplateColumns: '1fr 1fr', gap: '1rem', marginBottom: '1rem' }}>
            <UploadSlot tag="Current Month" tagClass="curr" icon="📅" label="Click to upload" file={curr?.file} onFile={f => loadFile(f, setCurr)} />
            <UploadSlot tag="Comparison Month" tagClass="comp" icon="📆" label="Click to upload" file={comp?.file} onFile={f => loadFile(f, setComp)} />
          </div>

          <div className="logic-box" style={{ marginBottom: '1.25rem' }}>
            📌 <strong>Logic:</strong> Targets = players active in comparison month with <strong>zero</strong> in current month.<br />
            ⚽ Sport — had tickets → now 0 &nbsp;|&nbsp; 🎰 Casino — had payin → now 0 &nbsp;|&nbsp; 💳 Deposit — had deposit → now 0
          </div>

          <div style={{ display: 'flex', gap: 8, marginBottom: '2rem' }}>
            <button className="btn primary" disabled={!curr || !comp} onClick={doRun} style={{ fontSize: 14, padding: '10px 24px' }}>
              → Generate Retention Target Lists
            </button>
            <button className="btn" onClick={() => { setCurr(null); setComp(null); setResults(null); }} style={{ color: 'var(--danger)', borderColor: 'rgba(239,68,68,0.3)' }}>
              ✕ Clear All
            </button>
          </div>

          <hr className="divider" />

          {/* Bonus Validator */}
          <p className="section-label">🎁 Bonus Validator</p>
          <div style={{ fontSize: 13, color: 'var(--muted)', marginBottom: '1rem' }}>
            Upload your target list and offer period report. Set conditions to find who qualifies for the bonus.
          </div>

          <div style={{ display: 'grid', gridTemplateColumns: '1fr 1fr', gap: '1rem', marginBottom: '1.25rem' }}>
            <UploadSlot tag="Target List" tagClass="curr" icon="👥" label="Upload exported target list" file={bTargets?.file} onFile={f => loadFile(f, setBTargets)} />
            <UploadSlot tag="Offer Period Report" tagClass="comp" icon="📊" label="Player report for offer period" file={bOffer?.file} onFile={f => loadFile(f, setBOffer)} />
          </div>

          <div className="card" style={{ marginBottom: '1.25rem' }}>
            <p className="section-label">Offer Type</p>
            <div style={{ display: 'flex', gap: 8, marginBottom: '1rem' }}>
              {['sport', 'casino'].map(t => (
                <button key={t} onClick={() => setBType(t)} style={{
                  flex: 1, padding: '9px', borderRadius: 8, fontSize: 13, fontWeight: 600, cursor: 'pointer',
                  border: `1px solid ${bType === t ? (t === 'sport' ? 'var(--sport)' : '#D97706') : 'var(--border)'}`,
                  background: bType === t ? (t === 'sport' ? 'var(--sport-bg)' : 'var(--casino-bg)') : 'transparent',
                  color: bType === t ? (t === 'sport' ? 'var(--sport)' : '#D97706') : 'var(--muted)',
                }}>
                  {t === 'sport' ? '⚽ Sport Retention' : '🎰 Casino Retention'}
                </button>
              ))}
            </div>

            <p className="section-label">Conditions — ALL must be met (0 = skip)</p>
            <div style={{ display: 'grid', gridTemplateColumns: '1fr 1fr', gap: '1rem', marginBottom: '1rem' }}>
              <label style={{ display: 'flex', flexDirection: 'column', gap: 5 }}>
                <span style={{ fontSize: 11, color: 'var(--muted)', fontWeight: 600 }}>Min. Deposit Amount</span>
                <input type="number" min="0" value={bMinD} onChange={e => setBMinD(Number(e.target.value))} className="field-input" />
              </label>
              <label style={{ display: 'flex', flexDirection: 'column', gap: 5 }}>
                <span style={{ fontSize: 11, color: 'var(--muted)', fontWeight: 600 }}>Min. {bType === 'sport' ? 'Sport' : 'Casino'} Payin Amount</span>
                <input type="number" min="0" value={bMinP} onChange={e => setBMinP(Number(e.target.value))} className="field-input" />
              </label>
            </div>
            <label style={{ display: 'flex', flexDirection: 'column', gap: 5 }}>
              <span style={{ fontSize: 11, color: 'var(--muted)', fontWeight: 600 }}>Offer / Bonus Label (optional)</span>
              <input type="text" value={bLabel} onChange={e => setBLabel(e.target.value)} placeholder="e.g. Sport Retention 100% up to 50,000 XAF" className="field-input" style={{ width: '100%' }} />
            </label>
          </div>

          <div style={{ display: 'flex', gap: 8 }}>
            <button className="btn primary" disabled={!bTargets || !bOffer} onClick={doBonus} style={{ fontSize: 14, padding: '10px 24px' }}>
              → Check Who Qualifies for Bonus
            </button>
            <button className="btn" onClick={() => { setBTargets(null); setBOffer(null); setBResults(null); }} style={{ color: 'var(--danger)', borderColor: 'rgba(239,68,68,0.3)' }}>
              ✕ Clear
            </button>
          </div>

          {/* Bonus Results */}
          {bResults && (
            <div style={{ marginTop: '1.5rem' }}>
              {bLabel && <div style={{ fontSize: 14, fontWeight: 600, marginBottom: 10 }}>{bLabel}</div>}
              <div className="kpi-banner" style={{ marginBottom: '1rem' }}>
                <div className="kpi-item"><div className="kpi-label">Targeted</div><div className="kpi-value" style={{ fontSize: 22 }}>{(bResults.valid.length + bResults.invalid.length).toLocaleString()}</div></div>
                <div className="kpi-item"><div className="kpi-label" style={{ color: 'var(--deposit)' }}>✓ Valid</div><div className="kpi-value" style={{ fontSize: 22, color: 'var(--deposit)' }}>{bResults.valid.length.toLocaleString()}</div></div>
                <div className="kpi-item"><div className="kpi-label" style={{ color: 'var(--danger)' }}>✗ Failed</div><div className="kpi-value" style={{ fontSize: 22, color: 'var(--danger)' }}>{bResults.invalid.length.toLocaleString()}</div></div>
              </div>
              <div className="card" style={{ padding: 0, overflow: 'hidden' }}>
                <div style={{ display: 'flex', alignItems: 'center', gap: 8, padding: '10px 14px', borderBottom: '1px solid var(--border)' }}>
                  <input className="field-input" placeholder="Search ID, phone..." value={bSearch} onChange={e => { setBSearch(e.target.value); setBPage(1); }} style={{ flex: 1 }} />
                  <span style={{ fontSize: 12, color: 'var(--muted)' }}>{bFiltered.length.toLocaleString()} valid</span>
                  <button className="btn sm" onClick={exportBonus}>⬇ Export</button>
                </div>
                <table className="hist-table">
                  <thead><tr><th>Account ID</th><th>Country</th><th>Phone</th><th>Deposit</th><th>{bResults.isSport ? 'Sport Payin' : 'Casino Payin'}</th><th>Status</th></tr></thead>
                  <tbody>
                    {bPageRows.map((r, i) => (
                      <tr key={i}>
                        <td style={{ fontWeight: 600, fontFamily: 'JetBrains Mono, monospace' }}>{r.id}</td>
                        <td style={{ fontSize: 12 }}>{r.country}</td>
                        <td style={{ fontFamily: 'JetBrains Mono, monospace', fontSize: 12 }}>{r.phone}</td>
                        <td style={{ fontFamily: 'JetBrains Mono, monospace' }}>{Math.round(r.dep).toLocaleString()} {r.currency}</td>
                        <td style={{ fontFamily: 'JetBrains Mono, monospace' }}>{Math.round(r.payin).toLocaleString()} {r.currency}</td>
                        <td><span style={{ fontSize: 10, fontWeight: 700, padding: '2px 8px', borderRadius: 5, background: 'var(--deposit-bg)', color: 'var(--deposit)' }}>✓ VALID</span></td>
                      </tr>
                    ))}
                  </tbody>
                </table>
                <Pagination page={bPage} total={bFiltered.length} totalPages={bTotalPages} onPage={p => { setBPage(p); }} />
              </div>
            </div>
          )}
        </>
      ) : (
        <>
          {/* Results view */}
          <div style={{ display: 'flex', alignItems: 'center', justifyContent: 'space-between', marginBottom: '1rem' }}>
            <p className="section-label" style={{ margin: 0 }}>Retention Target Lists</p>
            <button className="btn sm" onClick={() => setView('upload')}>↩ Back</button>
          </div>

          <div className="kpi-banner" style={{ marginBottom: '1.25rem' }}>
            <div className="kpi-item"><div className="kpi-label">Total Unique</div><div className="kpi-value" style={{ fontSize: 22 }}>{uniqueTotal.toLocaleString()}</div><div className="kpi-sub">Zero activity</div></div>
            {TABS.map(t => (
              <div key={t.k} className="kpi-item" style={{ cursor: 'pointer' }} onClick={() => { setTab(t.k); setPage(1); }}>
                <div className="kpi-label">{t.label}</div>
                <div className="kpi-value" style={{ fontSize: 22, color: t.color }}>{t.count.toLocaleString()}</div>
              </div>
            ))}
          </div>

          {/* Tabs */}
          <div style={{ display: 'flex', gap: 6, marginBottom: '1rem' }}>
            {TABS.map(t => (
              <button key={t.k} onClick={() => { setTab(t.k); setPage(1); }}
                style={{ padding: '7px 14px', borderRadius: 8, fontSize: 12, fontWeight: 600, cursor: 'pointer', border: `1px solid ${tab === t.k ? t.color : 'var(--border)'}`, background: tab === t.k ? t.bg : 'transparent', color: tab === t.k ? t.color : 'var(--muted)' }}>
                {t.label} <span style={{ fontSize: 11, fontWeight: 700, padding: '1px 6px', borderRadius: 8, background: 'rgba(0,0,0,0.08)', marginLeft: 4 }}>{t.count}</span>
              </button>
            ))}
          </div>

          <div className="card" style={{ padding: 0, overflow: 'hidden' }}>
            <div style={{ display: 'flex', alignItems: 'center', gap: 8, padding: '10px 14px', borderBottom: '1px solid var(--border)' }}>
              <input className="field-input" placeholder="Search ID, city, phone..." value={search} onChange={e => { setSearch(e.target.value); setPage(1); }} style={{ flex: 1 }} />
              <span style={{ fontSize: 12, color: 'var(--muted)' }}>{filtered.length.toLocaleString()} players</span>
              <button className="btn sm" onClick={exportRet}>⬇ Export</button>
            </div>
            <RetTable rows={pageRows} tab={tab} />
            <Pagination page={page} total={filtered.length} totalPages={totalPages} onPage={p => setPage(p)} />
          </div>
        </>
      )}
    </div>
  );
}

function Pagination({ page, total, totalPages, onPage }) {
  if (totalPages <= 1) return null;
  const start = (page - 1) * PAGE + 1;
  const end = Math.min(page * PAGE, total);
  const pages = [];
  for (let i = 1; i <= totalPages; i++) {
    if (i === 1 || i === totalPages || Math.abs(i - page) <= 2) pages.push(i);
    else if (pages[pages.length - 1] !== '…') pages.push('…');
  }
  return (
    <div style={{ display: 'flex', alignItems: 'center', justifyContent: 'space-between', padding: '10px 14px', borderTop: '1px solid var(--border)', fontSize: 12, color: 'var(--muted)' }}>
      <span>{start}–{end} of {total.toLocaleString()}</span>
      <div style={{ display: 'flex', gap: 3 }}>
        <button className="btn sm" onClick={() => onPage(page - 1)} disabled={page === 1}>‹</button>
        {pages.map((p, i) => typeof p === 'number'
          ? <button key={i} className="btn sm" onClick={() => onPage(p)} style={{ background: p === page ? 'var(--accent)' : '', color: p === page ? '#fff' : '', borderColor: p === page ? 'var(--accent)' : '' }}>{p}</button>
          : <span key={i} style={{ padding: '0 4px', lineHeight: '28px' }}>…</span>
        )}
        <button className="btn sm" onClick={() => onPage(page + 1)} disabled={page === totalPages}>›</button>
      </div>
    </div>
  );
}
