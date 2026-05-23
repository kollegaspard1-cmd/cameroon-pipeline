// src/pages/CMReportPage.js
import React, { useState, useRef } from 'react';
import * as XLSX from 'xlsx';
import Papa from 'papaparse';
import { saveAs } from 'file-saver';

const XAF_EUR_RATE = 657.9;
const eur = (xaf) => Math.round(xaf / XAF_EUR_RATE);

export default function CMReportPage() {
  const [files, setFiles] = useState({ t1: null, t2: null, a1: null, a2: null });
  const [running, setRunning]   = useState(false);
  const [logs, setLogs]         = useState([]);
  const [result, setResult]     = useState(null);
  const refs = { t1: useRef(), t2: useRef(), a1: useRef(), a2: useRef() };

  const log = (msg, type = 'normal') =>
    setLogs(p => [...p, { msg, type, t: new Date().toLocaleTimeString() }]);

  const setFile = (key, f) => setFiles(p => ({ ...p, [key]: f }));

  // ── CSV loader ───────────────────────────────────
  const loadCSV = (file) => new Promise((res, rej) => {
    Papa.parse(file, {
      header: true, skipEmptyLines: true, dynamicTyping: true,
      complete: ({ data }) => res(data),
      error: rej,
    });
  });

  // ── Compute KPIs ─────────────────────────────────
  const computeKPIs = (t1, t2, a1, a2) => {
    const newIdsYest = new Set(a1.map(r => r['Account ID']));
    const newIdsMTD  = new Set(a2.map(r => r['Account ID']));
    const t1New = t1.filter(r => newIdsYest.has(r['Account ID']));
    const t2New = t2.filter(r => newIdsMTD.has(r['Account ID']));

    const sum = (arr, col) => arr.reduce((s, r) => s + (Number(r[col]) || 0), 0);

    return {
      sport_payin_yesterday:      eur(sum(t1, 'Standard Payin Money')),
      sport_gross_yesterday:      eur(sum(t1, 'SportBetting Gross')),
      deposit_yesterday:          eur(sum(t1, 'Deposit Money')),
      sport_payin_mtd:            eur(sum(t2, 'Standard Payin Money')),
      sport_gross_mtd:            eur(sum(t2, 'SportBetting Gross')),
      deposit_mtd:                eur(sum(t2, 'Deposit Money')),
      registered_yesterday:       a1.length,
      registered_mtd:             a2.length,
      active_sport_yesterday:     t1.filter(r => (Number(r['SportBetting Gross']) || 0) !== 0).length,
      active_sport_mtd:           t2.filter(r => (Number(r['SportBetting Gross']) || 0) !== 0).length,
      bonus_new_payin_yesterday:  Math.round(sum(t1New, 'Bonus Payin Money EUR')),
      bonus_new_payout_yesterday: Math.round(sum(t1New, 'Bonus Payout Amount EUR')),
      bonus_new_payin_mtd:        Math.round(sum(t2New, 'Bonus Payin Money EUR')),
      bonus_new_payout_mtd:       Math.round(sum(t2New, 'Bonus Payout Amount EUR')),
      bonus_total_payin_mtd:      Math.round(sum(t2, 'Bonus Payin Money EUR')),
      bonus_total_payout_mtd:     Math.round(sum(t2, 'Bonus Payout Amount EUR')),
    };
  };

  // ── Top players ──────────────────────────────────
  const topPlayers = (t1, n = 5) => {
    const rows = t1.map(r => ({
      id:    r['Account ID'],
      name:  `${r['First Name'] || ''} ${r['Last Name'] || ''}`.trim(),
      gross: Number(r['SportBetting Gross EUR']) || 0,
      type:  r['Player Type Risk'] || '',
      bets:  Number(r['Standard Payin Tickets']) || 0,
    }));
    const losses = [...rows].filter(r => r.gross > 0).sort((a, b) => b.gross - a.gross).slice(0, n);
    const wins   = [...rows].filter(r => r.gross < 0).sort((a, b) => a.gross - b.gross).slice(0, n);
    return { losses, wins };
  };

  // ── Build XLSX from template structure ───────────
  const buildReport = (kpis, losses, wins) => {
    const wb = XLSX.utils.book_new();

    // ── Sheet1: Report ──
    const ws = {};
    const merges = [];

    const hdr = { font: { bold: true, sz: 11, color: { rgb: 'FFFFFF' } }, fill: { fgColor: { rgb: '1F3864' } }, alignment: { horizontal: 'center', vertical: 'center', wrapText: true }, border: bdr() };
    const lbl = { font: { bold: true, sz: 10 }, fill: { fgColor: { rgb: 'D9E1F2' } }, border: bdr() };
    const val = { font: { sz: 10 }, fill: { fgColor: { rgb: 'FFFFFF' } }, border: bdr(), alignment: { horizontal: 'right' }, numFmt: '#,##0' };
    const fml = { font: { sz: 10, color: { rgb: '155724' } }, fill: { fgColor: { rgb: 'D4EDDA' } }, border: bdr(), alignment: { horizontal: 'right' } };
    const sec = { font: { bold: true, sz: 11, color: { rgb: 'FFFFFF' } }, fill: { fgColor: { rgb: '2E75B6' } }, border: bdr() };
    const alt = { font: { sz: 10 }, fill: { fgColor: { rgb: 'EBF3FB' } }, border: bdr(), alignment: { horizontal: 'right' } };

    function sc(r, c, v, style, numFmt) {
      const addr = XLSX.utils.encode_cell({ r: r - 1, c: c - 1 });
      ws[addr] = { v, t: typeof v === 'number' ? 'n' : 's', s: style };
      if (numFmt) ws[addr].z = numFmt;
    }
    function sf(r, c, formula, style) {
      const addr = XLSX.utils.encode_cell({ r: r - 1, c: c - 1 });
      ws[addr] = { f: formula.replace('=', ''), t: 'n', s: style };
    }

    // Row 3 — Headers
    ['', 'Period', 'Sport Payin (€)', 'Sport Gross (€)', 'Deposit (€)', 'Margin'].forEach((h, i) => {
      if (h) sc(3, i + 1, h, hdr);
    });
    merges.push({ s: { r: 2, c: 0 }, e: { r: 2, c: 0 } });

    // Row 4 — Yesterday
    sc(4, 2, 'Yesterday', lbl);
    sc(4, 3, kpis.sport_payin_yesterday, val, '#,##0');
    sc(4, 4, kpis.sport_gross_yesterday, val, '#,##0');
    sc(4, 5, kpis.deposit_yesterday,     val, '#,##0');
    sf(4, 6, '=D4/C4', { ...fml, z: '0.00%' });

    // Row 6 — MTD
    sc(6, 2, 'MTD', lbl);
    sc(6, 3, kpis.sport_payin_mtd,   val, '#,##0');
    sc(6, 4, kpis.sport_gross_mtd,   val, '#,##0');
    sc(6, 5, kpis.deposit_mtd,       val, '#,##0');
    sf(6, 6, '=D6/C6', { ...fml, z: '0.00%' });

    // Row 8 — Target
    sc(8, 2, 'TARGET (149.044)', lbl);
    sf(8, 3, '=D6/149044', { ...fml, z: '0.00%' });

    // Row 10 — Players header
    sc(10, 2, 'Players',    hdr); sc(10, 3, 'Yesterday', hdr); sc(10, 4, 'MTD', hdr);
    merges.push({ s:{r:9,c:1}, e:{r:9,c:1} });

    sc(11, 2, 'Registered players',       lbl);
    sc(11, 3, kpis.registered_yesterday,  val, '#,##0');
    sc(11, 4, kpis.registered_mtd,        val, '#,##0');

    sc(12, 2, 'Total active players SPORT', lbl);
    sc(12, 3, kpis.active_sport_yesterday,  val, '#,##0');
    sc(12, 4, kpis.active_sport_mtd,        val, '#,##0');

    // Row 14 — Bonus Conversion header
    sc(14, 2, 'Bonus Conversion New Players', hdr); sc(14, 3, 'Bonus Payin', hdr); sc(14, 4, 'Bonus Payout', hdr); sc(14, 5, 'Conversion', hdr);
    merges.push({ s:{r:13,c:1}, e:{r:13,c:1} });

    sc(15, 2, 'Yesterday',                      lbl);
    sc(15, 3, kpis.bonus_new_payin_yesterday,   val, '#,##0');
    sc(15, 4, kpis.bonus_new_payout_yesterday,  val, '#,##0');
    sf(15, 5, '=IFERROR(D15/C15,0)', { ...fml, z: '0.00%' });

    sc(16, 2, 'MTD',                    lbl);
    sc(16, 3, kpis.bonus_new_payin_mtd, val, '#,##0');
    sc(16, 4, kpis.bonus_new_payout_mtd,val, '#,##0');
    sf(16, 5, '=IFERROR(D16/C16,0)', { ...fml, z: '0.00%' });

    // Row 18 — Total Bonus
    sc(18, 2, 'Total Bonus', hdr); sc(18, 3, 'Bonus Payin', hdr); sc(18, 4, 'Bonus Payout', hdr); sc(18, 5, 'Conversion', hdr);
    merges.push({ s:{r:17,c:1}, e:{r:17,c:1} });

    sc(19, 2, 'MTD',                      lbl);
    sc(19, 3, kpis.bonus_total_payin_mtd, val, '#,##0');
    sc(19, 4, kpis.bonus_total_payout_mtd,val, '#,##0');
    sf(19, 5, '=IFERROR(D19/C19,0)', { ...fml, z: '0.00%' });

    // Row 21 — Top Losses header
    sc(21, 2, 'Players with the biggest losses for the previous day', sec);
    merges.push({ s:{r:20,c:1}, e:{r:20,c:5} });
    sc(22, 2, 'Account ID', hdr); sc(22, 3, 'Name', hdr); sc(22, 4, 'Gross (€)', hdr); sc(22, 5, 'Player Type', hdr); sc(22, 6, 'Bets', hdr);

    losses.forEach((p, i) => {
      const r = 23 + i;
      const s = i % 2 === 0 ? alt : val;
      sc(r, 2, p.id,                    s);
      sc(r, 3, p.name,                  { ...s, alignment: { horizontal: 'left' } });
      sc(r, 4, Math.round(p.gross * 100) / 100, s, '#,##0.00');
      sc(r, 5, p.type,                  s);
      sc(r, 6, p.bets,                  s, '#,##0');
    });

    // Row 29 — Top Wins header
    sc(29, 2, 'Players with the biggest wins for the previous day', sec);
    merges.push({ s:{r:28,c:1}, e:{r:28,c:5} });
    sc(30, 2, 'Account ID', hdr); sc(30, 3, 'Name', hdr); sc(30, 4, 'Gross (€)', hdr); sc(30, 5, 'Player Type', hdr); sc(30, 6, 'Bets', hdr);

    wins.forEach((p, i) => {
      const r = 31 + i;
      const s = i % 2 === 0 ? alt : val;
      sc(r, 2, p.id,                    s);
      sc(r, 3, p.name,                  { ...s, alignment: { horizontal: 'left' } });
      sc(r, 4, Math.round(p.gross * 100) / 100, s, '#,##0.00');
      sc(r, 5, p.type,                  s);
      sc(r, 6, p.bets,                  s, '#,##0');
    });

    ws['!cols'] = [{ wch: 3 }, { wch: 32 }, { wch: 16 }, { wch: 16 }, { wch: 14 }, { wch: 12 }, { wch: 10 }];
    ws['!merges'] = merges;
    ws['!ref'] = XLSX.utils.encode_range({ s: { r: 0, c: 0 }, e: { r: 36, c: 6 } });
    XLSX.utils.book_append_sheet(wb, ws, 'CM Daily Report');

    return wb;
  };

  function bdr() {
    const t = { style: 'thin', color: { rgb: 'BFBFBF' } };
    return { top: t, bottom: t, left: t, right: t };
  }

  // ── Pipeline ─────────────────────────────────────
  const runPipeline = async () => {
    setRunning(true); setLogs([]); setResult(null);
    try {
      log('📂 Chargement des fichiers CSV...');
      const [t1, t2, a1, a2] = await Promise.all([
        loadCSV(files.t1), loadCSV(files.t2),
        loadCSV(files.a1), loadCSV(files.a2),
      ]);
      log(`  ✓ Transactions yesterday : ${t1.length} lignes`, 'info');
      log(`  ✓ Transactions MTD       : ${t2.length} lignes`, 'info');
      log(`  ✓ Accounts yesterday     : ${a1.length} joueurs`, 'info');
      log(`  ✓ Accounts MTD           : ${a2.length} joueurs`, 'info');

      log('🔢 Calcul des KPIs...');
      const kpis = computeKPIs(t1, t2, a1, a2);
      log(`  Sport Payin Yesterday : ${kpis.sport_payin_yesterday.toLocaleString()} €`, 'info');
      log(`  Sport Gross Yesterday : ${kpis.sport_gross_yesterday.toLocaleString()} €`, 'info');
      log(`  Registered Yesterday  : ${kpis.registered_yesterday}`, 'info');

      log('🏆 Calcul top joueurs...');
      const { losses, wins } = topPlayers(t1);
      log(`  ${losses.length} top pertes · ${wins.length} top gains`, 'info');

      log('📊 Génération du rapport Excel...');
      const wb = buildReport(kpis, losses, wins);
      const yesterday = new Date(); yesterday.setDate(yesterday.getDate() - 1);
      const dd = String(yesterday.getDate()).padStart(2,'0');
      const mm = String(yesterday.getMonth()+1).padStart(2,'0');
      const yy = String(yesterday.getFullYear()).slice(-2);
      const filename = `CM_DAILY_REPORT_${dd}${mm}${yy}.xlsx`;
      const blob = new Blob([XLSX.write(wb, { bookType: 'xlsx', type: 'array' })], { type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet' });

      setResult({ blob, filename, kpis, losses, wins });
      log(`✅ Rapport généré : ${filename}`, 'info');
    } catch (err) {
      log(`❌ Erreur : ${err.message}`, 'error');
    }
    setRunning(false);
  };

  const allFilesReady = files.t1 && files.t2 && files.a1 && files.a2;

  const FILE_DEFS = [
    { key: 't1', label: 'PLAYER_TRANSACTION_REPORT_1.csv', sub: 'Transactions J-1', icon: '📋', color: 'var(--sport)' },
    { key: 't2', label: 'PLAYER_TRANSACTION_REPORT_2.csv', sub: 'Transactions MTD',  icon: '📋', color: 'var(--deposit)' },
    { key: 'a1', label: 'player_account_report_1.csv',     sub: 'Comptes J-1',       icon: '👤', color: 'var(--sport)' },
    { key: 'a2', label: 'player_account_report_2.csv',     sub: 'Comptes MTD',        icon: '👤', color: 'var(--deposit)' },
  ];

  return (
    <div>
      {/* Upload grid */}
      <p className="section-label">Upload des 4 fichiers CSV</p>
      <div style={{ display: 'grid', gridTemplateColumns: '1fr 1fr', gap: '1rem', marginBottom: '1.5rem' }}>
        {FILE_DEFS.map(({ key, label, sub, icon, color }) => (
          <div key={key}
            onClick={() => refs[key].current.click()}
            onDragOver={e => e.preventDefault()}
            onDrop={e => { e.preventDefault(); setFile(key, e.dataTransfer.files[0]); }}
            style={{
              border: `2px dashed ${files[key] ? color : 'var(--border2)'}`,
              borderRadius: 'var(--radius-lg)', padding: '1.5rem', textAlign: 'center',
              cursor: 'pointer', background: files[key] ? `rgba(14,165,233,0.05)` : 'var(--bg3)',
              transition: 'all 0.2s',
            }}>
            <input ref={refs[key]} type="file" accept=".csv" hidden onChange={e => setFile(key, e.target.files[0])} />
            <div style={{ fontSize: 24, marginBottom: 8 }}>{icon}</div>
            <div style={{ fontFamily: 'JetBrains Mono, monospace', fontSize: 11, color: 'var(--text)', fontWeight: 600, marginBottom: 4 }}>{label}</div>
            <div style={{ fontSize: 11, color: 'var(--muted)' }}>{sub}</div>
            {files[key] && (
              <div style={{ marginTop: 8, fontSize: 11, color: color, fontWeight: 600 }}>✓ {files[key].name}</div>
            )}
          </div>
        ))}
      </div>

      <button onClick={runPipeline} disabled={!allFilesReady || running}
        className="btn primary" style={{ fontSize: 14, padding: '11px 28px', marginBottom: '1.5rem' }}>
        {running ? '⟳ Génération…' : '▶  Générer CM Daily Report'}
      </button>

      {/* Logs */}
      {logs.length > 0 && (
        <>
          <p className="section-label" style={{ marginTop: '1rem' }}>Console</p>
          <div className="log-console" style={{ marginBottom: '1.5rem' }}>
            {logs.map((l, i) => (
              <span key={i} className={`log-line ${l.type}`}>
                <span style={{ color: '#475569', marginRight: 8 }}>[{l.t}]</span>{l.msg}{'\n'}
              </span>
            ))}
          </div>
        </>
      )}

      {/* Results */}
      {result && (
        <div>
          {/* KPI Preview */}
          <p className="section-label">Aperçu des KPIs</p>
          <div className="kpi-banner" style={{ marginBottom: '1.5rem' }}>
            {[
              { label: 'Sport Payin Yesterday', value: `${result.kpis.sport_payin_yesterday.toLocaleString()} €`, color: 'var(--sport)' },
              { label: 'Sport Gross Yesterday', value: `${result.kpis.sport_gross_yesterday.toLocaleString()} €`, color: 'var(--accent)' },
              { label: 'Deposit Yesterday',     value: `${result.kpis.deposit_yesterday.toLocaleString()} €`,     color: 'var(--deposit)' },
              { label: 'Registered Yesterday',  value: result.kpis.registered_yesterday,  color: 'var(--text)' },
              { label: 'Active Sport Yest.',    value: result.kpis.active_sport_yesterday, color: 'var(--text)' },
              { label: 'Sport Payin MTD',       value: `${result.kpis.sport_payin_mtd.toLocaleString()} €`,      color: 'var(--sport)' },
              { label: 'Sport Gross MTD',       value: `${result.kpis.sport_gross_mtd.toLocaleString()} €`,      color: 'var(--accent)' },
              { label: 'Deposit MTD',           value: `${result.kpis.deposit_mtd.toLocaleString()} €`,          color: 'var(--deposit)' },
            ].map(k => (
              <div key={k.label} className="kpi-item">
                <div className="kpi-label">{k.label}</div>
                <div className="kpi-value" style={{ color: k.color, fontSize: 18 }}>{k.value}</div>
              </div>
            ))}
          </div>

          {/* Top 5 players */}
          <div style={{ display: 'grid', gridTemplateColumns: '1fr 1fr', gap: '1rem', marginBottom: '1.5rem' }}>
            {[
              { title: '🔴 Top 5 Pertes', players: result.losses, color: 'var(--danger)' },
              { title: '🟢 Top 5 Gains',  players: result.wins,   color: 'var(--deposit)' },
            ].map(({ title, players, color }) => (
              <div key={title} className="card" style={{ padding: 0, overflow: 'hidden' }}>
                <div style={{ padding: '10px 14px', background: 'var(--bg3)', borderBottom: '1px solid var(--border)', fontWeight: 600, fontSize: 13 }}>{title}</div>
                <table style={{ width: '100%', borderCollapse: 'collapse', fontSize: 12 }}>
                  <thead>
                    <tr>
                      {['ID', 'Nom', 'Gross (€)'].map(h => (
                        <th key={h} style={{ padding: '8px 12px', textAlign: 'left', color: 'var(--muted)', fontWeight: 600, fontSize: 10, textTransform: 'uppercase', borderBottom: '1px solid var(--border)' }}>{h}</th>
                      ))}
                    </tr>
                  </thead>
                  <tbody>
                    {players.map((p, i) => (
                      <tr key={i} style={{ background: i % 2 === 0 ? 'var(--bg2)' : 'var(--bg3)' }}>
                        <td style={{ padding: '8px 12px', fontFamily: 'JetBrains Mono, monospace' }}>{p.id}</td>
                        <td style={{ padding: '8px 12px' }}>{p.name || '—'}</td>
                        <td style={{ padding: '8px 12px', color, fontWeight: 700, fontFamily: 'JetBrains Mono, monospace' }}>{Math.round(p.gross).toLocaleString()}</td>
                      </tr>
                    ))}
                  </tbody>
                </table>
              </div>
            ))}
          </div>

          {/* Download */}
          <button onClick={() => saveAs(result.blob, result.filename)}
            className="btn success" style={{ fontSize: 14, padding: '11px 28px' }}>
            ⬇  Télécharger {result.filename}
          </button>
        </div>
      )}
    </div>
  );
}
