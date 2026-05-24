// src/pages/CMReportPage.js
import React, { useState, useRef } from 'react';
import XLSX from 'xlsx-js-style';
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

  const loadCSV = (file) => new Promise((res, rej) => {
    Papa.parse(file, {
      header: true, skipEmptyLines: true, dynamicTyping: true,
      complete: ({ data }) => res(data),
      error: rej,
    });
  });

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

  const topPlayers = (t1, n = 5) => {
    const rows = t1.map(r => ({
      id:    r['Account ID'],
      name:  ((r['First Name'] || '') + ' ' + (r['Last Name'] || '')).trim(),
      gross: Number(r['SportBetting Gross EUR']) || 0,
      type:  r['Player Type Risk'] || '',
      bets:  Number(r['Standard Payin Tickets']) || 0,
    }));
    const losses = [...rows].filter(r => r.gross > 0).sort((a, b) => b.gross - a.gross).slice(0, n);
    const wins   = [...rows].filter(r => r.gross < 0).sort((a, b) => a.gross - b.gross).slice(0, n);
    return { losses, wins };
  };

  const buildReport = (kpis, losses, wins) => {
    const wb = XLSX.utils.book_new();
    const ws = {};

    const cv = (addr, v, style, fmt) => {
      ws[addr] = { v: v === null || v === undefined ? '' : v, t: typeof v === 'number' ? 'n' : 's', s: style };
      if (fmt) ws[addr].z = fmt;
    };
    const cf = (addr, formula, style, fmt) => {
      ws[addr] = { f: formula, t: 'n', s: style };
      if (fmt) ws[addr].z = fmt;
    };

    const boldLg  = { font: { name: 'Calibri', bold: true,  sz: 15 }, alignment: { horizontal: 'center', vertical: 'center' } };
    const boldMd  = { font: { name: 'Calibri', bold: true,  sz: 14 }, alignment: { horizontal: 'center', vertical: 'center' } };
    const boldSm  = { font: { name: 'Calibri', bold: true,  sz: 12 }, alignment: { horizontal: 'center', vertical: 'center' } };
    const plainLg = { font: { name: 'Calibri', bold: false, sz: 15 }, alignment: { horizontal: 'center', vertical: 'center' } };
    const plainMd = { font: { name: 'Calibri', bold: false, sz: 12 }, alignment: { horizontal: 'center', vertical: 'center' } };
    const numFmt  = '#,##0';
    const pctFmt  = '0.00%';
    const redNote = { font: { name: 'Calibri', bold: false, sz: 10, color: { rgb: 'FF0000' } }, alignment: { horizontal: 'left' } };
    const redHdr  = { font: { name: 'Calibri', bold: true,  sz: 12, color: { rgb: 'FFFFFF' } }, fill: { fgColor: { rgb: 'B50000' } }, alignment: { horizontal: 'center', vertical: 'center' } };
    const redCol  = { font: { name: 'Calibri', bold: true,  sz: 14, color: { rgb: 'FFFFFF' } }, fill: { fgColor: { rgb: 'B50000' } }, alignment: { horizontal: 'center', vertical: 'center' } };
    const pinkSt  = { font: { name: 'Calibri', bold: true,  sz: 14 }, fill: { fgColor: { rgb: 'F8D3DB' } }, alignment: { horizontal: 'center', vertical: 'center' } };

    // Row 3
    cv('B3','Period',boldLg); cv('C3','Sport Payin',boldLg); cv('D3','Sport Gross',boldLg); cv('E3','Deposit',boldLg); cv('F3','Margin',boldLg);
    // Row 4
    cv('B4','Yesterday',boldLg);
    cv('C4',kpis.sport_payin_yesterday,boldLg,numFmt);
    cv('D4',kpis.sport_gross_yesterday,boldLg,numFmt);
    cv('E4',kpis.deposit_yesterday,boldLg,numFmt);
    cf('F4','D4/C4',boldLg,pctFmt);
    // Row 5
    ['B5','C5','D5','E5','F5'].forEach(a=>cv(a,'',plainLg));
    // Row 6
    cv('B6','MTD',boldLg);
    cv('C6',kpis.sport_payin_mtd,boldLg,numFmt);
    cv('D6',kpis.sport_gross_mtd,boldLg,numFmt);
    cv('E6',kpis.deposit_mtd,boldLg,numFmt);
    cf('F6','D6/C6',boldLg,pctFmt);
    // Row 7
    ['C7','D7','E7','F7'].forEach(a=>cv(a,'',plainMd));
    // Row 8
    cv('B8','TARGET (149.044)',boldLg);
    cf('C8','D6/149044',boldLg,pctFmt);
    ['D8','E8','F8'].forEach(a=>cv(a,'',plainMd));
    // Row 9
    cv('E9','',plainMd);
    // Row 10
    cv('B10','Players',boldLg); cv('C10','Yesterday',boldLg); cv('D10','MTD',boldLg);
    // Row 11
    cv('B11','Registered players',boldLg);
    cv('C11',kpis.registered_yesterday,boldLg,numFmt);
    cv('D11',kpis.registered_mtd,boldLg,numFmt);
    cv('E11','registered,actiity is not important',redNote);
    // Row 12
    cv('B12','Total active players SPORT',boldLg);
    cv('C12',kpis.active_sport_yesterday,boldLg,numFmt);
    cv('D12',kpis.active_sport_mtd,boldLg,numFmt);
    // Row 13
    ['C13','D13','E13','F13'].forEach(a=>cv(a,'',plainMd));
    // Row 14
    cv('B14','Bonus Conversion New Players',boldLg); cv('C14','Bonus Payin',boldLg); cv('D14','Bonus payout',boldLg); cv('E14','Conversion',boldLg); cv('F14','',plainMd);
    // Row 15
    cv('B15','Yesterday',boldLg);
    cv('C15',kpis.bonus_new_payin_yesterday,boldLg,numFmt);
    cv('D15',kpis.bonus_new_payout_yesterday,boldLg,numFmt);
    cf('E15','D15/C15',boldLg,pctFmt); cv('F15','',plainMd);
    // Row 16
    cv('B16','MTD',boldLg);
    cv('C16',kpis.bonus_new_payin_mtd,boldLg,numFmt);
    cv('D16',kpis.bonus_new_payout_mtd,boldLg,numFmt);
    cf('E16','D16/C16',boldLg,pctFmt); cv('F16','',plainMd);
    // Row 17
    ['B17','C17','D17','E17','F17'].forEach(a=>cv(a,'',plainLg));
    // Row 18
    cv('B18','Total Bonus',boldLg); cv('C18','Bonus Payin',boldLg); cv('D18','Bonus payout',boldLg); cv('E18','Conversion',boldLg); cv('F18','',plainMd);
    // Row 19
    cv('B19','MTD',boldLg);
    cv('C19',kpis.bonus_total_payin_mtd,boldLg,numFmt);
    cv('D19',kpis.bonus_total_payout_mtd,boldLg,numFmt);
    cf('E19','D19/C19',boldLg,pctFmt); cv('F19','',plainMd);
    // Row 21 — losses header (merged)
    cv('B21','Players with the biggest losses for the previous day',boldSm);
    ['C21','D21','E21','F21'].forEach(a=>cv(a,'',boldSm));
    // Row 22
    cv('B22','Account ID',boldMd); cv('C22','Name',boldMd); cv('D22','Gross',boldMd); cv('E22','Player type',boldMd); cv('F22','Bets:',boldMd);
    // Rows 23–27
    for (let i = 0; i < 5; i++) {
      const r = 23+i, p = losses[i];
      cv(`B${r}`, p ? p.id   : '', boldMd);
      cv(`C${r}`, p ? p.name : '', boldMd);
      cv(`D${r}`, p ? Math.round(p.gross*100)/100 : '', boldMd, '#,##0.00');
      cv(`E${r}`, p ? p.type : '', boldMd);
      cv(`F${r}`, p ? p.bets : '', boldMd, numFmt);
    }
    // Row 28
    ['B28','C28','D28','E28','F28'].forEach(a=>cv(a,'',{font:{name:'Calibri',bold:true,sz:13},alignment:{horizontal:'center'}}));
    // Row 29 — wins header (red bg, merged)
    cv('B29','Players with the biggest wins for the previous day',redHdr);
    ['C29','D29','E29','F29'].forEach(a=>cv(a,'',redHdr));
    // Row 30
    cv('B30','Account ID',redCol); cv('C30','Name',redCol); cv('D30','Gross',redCol); cv('E30','Player type',redCol); cv('F30','Bets:',redCol);
    // Rows 31–35
    for (let i = 0; i < 5; i++) {
      const r = 31+i, p = wins[i];
      cv(`B${r}`, p ? p.id   : '', pinkSt);
      cv(`C${r}`, p ? p.name : '', pinkSt);
      cv(`D${r}`, p ? Math.round(p.gross*100)/100 : '', pinkSt, '#,##0.00');
      cv(`E${r}`, p ? p.type : '', pinkSt);
      cv(`F${r}`, p ? p.bets : '', pinkSt, numFmt);
    }

    ws['!merges'] = [
      { s:{r:20,c:1}, e:{r:20,c:5} },
      { s:{r:28,c:1}, e:{r:28,c:5} },
    ];
    ws['!cols'] = [{wch:5.33},{wch:33.5},{wch:21.5},{wch:15.66},{wch:17.5},{wch:76.33},{wch:10.83}];
    const hts = {3:20,4:20,5:8,6:17,8:21,10:20,11:20,12:20,14:20,15:20,16:20,17:20,18:20,19:21,21:20,22:19,23:19,24:19,25:19,26:19,27:19,28:17,30:19,31:19,32:19,33:19,34:19,35:19};
    ws['!rows'] = [];
    for (let i=0;i<=35;i++) ws['!rows'].push(hts[i+1]?{hpt:hts[i+1]}:{});
    ws['!ref'] = 'A1:G36';
    XLSX.utils.book_append_sheet(wb, ws, 'Sheet1');
    return wb;
  };

  const runPipeline = async () => {
    setRunning(true); setLogs([]); setResult(null);
    try {
      log('📂 Chargement des fichiers CSV...');
      const [t1, t2, a1, a2] = await Promise.all([loadCSV(files.t1), loadCSV(files.t2), loadCSV(files.a1), loadCSV(files.a2)]);
      log(`  ✓ Transactions yesterday : ${t1.length} lignes`, 'info');
      log(`  ✓ Transactions MTD       : ${t2.length} lignes`, 'info');
      log(`  ✓ Accounts yesterday     : ${a1.length} joueurs`, 'info');
      log(`  ✓ Accounts MTD           : ${a2.length} joueurs`, 'info');
      log('🔢 Calcul des KPIs...');
      const kpis = computeKPIs(t1, t2, a1, a2);
      log(`  Sport Payin Yesterday : ${kpis.sport_payin_yesterday.toLocaleString()} €`, 'info');
      log('🏆 Calcul top joueurs...');
      const { losses, wins } = topPlayers(t1);
      log('📊 Génération du rapport Excel...');
      const wb = buildReport(kpis, losses, wins);
      const yesterday = new Date(); yesterday.setDate(yesterday.getDate()-1);
      const dd = String(yesterday.getDate()).padStart(2,'0');
      const mm = String(yesterday.getMonth()+1).padStart(2,'0');
      const yy = String(yesterday.getFullYear()).slice(-2);
      const filename = `CM_DAILY_REPORT_${dd}${mm}${yy}.xlsx`;
      const blob = new Blob([XLSX.write(wb,{bookType:'xlsx',type:'array'})],{type:'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'});
      setResult({ blob, filename, kpis, losses, wins });
      log(`✅ Rapport généré : ${filename}`, 'info');
    } catch(err) { log(`❌ Erreur : ${err.message}`, 'error'); }
    setRunning(false);
  };

  const FILE_DEFS = [
    { key:'t1', label:'PLAYER_TRANSACTION_REPORT_1.csv', sub:'Transactions J-1', icon:'📋' },
    { key:'t2', label:'PLAYER_TRANSACTION_REPORT_2.csv', sub:'Transactions MTD',  icon:'📋' },
    { key:'a1', label:'player_account_report_1.csv',     sub:'Comptes J-1',       icon:'👤' },
    { key:'a2', label:'player_account_report_2.csv',     sub:'Comptes MTD',        icon:'👤' },
  ];

  return (
    <div>
      <p className="section-label">Upload des 4 fichiers CSV</p>
      <div style={{ display:'grid', gridTemplateColumns:'1fr 1fr', gap:'1rem', marginBottom:'1.5rem' }}>
        {FILE_DEFS.map(({ key, label, sub, icon }) => (
          <div key={key} onClick={() => refs[key].current.click()}
            onDragOver={e => e.preventDefault()}
            onDrop={e => { e.preventDefault(); setFile(key, e.dataTransfer.files[0]); }}
            className={`drop-zone ${files[key] ? 'has-file' : ''}`}>
            <input ref={refs[key]} type="file" accept=".csv" hidden onChange={e => setFile(key, e.target.files[0])} />
            <div className="drop-icon">{icon}</div>
            <div className="drop-label">{label}</div>
            <div className="drop-sub">{sub}</div>
            {files[key] && <div className="file-chip">✓ {files[key].name}</div>}
          </div>
        ))}
      </div>

      <button onClick={runPipeline} disabled={!files.t1||!files.t2||!files.a1||!files.a2||running}
        className="btn primary" style={{ fontSize:14, padding:'11px 28px', marginBottom:'1.5rem' }}>
        {running ? '⟳ Génération…' : '▶  Générer Daily Sport Report'}
      </button>

      {logs.length > 0 && (
        <>
          <p className="section-label" style={{ marginTop:'1rem' }}>Console</p>
          <div className="log-console" style={{ marginBottom:'1.5rem' }}>
            {logs.map((l,i) => (
              <span key={i} className={`log-line ${l.type}`} style={{ display:'block' }}>
                <span style={{ color:'#475569', marginRight:8 }}>[{l.t}]</span>{l.msg}
              </span>
            ))}
          </div>
        </>
      )}

      {result && (
        <div>
          <div className="kpi-banner" style={{ marginBottom:'1.5rem' }}>
            {[
              {label:'Sport Payin Yesterday', value:`${result.kpis.sport_payin_yesterday.toLocaleString()} €`, color:'var(--sport)'},
              {label:'Sport Gross Yesterday', value:`${result.kpis.sport_gross_yesterday.toLocaleString()} €`, color:'var(--accent)'},
              {label:'Deposit Yesterday',     value:`${result.kpis.deposit_yesterday.toLocaleString()} €`,     color:'var(--deposit)'},
              {label:'Registered Yesterday',  value:result.kpis.registered_yesterday,  color:'var(--text)'},
              {label:'Active Sport Yest.',    value:result.kpis.active_sport_yesterday, color:'var(--text)'},
              {label:'Sport Payin MTD',       value:`${result.kpis.sport_payin_mtd.toLocaleString()} €`, color:'var(--sport)'},
              {label:'Sport Gross MTD',       value:`${result.kpis.sport_gross_mtd.toLocaleString()} €`, color:'var(--accent)'},
              {label:'Deposit MTD',           value:`${result.kpis.deposit_mtd.toLocaleString()} €`,     color:'var(--deposit)'},
            ].map(k => (
              <div key={k.label} className="kpi-item">
                <div className="kpi-label">{k.label}</div>
                <div className="kpi-value" style={{ color:k.color, fontSize:18 }}>{k.value}</div>
              </div>
            ))}
          </div>
          <button onClick={() => saveAs(result.blob, result.filename)}
            className="btn success" style={{ fontSize:14, padding:'11px 28px' }}>
            ⬇  Télécharger {result.filename}
          </button>
        </div>
      )}
    </div>
  );
}
