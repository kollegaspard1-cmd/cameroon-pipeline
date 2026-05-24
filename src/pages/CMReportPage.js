// src/pages/CMReportPage.js
import React, { useState, useRef } from 'react';
import XLSX from 'xlsx-js-style';
import Papa from 'papaparse';
import { saveAs } from 'file-saver';

export default function CMReportPage() {
  const [files, setFiles] = useState({ t1: null, t2: null, a1: null, a2: null });
  const [running, setRunning] = useState(false);
  const [logs, setLogs]       = useState([]);
  const [result, setResult]   = useState(null);
  const refs = { t1: useRef(), t2: useRef(), a1: useRef(), a2: useRef() };

  const log = (msg, type = 'normal') =>
    setLogs(p => [...p, { msg, type, t: new Date().toLocaleTimeString() }]);
  const setFile = (key, f) => setFiles(p => ({ ...p, [key]: f }));

  const loadCSV = (file) => new Promise((res, rej) =>
    Papa.parse(file, { header: true, skipEmptyLines: true, dynamicTyping: true,
      complete: ({ data }) => res(data), error: rej }));

  const computeKPIs = (t1, t2, a1, a2) => {
    const yIds = new Set(a1.map(r => r['Account ID']));
    const mIds = new Set(a2.map(r => r['Account ID']));
    const t1y = t1.filter(r => yIds.has(r['Account ID']));
    const t2m = t2.filter(r => mIds.has(r['Account ID']));
    const sum = (arr, col) => arr.reduce((s, r) => s + (Number(r[col]) || 0), 0);
    return {
      sp_y: Math.round(sum(t1, 'Standard Payin Money EUR')),
      sg_y: Math.round(sum(t1, 'SportBetting Gross EUR')),
      dp_y: Math.round(sum(t1, 'Deposit Money EUR')),
      sp_m: Math.round(sum(t2, 'Standard Payin Money EUR')),
      sg_m: Math.round(sum(t2, 'SportBetting Gross EUR')),
      dp_m: Math.round(sum(t2, 'Deposit Money EUR')),
      reg_y: a1.length, reg_m: a2.length,
      act_y: t1.filter(r => (Number(r['SportBetting Gross']) || 0) !== 0).length,
      act_m: t2.filter(r => (Number(r['SportBetting Gross']) || 0) !== 0).length,
      bnp_y: Math.round(sum(t1y, 'Bonus Payin Money EUR')),
      bno_y: Math.round(sum(t1y, 'Bonus Payout Amount EUR')),
      bnp_m: Math.round(sum(t2m, 'Bonus Payin Money EUR')),
      bno_m: Math.round(sum(t2m, 'Bonus Payout Amount EUR')),
      btp_m: Math.round(sum(t2, 'Bonus Payin Money EUR')),
      bto_m: Math.round(sum(t2, 'Bonus Payout Amount EUR')),
    };
  };

  const topPlayers = (t1) => {
    const rows = t1.map(r => ({
      id: r['Account ID'],
      name: ((r['First Name']||'') + ' ' + (r['Last Name']||'')).trim(),
      gross: Number(r['SportBetting Gross EUR']) || 0,
      type: r['Player Type Risk'] || '',
      bets: Number(r['Standard Payin Tickets']) || 0,
    }));
    return {
      losses: [...rows].filter(r=>r.gross>0).sort((a,b)=>b.gross-a.gross).slice(0,5),
      wins:   [...rows].filter(r=>r.gross<0).sort((a,b)=>a.gross-b.gross).slice(0,5),
    };
  };

  const buildReport = (k, losses, wins) => {
    const wb = XLSX.utils.book_new();
    const ws = {};

    // ── Style helpers ──────────────────────────────
    const NONE = { patternType: 'solid', fgColor: { rgb: 'FFFFFF' } };
    const DARK = { patternType: 'solid', fgColor: { rgb: '1F3864' } };
    const RED  = { patternType: 'solid', fgColor: { rgb: 'B50000' } };
    const PINK = { patternType: 'solid', fgColor: { rgb: 'F8D3DB' } };

    const fBoldLg  = { name:'Calibri', bold:true,  sz:15, color:{rgb:'000000'} };
    const fBoldLgW = { name:'Calibri', bold:true,  sz:15, color:{rgb:'FFFFFF'} };
    const fBoldMd  = { name:'Calibri', bold:true,  sz:14, color:{rgb:'000000'} };
    const fBoldMdW = { name:'Calibri', bold:true,  sz:14, color:{rgb:'FFFFFF'} };
    const fBoldSm  = { name:'Calibri', bold:true,  sz:12, color:{rgb:'FFFFFF'} };
    const fRedSm   = { name:'Calibri', bold:false, sz:10, color:{rgb:'FF0000'} };

    const aCtr = { horizontal:'center', vertical:'center' };
    const aLft = { horizontal:'left',   vertical:'center' };

    const hair = { style:'hair' };
    const thin = { style:'thin' };

    // ── Cell setter ────────────────────────────────
    const cv = (addr, v, font, align, border, fill, fmt) => {
      ws[addr] = {
        v: v ?? '', t: typeof v === 'number' ? 'n' : 's',
        s: { font, alignment: align||aCtr, border: border||{}, fill: fill||NONE }
      };
      if (fmt) ws[addr].z = fmt;
    };
    const cf = (addr, formula, font, align, border, fill, fmt) => {
      ws[addr] = {
        f: formula, t: 'n',
        s: { font, alignment: align||aCtr, border: border||{}, fill: fill||NONE }
      };
      if (fmt) ws[addr].z = fmt;
    };

    // ── Row 3: Headers ─────────────────────────────
    cv('B3','Period',      fBoldLgW, aCtr, {right:thin},           DARK);
    cv('C3','Sport Payin', fBoldLgW, aCtr, {left:thin,right:thin}, DARK);
    cv('D3','Sport Gross', fBoldLgW, aCtr, {left:thin,right:thin}, DARK);
    cv('E3','Deposit',     fBoldLgW, aCtr, {left:thin},            DARK);
    cv('F3','Margin',      fBoldLgW, aCtr, {left:thin,right:hair}, DARK);

    // ── Row 4: Yesterday ───────────────────────────
    cv('B4','Yesterday', fBoldLgW, aCtr, {top:hair,bottom:hair,left:hair,right:hair}, DARK);
    cv('C4', k.sp_y,     fBoldLg,  aCtr, {bottom:hair,left:hair,right:hair}, NONE, '#,##0');
    cv('D4', k.sg_y,     fBoldLg,  aCtr, {bottom:hair,left:hair,right:hair}, NONE, '#,##0');
    cv('E4', k.dp_y,     fBoldLg,  aCtr, {bottom:hair,left:hair,right:hair}, NONE, '#,##0');
    cf('F4','D4/C4',     fBoldLg,  aCtr, {bottom:hair,left:hair,right:hair}, NONE, '0.00%');

    // ── Row 5: empty ───────────────────────────────
    cv('C5','', fBoldLg, aCtr, {bottom:hair,left:hair,right:hair}, NONE);

    // ── Row 6: MTD ─────────────────────────────────
    cv('B6','MTD',  fBoldLgW, aCtr, {top:hair,bottom:hair,left:hair,right:hair}, DARK);
    cv('C6', k.sp_m, fBoldLg, aCtr, {bottom:hair,left:hair,right:hair}, NONE, '#,##0');
    cv('D6', k.sg_m, fBoldLg, aCtr, {bottom:hair,left:hair,right:hair}, NONE, '#,##0');
    cv('E6', k.dp_m, fBoldLg, aCtr, {bottom:hair,left:hair,right:hair}, NONE, '#,##0');
    cf('F6','D6/C6', fBoldLg, aCtr, {top:hair,bottom:hair,left:hair,right:hair}, NONE, '0.00%');

    // ── Row 8: Target ──────────────────────────────
    cv('B8','TARGET (149.044)', fBoldLgW, aCtr, {top:hair,bottom:hair,left:hair,right:hair}, DARK);
    cf('C8','D6/149044',        fBoldLg,  aCtr, {top:hair,bottom:hair,left:hair,right:hair}, NONE, '0.00%');

    // ── Row 10: Players header ─────────────────────
    cv('B10','Players',   fBoldLgW, aCtr, {right:thin},           DARK);
    cv('C10','Yesterday', fBoldLgW, aCtr, {left:thin,right:thin}, DARK);
    cv('D10','MTD',       fBoldLgW, aCtr, {left:thin,right:thin}, DARK);

    // ── Row 11: Registered ─────────────────────────
    cv('B11','Registered players',   fBoldLgW, aCtr, {top:hair,bottom:hair,left:hair,right:hair}, DARK);
    cv('C11', k.reg_y,               fBoldLg,  aCtr, {bottom:hair,left:hair,right:hair},          NONE, '#,##0');
    cv('D11', k.reg_m,               fBoldLg,  aCtr, {bottom:hair,left:hair,right:hair},          NONE, '#,##0');
    cv('E11','registered, activity is not important', fRedSm, aLft, {}, NONE);

    // ── Row 12: Active ─────────────────────────────
    cv('B12','Total active players SPORT', fBoldLgW, aCtr, {top:hair,bottom:hair,left:hair,right:hair}, DARK);
    cv('C12', k.act_y, fBoldLg, aCtr, {bottom:hair,left:hair,right:hair}, NONE, '#,##0');
    cv('D12', k.act_m, fBoldLg, aCtr, {bottom:hair,left:hair,right:hair}, NONE, '#,##0');

    // ── Row 14: Bonus header ───────────────────────
    cv('B14','Bonus Conversion New Players', fBoldLgW, aCtr, {right:thin},           DARK);
    cv('C14','Bonus Payin',                  fBoldLgW, aCtr, {left:thin,right:thin}, DARK);
    cv('D14','Bonus payout',                 fBoldLgW, aCtr, {left:thin,right:thin}, DARK);
    cv('E14','Conversion',                   fBoldLgW, aCtr, {left:thin},            DARK);

    // ── Row 15: Bonus Yesterday ────────────────────
    cv('B15','Yesterday', fBoldLgW, aCtr, {left:thin,right:thin}, DARK);
    cv('C15', k.bnp_y,   fBoldLg,  aCtr, {bottom:hair,left:hair,right:hair}, NONE, '#,##0');
    cv('D15', k.bno_y,   fBoldLg,  aCtr, {bottom:hair,left:hair,right:hair}, NONE, '#,##0');
    cf('E15','IFERROR(D15/C15,0)', fBoldLg, aCtr, {bottom:hair,left:hair,right:hair}, NONE, '0.00%');

    // ── Row 16: Bonus MTD ──────────────────────────
    cv('B16','MTD',    fBoldLgW, aCtr, {left:thin,right:thin}, DARK);
    cv('C16', k.bnp_m, fBoldLg,  aCtr, {bottom:hair,left:hair,right:hair}, NONE, '#,##0');
    cv('D16', k.bno_m, fBoldLg,  aCtr, {bottom:hair,left:hair,right:hair}, NONE, '#,##0');
    cf('E16','IFERROR(D16/C16,0)', fBoldLg, aCtr, {bottom:hair,left:hair,right:hair}, NONE, '0.00%');

    // ── Row 18: Total Bonus header ─────────────────
    cv('B18','Total Bonus',  fBoldLgW, aCtr, {right:thin},           DARK);
    cv('C18','Bonus Payin',  fBoldLgW, aCtr, {left:thin,right:thin}, DARK);
    cv('D18','Bonus payout', fBoldLgW, aCtr, {left:thin,right:thin}, DARK);
    cv('E18','Conversion',   fBoldLgW, aCtr, {left:thin},            DARK);

    // ── Row 19: Total Bonus MTD ────────────────────
    cv('B19','MTD',    fBoldLgW, aCtr, {top:hair,bottom:hair,left:hair,right:hair}, DARK);
    cv('C19', k.btp_m, fBoldLg,  aCtr, {bottom:hair,left:hair,right:hair}, NONE, '#,##0');
    cv('D19', k.bto_m, fBoldLg,  aCtr, {bottom:hair,left:hair,right:hair}, NONE, '#,##0');
    cf('E19','IFERROR(D19/C19,0)', fBoldLg, aCtr, {bottom:hair,left:hair,right:hair}, NONE, '0.00%');

    // ── Row 21: Losses header (merged) ────────────
    ['B21','C21','D21','E21','F21'].forEach(a =>
      cv(a, a==='B21' ? 'Players with the biggest losses for the previous day' : '',
        fBoldSm, aCtr, {bottom:hair}, DARK));

    // ── Row 22: Losses col headers ─────────────────
    cv('B22','Account ID',  fBoldMdW, aCtr, {top:hair,right:hair},            DARK);
    cv('C22','Name',        fBoldMdW, aCtr, {top:hair,left:hair,right:hair},  DARK);
    cv('D22','Gross',       fBoldMdW, aCtr, {top:hair,left:hair,right:hair},  DARK);
    cv('E22','Player type', fBoldMdW, aCtr, {top:hair,left:hair,right:hair},  DARK);
    cv('F22','Bets:',       fBoldMdW, aCtr, {top:hair,left:hair},             DARK);

    // ── Rows 23-27: Losses data ────────────────────
    for (let i=0; i<5; i++) {
      const r = 23+i, p = losses[i];
      cv(`B${r}`, p?p.id:'',   fBoldMd, aCtr, {}, NONE);
      cv(`C${r}`, p?p.name:'', fBoldMd, aCtr, {}, NONE);
      cv(`D${r}`, p?Math.round(p.gross*100)/100:'', fBoldMd, aCtr, {bottom:hair,left:hair,right:hair}, NONE, '#,##0.00');
      cv(`E${r}`, p?p.type:'', fBoldMd, aCtr, {}, NONE);
      cv(`F${r}`, p?p.bets:'', fBoldMd, aCtr, {top:hair,bottom:hair,left:hair,right:hair}, NONE, '#,##0');
    }

    // ── Row 29: Wins header (merged, red bg) ───────
    ['B29','C29','D29','E29','F29'].forEach(a =>
      cv(a, a==='B29' ? 'Players with the biggest wins for the previous day' : '',
        fBoldSm, aCtr, {bottom:hair}, RED));

    // ── Row 30: Wins col headers (red bg) ──────────
    cv('B30','Account ID',  fBoldMdW, aCtr, {top:hair,right:hair},           RED);
    cv('C30','Name',        fBoldMdW, aCtr, {top:hair,left:hair,right:hair}, RED);
    cv('D30','Gross',       fBoldMdW, aCtr, {top:hair,left:hair,right:hair}, RED);
    cv('E30','Player type', fBoldMdW, aCtr, {top:hair,left:hair,right:hair}, RED);
    cv('F30','Bets:',       fBoldMdW, aCtr, {top:hair,left:hair},            RED);

    // ── Rows 31-35: Wins data (pink bg) ────────────
    for (let i=0; i<5; i++) {
      const r = 31+i, p = wins[i];
      cv(`B${r}`, p?p.id:'',   fBoldMd, aCtr, {}, PINK);
      cv(`C${r}`, p?p.name:'', fBoldMd, aCtr, {}, PINK);
      cv(`D${r}`, p?Math.round(p.gross*100)/100:'', fBoldMd, aCtr, {top:hair,bottom:hair,left:hair,right:hair}, PINK, '#,##0.00');
      cv(`E${r}`, p?p.type:'', fBoldMd, aCtr, {}, PINK);
      cv(`F${r}`, p?p.bets:'', fBoldMd, aCtr, {top:hair,bottom:hair,left:hair,right:hair}, PINK, '#,##0');
    }

    ws['!merges'] = [
      {s:{r:20,c:1}, e:{r:20,c:5}},
      {s:{r:28,c:1}, e:{r:28,c:5}},
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
      const [t1,t2,a1,a2] = await Promise.all([loadCSV(files.t1),loadCSV(files.t2),loadCSV(files.a1),loadCSV(files.a2)]);
      log(`  ✓ Transactions yesterday : ${t1.length} lignes`,'info');
      log(`  ✓ Transactions MTD       : ${t2.length} lignes`,'info');
      log(`  ✓ Accounts yesterday     : ${a1.length} joueurs`,'info');
      log(`  ✓ Accounts MTD           : ${a2.length} joueurs`,'info');
      log('🔢 Calcul des KPIs...');
      const k = computeKPIs(t1,t2,a1,a2);
      log(`  Sport Payin Yesterday : ${k.sp_y.toLocaleString()} €`,'info');
      log(`  Sport Payin MTD       : ${k.sp_m.toLocaleString()} €`,'info');
      log('🏆 Calcul top joueurs...');
      const {losses,wins} = topPlayers(t1);
      log('📊 Génération du rapport Excel...');
      const wb = buildReport(k,losses,wins);
      const d = new Date(); d.setDate(d.getDate()-1);
      const dd=String(d.getDate()).padStart(2,'0'), mm=String(d.getMonth()+1).padStart(2,'0'), yy=String(d.getFullYear()).slice(-2);
      const filename = `CM_DAILY_REPORT_${dd}${mm}${yy}.xlsx`;
      const blob = new Blob([XLSX.write(wb,{bookType:'xlsx',type:'array'})],
        {type:'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'});
      setResult({blob,filename,kpis:k,losses,wins});
      log(`✅ Rapport généré : ${filename}`,'info');
    } catch(err) { log(`❌ Erreur : ${err.message}`,'error'); }
    setRunning(false);
  };

  const FILE_DEFS = [
    {key:'t1', label:'Transaction Report J-1', sub:'Transactions d\'hier (Standard Payin Money EUR, SportBetting Gross EUR...)', icon:'📋'},
    {key:'t2', label:'Transaction Report MTD',  sub:'Transactions du mois en cours', icon:'📋'},
    {key:'a1', label:'Account Report J-1',      sub:'Nouveaux joueurs hier (colonne Account ID)', icon:'👤'},
    {key:'a2', label:'Account Report MTD',       sub:'Nouveaux joueurs du mois en cours', icon:'👤'},
  ];

  return (
    <div>
      <p className="section-label">Upload des 4 fichiers CSV</p>
      <div style={{display:'grid',gridTemplateColumns:'1fr 1fr',gap:'1rem',marginBottom:'1.5rem'}}>
        {FILE_DEFS.map(({key,label,sub,icon}) => (
          <div key={key} onClick={()=>refs[key].current.click()}
            onDragOver={e=>e.preventDefault()}
            onDrop={e=>{e.preventDefault();setFile(key,e.dataTransfer.files[0]);}}
            className={`drop-zone ${files[key]?'has-file':''}`}>
            <input ref={refs[key]} type="file" accept=".csv" hidden onChange={e=>setFile(key,e.target.files[0])}/>
            <div className="drop-icon">{icon}</div>
            <div className="drop-label">{label}</div>
            <div className="drop-sub">{sub}</div>
            {files[key]&&<div className="file-chip">✓ {files[key].name}</div>}
          </div>
        ))}
      </div>
      <button onClick={runPipeline} disabled={!files.t1||!files.t2||!files.a1||!files.a2||running}
        className="btn primary" style={{fontSize:14,padding:'11px 28px',marginBottom:'1.5rem'}}>
        {running?'⟳ Génération…':'▶  Générer Daily Sport Report'}
      </button>
      {logs.length>0&&(
        <>
          <p className="section-label" style={{marginTop:'1rem'}}>Console</p>
          <div className="log-console" style={{marginBottom:'1.5rem'}}>
            {logs.map((l,i)=>(
              <span key={i} className={`log-line ${l.type}`} style={{display:'block'}}>
                <span style={{color:'#475569',marginRight:8}}>[{l.t}]</span>{l.msg}
              </span>
            ))}
          </div>
        </>
      )}
      {result&&(
        <div>
          <div className="kpi-banner" style={{marginBottom:'1.5rem'}}>
            {[
              {label:'Sport Payin Yesterday',value:`${result.kpis.sp_y.toLocaleString()} €`,color:'var(--sport)'},
              {label:'Sport Gross Yesterday',value:`${result.kpis.sg_y.toLocaleString()} €`,color:'var(--accent)'},
              {label:'Deposit Yesterday',    value:`${result.kpis.dp_y.toLocaleString()} €`,color:'var(--deposit)'},
              {label:'Registered Yesterday', value:result.kpis.reg_y, color:'var(--text)'},
              {label:'Active Sport Yest.',   value:result.kpis.act_y, color:'var(--text)'},
              {label:'Sport Payin MTD',      value:`${result.kpis.sp_m.toLocaleString()} €`,color:'var(--sport)'},
              {label:'Sport Gross MTD',      value:`${result.kpis.sg_m.toLocaleString()} €`,color:'var(--accent)'},
              {label:'Deposit MTD',          value:`${result.kpis.dp_m.toLocaleString()} €`,color:'var(--deposit)'},
            ].map(kpi=>(
              <div key={kpi.label} className="kpi-item">
                <div className="kpi-label">{kpi.label}</div>
                <div className="kpi-value" style={{color:kpi.color,fontSize:18}}>{kpi.value}</div>
              </div>
            ))}
          </div>
          <button onClick={()=>saveAs(result.blob,result.filename)}
            className="btn success" style={{fontSize:14,padding:'11px 28px'}}>
            ⬇  Télécharger {result.filename}
          </button>
        </div>
      )}
    </div>
  );
}
