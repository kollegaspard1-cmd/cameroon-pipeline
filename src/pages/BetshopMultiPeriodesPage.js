// src/pages/BetshopMultiPeriodesPage.js
import React, { useState, useRef } from 'react';
import XLSX from 'xlsx-js-style';
import { saveAs } from 'file-saver';
import { useAppContext } from '../AppContext';

const C = {
  dark_blue:'1F3864', mid_blue:'2E75B6', green_med:'375623', green_bg:'E2EFDA',
  green_fg:'375623', red_med:'7B2C2C', red_bg:'FCE4D6', red_fg:'C55A11',
  gold_bg:'FFF2CC', gold_fg:'7F6000', gold_dark:'BF8F00',
  grey:'F2F2F2', white:'FFFFFF', light_blue:'DEEAF1',
};
const FMT_NB='#,##0', FMT_DIF='#,##0;[Red]-#,##0';

function bdr() { const t={style:'thin',color:{rgb:'BFBFBF'}}; return {top:t,bottom:t,left:t,right:t}; }
function fill(rgb) { return {patternType:'solid',fgColor:{rgb}}; }
function fnt(bold=false,sz=10,rgb='000000') { return {name:'Arial',bold,sz,color:{rgb}}; }
function aln(h='center',v='center',wrap=false) { return {horizontal:h,vertical:v,wrapText:wrap}; }
function sc(ws,r,c,v,s) { const a=XLSX.utils.encode_cell({r:r-1,c:c-1}); ws[a]={v,t:typeof v==='number'?'n':'s',s}; }
function scn(ws,r,c,v,s,z) { const a=XLSX.utils.encode_cell({r:r-1,c:c-1}); ws[a]={v,t:'n',s,z}; }
function mg(ws,r1,c1,r2,c2) { if(!ws['!merges'])ws['!merges']=[]; ws['!merges'].push({s:{r:r1-1,c:c1-1},e:{r:r2-1,c:c2-1}}); }
function hS(bg,fg='FFFFFF',sz=10) { return {font:fnt(true,sz,fg),fill:fill(bg),border:bdr(),alignment:aln('center','center',true)}; }
function dS(bg='FFFFFF',bold=false,fg='000000',h='right') { return {font:fnt(bold,10,fg),fill:fill(bg),border:bdr(),alignment:aln(h,'center')}; }
function tS(bg) { return {font:fnt(true,13,'FFFFFF'),fill:fill(bg),alignment:aln('center','center')}; }
function sS(bg) { return {font:fnt(true,11,'FFFFFF'),fill:fill(bg),border:bdr(),alignment:aln('center','center')}; }

function detectPeriode(n, moisList) {
  if (n === 2) return 'Analyse Bimestrielle';
  if (n === 3) {
    const m0 = moisList[0].toLowerCase();
    if (m0.includes('janv') || m0.includes('jan')) return 'Analyse Trimestrielle — T1';
    if (m0.includes('avr')  || m0.includes('apr')) return 'Analyse Trimestrielle — T2';
    if (m0.includes('juil') || m0.includes('jul')) return 'Analyse Trimestrielle — T3';
    if (m0.includes('oct'))                         return 'Analyse Trimestrielle — T4';
    return 'Analyse Trimestrielle';
  }
  if (n === 6) return moisList[0].toLowerCase().includes('janv') ? 'Analyse Semestrielle — S1' : 'Analyse Semestrielle — S2';
  if (n === 12) return 'Analyse Annuelle';
  return `Analyse Personnalisée (${n} mois)`;
}

function detectLabel(n, moisList) {
  const first = moisList[0].split(' ').pop();
  const last  = moisList[moisList.length-1];
  if (n === 12) return `Annuel_${first}`;
  if (n === 6)  return moisList[0].toLowerCase().includes('janv') ? `S1_${first}` : `S2_${first}`;
  if (n === 3) {
    const m = moisList[0].toLowerCase();
    const q = m.includes('janv')||m.includes('jan') ? 'T1' : m.includes('avr')||m.includes('apr') ? 'T2' : m.includes('juil')||m.includes('jul') ? 'T3' : 'T4';
    return `${q}_${first}`;
  }
  return `${moisList[0].split(' ')[0].slice(0,3)}_${last.replace(' ','_')}`;
}

function loadCampresj(rows, moisLabel) {
  return rows.map(row => {
    const n = k => parseFloat(row[k]) || 0;
    return {
      Shop: row['Shop'] || row['Betshop name'] || '',
      Tickets: n('Bet tickets MTD') || n('Tickets') || n('Bet Tickets MTD') || 0,
      Payin:   n('Payin Bet MTD')   || n('Payin')   || 0,
      GGR:     n('GGR Bet MTD')     || n('GGR')     || 0,
      Mois:    moisLabel,
    };
  }).filter(r => r.Shop);
}

function computeMulti(dfs, moisList) {
  const shops = [...new Set(dfs.flatMap(df => df.map(r => r.Shop)))].sort();
  const byShop = {};
  shops.forEach(shop => {
    byShop[shop] = moisList.map((m, i) => {
      const row = dfs[i].find(r => r.Shop === shop);
      return row ? { Tickets: row.Tickets, Payin: row.Payin, GGR: row.GGR } : { Tickets: 0, Payin: 0, GGR: 0 };
    });
  });
  const totals = moisList.map((m, i) => ({
    Tickets: dfs[i].reduce((s,r) => s+r.Tickets, 0),
    Payin:   dfs[i].reduce((s,r) => s+r.Payin,   0),
    GGR:     dfs[i].reduce((s,r) => s+r.GGR,     0),
  }));
  return { shops, byShop, totals };
}

// ── Sheet: Comparatif ─────────────────────────────
function buildComparatif(wb, moisList, data) {
  const ws = {};
  const n = moisList.length;
  const totalCols = 1 + n * 3;

  // Title
  sc(ws,2,1, `COMPARATIF CAMPRESJ — ${moisList[0]} → ${moisList[n-1]}`, tS(C.dark_blue));
  mg(ws,2,1,2,totalCols);

  // Month headers
  let col = 2;
  moisList.forEach(m => {
    sc(ws,3,col,m.toUpperCase().slice(0,3)+'.',hS(C.mid_blue,'FFFFFF',9));
    mg(ws,3,col,3,col+2);
    ['Tickets','Payin','GGR'].forEach((h,i) => sc(ws,4,col+i,h,hS('2E5F8A','FFFFFF',8)));
    col += 3;
  });
  sc(ws,3,1,'SALLE',hS(C.dark_blue));
  mg(ws,3,1,4,1);

  // Data rows
  data.shops.forEach((shop, si) => {
    const r = 5 + si;
    const bg = si%2===0 ? C.light_blue : C.white;
    sc(ws,r,1,shop,dS(bg,false,'000000','left'));
    let c = 2;
    data.byShop[shop].forEach(month => {
      scn(ws,r,c,   month.Tickets, dS(bg), FMT_NB);
      scn(ws,r,c+1, month.Payin,   dS(bg), FMT_NB);
      scn(ws,r,c+2, month.GGR,     dS(bg), FMT_NB);
      c += 3;
    });
  });

  // Total row
  const tr = 5 + data.shops.length;
  const tSt = { font:fnt(true,10,'FFFFFF'), fill:fill(C.dark_blue), border:bdr(), alignment:aln('right','center') };
  sc(ws,tr,1,'TOTAL',{...tSt,alignment:aln('left','center')});
  let c = 2;
  data.totals.forEach(t => {
    scn(ws,tr,c,  t.Tickets,tSt,FMT_NB);
    scn(ws,tr,c+1,t.Payin,  tSt,FMT_NB);
    scn(ws,tr,c+2,t.GGR,    tSt,FMT_NB);
    c += 3;
  });

  ws['!cols'] = [30, ...Array(n*3).fill(13)].map(w=>({wch:w}));
  ws['!ref'] = XLSX.utils.encode_range({s:{r:0,c:0},e:{r:tr,c:totalCols-1}});
  XLSX.utils.book_append_sheet(wb, ws, '① Comparatif');
}

// ── Sheet: Évolution ──────────────────────────────
function buildEvolution(wb, moisList, data) {
  const ws = {};
  const n = moisList.length;
  sc(ws,2,1,'ÉVOLUTION GLOBALE CAMPRESJ', tS(C.dark_blue));
  mg(ws,2,1,2,5);

  ['Mois','Tickets','Var vs M-1','Payin (FCFA)','GGR (FCFA)'].forEach((h,i) => sc(ws,4,i+1,h,hS(C.mid_blue)));

  data.totals.forEach((t,i) => {
    const r = 5 + i;
    const bg = i%2===0 ? C.light_blue : C.white;
    const prev = i > 0 ? data.totals[i-1].Tickets : null;
    const diff = prev !== null ? t.Tickets - prev : null;
    const isGood = diff !== null ? diff >= 0 : true;
    sc(ws,r,1, moisList[i], dS(bg,true,'000000','left'));
    scn(ws,r,2, t.Tickets, dS(bg), FMT_NB);
    if (diff !== null) {
      scn(ws,r,3, diff, {font:fnt(true,10,isGood?C.green_fg:C.red_fg),fill:fill(isGood?C.green_bg:C.red_bg),border:bdr(),alignment:aln('right','center')}, FMT_DIF);
    } else {
      sc(ws,r,3,'—',dS(bg,false,'888888','center'));
    }
    scn(ws,r,4, t.Payin, dS(bg), FMT_NB);
    scn(ws,r,5, t.GGR,   dS(bg), FMT_NB);
  });

  ws['!cols'] = [22,14,14,16,14].map(w=>({wch:w}));
  ws['!ref'] = XLSX.utils.encode_range({s:{r:0,c:0},e:{r:4+n,c:4}});
  XLSX.utils.book_append_sheet(wb, ws, '② Évolution globale');
}

// ── Sheet: Top 10 période ─────────────────────────
function buildTop10(wb, moisList, data) {
  const ws = {};
  const n = moisList.length;
  sc(ws,2,1,`TOP 10 — PÉRIODE (${moisList[0]} → ${moisList[n-1]})`, tS(C.dark_blue));
  mg(ws,2,1,2,4);

  sc(ws,4,1,'MEILLEURES PERFORMANCES — Total tickets sur la période', sS(C.green_med));
  mg(ws,4,1,4,4);
  ['Rang','Salle','Total Tickets','Moy/Mois'].forEach((h,i) => sc(ws,5,i+1,h,hS('4A7C3F')));

  const ranked = data.shops.map(shop => ({
    shop,
    total: data.byShop[shop].reduce((s,m)=>s+m.Tickets,0),
  })).sort((a,b)=>b.total-a.total).slice(0,10);

  ranked.forEach(({shop,total},i) => {
    const r = 6+i;
    const bg = i%2===0 ? C.green_bg : C.white;
    const medal = ['🥇','🥈','🥉'][i]||'';
    sc(ws,r,1,`${medal} ${i+1}`,{font:fnt(true,10,'FFFFFF'),fill:fill(C.green_med),border:bdr(),alignment:aln('center','center')});
    sc(ws,r,2,shop,dS(bg,i<3,'000000','left'));
    scn(ws,r,3,total,dS(bg,true),FMT_NB);
    scn(ws,r,4,Math.round(total/n),dS(bg),FMT_NB);
  });

  // Bottom 10
  let r = 18;
  sc(ws,r,1,'SALLES EN DIFFICULTÉS — Plus faibles totaux sur la période', sS(C.red_med));
  mg(ws,r,1,r,4); r++;
  ['Rang','Salle','Total Tickets','Mois actifs'].forEach((h,i) => sc(ws,r,i+1,h,hS('A04040'))); r++;

  const bottom = [...data.shops.map(shop=>({
    shop,
    total: data.byShop[shop].reduce((s,m)=>s+m.Tickets,0),
    active: data.byShop[shop].filter(m=>m.Tickets>0).length,
  }))].sort((a,b)=>a.total-b.total).slice(0,10);

  bottom.forEach(({shop,total,active},i) => {
    const bg = i%2===0 ? C.red_bg : C.white;
    sc(ws,r,1,`${i+1}`,{font:fnt(true,10,'FFFFFF'),fill:fill(C.red_med),border:bdr(),alignment:aln('center','center')});
    sc(ws,r,2,shop,dS(bg,false,'000000','left'));
    scn(ws,r,3,total,dS(bg,false,total===0?C.red_fg:'000000'),FMT_NB);
    scn(ws,r,4,active,dS(bg),FMT_NB);
    r++;
  });

  ws['!cols'] = [6,30,16,12].map(w=>({wch:w}));
  ws['!ref'] = XLSX.utils.encode_range({s:{r:0,c:0},e:{r:r,c:4}});
  XLSX.utils.book_append_sheet(wb, ws, '③ Top 10 — Période');
}

// ── Sheet: Tendances & Alertes ────────────────────
function buildTendances(wb, moisList, data) {
  const ws = {};
  const n = moisList.length;
  sc(ws,2,1,'TENDANCES & ALERTES', tS(C.dark_blue));
  mg(ws,2,1,2,5);

  // Progression générale
  const first = data.totals[0].Tickets;
  const last  = data.totals[n-1].Tickets;
  const diff  = last - first;
  const pct   = first !== 0 ? diff / first * 100 : 0;
  const isGood = diff >= 0;

  sc(ws,4,1,'📊 TENDANCE GÉNÉRALE DE LA PÉRIODE', sS(C.dark_blue));
  mg(ws,4,1,4,5);

  const items = [
    [`${isGood?'📈':'📉'} Tickets ${moisList[0]} → ${moisList[n-1]}`, `${first.toLocaleString()} → ${last.toLocaleString()}`, `${diff>=0?'+':''}${diff.toLocaleString()} tickets (${pct>0?'+':''}${pct.toFixed(1)}%)`, isGood ? C.green_bg : C.red_bg, isGood ? C.green_fg : C.red_fg],
    [`📊 Meilleur mois`, moisList[data.totals.indexOf(data.totals.reduce((a,b)=>b.Tickets>a.Tickets?b:a))], `${Math.max(...data.totals.map(t=>t.Tickets)).toLocaleString()} tickets`, C.green_bg, C.green_fg],
    [`📉 Mois le plus faible`, moisList[data.totals.indexOf(data.totals.reduce((a,b)=>b.Tickets<a.Tickets?b:a))], `${Math.min(...data.totals.map(t=>t.Tickets)).toLocaleString()} tickets`, C.red_bg, C.red_fg],
    [`🏪 Salles suivies`, `${data.shops.length} salles CAMPRESJ`, `Sur ${n} mois`, C.light_blue, '1F3864'],
  ];

  items.forEach(([label,valeur,detail,bg,fg],i) => {
    const r = 5+i;
    sc(ws,r,1,label,{font:fnt(true,10,'000000'),fill:fill(bg),border:bdr(),alignment:aln('left','center',true)});
    mg(ws,r,1,r,2);
    sc(ws,r,3,valeur,{font:fnt(true,10,fg),fill:fill(bg),border:bdr(),alignment:aln('center','center',true)});
    mg(ws,r,3,r,4);
    sc(ws,r,5,detail,{font:fnt(false,9,'555555',false),fill:fill(bg),border:bdr(),alignment:aln('left','center',true)});
  });

  // Alertes salles inactives
  let r = 11;
  sc(ws,r,1,'⚠️ ALERTES — Salles avec mois inactifs (0 tickets)', sS(C.red_med));
  mg(ws,r,1,r,5); r++;
  ['Salle','Mois inactifs','Dernier actif','Total tickets','Tendance'].forEach((h,i) => sc(ws,r,i+1,h,hS('A04040'))); r++;

  const alertes = data.shops.map(shop => {
    const months = data.byShop[shop];
    const inactifs = months.filter(m=>m.Tickets===0).length;
    const lastActive = months.map((m,i)=>m.Tickets>0?moisList[i]:null).filter(Boolean).pop() || '—';
    return { shop, inactifs, lastActive, total: months.reduce((s,m)=>s+m.Tickets,0) };
  }).filter(a=>a.inactifs>0).sort((a,b)=>b.inactifs-a.inactifs);

  if (alertes.length === 0) {
    sc(ws,r,1,'✅ Aucune salle inactive détectée sur la période.',{font:fnt(false,10,C.green_fg),fill:fill(C.green_bg),border:bdr(),alignment:aln('left','center')});
    mg(ws,r,1,r,5); r++;
  } else {
    alertes.slice(0,15).forEach((a,i) => {
      const bg = i%2===0 ? C.red_bg : C.white;
      const trend = a.inactifs >= n/2 ? '🔴 Critique' : a.inactifs >= 2 ? '🟠 Attention' : '🟡 Surveiller';
      sc(ws,r,1,a.shop,dS(bg,false,'000000','left'));
      scn(ws,r,2,a.inactifs,{font:fnt(true,10,C.red_fg),fill:fill(bg),border:bdr(),alignment:aln('center','center')},FMT_NB);
      sc(ws,r,3,a.lastActive,dS(bg,false,'555555','center'));
      scn(ws,r,4,a.total,dS(bg),FMT_NB);
      sc(ws,r,5,trend,dS(bg,true,'000000','center'));
      r++;
    });
  }

  ws['!cols'] = [30,14,16,14,14].map(w=>({wch:w}));
  ws['!ref'] = XLSX.utils.encode_range({s:{r:0,c:0},e:{r:r,c:4}});
  XLSX.utils.book_append_sheet(wb, ws, '④ Tendances & Alertes');
}

// ── Main Component ────────────────────────────────
const MOIS_OPTIONS = [
  'Janvier','Février','Mars','Avril','Mai','Juin',
  'Juillet','Août','Septembre','Octobre','Novembre','Décembre'
];

export default function BetshopMultiPeriodesPage({ onNavigate }) {
  const [annee, setAnnee]   = useState('2026');
  const { betshopForMulti } = useAppContext();
  const [files, setFiles]   = useState([]); // [{mois, file, data}]
  const [running, setRunning] = useState(false);
  const [logs, setLogs]     = useState([]);
  const [result, setResult] = useState(null);
  const fileRefs = useRef({});

  const log = (msg) => setLogs(p => [...p, msg]);

  const n = files.filter(f => f.file).length;
  const periode = n >= 2 ? detectPeriode(n, files.filter(f=>f.file).map(f=>f.mois)) : null;

  const toggleMois = (mois) => {
    setFiles(prev => {
      const exists = prev.find(f => f.mois === mois);
      if (exists) return prev.filter(f => f.mois !== mois);
      return [...prev, { mois: `${mois} ${annee}`, file: null, data: null }].sort((a,b) => {
        const ia = MOIS_OPTIONS.indexOf(a.mois.split(' ')[0]);
        const ib = MOIS_OPTIONS.indexOf(b.mois.split(' ')[0]);
        return ia - ib;
      });
    });
  };

  const isSelected = (mois) => files.some(f => f.mois.startsWith(mois));

  const handleFile = async (moisLabel, file) => {
    const reader = new FileReader();
    reader.onload = e => {
      const wb = XLSX.read(e.target.result, { type:'array' });
      const ws = wb.Sheets['CAMPRESJ'] || wb.Sheets[wb.SheetNames[0]];
      const rows = XLSX.utils.sheet_to_json(ws, { defval: 0 });
      setFiles(prev => prev.map(f => f.mois === moisLabel ? { ...f, file, data: rows } : f));
      log(`  ✓ ${moisLabel} : ${rows.length} salles chargées`);
    };
    reader.readAsArrayBuffer(file);
  };

  const runPipeline = async () => {
    setRunning(true); setLogs([]); setResult(null);
    try {
      const loaded = files.filter(f => f.file && f.data);
      if (loaded.length < 2) { log('❌ Minimum 2 fichiers requis'); setRunning(false); return; }

      log(`📊 Analyse ${periode}...`);
      log(`   ${loaded.length} mois : ${loaded.map(f=>f.mois).join(' → ')}`);

      const moisList = loaded.map(f => f.mois);
      const dfs = loaded.map(f => loadCampresj(f.data, f.mois));
      dfs.forEach((df,i) => log(`  ✓ ${moisList[i]} : ${df.length} salles | ${df.reduce((s,r)=>s+r.Tickets,0).toLocaleString()} tickets`));

      log('🔢 Calcul des métriques...');
      const data = computeMulti(dfs, moisList);
      log(`  ${data.shops.length} salles suivies`);

      log('🔨 Construction du rapport Excel...');
      const wb = XLSX.utils.book_new();
      buildComparatif(wb, moisList, data);  log('  ✅ ① Comparatif');
      buildEvolution(wb, moisList, data);   log('  ✅ ② Évolution');
      buildTop10(wb, moisList, data);       log('  ✅ ③ Top 10');
      buildTendances(wb, moisList, data);   log('  ✅ ④ Tendances & Alertes');

      const label = detectLabel(loaded.length, moisList);
      const filename = `Analyse_CAMPRESJ_${label}.xlsx`;
      const blob = new Blob([XLSX.write(wb, { bookType:'xlsx', type:'array' })], {
        type:'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'
      });
      setResult({ blob, filename });
      log(`\n✅ Terminé — ${filename}`);
    } catch(err) {
      log(`❌ Erreur : ${err.message}`);
    }
    setRunning(false);
  };

  return (
    <div>
      {/* Jumelage B: Betshop Marchés → Multi-Périodes */}
      {betshopForMulti && (
        <div style={{ background:'var(--accent-dim)', border:'1px solid var(--accent)', borderRadius:'var(--radius)', padding:'10px 16px', marginBottom:'1.25rem', display:'flex', alignItems:'center', gap:'1rem', flexWrap:'wrap' }}>
          <span style={{ fontSize:20 }}>🔗</span>
          <div style={{ flex:1 }}>
            <div style={{ fontSize:13, fontWeight:600, color:'var(--accent)' }}>Fichier CAMPRESJ disponible depuis Betshop Par Marché</div>
            <div style={{ fontSize:11, color:'var(--muted)', marginTop:2 }}>{betshopForMulti.file.name} · {betshopForMulti.mois}</div>
          </div>
          <button onClick={() => {
            const moisLabel = betshopForMulti.mois;
            const moisNom = moisLabel.split(' ')[0];
            const anneeN = moisLabel.split(' ')[1] || annee;
            setAnnee(anneeN);
            const fullMois = `${moisNom} ${anneeN}`;
            // Add the file to the files list
            setFiles(prev => {
              const exists = prev.find(f => f.mois === fullMois);
              if (exists) return prev.map(f => f.mois === fullMois ? {...f, file: betshopForMulti.file, data: null} : f);
              return [...prev, { mois: fullMois, file: betshopForMulti.file, data: null }];
            });
            // Load the data
            handleFile(fullMois, betshopForMulti.file);
          }}
            style={{ padding:'7px 14px', fontSize:12, background:'var(--accent)', color:'#fff', border:'none', borderRadius:'var(--radius)', cursor:'pointer', fontWeight:600, whiteSpace:'nowrap' }}>
            ➕ Pré-charger {betshopForMulti.mois}
          </button>
        </div>
      )}

  {/* Sélection mois */}
      <p className="section-label">1. Sélectionner les mois à analyser</p>
      <div className="card" style={{ marginBottom:'1.25rem' }}>
        <div style={{ display:'flex', alignItems:'center', gap:'1rem', marginBottom:'1rem' }}>
          <label style={{ fontSize:13, color:'var(--muted)', fontWeight:600 }}>Année :</label>
          <input value={annee} onChange={e=>setAnnee(e.target.value)} className="field-input" style={{ width:90 }} />
          {periode && (
            <span style={{ marginLeft:'auto', fontSize:12, fontWeight:700, padding:'4px 14px', borderRadius:100, background:'var(--accent-dim)', color:'var(--accent)' }}>
              {periode}
            </span>
          )}
        </div>
        <div style={{ display:'grid', gridTemplateColumns:'repeat(6,1fr)', gap:6 }}>
          {MOIS_OPTIONS.map(m => (
            <button key={m} onClick={() => toggleMois(m)}
              style={{
                padding:'7px 4px', borderRadius:6, fontSize:11, fontWeight:600, cursor:'pointer',
                border:`1px solid ${isSelected(m) ? 'var(--accent)' : 'var(--border)'}`,
                background: isSelected(m) ? 'var(--accent-dim)' : 'var(--bg3)',
                color: isSelected(m) ? 'var(--accent)' : 'var(--muted)',
              }}>
              {m.slice(0,3)}.
            </button>
          ))}
        </div>
      </div>

      {/* Upload fichiers */}
      {files.length > 0 && (
        <>
          <p className="section-label">2. Uploader les fichiers CAMPRESJ mensuels</p>
          <div style={{ display:'grid', gridTemplateColumns:'1fr 1fr 1fr', gap:'0.75rem', marginBottom:'1.25rem' }}>
            {files.map(f => (
              <div key={f.mois}
                onClick={() => { if (!fileRefs.current[f.mois]) fileRefs.current[f.mois] = document.createElement('input'); const inp = fileRefs.current[f.mois]; inp.type='file'; inp.accept='.xlsx,.xls'; inp.onchange=e=>handleFile(f.mois,e.target.files[0]); inp.click(); }}
                style={{
                  border: `2px dashed ${f.file ? 'var(--deposit)' : 'var(--border2)'}`,
                  borderRadius:'var(--radius)', padding:'1rem', textAlign:'center', cursor:'pointer',
                  background: f.file ? 'var(--deposit-bg)' : 'var(--bg3)',
                }}>
                <div style={{ fontSize:20, marginBottom:6 }}>{f.file ? '✅' : '📊'}</div>
                <div style={{ fontSize:12, fontWeight:700, fontFamily:'JetBrains Mono, monospace', color:'var(--text)' }}>{f.mois}</div>
                {f.file
                  ? <div style={{ fontSize:11, color:'var(--deposit)', marginTop:4, fontWeight:600 }}>✓ {f.data?.length} salles</div>
                  : <div style={{ fontSize:11, color:'var(--muted)', marginTop:4 }}>Cliquer pour uploader</div>
                }
              </div>
            ))}
          </div>
        </>
      )}

      <button className="btn primary" disabled={n < 2 || running} onClick={runPipeline}
        style={{ fontSize:14, padding:'11px 28px', marginBottom:'1.5rem' }}>
        {running ? '⟳ Génération...' : `▶  Générer l'analyse ${periode||''}`}
      </button>

      {/* Logs */}
      {logs.length > 0 && (
        <>
          <p className="section-label">Console</p>
          <div className="log-console" style={{ marginBottom:'1.5rem' }}>
            {logs.map((l,i) => <span key={i} className="log-line" style={{ display:'block' }}>{l}</span>)}
          </div>
        </>
      )}

      {/* Result */}
      {result && (
        <div className="card">
          <button className="btn success" onClick={() => saveAs(result.blob, result.filename)}
            style={{ fontSize:14, padding:'10px 24px' }}>
            ⬇  Télécharger {result.filename}
          </button>
        </div>
      )}
    </div>
  );
}
