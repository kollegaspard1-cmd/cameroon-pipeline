// src/pages/RapportsPage.js
import React, { useState, useRef } from 'react';
import * as XLSX from 'xlsx';
import { saveAs } from 'file-saver';
import JSZip from 'jszip';

export default function RapportsPage() {
  const [betshopFile, setBetshopFile] = useState(null);
  const [objectifsFile, setObjectifsFile] = useState(null);
  const [recapFile, setRecapFile]     = useState(null);
  const [mois, setMois]               = useState('MARS 2026');
  const [running, setRunning]         = useState(false);
  const [logs, setLogs]               = useState([]);
  const [results, setResults]         = useState(null);
  const betRef  = useRef();
  const objRef  = useRef();
  const recRef  = useRef();

  const log = (msg) => setLogs(p => [...p, msg]);

  // ── Helpers couleurs ────────────────────────────
  const C = {
    DARK:'1F3864', BLUE:'1F4E79', MID:'2E5F8A', GREY:'404040',
    GREEN:'375623', RED:'7B2C2C', GOLD:'BF8F00', TTOTAL:'2E4B7A',
    OK:'E2EFDA', OKF:'375623', KO:'FCE4D6', KOF:'C55A11',
    OK15:'FFF2CC', OK15F:'7F6000', ALT:'EBF3FB', KPI:'F2F2F2',
  };

  function fill(rgb)   { return { patternType: 'solid', fgColor: { rgb } }; }
  function bdr()       { const t = { style: 'thin', color: { rgb: 'BFBFBF' } }; return { top: t, bottom: t, left: t, right: t }; }
  function fnt(bold=false, sz=10, rgb='000000') { return { name: 'Arial', bold, sz, color: { rgb } }; }
  function aln(h='center', v='center', wrap=false) { return { horizontal: h, vertical: v, wrapText: wrap }; }

  function sc(ws, r, c, v, bold=false, sz=10, fc='000000', bg=null, h='left', numFmt=null) {
    const addr = XLSX.utils.encode_cell({ r: r-1, c: c-1 });
    ws[addr] = { v, t: typeof v === 'number' ? 'n' : 's', s: { font: fnt(bold,sz,fc), fill: bg?fill(bg):undefined, alignment: aln(h), border: bdr() } };
    if (numFmt) ws[addr].z = numFmt;
  }

  // ── Chargement Betshop ──────────────────────────
  const loadBetshop = (file) => new Promise((res, rej) => {
    const reader = new FileReader();
    reader.onload = e => {
      try {
        const wb = XLSX.read(e.target.result, { type: 'array' });
        const ws = wb.Sheets[wb.SheetNames[0]];
        const raw = XLSX.utils.sheet_to_json(ws, { header: 1, defval: '' });
        let hr = null;
        for (let i = 0; i < raw.length; i++) {
          const vs = raw[i].map(v => String(v).trim().toUpperCase());
          if (vs.includes('SALLE')) {
            if (i + 1 < raw.length) {
              const nxt = raw[i+1].map(v => String(v).trim().toUpperCase());
              if (nxt.some(v => v.includes('MTD'))) { hr = i + 1; break; }
            }
            if (vs.some(v => v.includes('MTD'))) { hr = i; break; }
          }
        }
        if (hr === null) throw new Error('En-tête betshop introuvable');
        const headers = raw[hr];
        const rows = [];
        for (let i = hr + 1; i < raw.length; i++) {
          const row = raw[i];
          const salle = String(row[0] || '').trim();
          if (!salle || salle.toLowerCase().startsWith('total') || salle.toLowerCase().includes('campresj')) continue;
          rows.push({
            _k: salle.toUpperCase().replace(/ /g,''),
            Tickets_MTD: parseFloat(row[1]) || 0,
            Payin_MTD:   parseFloat(row[10]) || 0,
            GGR_MTD:     parseFloat(row[16]) || 0,
          });
        }
        log(`  Betshop chargé : ${rows.length} salles`);
        res(rows);
      } catch(err) { rej(err); }
    };
    reader.onerror = rej;
    reader.readAsArrayBuffer(file);
  });

  // ── Chargement Objectifs ────────────────────────
  const loadObjectifs = (file) => new Promise((res, rej) => {
    const reader = new FileReader();
    reader.onload = e => {
      try {
        const wb = XLSX.read(e.target.result, { type: 'array' });
        const ws = wb.Sheets[wb.SheetNames[0]];
        const raw = XLSX.utils.sheet_to_json(ws, { header: 1, defval: '' });
        let hr = null;
        for (let i = 0; i < raw.length; i++) {
          const vs = raw[i].map(v => String(v));
          if (vs.some(v => v.toLowerCase().includes('tickets objectif'))) { hr = i; break; }
        }
        if (hr === null) throw new Error("Colonne 'Tickets Objectif' introuvable");
        const headers = raw[hr].map(v => String(v).trim());
        const ocIdx  = headers.findIndex(h => h.toLowerCase().includes('tickets objectif'));
        const tc2025 = headers.findIndex(h => h.toLowerCase().includes('tickets') && !h.toLowerCase().includes('objectif') && h.includes('2025'));
        const rows = [];
        for (let i = hr + 1; i < raw.length; i++) {
          const row = raw[i];
          const salle = String(row[0] || '').trim();
          if (!salle || /salle|objectif|total/i.test(salle)) continue;
          let obj = parseFloat(row[ocIdx]) || 0;
          if (obj === 0 && tc2025 >= 0) obj = Math.round((parseFloat(row[tc2025]) || 0) * 1.15);
          rows.push({ _k: salle.toUpperCase().replace(/ /g,''), Tickets_Obj: obj });
        }
        log(`  Objectifs chargés : ${rows.length} salles`);
        res(rows);
      } catch(err) { rej(err); }
    };
    reader.onerror = rej;
    reader.readAsArrayBuffer(file);
  });

  // ── Chargement Recap ────────────────────────────
  const loadRecap = (file) => new Promise((res, rej) => {
    const reader = new FileReader();
    reader.onload = e => {
      try {
        const wb = XLSX.read(e.target.result, { type: 'array' });
        const result = {};
        wb.SheetNames.forEach(sheet => {
          const ws = wb.Sheets[sheet];
          const raw = XLSX.utils.sheet_to_json(ws, { header: 1, defval: '' });
          let hr = null;
          for (let i = 0; i < raw.length; i++) {
            const vs = raw[i].map(v => String(v).trim().toUpperCase());
            if (vs.includes('ZONES') && vs.includes('SUPER EN CHARGE')) { hr = i; break; }
          }
          if (hr === null) return;
          const headers = raw[hr].map(v => String(v).trim().toUpperCase());
          const ci = {
            ZONES: headers.findIndex(h => h === 'ZONES' || h === 'ZONE'),
            SUPER: headers.findIndex(h => h.includes('SUPER EN CHARGE')),
            SALLES: headers.findIndex(h => h === 'SALLES' || h === 'SALLE'),
            TECH: headers.findIndex(h => h.includes('NOMS TECHNIQUE') || h.includes('NOM TECHNIQUE')),
          };
          if (Object.values(ci).some(v => v < 0)) return;
          const rows = []; let lastZone = '', lastSuper = '';
          for (let i = hr + 1; i < raw.length; i++) {
            const row = raw[i];
            const zone  = String(row[ci.ZONES] || '').trim() || lastZone;
            const sup   = String(row[ci.SUPER] || '').trim() || lastSuper;
            const salle = String(row[ci.SALLES] || '').trim();
            const tech  = String(row[ci.TECH]   || '').trim();
            if (!salle && !tech) continue;
            if (zone.toUpperCase() === 'ZONES' || zone.toUpperCase() === 'NAN') continue;
            if (zone) lastZone = zone;
            if (sup)  lastSuper = sup;
            if (salle) {
              let k = tech.toUpperCase().replace(/ /g,'');
              if (k === 'ZOETELECAMTEL') k = 'ZOATELECAMTEL';
              rows.push({ ZONES: lastZone, SUPER: lastSuper, SALLES: salle, NOMS_TECH: tech, _k: k });
            }
          }
          result[sheet] = rows;
          log(`  Zone ${sheet} : ${rows.length} salles`);
        });
        res(result);
      } catch(err) { rej(err); }
    };
    reader.onerror = rej;
    reader.readAsArrayBuffer(file);
  });

  // ── Build master ────────────────────────────────
  const buildMaster = (recap, betRows, objRows) => {
    const betMap = {}; betRows.forEach(r => { betMap[r._k] = r; });
    const objMap = {}; objRows.forEach(r => { objMap[r._k] = r; });
    const out = {};
    Object.entries(recap).forEach(([ville, rows]) => {
      out[ville] = rows.map(r => {
        const b = betMap[r._k] || {};
        const o = objMap[r._k] || {};
        const mtd = b.Tickets_MTD || 0;
        const obj = o.Tickets_Obj || 0;
        const diff = mtd - obj;
        const pct  = obj > 0 ? (diff / obj) * 100 : null;
        let statut = '❌ Non atteint';
        if (pct !== null && pct >= 15) statut = '🏆 Obj+15%';
        else if (pct !== null && pct >= 0) statut = '✅ Atteint';
        return { ...r, Tickets_MTD: mtd, Payin_MTD: b.Payin_MTD||0, GGR_MTD: b.GGR_MTD||0, Tickets_Obj: obj, Diff: diff, Pct: pct, Statut: statut };
      });
    });
    return out;
  };

  // ── Classement Supers ───────────────────────────
  const makeSupers = (dfV) => {
    const supMap = {};
    dfV.forEach(r => {
      if (!supMap[r.SUPER]) supMap[r.SUPER] = [];
      supMap[r.SUPER].push(r);
    });
    const supers = Object.entries(supMap).map(([sup, rows]) => {
      const zones = [...new Set(rows.map(r => r.ZONES))];
      const att   = rows.filter(r => r.Statut !== '❌ Non atteint').length;
      const natt  = rows.filter(r => r.Statut === '❌ Non atteint').length;
      return {
        zone: zones.join(' / '), super: sup, nb: rows.length, att, natt,
        mtd: rows.reduce((s,r) => s + r.Tickets_MTD, 0),
        obj: rows.reduce((s,r) => s + r.Tickets_Obj, 0),
        ggr: rows.reduce((s,r) => s + r.GGR_MTD, 0),
        rows: rows.map(r => ({ zone: r.ZONES, salle: r.SALLES, tech: r.NOMS_TECH, mtd: r.Tickets_MTD, obj: r.Tickets_Obj, diff: r.Diff, pct: r.Pct, ggr: r.GGR_MTD, statut: r.Statut }))
      };
    });
    return supers.sort((a,b) => {
      const pa = a.obj > 0 ? (a.mtd-a.obj)/a.obj*100 : 0;
      const pb = b.obj > 0 ? (b.mtd-b.obj)/b.obj*100 : 0;
      return pb - pa;
    });
  };

  // ── Write Classement ────────────────────────────
  const writeClassement = (supers, moisLabel) => {
    const ws = {}; const merges = [];
    let R = 1;
    // Titre
    sc(ws,R,1,`CLASSEMENT SUPERS — ${moisLabel}`,true,13,'FFFFFF',C.DARK,'center');
    merges.push({ s:{r:R-1,c:0}, e:{r:R-1,c:7} });
    R++;
    // Headers
    ['Rang','Super','Zone','Salles','Atteints','Non att.','Tickets MTD','% vs Obj'].forEach((h,i) => {
      sc(ws,R,i+1,h,true,10,'FFFFFF',C.BLUE,'center');
    });
    R++;
    // Rows
    supers.forEach((sd, idx) => {
      const pct = sd.obj > 0 ? (sd.mtd-sd.obj)/sd.obj*100 : 0;
      const bg  = pct >= 15 ? C.OK15 : (pct >= 0 ? C.OK : C.KO);
      const fc  = pct >= 15 ? C.OK15F : (pct >= 0 ? C.OKF : C.KOF);
      const medal = ['🥇','🥈','🥉'][idx] || '';
      [
        [1, `${medal} ${idx+1}`, true,10,'000000',bg,'center'],
        [2, sd.super,             false,10,'000000',bg,'left'],
        [3, sd.zone,              false,9,'555555',bg,'left'],
        [4, sd.nb,                false,10,'000000',bg,'center'],
        [5, sd.att,               false,10,C.OKF,bg,'center'],
        [6, sd.natt,              false,10,C.KOF,bg,'center'],
        [7, sd.mtd,               true,10,'000000',bg,'right'],
        [8, `${pct>=0?'+':''}${pct.toFixed(1)}%`, true,10,fc,bg,'center'],
      ].forEach(([c,v,b,s,fc2,bg2,h]) => sc(ws,R,c,v,b,s,fc2,bg2,h));
      R++;
    });
    ws['!cols'] = [6,28,20,8,8,8,14,12].map(w => ({ wch: w }));
    ws['!merges'] = merges;
    ws['!ref'] = XLSX.utils.encode_range({ s:{r:0,c:0}, e:{r:R-1,c:7} });
    return ws;
  };

  // ── Write Detail ────────────────────────────────
  const writeDetail = (master, moisLabel) => {
    const ws = {}; let R = 1;
    sc(ws,R,1,`DÉTAIL PAR ZONE & SUPER — ${moisLabel}`,true,12,'FFFFFF',C.DARK,'center');
    ws['!merges'] = [{ s:{r:0,c:0}, e:{r:0,c:7} }];
    R++;
    ['Zone','Super','Salle','Nom technique','Tickets MTD','Objectif','Diff','% vs Obj'].forEach((h,i) => {
      sc(ws,R,i+1,h,true,10,'FFFFFF',C.BLUE,'center');
    });
    R++;
    Object.entries(master).forEach(([ville, rows]) => {
      sc(ws,R,1,`── ${ville}`,true,10,'FFFFFF',C.MID,'left');
      for (let c=2;c<=8;c++) sc(ws,R,c,'',false,10,'FFFFFF',C.MID,'left');
      R++;
      rows.forEach((r, i) => {
        const bg  = r.Pct === null ? C.KPI : (r.Pct >= 15 ? C.OK15 : (r.Pct >= 0 ? C.OK : C.KO));
        const dfc = (r.Diff||0) >= 0 ? C.OKF : C.KOF;
        const pfc = (r.Pct === null || r.Pct < 0) ? C.KOF : C.OKF;
        sc(ws,R,1,r.ZONES,  false,9,'555555',bg,'left');
        sc(ws,R,2,r.SUPER,  false,9,'000000',bg,'left');
        sc(ws,R,3,r.SALLES, false,9,'000000',bg,'left');
        sc(ws,R,4,r.NOMS_TECH,false,9,'000000',bg,'left');
        sc(ws,R,5,r.Tickets_MTD,true,10,'000000',bg,'right');
        sc(ws,R,6,r.Tickets_Obj,false,10,'000000',bg,'right');
        sc(ws,R,7,r.Diff,   true,10,dfc,bg,'right');
        sc(ws,R,8,r.Pct !== null ? parseFloat(r.Pct.toFixed(2)) : '',true,10,pfc,bg,'right');
        if (ws[XLSX.utils.encode_cell({r:R-1,c:7})]) ws[XLSX.utils.encode_cell({r:R-1,c:7})].z = '0.00"%"';
        R++;
      });
    });
    ws['!cols'] = [14,22,26,24,14,14,10,12].map(w => ({ wch: w }));
    ws['!ref'] = XLSX.utils.encode_range({ s:{r:0,c:0}, e:{r:R-1,c:7} });
    return ws;
  };

  // ── Write Super sheet ───────────────────────────
  const writeSuperSheet = (sd, moisLabel, rang, total) => {
    const ws = {}; const merges = []; let R = 1;
    const pct = sd.obj > 0 ? (sd.mtd-sd.obj)/sd.obj*100 : 0;
    sc(ws,R,1,`SUPER : ${sd.super}  ·  ${sd.zone}  ·  ${moisLabel}  —  Rang ${rang}/${total}`,true,13,'FFFFFF',C.BLUE,'center');
    merges.push({ s:{r:0,c:0}, e:{r:0,c:7} });
    R++;
    // KPIs
    const kpis = [
      ['Salles en charge',`${sd.nb} salles`,'Objectifs atteints',`${sd.att} / ${sd.nb}`],
      ['Non atteints',`${sd.natt} / ${sd.nb}`,'Tickets réalisés (MTD)',sd.mtd.toLocaleString('fr-FR')],
      ['Tickets objectif',sd.obj.toLocaleString('fr-FR'),'Réalisation vs objectif',`${pct>=0?'▲ ':'▼ '}${pct>0?'+':''}${pct.toFixed(1)}%`],
    ];
    kpis.forEach(([la,lv,ra,rv]) => {
      for (let c=1;c<=8;c++) sc(ws,R,c,'',false,10,'000000',C.KPI,'left');
      sc(ws,R,1,la,true,10);
      merges.push({ s:{r:R-1,c:0}, e:{r:R-1,c:1} });
      sc(ws,R,3,lv,true,11,'000000',C.KPI,'center');
      merges.push({ s:{r:R-1,c:2}, e:{r:R-1,c:3} });
      sc(ws,R,5,ra,true,10);
      merges.push({ s:{r:R-1,c:4}, e:{r:R-1,c:5} });
      const vbg = rv.includes('%') ? (pct>=0?C.OK:C.KO) : 'DEEAF1';
      const vfc = pct>=0 ? C.OKF : C.KOF;
      sc(ws,R,7,rv,true,11,vfc,vbg,'center');
      merges.push({ s:{r:R-1,c:6}, e:{r:R-1,c:7} });
      R++;
    });
    // En-têtes
    ['Salle (terrain)','Nom technique','Tickets MTD','Objectif (+15%)','Diff','% vs Obj','GGR MTD','Statut'].forEach((h,i) => {
      sc(ws,R,i+1,h,true,9,'FFFFFF',[C.BLUE,C.BLUE,C.BLUE,C.MID,C.MID,C.MID,C.RED,C.GREY][i],'center');
    });
    R++;
    // Lignes triées
    const sorted = [...sd.rows].sort((a,b) => {
      const aNo = a.mtd===0&&a.obj===0; const bNo = b.mtd===0&&b.obj===0;
      if (aNo!==bNo) return aNo?1:-1;
      return (b.pct||0)-(a.pct||0);
    });
    sorted.forEach(r => {
      const noData = r.mtd===0&&r.obj===0;
      const bg  = noData ? C.KPI : (r.statut==='🏆 Obj+15%' ? C.OK15 : (r.statut==='✅ Atteint' ? C.OK : C.KO));
      const dfc = (r.diff||0)>=0 ? C.OKF : C.KOF;
      const pfc = (!r.pct||r.pct<0) ? C.KOF : C.OKF;
      sc(ws,R,1,r.salle,  false,10,'000000',bg,'left');
      sc(ws,R,2,r.tech,   false,10,'000000',bg,'left');
      if (noData) {
        for (let c=3;c<=7;c++) { const a=XLSX.utils.encode_cell({r:R-1,c:c-1}); ws[a]={v:'—',t:'s',s:{font:fnt(false,10,'888888'),fill:fill(bg),border:bdr(),alignment:aln('center')}}; }
        sc(ws,R,8,'Données absentes',false,9,'888888',C.KPI,'center');
      } else {
        sc(ws,R,3,r.mtd,  true,10,'000000',bg,'right');
        sc(ws,R,4,r.obj,  false,10,'000000',bg,'right');
        sc(ws,R,5,r.diff, true,10,dfc,bg,'right');
        if (r.pct !== null) { sc(ws,R,6,parseFloat(r.pct.toFixed(2)),true,10,pfc,bg,'right'); if(ws[XLSX.utils.encode_cell({r:R-1,c:5})]) ws[XLSX.utils.encode_cell({r:R-1,c:5})].z='0.00"%"'; }
        sc(ws,R,7,Math.round(r.ggr),false,10,'000000',bg,'right');
        sc(ws,R,8,r.statut,true,9,r.statut==='🏆 Obj+15%'?C.OK15F:(r.statut==='✅ Atteint'?C.OKF:C.KOF),bg,'center');
      }
      R++;
    });
    // Total
    sc(ws,R,1,`TOTAL — ${sd.super}`,true,10,'FFFFFF',C.BLUE,'left');
    merges.push({ s:{r:R-1,c:0}, e:{r:R-1,c:1} });
    sc(ws,R,3,sd.mtd,true,10,'FFFFFF',C.BLUE,'right');
    sc(ws,R,4,sd.obj,true,10,'FFFFFF',C.BLUE,'right');
    sc(ws,R,5,sd.mtd-sd.obj,true,10,'FFFFFF',C.BLUE,'right');
    sc(ws,R,6,`${pct>0?'+':''}${pct.toFixed(2)}%`,true,10,'FFFFFF',C.BLUE,'center');
    sc(ws,R,7,Math.round(sd.ggr),true,10,'FFFFFF',C.BLUE,'right');
    sc(ws,R,8,'',true,10,'FFFFFF',C.BLUE,'center');
    ws['!cols'] = [26,24,14,16,12,10,14,14].map(w => ({ wch: w }));
    ws['!merges'] = merges;
    ws['!ref'] = XLSX.utils.encode_range({ s:{r:0,c:0}, e:{r:R-1,c:7} });
    return ws;
  };

  // ── Build workbook ──────────────────────────────
  const buildWb = (supers, masterSub, moisLabel) => {
    const wb = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(wb, writeClassement(supers, moisLabel), '① Classement Supers');
    XLSX.utils.book_append_sheet(wb, writeDetail(masterSub, moisLabel), '② Détail par Zone & Super');
    const medals = {1:'🥇',2:'🥈',3:'🥉'};
    supers.forEach((sd, i) => {
      const em = medals[i+1] || '  ';
      const name = `${em} ${sd.super}`.slice(0,31);
      XLSX.utils.book_append_sheet(wb, writeSuperSheet(sd, moisLabel, i+1, supers.length), name);
    });
    return wb;
  };

  // ── Pipeline principal ───────────────────────────
  const runPipeline = async () => {
    setRunning(true); setLogs([]); setResults(null);
    try {
      const moisUp = mois.toUpperCase();
      log('📂 Chargement des données...');
      const [betRows, objRows, recap] = await Promise.all([
        loadBetshop(betshopFile),
        loadObjectifs(objectifsFile),
        loadRecap(recapFile),
      ]);
      const villes = Object.keys(recap);
      log(`  Villes détectées : ${villes.join(', ')}`);
      const master = buildMaster(recap, betRows, objRows);
      const zip = new JSZip();

      // Fichier par ville
      for (const ville of villes) {
        log(`  Génération ${ville}...`);
        const supers = makeSupers(master[ville]);
        const wb = buildWb(supers, { [ville]: master[ville] }, moisUp);
        const fname = `ANALYSE_${ville.toUpperCase().replace(/ /g,'_')}_${moisUp.replace(/ /g,'_')}.xlsx`;
        zip.file(fname, XLSX.write(wb, { bookType: 'xlsx', type: 'array' }));
        log(`    ✅ ${fname}`);
      }

      // Fichier GLOBAL
      log('  Génération GLOBAL...');
      const allSupers = villes.flatMap(v => makeSupers(master[v]));
      allSupers.sort((a,b) => {
        const pa = a.obj>0?(a.mtd-a.obj)/a.obj*100:0;
        const pb = b.obj>0?(b.mtd-b.obj)/b.obj*100:0;
        return pb-pa;
      });
      const wbG = buildWb(allSupers, master, moisUp);
      const fnameG = `ANALYSE_GLOBAL_${moisUp.replace(/ /g,'_')}.xlsx`;
      zip.file(fnameG, XLSX.write(wbG, { bookType: 'xlsx', type: 'array' }));
      log(`    ✅ ${fnameG}`);

      const blob = await zip.generateAsync({ type: 'blob' });
      setResults({ blob, filename: `Rapports_${moisUp.replace(/ /g,'_')}.zip` });
      log(`\n✅ Terminé — ${villes.length + 1} fichiers générés`);
    } catch(err) {
      log(`❌ Erreur : ${err.message}`);
    }
    setRunning(false);
  };

  return (
    <div style={{ maxWidth: 860, margin: '0 auto', padding: '2rem 1rem' }}>
      <h2 style={{ fontFamily: 'monospace', color: 'var(--accent)', marginBottom: '1.5rem' }}>
        RAPPORTS RÉALISATION vs OBJECTIF
      </h2>

      {/* Upload 3 fichiers */}
      <div style={{ display: 'grid', gridTemplateColumns: '1fr 1fr 1fr', gap: '1rem', marginBottom: '1.5rem' }}>
        {[
          { label: 'Betshop_CAMPRESJ.xlsx', sub: 'Colonnes MTD : Tickets · Payin · GGR', ref: betRef, setter: setBetshopFile, file: betshopFile },
          { label: 'OBJECTIFS_CAMPRESJ.xlsx', sub: 'Colonne : Tickets Objectif', ref: objRef, setter: setObjectifsFile, file: objectifsFile },
          { label: 'RECAP_DES_SALLES.xlsx', sub: 'Feuilles par ville · Zones · Supers', ref: recRef, setter: setRecapFile, file: recapFile },
        ].map(({ label, sub, ref, setter, file }) => (
          <div key={label}
            onClick={() => ref.current.click()}
            onDragOver={e => e.preventDefault()}
            onDrop={e => { e.preventDefault(); setter(e.dataTransfer.files[0]); }}
            style={{ border: `2px dashed ${file ? 'var(--accent)' : '#444'}`, borderRadius: 8, padding: '1rem', textAlign: 'center', cursor: 'pointer', background: 'var(--surface)' }}>
            <input ref={ref} type="file" accept=".xlsx" hidden onChange={e => setter(e.target.files[0])} />
            <div style={{ fontSize: 22, marginBottom: 6 }}>📊</div>
            <div style={{ fontFamily: 'monospace', fontSize: 11, color: 'var(--fg)' }}>{label}</div>
            <div style={{ fontSize: 10, color: 'var(--muted)', marginTop: 4 }}>{sub}</div>
            {file && <div style={{ marginTop: 6, fontSize: 11, color: 'var(--accent)' }}>✓ {file.name}</div>}
          </div>
        ))}
      </div>

      {/* Mois */}
      <div style={{ marginBottom: '1.5rem' }}>
        <label style={{ display: 'flex', flexDirection: 'column', gap: 4, fontSize: 13, color: 'var(--muted)', maxWidth: 200 }}>
          Mois
          <input value={mois} onChange={e => setMois(e.target.value)}
            style={{ padding: '6px 10px', borderRadius: 6, border: '1px solid #444', background: 'var(--surface)', color: 'var(--fg)', fontFamily: 'monospace', fontSize: 13 }} />
        </label>
      </div>

      <button onClick={runPipeline} disabled={!betshopFile || !objectifsFile || !recapFile || running}
        style={{ padding: '10px 28px', background: running ? '#444' : 'var(--accent)', color: '#fff', border: 'none', borderRadius: 6, cursor: 'pointer', fontWeight: 'bold', fontSize: 14, marginBottom: '1.5rem' }}>
        {running ? '⏳ Génération...' : '▶ Générer les rapports'}
      </button>

      {/* Logs */}
      {logs.length > 0 && (
        <div style={{ background: '#111', borderRadius: 8, padding: '1rem', fontFamily: 'monospace', fontSize: 12, color: '#aaa', maxHeight: 240, overflowY: 'auto', marginBottom: '1rem', whiteSpace: 'pre-wrap' }}>
          {logs.join('\n')}
        </div>
      )}

      {/* Résultats */}
      {results && (
        <div style={{ background: 'var(--surface)', borderRadius: 8, padding: '1.5rem' }}>
          <button onClick={() => saveAs(results.blob, results.filename)}
            style={{ padding: '10px 24px', background: '#27ae60', color: '#fff', border: 'none', borderRadius: 6, cursor: 'pointer', fontWeight: 'bold', fontSize: 14 }}>
            ⬇ Télécharger {results.filename}
          </button>
        </div>
      )}
    </div>
  );
}
