// src/pages/ObjectifsPage.js
import React, { useState, useRef } from 'react';
import Papa from 'papaparse';
import XLSX from 'xlsx-js-style';
import { saveAs } from 'file-saver';
import JSZip from 'jszip';

export default function ObjectifsPage() {
  const [csvFile,     setCsvFile]     = useState(null);
  const [recapFile,   setRecapFile]   = useState(null);
  const [listFile,    setListFile]    = useState(null);
  const [listData,    setListData]    = useState(null);
  const [selectedProp,setSelectedProp]= useState('');
  const [propResult,  setPropResult]  = useState(null);
  const [mois,        setMois]        = useState('MAI');
  const [annee,       setAnnee]       = useState(2026);
  const [taux,        setTaux]        = useState(15);
  const [supers,      setSupers]      = useState('YAOUNDE,DOUALA,OUEST');
  const [running,     setRunning]     = useState(false);
  const [logs,        setLogs]        = useState([]);
  const [result,      setResult]      = useState(null);

  const csvRef   = useRef();
  const recapRef = useRef();
  const listRef  = useRef();

  const log = (msg) => setLogs(p => [...p, msg]);

  // ── Styles ──────────────────────────────────────
  const S = {
    HEADER: { patternType:'solid', fgColor:{ rgb:'1F3864' } },
    GREEN:  { patternType:'solid', fgColor:{ rgb:'375623' } },
    RED:    { patternType:'solid', fgColor:{ rgb:'7B2C2C' } },
    BLUE:   { patternType:'solid', fgColor:{ rgb:'1F4E79' } },
    ALT:    { patternType:'solid', fgColor:{ rgb:'EBF3FB' } },
    WHITE:  { patternType:'solid', fgColor:{ rgb:'FFFFFF' } },
    TOTAL:  { patternType:'solid', fgColor:{ rgb:'D9D9D9' } },
    ZONE:   { patternType:'solid', fgColor:{ rgb:'2E75B6' } },
  };

  function bdr() {
    const t = { style:'thin', color:{ rgb:'BFBFBF' } };
    return { top:t, bottom:t, left:t, right:t };
  }
  function fnt(bold, sz, rgb) {
    return { name:'Calibri', bold:!!bold, sz:sz||10, color:{ rgb:rgb||'000000' } };
  }
  function aln(h, wrap) {
    return { horizontal:h||'left', vertical:'center', wrapText:!!wrap };
  }
  function sc(ws, r, c, v, style) {
    const addr = XLSX.utils.encode_cell({ r:r-1, c:c-1 });
    ws[addr] = { v:v===null||v===undefined?'':v, t:typeof v==='number'?'n':'s', s:style };
  }
  function scn(ws, r, c, v, style, fmt) {
    const addr = XLSX.utils.encode_cell({ r:r-1, c:c-1 });
    ws[addr] = { v:typeof v==='number'?v:0, t:'n', s:style, z:fmt };
  }
  function mg(ws, r1, c1, r2, c2) {
    if (!ws['!merges']) ws['!merges'] = [];
    ws['!merges'].push({ s:{r:r1-1,c:c1-1}, e:{r:r2-1,c:c2-1} });
  }

  const normalize = (text) => text.replace(/\s+/g,' ').trim().toLowerCase();

  // ── Load CSV ────────────────────────────────────
  const loadCSV = (file) => new Promise((res, rej) => {
    Papa.parse(file, {
      header:true, skipEmptyLines:true,
      complete: ({ data }) => {
        const lookup = {};
        data.forEach(row => {
          const name  = row['Betshop name'] || '';
          const parts = name.split('.');
          const salle = parts.length >= 3 ? parts[parts.length-1].trim() : name.trim();
          lookup[normalize(salle)] = row;
        });
        log(`  ✓ CSV : ${Object.keys(lookup).length} salles`);
        res(lookup);
      },
      error: rej
    });
  });

  // ── Load RECAP ──────────────────────────────────
  const loadRecap = (file) => new Promise((res, rej) => {
    const reader = new FileReader();
    reader.onload = (e) => {
      try {
        const wb = XLSX.read(e.target.result, { type:'array' });
        const zoneData = {};
        wb.SheetNames.forEach(sheet => {
          const ws   = wb.Sheets[sheet];
          const rows = XLSX.utils.sheet_to_json(ws, { header:1, defval:'' });
          const data = [];
          for (let i = 2; i < rows.length; i++) {
            const [zone, sup, sd, st] = rows[i];
            if (zone && sup && sd && st && String(sup).trim() !== 'SUPER EN CHARGE') {
              data.push([String(zone).trim(), String(sup).trim(), String(sd).trim(), String(st).trim()]);
            }
          }
          zoneData[sheet] = data;
          log(`  ✓ Zone ${sheet} : ${data.length} salles`);
        });
        res(zoneData);
      } catch(err) { rej(err); }
    };
    reader.onerror = rej;
    reader.readAsArrayBuffer(file);
  });

  // ── Load LISTE DES SALLES ───────────────────────
  const loadListFile = (file) => {
    const reader = new FileReader();
    reader.onload = (e) => {
      try {
        const wb   = XLSX.read(e.target.result, { type:'array' });
        const ws   = wb.Sheets[wb.SheetNames[0]];
        const rows = XLSX.utils.sheet_to_json(ws, { header:1, defval:'' });
        const propMap = {};
        for (let i = 1; i < rows.length; i++) {
          const shop = String(rows[i][2] || '').trim();
          const prop = String(rows[i][5] || '').trim();
          if (shop && prop) {
            if (!propMap[prop]) propMap[prop] = [];
            propMap[prop].push(shop);
          }
        }
        setListData(propMap);
        log(`  ✓ Liste : ${Object.keys(propMap).length} propriétaires, ${Object.values(propMap).flat().length} salles`);
      } catch(err) { log(`❌ Erreur liste : ${err.message}`); }
    };
    reader.readAsArrayBuffer(file);
  };

  // ── Build sheet ─────────────────────────────────
  const buildSheet = (rowsData, csvLookup, titleText, moisUp, anneeN, hasZone=true, hasSuper=true, tauxVal=15) => {
    const ws = {};
    const anneePrec = anneeN - 1;

    const colDefs = [];
    if (hasZone)  colDefs.push({ h:'ZONE',                                           fill:S.HEADER, w:14 });
    if (hasSuper) colDefs.push({ h:'SUPER EN CHARGE',                                fill:S.HEADER, w:22 });
    colDefs.push(
      { h:'SALLE',                                              fill:S.HEADER, w:28 },
      { h:'NOM TECHNIQUE',                                      fill:S.HEADER, w:18 },
      { h:`Payin ${moisUp} ${anneePrec} (FCFA)`,               fill:S.GREEN,  w:22 },
      { h:`Payin Objectif (${moisUp} ${anneeN}) +${tauxVal}%`, fill:S.HEADER, w:26 },
      { h:`GGR ${moisUp} ${anneePrec} (FCFA)`,                 fill:S.RED,    w:20 },
      { h:`Tickets ${moisUp} ${anneePrec}`,                     fill:S.BLUE,   w:16 },
      { h:`Tickets Objectif (${moisUp} ${anneeN}) +${tauxVal}%`, fill:S.HEADER, w:24 },
    );
    const ncols = colDefs.length;
    const pCol  = (hasZone?1:0) + (hasSuper?1:0) + 3; // 1-based col index of Payin

    // Title row 2
    sc(ws, 2, 1, titleText, { font:fnt(true,11,'FFFFFF'), fill:S.HEADER, alignment:aln('left'), border:bdr() });
    mg(ws, 2, 1, 2, ncols);
    ws['!rows'] = [{}, {hpt:20}, {}, {hpt:42}];

    // Headers row 4
    colDefs.forEach((cd, i) => {
      sc(ws, 4, i+1, cd.h, { font:fnt(true,10,'FFFFFF'), fill:cd.fill, alignment:aln('center',true), border:bdr() });
    });

    // Data rows
    let dataRow = 5;
    let prevGroup = null;
    const totals = [0,0,0,0,0];
    const missing = [];
    let rowIdx = 0;

    rowsData.forEach(([zone, sup, sd, st]) => {
      const groupKey = hasSuper ? `${zone}|${sup}` : zone;

      if (groupKey !== prevGroup) {
        let label = hasSuper && hasZone ? `── ${zone}  ·  ${sup}` : hasZone ? `── ${zone}` : `── ${sup}`;
        sc(ws, dataRow, 1, label, { font:fnt(true,9,'FFFFFF'), fill:S.ZONE, border:bdr() });
        for (let c = 2; c <= ncols; c++) {
          sc(ws, dataRow, c, '', { font:fnt(false,9,'FFFFFF'), fill:S.ZONE, border:bdr() });
        }
        if (!ws['!rows'][dataRow-1]) ws['!rows'][dataRow-1] = {};
        ws['!rows'][dataRow-1] = { hpt:14 };
        dataRow++;
        prevGroup = groupKey;
      }

      const rowFill = rowIdx % 2 === 0 ? S.ALT : S.WHITE;
      let col = 1;
      if (hasZone)  { sc(ws, dataRow, col, zone, { font:fnt(false,9), fill:rowFill, border:bdr(), alignment:aln('left') }); col++; }
      if (hasSuper) { sc(ws, dataRow, col, sup,  { font:fnt(false,9), fill:rowFill, border:bdr(), alignment:aln('left') }); col++; }
      sc(ws, dataRow, col, sd, { font:fnt(false,9), fill:rowFill, border:bdr(), alignment:aln('left') }); col++;
      sc(ws, dataRow, col, st, { font:fnt(false,9), fill:rowFill, border:bdr(), alignment:aln('left') }); col++;

      const csvRow = csvLookup[normalize(st)];
      if (csvRow) {
        const payin      = parseInt(csvRow['Payin money'])  || 0;
        const tickets    = parseInt(csvRow['Payin count'])  || 0;
        const gross      = parseFloat(csvRow['Gross'])      || 0;
        const payinObj   = Math.round(payin   * (1 + tauxVal/100));
        const ticketsObj = Math.round(tickets * (1 + tauxVal/100));
        totals[0]+=payin; totals[1]+=payinObj; totals[2]+=gross; totals[3]+=tickets; totals[4]+=ticketsObj;
        scn(ws,dataRow,col,  payin,      { font:fnt(false,9), fill:rowFill, border:bdr(), alignment:aln('right') }, '#,##0'); col++;
        scn(ws,dataRow,col,  payinObj,   { font:fnt(true,9),  fill:S.ALT,   border:bdr(), alignment:aln('right') }, '#,##0'); col++;
        scn(ws,dataRow,col,  gross,      { font:fnt(false,9), fill:rowFill, border:bdr(), alignment:aln('right') }, '#,##0.00'); col++;
        scn(ws,dataRow,col,  tickets,    { font:fnt(false,9), fill:rowFill, border:bdr(), alignment:aln('right') }, '#,##0'); col++;
        scn(ws,dataRow,col,  ticketsObj, { font:fnt(true,9),  fill:S.ALT,   border:bdr(), alignment:aln('right') }, '#,##0');
      } else {
        missing.push(st);
        for (let c = col; c <= ncols; c++) {
          sc(ws, dataRow, c, '', { font:fnt(false,9), fill:rowFill, border:bdr() });
        }
      }
      dataRow++; rowIdx++;
    });

    // Total row
    const totalR = dataRow;
    sc(ws, totalR, 1, `TOTAL  (${rowsData.length} salles)`, { font:fnt(true,10), fill:S.TOTAL, border:bdr(), alignment:aln('left') });
    if (pCol > 1) { mg(ws, totalR, 1, totalR, pCol-1); }
    for (let c = 2; c < pCol; c++) { sc(ws, totalR, c, '', { font:fnt(true,10), fill:S.TOTAL, border:bdr() }); }
    [[pCol,totals[0],'#,##0'],[pCol+1,totals[1],'#,##0'],[pCol+2,Math.round(totals[2]*100)/100,'#,##0.00'],[pCol+3,totals[3],'#,##0'],[pCol+4,totals[4],'#,##0']].forEach(([c,v,fmt]) => {
      scn(ws, totalR, c, v, { font:fnt(true,10), fill:S.TOTAL, border:bdr(), alignment:aln('right') }, fmt);
    });

    ws['!cols'] = colDefs.map(cd => ({ wch:cd.w }));
    ws['!ref']  = XLSX.utils.encode_range({ s:{r:0,c:0}, e:{r:totalR-1,c:ncols-1} });
    return { ws, missing };
  };

  // ── Pipeline global ──────────────────────────────
  const runPipeline = async () => {
    setRunning(true); setLogs([]); setResult(null);
    try {
      const moisUp   = mois.toUpperCase().trim();
      const anneeN   = parseInt(annee);
      const tauxVal  = parseInt(taux);
      const prefix   = `${moisUp}${anneeN}`;
      const zonesSup = supers.split(',').map(s => s.trim().toUpperCase()).filter(Boolean);

      log('📂 Chargement des données...');
      const [csvLookup, zoneData] = await Promise.all([loadCSV(csvFile), loadRecap(recapFile)]);
      const allZones = Object.keys(zoneData);
      const allRows  = allZones.flatMap(z => zoneData[z]);
      log(`\n📋 ${allZones.length} zones · ${allRows.length} salles`);

      const zip = new JSZip();
      const allMissing = [];

      // GLOBAL
      log('\n🌍 Génération GLOBAL...');
      const { ws:wsG, missing:mG } = buildSheet(allRows, csvLookup,
        `OBJECTIFS CAMPRESJ — ${moisUp} ${anneeN}  ·  GLOBAL  ·  ${allRows.length} salles`,
        moisUp, anneeN, true, true, tauxVal);
      allMissing.push(...mG);
      const wbG = XLSX.utils.book_new();
      XLSX.utils.book_append_sheet(wbG, wsG, `GLOBAL ${moisUp} ${anneeN}`.slice(0,31));
      zip.file(`${prefix}_GLOBAL.xlsx`, XLSX.write(wbG,{bookType:'xlsx',type:'array'}));
      log(`  ✅ ${prefix}_GLOBAL.xlsx (${allRows.length} salles)`);

      // PAR ZONE
      log('\n🗺️  Génération par zone...');
      for (const zoneName of allZones) {
        const rows = zoneData[zoneName];
        const { ws, missing } = buildSheet(rows, csvLookup,
          `OBJECTIFS CAMPRESJ — ${moisUp} ${anneeN}  ·  ${zoneName}  ·  ${rows.length} salles`,
          moisUp, anneeN, true, true, tauxVal);
        allMissing.push(...missing);
        const wb = XLSX.utils.book_new();
        XLSX.utils.book_append_sheet(wb, ws, zoneName.slice(0,31));
        zip.file(`${prefix}_ZONE_${zoneName}.xlsx`, XLSX.write(wb,{bookType:'xlsx',type:'array'}));
        log(`  ✅ ${prefix}_ZONE_${zoneName}.xlsx (${rows.length} salles)`);
      }

      // PAR SUPER
      log(`\n👤 Génération par super (${zonesSup.join(',')})...`);
      for (const zoneName of zonesSup) {
        if (!zoneData[zoneName]) { log(`  ⚠️ Zone '${zoneName}' introuvable`); continue; }
        const superMap = {};
        zoneData[zoneName].forEach(([z,sup,sd,st]) => {
          if (!superMap[sup]) superMap[sup] = [];
          superMap[sup].push([z,sup,sd,st]);
        });
        for (const [sup, rows] of Object.entries(superMap)) {
          const { ws, missing } = buildSheet(rows, csvLookup,
            `OBJECTIFS ${moisUp} ${anneeN}  ·  ${zoneName}  ·  ${sup}  ·  ${rows.length} salles`,
            moisUp, anneeN, true, false, tauxVal);
          allMissing.push(...missing);
          const wb = XLSX.utils.book_new();
          XLSX.utils.book_append_sheet(wb, ws, sup.slice(0,31));
          const safe = sup.replace(/ /g,'_').replace(/\//g,'-');
          zip.file(`${prefix}_${zoneName}_${safe}.xlsx`, XLSX.write(wb,{bookType:'xlsx',type:'array'}));
          log(`  ✅ ${prefix}_${zoneName}_${safe}.xlsx`);
        }
      }

      const blob = await zip.generateAsync({ type:'blob' });
      const uniqueMissing = [...new Set(allMissing)];
      setResult({ blob, filename:`Objectifs_${moisUp}${anneeN}.zip`, missing:uniqueMissing });
      log(`\n✅ Terminé !${uniqueMissing.length ? ` ⚠️ ${uniqueMissing.length} salle(s) sans données CSV` : ''}`);
    } catch(err) { log(`❌ Erreur : ${err.message}`); }
    setRunning(false);
  };

  // ── Pipeline propriétaire ────────────────────────
  const runPropPipeline = async () => {
    if (!selectedProp || !listData || !csvFile) return;
    setPropResult(null);
    try {
      const moisUp  = mois.toUpperCase().trim();
      const anneeN  = parseInt(annee);
      const tauxVal = parseInt(taux);
      const propShops = new Set(listData[selectedProp].map(s => normalize(s)));

      log(`\n🏪 Génération objectifs ${selectedProp}...`);
      const csvLookup = await loadCSV(csvFile);

      let propRows = [];

      if (recapFile) {
        // Si RECAP disponible : filtrer les salles du propriétaire depuis le RECAP
        const zoneData = await loadRecap(recapFile);
        const allRows  = Object.values(zoneData).flat();
        propRows = allRows.filter(([z,sup,sd,st]) =>
          propShops.has(normalize(st)) || propShops.has(normalize(sd))
        );
        log(`  ${propRows.length} salles trouvées dans le RECAP pour ${selectedProp}`);
      } else {
        // Sinon : construire les lignes directement depuis la liste des salles
        listData[selectedProp].forEach(shopName => {
          propRows.push(['—', '—', shopName, shopName]);
        });
        log(`  ${propRows.length} salles depuis la liste pour ${selectedProp}`);
      }

      if (propRows.length === 0) {
        log(`  ⚠️ Aucune salle trouvée — vérifiez la correspondance des noms`);
        return;
      }

      const hasZone  = recapFile ? true  : false;
      const hasSuper = recapFile ? true  : false;

      const { ws, missing } = buildSheet(propRows, csvLookup,
        `OBJECTIFS ${selectedProp} — ${moisUp} ${anneeN}  ·  ${propRows.length} salles`,
        moisUp, anneeN, hasZone, hasSuper, tauxVal);
      const wb = XLSX.utils.book_new();
      XLSX.utils.book_append_sheet(wb, ws, selectedProp.slice(0,31));
      const filename = `${moisUp}${anneeN}_${selectedProp.replace(/ /g,'_')}.xlsx`;
      const blob = new Blob([XLSX.write(wb,{bookType:'xlsx',type:'array'})],
        {type:'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'});
      setPropResult({ blob, filename, count:propRows.length, missing });
      log(`  ✅ ${filename} (${propRows.length} salles)`);
    } catch(err) { log(`❌ Erreur : ${err.message}`); }
  };

  // ── UI ───────────────────────────────────────────
  return (
    <div>
      {/* Fichiers principaux */}
      <p className="section-label">Fichiers source</p>
      <div style={{ display:'grid', gridTemplateColumns:'1fr 1fr', gap:'1rem', marginBottom:'1.25rem' }}>
        {[
          { label:'shift_report_summary.csv', sub:'Colonnes : Betshop name · Payin money · Payin count · Gross', ref:csvRef, setter:setCsvFile, file:csvFile, accept:'.csv', icon:'📋' },
          { label:'RECAP_DES_SALLES_PAR_ZONE.xlsx', sub:'Colonnes : Zone · Super · Salle · Nom technique', ref:recapRef, setter:setRecapFile, file:recapFile, accept:'.xlsx', icon:'📊' },
        ].map(({ label, sub, ref, setter, file, accept, icon }) => (
          <div key={label} onClick={() => ref.current.click()}
            onDragOver={e => e.preventDefault()}
            onDrop={e => { e.preventDefault(); setter(e.dataTransfer.files[0]); }}
            className={`drop-zone ${file ? 'has-file' : ''}`}>
            <input ref={ref} type="file" accept={accept} hidden onChange={e => setter(e.target.files[0])} />
            <div className="drop-icon">{icon}</div>
            <div className="drop-label">{label}</div>
            <div className="drop-sub">{sub}</div>
            {file && <div className="file-chip">✓ {file.name}</div>}
          </div>
        ))}
      </div>

      {/* Paramètres */}
      <div className="card" style={{ marginBottom:'1.25rem' }}>
        <p className="section-label">Paramètres</p>
        <div style={{ display:'flex', gap:'1.5rem', flexWrap:'wrap' }}>
          <label style={{ display:'flex', flexDirection:'column', gap:5 }}>
            <span style={{ fontSize:12, color:'var(--muted)', fontWeight:600 }}>Mois (ex: MAI)</span>
            <input value={mois} onChange={e => setMois(e.target.value)} className="field-input" style={{ width:120 }} />
          </label>
          <label style={{ display:'flex', flexDirection:'column', gap:5 }}>
            <span style={{ fontSize:12, color:'var(--muted)', fontWeight:600 }}>Année</span>
            <input type="number" value={annee} onChange={e => setAnnee(Number(e.target.value))} className="field-input" style={{ width:100 }} />
          </label>
          <label style={{ display:'flex', flexDirection:'column', gap:5 }}>
            <span style={{ fontSize:12, color:'var(--muted)', fontWeight:600 }}>Taux objectif (%)</span>
            <input type="number" value={taux} onChange={e => setTaux(Number(e.target.value))} className="field-input" style={{ width:100 }} />
          </label>
          <label style={{ display:'flex', flexDirection:'column', gap:5, flex:1 }}>
            <span style={{ fontSize:12, color:'var(--muted)', fontWeight:600 }}>Zones avec fichiers par super</span>
            <input value={supers} onChange={e => setSupers(e.target.value)} className="field-input" />
          </label>
        </div>
      </div>

      {/* Liste des salles par propriétaire (optionnel) */}
      <p className="section-label">Fichier optionnel — Objectifs par propriétaire</p>
      <div onClick={() => listRef.current.click()}
        onDragOver={e => e.preventDefault()}
        onDrop={e => { e.preventDefault(); const f=e.dataTransfer.files[0]; setListFile(f); loadListFile(f); }}
        className={`drop-zone ${listFile ? 'has-file' : ''}`}
        style={{ marginBottom:'1rem' }}>
        <input ref={listRef} type="file" accept=".xlsx,.xls" hidden
          onChange={e => { setListFile(e.target.files[0]); loadListFile(e.target.files[0]); }} />
        <div className="drop-icon">🏪</div>
        <div className="drop-label">LISTE_DES_SALLES_CAMEROON.xlsx</div>
        <div className="drop-sub">Colonnes : Market · Shop · City · Address · Proprietaire — permet de générer les objectifs par propriétaire sans RECAP</div>
        {listFile && <div className="file-chip">✓ {listFile.name}</div>}
      </div>

      {listData && (
        <div className="card" style={{ marginBottom:'1.25rem' }}>
          <p className="section-label">Sélectionner un propriétaire</p>
          <div style={{ display:'flex', gap:'1rem', alignItems:'flex-end', flexWrap:'wrap' }}>
            <label style={{ display:'flex', flexDirection:'column', gap:5, flex:1 }}>
              <span style={{ fontSize:12, color:'var(--muted)', fontWeight:600 }}>Propriétaire</span>
              <select value={selectedProp} onChange={e => setSelectedProp(e.target.value)}
                className="field-input" style={{ cursor:'pointer' }}>
                <option value="">-- Choisir un propriétaire --</option>
                {Object.entries(listData).sort(([a],[b])=>a.localeCompare(b)).map(([prop, shops]) => (
                  <option key={prop} value={prop}>{prop} ({shops.length} salles)</option>
                ))}
              </select>
            </label>
            <button onClick={runPropPipeline}
              disabled={!selectedProp || !csvFile}
              className="btn primary" style={{ padding:'9px 20px', whiteSpace:'nowrap' }}>
              ▶  Objectifs {selectedProp || 'propriétaire'}
            </button>
          </div>
          {propResult && (
            <div style={{ marginTop:'1rem', display:'flex', alignItems:'center', gap:'1rem', flexWrap:'wrap' }}>
              <button onClick={() => saveAs(propResult.blob, propResult.filename)}
                className="btn success" style={{ fontSize:13 }}>
                ⬇  Télécharger {propResult.filename}
              </button>
              <span style={{ fontSize:12, color:'var(--muted)' }}>
                {propResult.count} salles
                {propResult.missing.length > 0 && ` · ⚠️ ${propResult.missing.length} sans données CSV`}
              </span>
            </div>
          )}
        </div>
      )}

      {/* Bouton principal */}
      <button onClick={runPipeline} disabled={!csvFile || !recapFile || running}
        className="btn primary" style={{ fontSize:14, padding:'11px 28px', marginBottom:'1.5rem' }}>
        {running ? '⟳ Génération...' : '▶  Générer tous les objectifs'}
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

      {/* Résultat global */}
      {result && (
        <div className="card">
          <button onClick={() => saveAs(result.blob, result.filename)}
            className="btn success" style={{ fontSize:14, padding:'10px 24px', marginBottom:'1rem' }}>
            ⬇  Télécharger {result.filename}
          </button>
          {result.missing.length > 0 && (
            <div style={{ marginTop:'1rem' }}>
              <p style={{ color:'var(--casino)', fontSize:13, marginBottom:8, fontWeight:600 }}>
                ⚠️ {result.missing.length} salle(s) sans données CSV :
              </p>
              <ul style={{ fontFamily:'JetBrains Mono, monospace', fontSize:12, color:'var(--muted)', paddingLeft:20 }}>
                {result.missing.map(s => <li key={s}>{s}</li>)}
              </ul>
            </div>
          )}
        </div>
      )}
    </div>
  );
}
