// src/pages/ObjectifsPage.js
import React, { useState, useRef } from 'react';
import Papa from 'papaparse';
import XLSX from 'xlsx-js-style';
import { saveAs } from 'file-saver';
import JSZip from 'jszip';

export default function ObjectifsPage() {
  const [csvFile, setCsvFile]     = useState(null);
  const [recapFile, setRecapFile] = useState(null);
  const [mois, setMois]           = useState('MAI 2026');
  const [taux, setTaux]           = useState(15);
  const [running, setRunning]     = useState(false);
  const [logs, setLogs]           = useState([]);
  const [results, setResults]     = useState(null);
  const csvRef   = useRef();
  const recapRef = useRef();

  const log = (msg) => setLogs(p => [...p, msg]);

  // ── Chargement CSV ──────────────────────────────
  const loadCSV = (file) => new Promise((res, rej) => {
    Papa.parse(file, {
      header: true, skipEmptyLines: true,
      complete: ({ data }) => {
        const lookup = {};
        data.forEach(row => {
          const name = row['Betshop name'] || '';
          const parts = name.split('.');
          const salle = parts.length >= 3 ? parts[parts.length - 1].trim() : name.trim();
          lookup[salle.toLowerCase()] = row;
        });
        log(`  CSV chargé : ${Object.keys(lookup).length} salles trouvées`);
        res(lookup);
      },
      error: rej
    });
  });

  // ── Chargement RECAP ────────────────────────────
  const loadRecap = (file) => new Promise((res, rej) => {
    const reader = new FileReader();
    reader.onload = (e) => {
      try {
        const wb = XLSX.read(e.target.result, { type: 'array' });
        const zoneData = {};
        wb.SheetNames.forEach(sheet => {
          const ws = wb.Sheets[sheet];
          const rows = XLSX.utils.sheet_to_json(ws, { header: 1, defval: '' });
          const data = [];
          for (let i = 2; i < rows.length; i++) {
            const [zone, sup, sd, st] = rows[i];
            if (zone && sup && sd && st && sup !== 'SUPER EN CHARGE') {
              data.push([String(zone).trim(), String(sup).trim(), String(sd).trim(), String(st).trim()]);
            }
          }
          zoneData[sheet] = data;
          log(`  Zone ${sheet} : ${data.length} salles`);
        });
        res(zoneData);
      } catch (err) { rej(err); }
    };
    reader.onerror = rej;
    reader.readAsArrayBuffer(file);
  });

  // ── Styles ──────────────────────────────────────
  function bdr() {
    const t = { style: 'thin', color: { argb: 'FFBFBFBF' } };
    return { top: t, bottom: t, left: t, right: t };
  }
  function fill(rgb) { return { patternType: 'solid', fgColor: { rgb } }; }
  function fnt(bold, sz, rgb) { return { name: 'Calibri', bold, sz: sz||10, color: { rgb: rgb||'000000' } }; }
  function aln(h, wrap) { return { horizontal: h||'left', vertical: 'center', wrapText: !!wrap }; }

  const HDR_DARK  = { font: fnt(true,10,'FFFFFF'), fill: fill('1F3864'), border: bdr(), alignment: aln('center', true) };
  const HDR_GREEN = { font: fnt(true,10,'FFFFFF'), fill: fill('375623'), border: bdr(), alignment: aln('center', true) };
  const HDR_RED   = { font: fnt(true,10,'FFFFFF'), fill: fill('7B2C2C'), border: bdr(), alignment: aln('center', true) };
  const HDR_BLUE  = { font: fnt(true,10,'FFFFFF'), fill: fill('1F4E79'), border: bdr(), alignment: aln('center', true) };

  // Data row styles — alternating white/light-blue
  // Cols C and F always light-blue (EBF3FB) + bold
  function rowStyle(colIdx, rowIsEven) {
    // colIdx: 0=A, 1=B, 2=C, 3=D, 4=E, 5=F
    const isObjCol = colIdx === 2 || colIdx === 5; // C or F
    const bg = isObjCol ? 'EBF3FB' : (rowIsEven ? 'EBF3FB' : 'FFFFFF');
    const bold = isObjCol;
    return { font: fnt(bold, 10, '000000'), fill: fill(bg), border: bdr(), alignment: aln('right') };
  }
  function nameStyle(rowIsEven) {
    const bg = rowIsEven ? 'EBF3FB' : 'FFFFFF';
    return { font: fnt(false, 10, '000000'), fill: fill(bg), border: bdr(), alignment: aln('left') };
  }
  const TOTAL_STYLE = { font: fnt(true, 10, '000000'), fill: fill('D9D9D9'), border: bdr(), alignment: aln('right') };
  const TOTAL_NAME  = { font: fnt(true, 10, '000000'), fill: fill('D9D9D9'), border: bdr(), alignment: aln('left') };
  const TITLE_STYLE = { font: fnt(true, 11, 'FFFFFF'), fill: fill('1F3864'), border: bdr(), alignment: aln('left') };

  // ── Build sheet matching exact template ──────────
  const buildSheet = (rowsData, csvLookup, sheetTitle, moisLabel) => {
    const ws = {};
    const nRows = rowsData.length;
    const totalRow    = 5 + nRows;

    // Row 2 — Title (merged A2:F2)
    const titleText = `${sheetTitle}  ·  ${moisLabel.toUpperCase()}  ·  ${nRows} salles`;
    ws['A2'] = { v: titleText, t: 's', s: TITLE_STYLE };
    for (let c = 1; c < 6; c++) {
      const addr = XLSX.utils.encode_cell({ r: 1, c });
      ws[addr] = { v: '', t: 's', s: TITLE_STYLE };
    }

    // Row 3 — empty

    // Row 4 — Headers (height 42)
    ws['A4'] = { v: 'Salle', t: 's', s: HDR_DARK };
    ws['B4'] = { v: `Payin ${moisLabel} (FCFA)`, t: 's', s: HDR_GREEN };
    ws['C4'] = { v: `Payin Objectif (${moisLabel}) +${taux}% (FCFA) `, t: 's', s: HDR_DARK };
    ws['D4'] = { v: `GGR ${moisLabel} (FCFA)`, t: 's', s: HDR_RED };
    ws['E4'] = { v: `Tickets ${moisLabel}`, t: 's', s: HDR_BLUE };
    ws['F4'] = { v: `Tickets Objectif (${moisLabel}) +${taux}%`, t: 's', s: HDR_DARK };

    // Data rows starting at row 5
    const missing = [];
    const sumB = []; const sumD = []; const sumE = [];

    rowsData.forEach(([zone, sup, sd, st], i) => {
      const rowNum = 5 + i;
      const even   = i % 2 === 0;
      const csvRow = csvLookup[st.toLowerCase()];

      const payin   = csvRow ? parseFloat(csvRow['Payin money'])  || 0 : 0;
      const ggr     = csvRow ? parseFloat(csvRow['Gross'])        || 0 : 0;
      const tickets = csvRow ? parseFloat(csvRow['Payin count'])  || 0 : 0;
      if (!csvRow) missing.push(st);

      sumB.push(rowNum); sumD.push(rowNum); sumE.push(rowNum);

      ws[`A${rowNum}`] = { v: sd, t: 's', s: nameStyle(even) };
      ws[`B${rowNum}`] = { v: payin, t: 'n', s: rowStyle(1, even), z: '#,##0' };
      // C — formula
      ws[`C${rowNum}`] = { f: `B${rowNum}+((B${rowNum}*${taux}%)/100)`, t: 'n', s: rowStyle(2, even), z: '#,##0' };
      ws[`D${rowNum}`] = { v: ggr, t: 'n', s: rowStyle(3, even), z: '#,##0.00' };
      ws[`E${rowNum}`] = { v: tickets, t: 'n', s: rowStyle(4, even), z: '#,##0' };
      // F — formula
      ws[`F${rowNum}`] = { f: `ROUND(E${rowNum}*${1 + taux/100},0)`, t: 'n', s: rowStyle(5, even), z: '#,##0' };
    });

    // Total row
    const firstData = 5, lastData = 4 + nRows;
    ws[`A${totalRow}`] = { v: `TOTAL  (${nRows} salles)`, t: 's', s: TOTAL_NAME };
    ws[`B${totalRow}`] = { f: `SUM(B${firstData}:B${lastData})`, t: 'n', s: TOTAL_STYLE, z: '#,##0' };
    ws[`C${totalRow}`] = { f: `B${totalRow}+((B${totalRow}*${taux}%)/100)`, t: 'n', s: TOTAL_STYLE, z: '#,##0' };
    ws[`D${totalRow}`] = { f: `SUM(D${firstData}:D${lastData})`, t: 'n', s: TOTAL_STYLE, z: '#,##0.00' };
    ws[`E${totalRow}`] = { f: `SUM(E${firstData}:E${lastData})`, t: 'n', s: TOTAL_STYLE, z: '#,##0' };
    ws[`F${totalRow}`] = { f: `SUM(F${firstData}:F${lastData})`, t: 'n', s: TOTAL_STYLE, z: '#,##0' };

    // Merges & dims
    ws['!merges'] = [{ s: { r: 1, c: 0 }, e: { r: 1, c: 5 } }];
    ws['!cols'] = [30, 24, 32, 22, 20, 26].map(w => ({ wch: w }));
    ws['!rows'] = [
      {},         // row 1
      { hpt: 20 }, // row 2 title
      {},         // row 3
      { hpt: 42 }, // row 4 headers
      ...Array(nRows + 1).fill({}),
    ];
    ws['!ref'] = XLSX.utils.encode_range({ s: { r: 0, c: 0 }, e: { r: totalRow - 1, c: 5 } });

    return { ws, missing };
  };

  // ── Pipeline ─────────────────────────────────────
  const runPipeline = async () => {
    setRunning(true); setLogs([]); setResults(null);
    try {
      const moisUp = mois.toUpperCase();
      const moisSafe = moisUp.replace(/ /g, '_');
      const moisLabel = moisUp; // e.g. "AVRIL 2026"

      log('📂 Chargement des fichiers...');
      const [csvLookup, zoneData] = await Promise.all([loadCSV(csvFile), loadRecap(recapFile)]);

      const allZones = Object.keys(zoneData);
      const allRows  = allZones.flatMap(z => zoneData[z]);

      log(`\n📋 ${allZones.length} zones · ${allRows.length} salles au total`);

      const zip = new JSZip();
      const allMissing = [];

      // Global
      log('🌍 Génération GLOBAL...');
      const sheetTitle = `OBJECTIFS CAMPRESJ — ${moisUp}`;
      const { ws: wsG, missing: mG } = buildSheet(allRows, csvLookup, sheetTitle, moisLabel);
      allMissing.push(...mG);
      const wbG = XLSX.utils.book_new();
      const sheetName = `OBJECTIFS CAMPRESJ — ${moisUp}`.slice(0, 31);
      XLSX.utils.book_append_sheet(wbG, wsG, sheetName);
      zip.file(`${moisSafe}_GLOBAL.xlsx`, XLSX.write(wbG, { bookType: 'xlsx', type: 'array' }));
      log(`  ✅ ${moisSafe}_GLOBAL.xlsx (${allRows.length} salles)`);

      // Par zone
      log('🗺️  Génération par zone...');
      for (const zoneName of allZones) {
        const rows = zoneData[zoneName];
        const { ws, missing } = buildSheet(rows, csvLookup, `OBJECTIFS CAMPRESJ — ${moisUp}  ·  ${zoneName}`, moisLabel);
        allMissing.push(...missing);
        const wb = XLSX.utils.book_new();
        XLSX.utils.book_append_sheet(wb, ws, zoneName.slice(0, 31));
        zip.file(`${moisSafe}_ZONE_${zoneName}.xlsx`, XLSX.write(wb, { bookType: 'xlsx', type: 'array' }));
        log(`  ✅ ${moisSafe}_ZONE_${zoneName}.xlsx (${rows.length} salles)`);
      }

      const blob = await zip.generateAsync({ type: 'blob' });
      const uniqueMissing = [...new Set(allMissing)];
      setResults({ blob, filename: `Objectifs_${moisSafe}.zip`, missing: uniqueMissing });
      log(`\n✅ Terminé !${uniqueMissing.length ? `\n⚠️  ${uniqueMissing.length} salle(s) sans données CSV` : ''}`);
    } catch (err) {
      log(`❌ Erreur : ${err.message}`);
    }
    setRunning(false);
  };

  return (
    <div>
      {/* Upload */}
      <p className="section-label">Fichiers source</p>
      <div style={{ display: 'grid', gridTemplateColumns: '1fr 1fr', gap: '1rem', marginBottom: '1.5rem' }}>
        {[
          { label: 'shift_report_summary.csv', sub: 'Colonnes : Betshop name · Payin money · Payin count · Gross', ref: csvRef, setter: setCsvFile, file: csvFile, accept: '.csv', icon: '📋' },
          { label: 'RECAP_DES_SALLES_PAR_ZONE.xlsx', sub: 'Colonnes : Zone · Super · Salle · Nom technique', ref: recapRef, setter: setRecapFile, file: recapFile, accept: '.xlsx', icon: '📊' },
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
      <div className="card" style={{ marginBottom: '1.25rem' }}>
        <p className="section-label">Paramètres</p>
        <div style={{ display: 'flex', gap: '2rem', flexWrap: 'wrap' }}>
          <label style={{ display: 'flex', flexDirection: 'column', gap: 5 }}>
            <span style={{ fontSize: 12, color: 'var(--muted)', fontWeight: 600 }}>Mois</span>
            <input value={mois} onChange={e => setMois(e.target.value)} className="field-input" />
          </label>
          <label style={{ display: 'flex', flexDirection: 'column', gap: 5 }}>
            <span style={{ fontSize: 12, color: 'var(--muted)', fontWeight: 600 }}>Taux objectif (%)</span>
            <input type="number" value={taux} onChange={e => setTaux(Number(e.target.value))} className="field-input" style={{ width: 100 }} />
          </label>
        </div>
      </div>

      <button onClick={runPipeline} disabled={!csvFile || !recapFile || running}
        className="btn primary" style={{ fontSize: 14, padding: '11px 28px', marginBottom: '1.5rem' }}>
        {running ? '⟳ Génération...' : '▶  Générer les objectifs'}
      </button>

      {/* Logs */}
      {logs.length > 0 && (
        <>
          <p className="section-label">Console</p>
          <div className="log-console" style={{ marginBottom: '1.5rem' }}>
            {logs.map((l, i) => <span key={i} className="log-line" style={{ display: 'block' }}>{l}</span>)}
          </div>
        </>
      )}

      {/* Résultats */}
      {results && (
        <div className="card">
          <button onClick={() => saveAs(results.blob, results.filename)}
            className="btn success" style={{ fontSize: 14, padding: '10px 24px', marginBottom: '1rem' }}>
            ⬇  Télécharger {results.filename}
          </button>
          {results.missing.length > 0 && (
            <div style={{ marginTop: '1rem' }}>
              <p style={{ color: 'var(--casino)', fontSize: 13, marginBottom: 8, fontWeight: 600 }}>
                ⚠️ {results.missing.length} salle(s) sans données CSV :
              </p>
              <ul style={{ fontFamily: 'JetBrains Mono, monospace', fontSize: 12, color: 'var(--muted)', paddingLeft: 20 }}>
                {results.missing.map(s => <li key={s}>{s}</li>)}
              </ul>
            </div>
          )}
        </div>
      )}
    </div>
  );
}
