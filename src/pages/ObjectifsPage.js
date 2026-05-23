// src/pages/ObjectifsPage.js
import React, { useState, useRef } from 'react';
import Papa from 'papaparse';
import * as XLSX from 'xlsx';
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

  const handleDrop = (e, setter) => {
    e.preventDefault();
    const f = e.dataTransfer.files[0];
    if (f) setter(f);
  };

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
              data.push([
                String(zone).trim(),
                String(sup).trim(),
                String(sd).trim(),
                String(st).trim()
              ]);
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
  const STYLES = {
    HEADER:  { font: { name: 'Calibri', bold: true, sz: 10, color: { rgb: 'FFFFFF' } }, fill: { fgColor: { rgb: '1F3864' } }, alignment: { horizontal: 'center', vertical: 'center', wrapText: true }, border: bdr() },
    GREEN:   { font: { name: 'Calibri', bold: true, sz: 10, color: { rgb: 'FFFFFF' } }, fill: { fgColor: { rgb: '375623' } }, alignment: { horizontal: 'center', vertical: 'center', wrapText: true }, border: bdr() },
    RED:     { font: { name: 'Calibri', bold: true, sz: 10, color: { rgb: 'FFFFFF' } }, fill: { fgColor: { rgb: '7B2C2C' } }, alignment: { horizontal: 'center', vertical: 'center', wrapText: true }, border: bdr() },
    BLUE:    { font: { name: 'Calibri', bold: true, sz: 10, color: { rgb: 'FFFFFF' } }, fill: { fgColor: { rgb: '1F4E79' } }, alignment: { horizontal: 'center', vertical: 'center', wrapText: true }, border: bdr() },
    ZONE_H:  { font: { name: 'Calibri', bold: true, sz: 9,  color: { rgb: 'FFFFFF' } }, fill: { fgColor: { rgb: '2E75B6' } }, border: bdr() },
    TOTAL:   { font: { name: 'Calibri', bold: true, sz: 10 }, fill: { fgColor: { rgb: 'D9D9D9' } }, border: bdr() },
    ALT:     { font: { name: 'Calibri', sz: 9 }, fill: { fgColor: { rgb: 'EBF3FB' } }, border: bdr() },
    WHITE:   { font: { name: 'Calibri', sz: 9 }, fill: { fgColor: { rgb: 'FFFFFF' } }, border: bdr() },
  };

  function bdr() {
    const t = { style: 'thin', color: { rgb: 'BFBFBF' } };
    return { top: t, bottom: t, left: t, right: t };
  }

  function sc(ws, r, c, v, style, numFmt) {
    const addr = XLSX.utils.encode_cell({ r: r - 1, c: c - 1 });
    ws[addr] = { v, t: typeof v === 'number' ? 'n' : 's', s: style };
    if (numFmt) ws[addr].z = numFmt;
  }

  // ── Construction d'un onglet ─────────────────────
  const buildSheet = (rowsData, csvLookup, titleText, hasZone = true, hasSuper = true) => {
    const ws = {};
    const moisLabel = titleText.includes('·') ? titleText.split('·')[1].trim() : 'MOIS';

    const colDefs = [];
    if (hasZone)  colDefs.push({ h: 'ZONE',                            s: STYLES.HEADER, w: 14 });
    if (hasSuper) colDefs.push({ h: 'SUPER EN CHARGE',                 s: STYLES.HEADER, w: 22 });
    colDefs.push(
      { h: 'SALLE',                                  s: STYLES.HEADER, w: 28 },
      { h: 'NOM TECHNIQUE',                          s: STYLES.HEADER, w: 18 },
      { h: `Payin ${moisLabel} (FCFA)`,              s: STYLES.GREEN,  w: 22 },
      { h: `Payin Objectif +${taux}%`,               s: STYLES.HEADER, w: 26 },
      { h: `GGR ${moisLabel} (FCFA)`,                s: STYLES.RED,    w: 20 },
      { h: `Tickets ${moisLabel}`,                   s: STYLES.BLUE,   w: 16 },
      { h: `Tickets Objectif +${taux}%`,             s: STYLES.HEADER, w: 24 },
    );
    const ncols = colDefs.length;

    // Titre row 2
    const titleAddr = XLSX.utils.encode_cell({ r: 1, c: 0 });
    ws[titleAddr] = { v: titleText, t: 's', s: { font: { name: 'Calibri', bold: true, sz: 11, color: { rgb: 'FFFFFF' } }, fill: { fgColor: { rgb: '1F3864' } }, alignment: { horizontal: 'left', vertical: 'center' } } };
    ws['!merges'] = [{ s: { r: 1, c: 0 }, e: { r: 1, c: ncols - 1 } }];

    // Headers row 4
    colDefs.forEach((cd, i) => {
      const addr = XLSX.utils.encode_cell({ r: 3, c: i });
      ws[addr] = { v: cd.h, t: 's', s: cd.s };
    });

    // Données
    let dataRow = 4; // 0-based = row 5
    let prevGroup = null;
    let totalPayin = 0, totalPayinObj = 0, totalGross = 0, totalTickets = 0, totalTicketsObj = 0;
    const missing = [];
    const pCol = colDefs.findIndex(c => c.h.startsWith('Payin') && !c.h.includes('Objectif'));

    rowsData.forEach(([ zone, sup, sd, st ], i) => {
      const groupKey = hasSuper ? `${zone}|${sup}` : zone;
      if (groupKey !== prevGroup) {
        const label = hasSuper && hasZone ? `── ${zone}  ·  ${sup}` : hasZone ? `── ${zone}` : `── ${sup}`;
        for (let c = 0; c < ncols; c++) {
          const addr = XLSX.utils.encode_cell({ r: dataRow, c });
          ws[addr] = { v: c === 0 ? label : '', t: 's', s: STYLES.ZONE_H };
        }
        dataRow++; prevGroup = groupKey;
      }

      const rowStyle = i % 2 === 0 ? STYLES.ALT : STYLES.WHITE;
      let col = 0;
      if (hasZone)  { sc(ws, dataRow + 1, col + 1, zone, rowStyle); col++; }
      if (hasSuper) { sc(ws, dataRow + 1, col + 1, sup,  rowStyle); col++; }
      sc(ws, dataRow + 1, col + 1, sd, rowStyle); col++;
      sc(ws, dataRow + 1, col + 1, st, rowStyle); col++;

      const csvRow = csvLookup[st.toLowerCase()];
      if (csvRow) {
        const payin      = parseFloat(csvRow['Payin money']) || 0;
        const tickets    = parseFloat(csvRow['Payin count']) || 0;
        const gross      = parseFloat(csvRow['Gross'])       || 0;
        const payinObj   = Math.round(payin * (1 + taux / 100));
        const ticketsObj = Math.round(tickets * (1 + taux / 100));
        totalPayin += payin; totalPayinObj += payinObj;
        totalGross += gross; totalTickets += tickets; totalTicketsObj += ticketsObj;

        sc(ws, dataRow + 1, col + 1, payin,      rowStyle, '#,##0');    col++;
        sc(ws, dataRow + 1, col + 1, payinObj,   STYLES.ALT, '#,##0'); col++;
        sc(ws, dataRow + 1, col + 1, gross,      rowStyle, '#,##0.00'); col++;
        sc(ws, dataRow + 1, col + 1, tickets,    rowStyle, '#,##0');    col++;
        sc(ws, dataRow + 1, col + 1, ticketsObj, STYLES.ALT, '#,##0');
      } else {
        missing.push(st);
      }
      dataRow++;
    });

    // Ligne TOTAL
    const mergeEnd = pCol - 1;
    if (!ws['!merges']) ws['!merges'] = [];
    ws['!merges'].push({ s: { r: dataRow, c: 0 }, e: { r: dataRow, c: mergeEnd } });
    const totalAddr = XLSX.utils.encode_cell({ r: dataRow, c: 0 });
    ws[totalAddr] = { v: `TOTAL  (${rowsData.length} salles)`, t: 's', s: STYLES.TOTAL };
    [
      [pCol,     totalPayin,            '#,##0'],
      [pCol + 1, totalPayinObj,         '#,##0'],
      [pCol + 2, Math.round(totalGross * 100) / 100, '#,##0.00'],
      [pCol + 3, totalTickets,          '#,##0'],
      [pCol + 4, totalTicketsObj,       '#,##0'],
    ].forEach(([c, v, fmt]) => {
      const addr = XLSX.utils.encode_cell({ r: dataRow, c });
      ws[addr] = { v, t: 'n', s: STYLES.TOTAL, z: fmt };
    });

    ws['!cols'] = colDefs.map(cd => ({ wch: cd.w }));
    ws['!ref'] = XLSX.utils.encode_range({ s: { r: 0, c: 0 }, e: { r: dataRow, c: ncols - 1 } });
    return { ws, missing };
  };

  // ── Pipeline principal ───────────────────────────
  const runPipeline = async () => {
    setRunning(true); setLogs([]); setResults(null);
    try {
      const moisUp   = mois.toUpperCase();
      const moisSafe = moisUp.replace(/ /g, '');

      log('📂 Chargement des fichiers...');
      const [csvLookup, zoneData] = await Promise.all([loadCSV(csvFile), loadRecap(recapFile)]);

      const allZones = Object.keys(zoneData);
      const allRows  = allZones.flatMap(z => zoneData[z]);
      const superZones = ['YAOUNDE', 'DOUALA', 'OUEST'].filter(z => allZones.includes(z));

      log(`\n📋 ${allZones.length} zones · ${allRows.length} salles au total`);

      const zip = new JSZip();
      const allMissing = [];

      // 1. GLOBAL
      log('🌍 Génération GLOBAL...');
      const { ws: wsG, missing: mG } = buildSheet(allRows, csvLookup,
        `OBJECTIFS CAMPRESJ — ${moisUp}  ·  GLOBAL  ·  ${allRows.length} salles`);
      allMissing.push(...mG);
      const wbG = XLSX.utils.book_new();
      XLSX.utils.book_append_sheet(wbG, wsG, 'GLOBAL');
      zip.file(`${moisSafe}_GLOBAL.xlsx`, XLSX.write(wbG, { bookType: 'xlsx', type: 'array' }));
      log(`  ✅ ${moisSafe}_GLOBAL.xlsx`);

      // 2. Par zone
      log('🗺️  Génération par zone...');
      for (const zoneName of allZones) {
        const rows = zoneData[zoneName];
        const { ws, missing } = buildSheet(rows, csvLookup,
          `OBJECTIFS CAMPRESJ — ${moisUp}  ·  ${zoneName}  ·  ${rows.length} salles`);
        allMissing.push(...missing);
        const wb = XLSX.utils.book_new();
        XLSX.utils.book_append_sheet(wb, ws, zoneName.slice(0, 31));
        zip.file(`${moisSafe}_ZONE_${zoneName}.xlsx`, XLSX.write(wb, { bookType: 'xlsx', type: 'array' }));
        log(`  ✅ ${moisSafe}_ZONE_${zoneName}.xlsx  (${rows.length} salles)`);
      }

      // 3. Par super
      log('👤 Génération par super...');
      for (const zoneName of superZones) {
        const supers = {};
        zoneData[zoneName].forEach(r => {
          if (!supers[r[1]]) supers[r[1]] = [];
          supers[r[1]].push(r);
        });
        for (const [sup, rows] of Object.entries(supers)) {
          const { ws, missing } = buildSheet(rows, csvLookup,
            `OBJECTIFS ${moisUp}  ·  ${zoneName}  ·  ${sup}  ·  ${rows.length} salles`,
            true, false);
          allMissing.push(...missing);
          const wb = XLSX.utils.book_new();
          XLSX.utils.book_append_sheet(wb, ws, sup.slice(0, 31));
          const safeSup = sup.replace(/ /g, '_').replace(/\//g, '-');
          zip.file(`${moisSafe}_${zoneName}_${safeSup}.xlsx`, XLSX.write(wb, { bookType: 'xlsx', type: 'array' }));
          log(`  ✅ ${moisSafe}_${zoneName}_${safeSup}.xlsx`);
        }
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
    <div style={{ maxWidth: 860, margin: '0 auto', padding: '2rem 1rem' }}>
      <h2 style={{ fontFamily: 'monospace', color: 'var(--accent)', marginBottom: '1.5rem' }}>
        OBJECTIFS CAMPRESJ
      </h2>

      {/* Upload */}
      <div style={{ display: 'grid', gridTemplateColumns: '1fr 1fr', gap: '1rem', marginBottom: '1.5rem' }}>
        {[
          { label: 'shift_report_summary.csv', sub: 'Colonnes : Betshop name · Payin money · Payin count · Gross', ref: csvRef, setter: setCsvFile, file: csvFile, accept: '.csv' },
          { label: 'RECAP_DES_SALLES_PAR_ZONE.xlsx', sub: 'Colonnes : Zone · Super en charge · Salle · Nom technique', ref: recapRef, setter: setRecapFile, file: recapFile, accept: '.xlsx' },
        ].map(({ label, sub, ref, setter, file, accept }) => (
          <div key={label}
            onClick={() => ref.current.click()}
            onDragOver={e => e.preventDefault()}
            onDrop={e => { e.preventDefault(); setter(e.dataTransfer.files[0]); }}
            style={{ border: `2px dashed ${file ? 'var(--accent)' : '#444'}`, borderRadius: 8, padding: '1.5rem', textAlign: 'center', cursor: 'pointer', background: 'var(--surface)' }}>
            <input ref={ref} type="file" accept={accept} hidden onChange={e => setter(e.target.files[0])} />
            <div style={{ fontSize: 24, marginBottom: 8 }}>{accept === '.csv' ? '📋' : '📊'}</div>
            <div style={{ fontFamily: 'monospace', fontSize: 12, color: 'var(--fg)' }}>{label}</div>
            <div style={{ fontSize: 11, color: 'var(--muted)', marginTop: 4 }}>{sub}</div>
            {file && <div style={{ marginTop: 8, fontSize: 12, color: 'var(--accent)' }}>✓ {file.name}</div>}
          </div>
        ))}
      </div>

      {/* Paramètres */}
      <div style={{ display: 'flex', gap: '2rem', marginBottom: '1.5rem', flexWrap: 'wrap' }}>
        <label style={{ display: 'flex', flexDirection: 'column', gap: 4, fontSize: 13, color: 'var(--muted)' }}>
          Mois
          <input value={mois} onChange={e => setMois(e.target.value)}
            style={{ padding: '6px 10px', borderRadius: 6, border: '1px solid #444', background: 'var(--surface)', color: 'var(--fg)', fontFamily: 'monospace', fontSize: 13 }} />
        </label>
        <label style={{ display: 'flex', flexDirection: 'column', gap: 4, fontSize: 13, color: 'var(--muted)' }}>
          Taux objectif (%)
          <input type="number" value={taux} onChange={e => setTaux(Number(e.target.value))}
            style={{ padding: '6px 10px', borderRadius: 6, border: '1px solid #444', background: 'var(--surface)', color: 'var(--fg)', fontFamily: 'monospace', fontSize: 13, width: 100 }} />
        </label>
      </div>

      <button onClick={runPipeline} disabled={!csvFile || !recapFile || running}
        style={{ padding: '10px 28px', background: running ? '#444' : 'var(--accent)', color: '#fff', border: 'none', borderRadius: 6, cursor: 'pointer', fontWeight: 'bold', fontSize: 14, marginBottom: '1.5rem' }}>
        {running ? '⏳ Génération...' : '▶ Générer les objectifs'}
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
            style={{ padding: '10px 24px', background: '#27ae60', color: '#fff', border: 'none', borderRadius: 6, cursor: 'pointer', fontWeight: 'bold', fontSize: 14, marginBottom: '1rem' }}>
            ⬇ Télécharger {results.filename}
          </button>
          {results.missing.length > 0 && (
            <div style={{ marginTop: '1rem' }}>
              <p style={{ color: '#e67e22', fontSize: 13, marginBottom: 8 }}>⚠️ Salles sans données CSV :</p>
              <ul style={{ fontFamily: 'monospace', fontSize: 12, color: 'var(--muted)', paddingLeft: 20 }}>
                {results.missing.map(s => <li key={s}>{s}</li>)}
              </ul>
            </div>
          )}
        </div>
      )}
    </div>
  );
}
