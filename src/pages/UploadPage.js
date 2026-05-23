// src/pages/UploadPage.js
import React, { useState, useCallback, useRef } from 'react';
import { runPipeline, readExcelSheets, readCSV, buildAllExports } from '../lib/pipeline';
import { saveRun } from '../lib/runHistory';

export default function UploadPage({ onResult }) {
  const [excelFile, setExcelFile] = useState(null);
  const [csvFile,   setCsvFile]   = useState(null);
  const [running,   setRunning]   = useState(false);
  const [progress,  setProgress]  = useState(0);
  const [logs,      setLogs]      = useState([]);
  const [error,     setError]     = useState('');
  const [saveToFb,  setSaveToFb]  = useState(true);
  const [threshold, setThreshold] = useState(2000);
  const [maxSport,  setMaxSport]  = useState(30000);

  const excelRef = useRef();
  const csvRef   = useRef();

  const log = useCallback((msg, type = 'normal') => {
    setLogs(l => [...l, { msg, type, t: new Date().toLocaleTimeString() }]);
  }, []);

  function handleDrop(e, setter) {
    e.preventDefault();
    const f = e.dataTransfer.files[0];
    if (f) setter(f);
  }

  async function runAll() {
    if (!excelFile || !csvFile) return;
    setError('');
    setLogs([]);
    setRunning(true);
    setProgress(0);

    try {
      log('Reading Excel file…');
      const excelBuf = await excelFile.arrayBuffer();
      const sheets = readExcelSheets(excelBuf);
      const sheetNames = Object.keys(sheets);
      log(`Sheets found: ${sheetNames.join(', ')}`, 'info');

      for (const required of ['MTD', 'YTD', 'Last 3 days']) {
        if (!sheets[required]) {
          throw new Error(`Missing sheet "${required}" in Excel file`);
        }
      }

      log('Reading CSV file…');
      const csvText = await csvFile.text();
      const csvRows = readCSV(csvText);
      log(`CSV rows: ${csvRows.length}`, 'info');

      const result = runPipeline(sheets, csvRows, (msg, pct) => {
        log(msg);
        setProgress(pct);
      }, threshold, maxSport);

      log(`Pipeline done — ${result.stats.totalPlayers} players`, 'info');
      log(`Phones extracted: ${result.stats.phonesExtracted}`, 'info');
      log(`Today < 0 removed: ${result.stats.removedToday}`, 'warn');

      // Build export blobs
      log('Building output files…');
      const outputFiles = buildAllExports(result);
      log(`${outputFiles.length} files ready`, 'info');

      // Save to Firebase
      if (saveToFb) {
        log('Uploading to Firebase…');
        try {
          const { runId } = await saveRun({
            stats: result.stats,
            fileNames: { excel: excelFile.name, csv: csvFile.name },
            outputFiles
          });
          log(`Saved as run ${runId}`, 'info');
        } catch (fbErr) {
          log(`Firebase save failed: ${fbErr.message}`, 'warn');
        }
      }

      onResult({ ...result, outputFiles });
    } catch (err) {
      setError(err.message);
      log(err.message, 'error');
    } finally {
      setRunning(false);
    }
  }

  return (
    <div>
      <p className="section-label">Step 1 — Upload files</p>

      <div style={{ display: 'grid', gridTemplateColumns: '1fr 1fr', gap: '1rem', marginBottom: '1rem' }}>

        {/* Excel drop */}
        <div
          className={`drop-zone ${excelFile ? 'has-file' : ''}`}
          onDragOver={e => e.preventDefault()}
          onDrop={e => handleDrop(e, setExcelFile)}
          onClick={() => excelRef.current.click()}
        >
          <input ref={excelRef} type="file" accept=".xlsx,.xls" hidden onChange={e => setExcelFile(e.target.files[0])} />
          <div className="drop-icon">📊</div>
          <div className="drop-label">Cameroon_mtd_players.xlsx</div>
          <div className="drop-sub">Drag & drop or click — requires MTD, YTD, Last 3 days sheets</div>
          {excelFile && <div className="file-chip">✓ {excelFile.name}</div>}
        </div>

        {/* CSV drop */}
        <div
          className={`drop-zone ${csvFile ? 'has-file' : ''}`}
          onDragOver={e => e.preventDefault()}
          onDrop={e => handleDrop(e, setCsvFile)}
          onClick={() => csvRef.current.click()}
        >
          <input ref={csvRef} type="file" accept=".csv" hidden onChange={e => setCsvFile(e.target.files[0])} />
          <div className="drop-icon">📋</div>
          <div className="drop-label">PLAYER_TRANSACTION_REPORT.csv</div>
          <div className="drop-sub">Drag & drop or click — needs Account ID + Total Gross columns</div>
          {csvFile && <div className="file-chip">✓ {csvFile.name}</div>}
        </div>
      </div>

      <div style={{ display: 'flex', alignItems: 'center', gap: '1rem', marginBottom: '1rem' }}>
        <label style={{ fontSize: '13px', color: 'var(--muted)' }}>Seuil Yesterday GGR ≥</label>
        <input
          type="number"
          value={threshold}
          onChange={e => setThreshold(Number(e.target.value))}
          style={{ width: '110px', padding: '6px 10px', background: 'var(--bg3)', border: '1px solid var(--border2)', borderRadius: '6px', color: 'var(--text)', fontFamily: 'inherit', fontSize: '13px' }}
        />
        <span style={{ fontSize: '12px', color: 'var(--muted)' }}>(défaut : 2000)</span>
      </div>

      <div style={{ display: 'flex', alignItems: 'center', gap: '1rem', marginBottom: '1rem' }}>
        <label style={{ fontSize: '13px', color: 'var(--muted)' }}>Plafond Sport Bonus ≤</label>
        <input
          type="number"
          value={maxSport}
          onChange={e => setMaxSport(Number(e.target.value))}
          style={{ width: '110px', padding: '6px 10px', background: 'var(--bg3)', border: '1px solid var(--border2)', borderRadius: '6px', color: 'var(--text)', fontFamily: 'inherit', fontSize: '13px' }}
        />
        <span style={{ fontSize: '12px', color: 'var(--muted)' }}>(défaut : 30 000)</span>
      </div>

      <div style={{ display: 'flex', alignItems: 'center', gap: '1rem', marginBottom: '1.5rem' }}>
        <label style={{ display: 'flex', alignItems: 'center', gap: '8px', fontSize: '13px', color: 'var(--muted)', cursor: 'pointer' }}>
          <input
            type="checkbox"
            checked={saveToFb}
            onChange={e => setSaveToFb(e.target.checked)}
            style={{ accentColor: 'var(--accent)' }}
          />
          Save run to Firebase (history + output files)
        </label>
      </div>

      <button
        className="btn primary"
        disabled={!excelFile || !csvFile || running}
        onClick={runAll}
        style={{ marginBottom: '1.5rem' }}
      >
        {running ? '⟳ Running…' : '▶ Run Pipeline'}
      </button>

      {running && (
        <div className="progress-wrap">
          <div className="progress-label">{progress}%</div>
          <div className="progress-track">
            <div className="progress-fill" style={{ width: `${progress}%` }} />
          </div>
        </div>
      )}

      {error && (
        <div style={{ padding: '10px 14px', background: 'rgba(255,77,77,0.1)', border: '1px solid var(--danger)', borderRadius: '8px', color: 'var(--danger)', fontSize: '13px', marginBottom: '1rem' }}>
          ✗ {error}
        </div>
      )}

      {logs.length > 0 && (
        <>
          <p className="section-label" style={{ marginTop: '1.5rem' }}>Console</p>
          <div className="log-console">
            {logs.map((l, i) => (
              <span key={i} className={`log-line ${l.type}`}>
                <span style={{ color: 'var(--muted)', marginRight: '8px' }}>[{l.t}]</span>
                {l.msg}
              </span>
            ))}
          </div>
        </>
      )}

      <hr className="divider" />

      <div className="card" style={{ marginTop: 0 }}>
        <p className="section-label">Pipeline overview</p>
        <div style={{ fontSize: '13px', color: 'var(--muted)', lineHeight: 2 }}>
          <b style={{ color: 'var(--text)' }}>Step 1</b> — Clean MTD, YTD, Last 3 days sheets, extract phones<br/>
          <b style={{ color: 'var(--text)' }}>Step 2</b> — Segment players: casino / sport bonus / MIX, assign offers &amp; amounts<br/>
          <b style={{ color: 'var(--text)' }}>Step 3</b> — Merge Today GGR from CSV, filter negative<br/>
          <b style={{ color: 'var(--text)' }}>Step 4</b> — Generate 4+ campaign files (casino bonus, daily sport, cashback, Aviator freebets)
        </div>
      </div>
    </div>
  );
}
