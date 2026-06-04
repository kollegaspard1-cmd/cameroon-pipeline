// src/pages/UploadPage.js
import React, { useState, useCallback, useRef } from 'react';
import { runPipeline, readExcelSheets, readCSV, buildAllExports } from '../lib/pipeline';
import { saveRun } from '../lib/runHistory';
import { loadRulesForCountry, DEFAULT_RULES, applyRules } from '../lib/roundingRules';
import RulesEditor from '../components/RulesEditor';

// ── Countries with flags, currencies and languages ──
const COUNTRIES = [
  { code: 'CM', name: 'Cameroon',     flag: '🇨🇲', currency: 'XAF', prefix: 'CM'  },
  { code: 'TZ', name: 'Tanzania',     flag: '🇹🇿', currency: 'TZS', prefix: 'TZ'  },
  { code: 'KE', name: 'Kenya',        flag: '🇰🇪', currency: 'KES', prefix: 'KE'  },
  { code: 'ZM', name: 'Zambia',       flag: '🇿🇲', currency: 'ZMW', prefix: 'ZM'  },
  { code: 'CD', name: 'DR Congo',     flag: '🇨🇩', currency: 'CDF', prefix: 'CD'  },
  { code: 'CG', name: 'Congo',        flag: '🇨🇬', currency: 'XAF', prefix: 'CG'  },
  { code: 'SN', name: 'Senegal',      flag: '🇸🇳', currency: 'XOF', prefix: 'SN'  },
  { code: 'CI', name: 'Côte d\'Ivoire',flag:'🇨🇮', currency: 'XOF', prefix: 'CI'  },
  { code: 'GH', name: 'Ghana',        flag: '🇬🇭', currency: 'GHS', prefix: 'GH'  },
  { code: 'NG', name: 'Nigeria',      flag: '🇳🇬', currency: 'NGN', prefix: 'NG'  },
  { code: 'UG', name: 'Uganda',       flag: '🇺🇬', currency: 'UGX', prefix: 'UG'  },
  { code: 'RW', name: 'Rwanda',       flag: '🇷🇼', currency: 'RWF', prefix: 'RW'  },
  { code: 'ET', name: 'Ethiopia',     flag: '🇪🇹', currency: 'ETB', prefix: 'ET'  },
  { code: 'MA', name: 'Morocco',      flag: '🇲🇦', currency: 'MAD', prefix: 'MA'  },
  { code: 'ZA', name: 'South Africa', flag: '🇿🇦', currency: 'ZAR', prefix: 'ZA'  },
];

const CURRENCIES = [
  'XAF','XOF','USD','EUR','GBP','TZS','KES','ZMW','CDF','GHS',
  'NGN','UGX','RWF','ETB','MAD','ZAR','CHF','JPY','CAD','AUD',
  'CNY','BRL','INR','MXN','SCR','MUR','KMF','STN','CVE','BWP',
  'NAD','MWK','SZL','LSL','AOA','MZN','BIF','GMD','SLL','LRD',
  'GNF','MRU','HTG','XCD',
];

const TRANSLATIONS = {
  FR: {
    title:           'Daily Retention',
    uploadFiles:     'UPLOAD FILES',
    country:         'Pays',
    parameters:      'PARAMETERS',
    seuil:           'Seuil Yesterday GGR ≥',
    plafond:         'Plafond Sport Bonus ≤',
    defaut:          'défaut',
    saveFirebase:    'Sauvegarder sur Firebase',
    balise:          '🏁 Balise de fin',
    active:          'ACTIVE',
    disabled:        'DÉSACTIVÉE',
    baliseDesc:      'Ajoutée en fin de chaque fichier généré',
    compteBalise:    'Compte balise',
    montant:         'Montant',
    monnaie:         'Monnaie',
    runBtn:          '▶  Lancer le Pipeline',
    running:         '⟳ En cours…',
    console:         'Console',
    overview:        'Pipeline overview',
    step1:           'Nettoyer MTD, YTD, Last 3 days, extraire les téléphones',
    step2:           'Segmenter : casino / sport bonus / MIX, calculer les montants',
    step3:           'Fusionner Today GGR depuis CSV, filtrer les négatifs',
    step4:           'Générer les fichiers campagnes (casino bonus, sport, cashback, Aviator)',
    fileMtd:         'Mtd_players.xlsx',
    fileMtdSub:      'Requiert les onglets MTD, YTD, Last 3 days',
    fileCsv:         'PLAYER_TRANSACTION_REPORT.csv',
    fileCsvSub:      'Colonnes Account ID + Total Gross',
  },
  EN: {
    title:           'Daily Retention',
    uploadFiles:     'UPLOAD FILES',
    country:         'Country',
    parameters:      'PARAMETERS',
    seuil:           'Yesterday GGR threshold ≥',
    plafond:         'Sport Bonus cap ≤',
    defaut:          'default',
    saveFirebase:    'Save run to Firebase',
    balise:          '🏁 End marker',
    active:          'ACTIVE',
    disabled:        'DISABLED',
    baliseDesc:      'Appended to end of each generated file',
    compteBalise:    'Marker account',
    montant:         'Amount',
    monnaie:         'Currency',
    runBtn:          '▶  Run Pipeline',
    running:         '⟳ Running…',
    console:         'Console',
    overview:        'Pipeline overview',
    step1:           'Clean MTD, YTD, Last 3 days sheets, extract phones',
    step2:           'Segment players: casino / sport bonus / MIX, assign offers & amounts',
    step3:           'Merge Today GGR from CSV, filter negative',
    step4:           'Generate campaign files (casino bonus, daily sport, cashback, Aviator freebets)',
    fileMtd:         'Mtd_players.xlsx',
    fileMtdSub:      'Requires MTD, YTD, Last 3 days sheets',
    fileCsv:         'PLAYER_TRANSACTION_REPORT.csv',
    fileCsvSub:      'Needs Account ID + Total Gross columns',
  },
};

export default function UploadPage({ onResult }) {
  const [lang,      setLang]      = useState('FR');
  const [country,   setCountry]   = useState('CM');
  const [excelFile, setExcelFile] = useState(null);
  const [csvFile,   setCsvFile]   = useState(null);
  const [running,   setRunning]   = useState(false);
  const [progress,  setProgress]  = useState(0);
  const [logs,      setLogs]      = useState([]);
  const [error,     setError]     = useState('');
  const [saveToFb,  setSaveToFb]  = useState(true);
  const [threshold, setThreshold] = useState(2000);
  const [maxSport,  setMaxSport]  = useState(30000);

  // Balise de fin
  const [sentinelActive,   setSentinelActive]   = useState(true);
  const [sentinelAccount,  setSentinelAccount]  = useState(1326274);
  const [sentinelAmount,   setSentinelAmount]   = useState(1);
  const [sentinelCurrency, setSentinelCurrency] = useState('XAF');

  const [roundingRules, setRoundingRules] = useState([...DEFAULT_RULES]);
  const [rulesLoaded,   setRulesLoaded]   = useState(false);
  const [showRulesEditor, setShowRulesEditor] = useState(false);

  const excelRef = useRef();
  const csvRef   = useRef();

  const T = TRANSLATIONS[lang];
  const currentCountry = COUNTRIES.find(c => c.code === country) || COUNTRIES[0];

  // Auto-update currency and load rules when country changes
  const handleCountryChange = async (code) => {
    setCountry(code);
    const c = COUNTRIES.find(x => x.code === code);
    if (c) setSentinelCurrency(c.currency);
    // Load rules from Firebase
    try {
      const saved = await loadRulesForCountry(code);
      setRoundingRules(saved || [...DEFAULT_RULES]);
      setRulesLoaded(true);
    } catch(e) {
      setRoundingRules([...DEFAULT_RULES]);
      setRulesLoaded(true);
    }
  };

  // Load rules for default country on mount
  React.useEffect(() => {
    handleCountryChange('CM');
  }, []); // eslint-disable-line

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
    setError(''); setLogs([]); setRunning(true); setProgress(0);
    try {
      log('Reading Excel file…');
      const excelBuf = await excelFile.arrayBuffer();
      const sheets = readExcelSheets(excelBuf);
      const sheetNames = Object.keys(sheets);
      log(`Sheets found: ${sheetNames.join(', ')}`, 'info');
      for (const required of ['MTD', 'YTD', 'Last 3 days']) {
        if (!sheets[required]) throw new Error(`Missing sheet "${required}" in Excel file`);
      }
      log('Reading CSV file…');
      const csvText = await csvFile.text();
      const csvRows = readCSV(csvText);
      log(`CSV rows: ${csvRows.length}`, 'info');

      const sentinelConfig = sentinelActive ? {
        account: sentinelAccount, amount: sentinelAmount, currency: sentinelCurrency,
      } : null;

      const countryConfig = { code: country, prefix: currentCountry.prefix, currency: sentinelCurrency };

      if (sentinelActive) {
        log(`Balise de fin — Compte: ${sentinelAccount}, Montant: ${sentinelAmount} ${sentinelCurrency}`, 'info');
      }

      const result = runPipeline(sheets, csvRows, (msg, pct) => { log(msg); setProgress(pct); }, threshold, maxSport, sentinelConfig, roundingRules);
      log(`Pipeline done — ${result.stats.totalPlayers} players`, 'info');
      log(`Phones extracted: ${result.stats.phonesExtracted}`, 'info');
      log(`Today < 0 removed: ${result.stats.removedToday}`, 'warn');
      log('Building output files…');
      const outputFiles = buildAllExports(result, sentinelConfig, countryConfig);
      log(`${outputFiles.length} files ready`, 'info');

      if (saveToFb) {
        log('Uploading to Firebase…');
        try {
          const { runId } = await saveRun({ stats: result.stats, fileNames: { excel: excelFile.name, csv: csvFile.name }, outputFiles });
          log(`Saved as run ${runId}`, 'info');
        } catch (fbErr) { log(`Firebase save failed: ${fbErr.message}`, 'warn'); }
      }
      onResult({ ...result, outputFiles });
    } catch (err) {
      setError(err.message);
      log(err.message, 'error');
    } finally { setRunning(false); }
  }

  return (
    <div>
      {/* Language selector - top right */}
      <div style={{ display: 'flex', justifyContent: 'flex-end', marginBottom: '1rem', gap: 6 }}>
        {['FR','EN'].map(l => (
          <button key={l} onClick={() => setLang(l)}
            style={{ padding: '4px 12px', borderRadius: 6, fontSize: 12, fontWeight: 700, cursor: 'pointer',
              border: `1px solid ${lang===l ? 'var(--accent)' : 'var(--border)'}`,
              background: lang===l ? 'var(--accent)' : 'transparent',
              color: lang===l ? '#fff' : 'var(--muted)' }}>
            {l}
          </button>
        ))}
      </div>

      {/* Upload section with country selector */}
      <div style={{ display: 'flex', alignItems: 'center', justifyContent: 'space-between', marginBottom: '0.75rem' }}>
        <p className="section-label" style={{ margin: 0 }}>{T.uploadFiles}</p>
        <div style={{ display: 'flex', alignItems: 'center', gap: 8 }}>
          <span style={{ fontSize: 12, color: 'var(--muted)', fontWeight: 600 }}>{T.country} :</span>
          <div style={{ position: 'relative' }}>
            <select value={country} onChange={e => handleCountryChange(e.target.value)}
              className="field-input"
              style={{ paddingLeft: 28, cursor: 'pointer', minWidth: 160, appearance: 'none', WebkitAppearance: 'none' }}>
              {COUNTRIES.map(c => (
                <option key={c.code} value={c.code}>{c.flag} {c.name}</option>
              ))}
            </select>
            <span style={{ position: 'absolute', left: 8, top: '50%', transform: 'translateY(-50%)', pointerEvents: 'none', fontSize: 14 }}>
              {currentCountry.flag}
            </span>
          </div>
        </div>
      </div>

      {/* Upload cards */}
      <div style={{ display: 'grid', gridTemplateColumns: '1fr 1fr', gap: '1rem', marginBottom: '1.5rem' }}>
        <div className={`drop-zone ${excelFile ? 'has-file' : ''}`}
          onDragOver={e => e.preventDefault()}
          onDrop={e => handleDrop(e, setExcelFile)}
          onClick={() => excelRef.current.click()}>
          <input ref={excelRef} type="file" accept=".xlsx,.xls" hidden onChange={e => setExcelFile(e.target.files[0])} />
          <div className="drop-icon">📊</div>
          <div className="drop-label">{T.fileMtd}</div>
          <div className="drop-sub">{T.fileMtdSub}</div>
          {excelFile && <div className="file-chip">✓ {excelFile.name}</div>}
        </div>
        <div className={`drop-zone ${csvFile ? 'has-file' : ''}`}
          onDragOver={e => e.preventDefault()}
          onDrop={e => handleDrop(e, setCsvFile)}
          onClick={() => csvRef.current.click()}>
          <input ref={csvRef} type="file" accept=".csv" hidden onChange={e => setCsvFile(e.target.files[0])} />
          <div className="drop-icon">📋</div>
          <div className="drop-label">{T.fileCsv}</div>
          <div className="drop-sub">{T.fileCsvSub}</div>
          {csvFile && <div className="file-chip">✓ {csvFile.name}</div>}
        </div>
      </div>

      {/* Parameters */}
      <div className="card" style={{ marginBottom: '1.25rem' }}>
        <p className="section-label">{T.parameters}</p>
        <div style={{ display: 'flex', gap: '2rem', flexWrap: 'wrap', alignItems: 'flex-end' }}>
          <label style={{ display: 'flex', flexDirection: 'column', gap: 6 }}>
            <span style={{ fontSize: 12, color: 'var(--muted)', fontWeight: 600 }}>{T.seuil}</span>
            <div style={{ display: 'flex', alignItems: 'center', gap: 8 }}>
              <input type="number" value={threshold} onChange={e => setThreshold(Number(e.target.value))}
                className="field-input" style={{ width: 110 }} />
              <span style={{ fontSize: 11, color: 'var(--muted2)' }}>{T.defaut} : 2 000</span>
            </div>
          </label>
          <label style={{ display: 'flex', flexDirection: 'column', gap: 6 }}>
            <span style={{ fontSize: 12, color: 'var(--muted)', fontWeight: 600 }}>{T.plafond}</span>
            <div style={{ display: 'flex', alignItems: 'center', gap: 8 }}>
              <input type="number" value={maxSport} onChange={e => setMaxSport(Number(e.target.value))}
                className="field-input" style={{ width: 110 }} />
              <span style={{ fontSize: 11, color: 'var(--muted2)' }}>{T.defaut} : 30 000</span>
            </div>
          </label>
          <label style={{ display: 'flex', alignItems: 'center', gap: 8, fontSize: 13, color: 'var(--muted)', cursor: 'pointer', paddingBottom: 4 }}>
            <input type="checkbox" checked={saveToFb} onChange={e => setSaveToFb(e.target.checked)}
              style={{ accentColor: 'var(--accent)', width: 15, height: 15 }} />
            {T.saveFirebase}
          </label>
          <button onClick={() => setShowRulesEditor(true)}
            style={{ padding:'7px 14px', borderRadius:6, fontSize:12, fontWeight:600, cursor:'pointer',
              border:'1px solid var(--accent)', background:'var(--accent-dim)', color:'var(--accent)',
              marginLeft:'auto' }}>
            ⚙️ {lang === 'FR' ? 'Regles arrondi' : 'Rounding rules'}
            <span style={{ marginLeft:6, fontSize:10, padding:'1px 6px', borderRadius:10,
              background:'var(--accent)', color:'#fff' }}>
              {roundingRules.length}
            </span>
          </button>
        </div>
      </div>

      {/* Rules Editor Modal */}
      {showRulesEditor && (
        <RulesEditor
          rules={roundingRules}
          onChange={setRoundingRules}
          countryCode={country}
          lang={lang}
          onClose={() => setShowRulesEditor(false)}
        />
      )}

      {/* Balise de fin */}
      <div className="card" style={{ marginBottom: '1.25rem', borderColor: sentinelActive ? 'var(--deposit)' : 'var(--border)' }}>
        <div style={{ display: 'flex', alignItems: 'center', justifyContent: 'space-between', marginBottom: sentinelActive ? '1rem' : 0 }}>
          <div style={{ display: 'flex', alignItems: 'center', gap: 10 }}>
            <label style={{ display: 'flex', alignItems: 'center', gap: 8, cursor: 'pointer' }}>
              <input type="checkbox" checked={sentinelActive} onChange={e => setSentinelActive(e.target.checked)}
                style={{ accentColor: 'var(--deposit)', width: 16, height: 16 }} />
              <span style={{ fontSize: 13, fontWeight: 700, color: sentinelActive ? 'var(--deposit)' : 'var(--muted)' }}>
                {T.balise}
              </span>
            </label>
            <span style={{ fontSize: 10, fontWeight: 700, padding: '2px 8px', borderRadius: 100,
              background: sentinelActive ? 'var(--deposit-bg)' : 'var(--bg3)',
              color: sentinelActive ? 'var(--deposit)' : 'var(--muted)' }}>
              {sentinelActive ? T.active : T.disabled}
            </span>
          </div>
          <span style={{ fontSize: 11, color: 'var(--muted)', fontStyle: 'italic' }}>{T.baliseDesc}</span>
        </div>

        {sentinelActive && (
          <div style={{ display: 'flex', gap: '1.5rem', flexWrap: 'wrap', alignItems: 'flex-end' }}>
            <label style={{ display: 'flex', flexDirection: 'column', gap: 6 }}>
              <span style={{ fontSize: 12, color: 'var(--muted)', fontWeight: 600 }}>{T.compteBalise}</span>
              <div style={{ display: 'flex', alignItems: 'center', gap: 8 }}>
                <input type="number" value={sentinelAccount} onChange={e => setSentinelAccount(Number(e.target.value))}
                  className="field-input" style={{ width: 130 }} />
                <span style={{ fontSize: 11, color: 'var(--muted2)' }}>{T.defaut} : 1 326 274</span>
              </div>
            </label>
            <label style={{ display: 'flex', flexDirection: 'column', gap: 6 }}>
              <span style={{ fontSize: 12, color: 'var(--muted)', fontWeight: 600 }}>{T.montant}</span>
              <div style={{ display: 'flex', alignItems: 'center', gap: 8 }}>
                <input type="number" value={sentinelAmount} onChange={e => setSentinelAmount(Number(e.target.value))}
                  className="field-input" style={{ width: 90 }} />
                <span style={{ fontSize: 11, color: 'var(--muted2)' }}>{T.defaut} : 1</span>
              </div>
            </label>
            <label style={{ display: 'flex', flexDirection: 'column', gap: 6 }}>
              <span style={{ fontSize: 12, color: 'var(--muted)', fontWeight: 600 }}>{T.monnaie}</span>
              <select value={sentinelCurrency} onChange={e => setSentinelCurrency(e.target.value)}
                className="field-input" style={{ width: 100, cursor: 'pointer' }}>
                {CURRENCIES.map(c => (
                  <option key={c} value={c}>{c}</option>
                ))}
              </select>
            </label>
          </div>
        )}
      </div>

      <button className="btn primary" disabled={!excelFile || !csvFile || running} onClick={runAll}
        style={{ marginBottom: '1.5rem', padding: '11px 28px', fontSize: 14 }}>
        {running ? T.running : T.runBtn}
      </button>

      {running && (
        <div className="progress-wrap">
          <div className="progress-label">{progress}%</div>
          <div className="progress-track"><div className="progress-fill" style={{ width: `${progress}%` }} /></div>
        </div>
      )}

      {error && (
        <div style={{ padding: '10px 14px', background: 'var(--danger-bg)', border: '1px solid var(--danger)', borderRadius: 'var(--radius)', color: 'var(--danger)', fontSize: 13, marginBottom: '1rem' }}>
          ✗ {error}
        </div>
      )}

      {logs.length > 0 && (
        <>
          <p className="section-label" style={{ marginTop: '1.5rem' }}>{T.console}</p>
          <div className="log-console">
            {logs.map((l, i) => (
              <span key={i} className={`log-line ${l.type}`}>
                <span style={{ color: '#475569', marginRight: 8 }}>[{l.t}]</span>{l.msg}{'\n'}
              </span>
            ))}
          </div>
        </>
      )}

      <hr className="divider" />

      <div className="card" style={{ marginTop: 0 }}>
        <p className="section-label">{T.overview}</p>
        <div style={{ display: 'flex', flexDirection: 'column', gap: 10 }}>
          {[
            ['Step 1', T.step1],
            ['Step 2', T.step2],
            ['Step 3', T.step3],
            ['Step 4', T.step4],
          ].map(([step, desc]) => (
            <div key={step} style={{ display: 'flex', gap: 12, alignItems: 'flex-start' }}>
              <span style={{ fontSize: 11, fontWeight: 700, color: 'var(--accent)', background: 'var(--accent-dim)', padding: '2px 10px', borderRadius: 100, whiteSpace: 'nowrap', flexShrink: 0 }}>{step}</span>
              <span style={{ fontSize: 13, color: 'var(--text2)', lineHeight: 1.6 }}>{desc}</span>
            </div>
          ))}
        </div>
      </div>
    </div>
  );
}
