// src/pages/BonusAutoPage.js
// Envoi automatique des bonus via le serveur local Playwright
import React, { useState, useRef, useEffect } from 'react';

const SERVER = "http://localhost:5001";

const BONUS_TYPES = [
  { id: 'casino',   label: '🎰 Casino',   desc: 'casino_bonus_cm.csv',         color: '#818CF8' },
  { id: 'cashback', label: '💸 Cashback', desc: 'Daily_Cashback_Template.csv', color: '#34D399' },
];

function LogLine({ entry }) {
  const colors = { info: 'var(--text)', success: '#34D399', warn: '#FBBF24', error: '#F87171' };
  const icons  = { info: '·', success: '✓', warn: '⚠', error: '✗' };
  return (
    <div style={{ display:'flex', gap:8, fontSize:11, fontFamily:'monospace', padding:'2px 0' }}>
      <span style={{ color:'var(--muted)', flexShrink:0 }}>{entry.ts}</span>
      <span style={{ color: colors[entry.level] || colors.info, flexShrink:0 }}>{icons[entry.level]}</span>
      <span style={{ color: colors[entry.level] || colors.info }}>{entry.msg}</span>
    </div>
  );
}


function SportBonusSection({ serverOk, onFilesLoaded, initialFiles }) {
  const [files, setFiles] = useState(initialFiles && initialFiles.length > 0 ? initialFiles : []);
  const [status, setStatus] = useState('idle'); // idle | running | done | error
  const [logs,   setLogs]   = useState([]);
  const fileRef = useRef();
  const pollRef = useRef(null);

  const startPolling = (tid) => {
    pollRef.current = setInterval(async () => {
      try {
        const r = await fetch(`${SERVER}/task/${tid}`);
        const data = await r.json();
        setLogs([...data.logs]);
        if (data.done) {
          clearInterval(pollRef.current);
          setStatus(data.status === 'done' ? 'done' : 'error');
        }
      } catch {}
    }, 800);
  };

  const sendSport = async (f) => {
    setStatus('running');
    setLogs([]);
    try {
      const r = await fetch(`${SERVER}/send/sport`, {
        method: 'POST',
        headers: { 'Content-Type': 'application/json' },
        body: JSON.stringify({ csv_content: f.content, amount_label: f.label }),
      });
      const data = await r.json();
      startPolling(data.task_id);
    } catch(e) {
      setStatus('error');
      setLogs([{ ts:'--:--:--', msg: `Erreur: ${e.message}`, level:'error' }]);
    }
  };

  const resetSport = () => {
    clearInterval(pollRef.current);
    setStatus('idle');
    setLogs([]);
  };

  const addFile = (f) => {
    const reader = new FileReader();
    reader.onload = e => {
      const content = e.target.result;
      // Détecter le label depuis le nom du fichier (100, 200, etc.)
      const match = f.name.match(/(\d+)/);
      const label = match ? match[1] : String(Date.now());
      const newFile = { id: Date.now(), name: f.name, content, label };
      setFiles(prev => {
        const updated = [...prev, newFile];
        onFilesLoaded(updated);
        return updated;
      });
    };
    reader.readAsText(f);
  };

  const removeFile = (id) => {
    setFiles(prev => {
      const updated = prev.filter(f => f.id !== id);
      onFilesLoaded(updated);
      return updated;
    });
  };

  return (
    <div style={{
      background:'var(--bg2)', border:'1px solid var(--border)', borderRadius:12,
      padding:18, borderTop:'3px solid #F97316', marginBottom:14,
    }}>
      <div style={{ display:'flex', justifyContent:'space-between', alignItems:'center', marginBottom:12 }}>
        <div>
          <div style={{ fontSize:14, fontWeight:700, color:'#F97316' }}>⚽ Sport Bonus</div>
          <div style={{ fontSize:11, color:'var(--muted)' }}>
            Un fichier par type (Sport_Daily_100.csv, Sport_Daily_200.csv, etc.)
          </div>
        </div>
        <button
          onClick={() => fileRef.current.click()}
          style={{
            fontSize:11, padding:'6px 12px', borderRadius:7, border:'1px dashed #F97316',
            background:'rgba(249,115,22,.08)', color:'#F97316', cursor:'pointer', fontWeight:600,
          }}>
          + Ajouter un fichier
        </button>
        <input ref={fileRef} type="file" accept=".csv" hidden multiple
          onChange={e => { Array.from(e.target.files).forEach(addFile); e.target.value=''; }}/>
      </div>

      {files.length === 0 ? (
        <div
          onClick={() => fileRef.current.click()}
          onDragOver={e => e.preventDefault()}
          onDrop={e => { e.preventDefault(); Array.from(e.dataTransfer.files).forEach(addFile); }}
          style={{
            border:'2px dashed var(--border)', borderRadius:8, padding:'20px 16px',
            textAlign:'center', cursor:'pointer', color:'var(--muted)', fontSize:12,
          }}>
          <div style={{ fontSize:22, marginBottom:6 }}>📂</div>
          <div>Glisse les fichiers Sport ici ou <span style={{ color:'#F97316', fontWeight:600 }}>clique</span></div>
          <div style={{ fontSize:10, marginTop:4 }}>Sport_Daily_100.csv · Sport_Daily_200.csv · etc.</div>
        </div>
      ) : (
        <div style={{ display:'flex', flexDirection:'column', gap:8, maxHeight:300, overflowY:'auto', paddingRight:4 }}>
          {files.map(f => (
            <div key={f.id} style={{
              background:'var(--bg3)', borderRadius:8, padding:'10px 12px',
            }}>
              <div style={{ display:'flex', alignItems:'center', gap:10 }}>
                <span style={{ fontSize:16 }}>⚽</span>
                <div style={{ flex:1 }}>
                  <div style={{ fontSize:12, fontWeight:600, color:'var(--text)' }}>{f.name}</div>
                  <div style={{ fontSize:10, color:'var(--muted)' }}>
                    Label détecté : <span style={{ color:'#F97316', fontWeight:700 }}>{f.label} XAF</span>
                    {' · '}Template : CM Daily sport bonus {f.label} XAF
                  </div>
                </div>
                {status === 'idle' && (
                  <button onClick={() => sendSport(f)} disabled={!serverOk} style={{
                    fontSize:11, padding:'5px 12px', borderRadius:6, border:'none',
                    cursor: serverOk ? 'pointer' : 'not-allowed',
                    background: serverOk ? '#F97316' : 'var(--bg)',
                    color: serverOk ? '#fff' : 'var(--muted)', fontWeight:700,
                  }}>🚀 Envoyer</button>
                )}
                {status === 'running' && (
                  <span style={{ fontSize:11, color:'#FBBF24', fontWeight:600 }}>⏳ En cours…</span>
                )}
                {status === 'done' && (
                  <span style={{ fontSize:11, color:'#34D399', fontWeight:600 }}>✅ Envoyé</span>
                )}
                {status === 'error' && (
                  <span style={{ fontSize:11, color:'#F87171', fontWeight:600 }}>❌ Erreur</span>
                )}
                {status === 'idle' && (
                  <button onClick={() => removeFile(f.id)} style={{
                    background:'none', border:'none', cursor:'pointer',
                    color:'var(--muted)', fontSize:16, padding:'0 4px',
                  }}>✕</button>
                )}
              </div>
              {/* Logs pour ce fichier */}
              {logs.length > 0 && (
                <div style={{
                  marginTop:8, background:'var(--bg2)', borderRadius:6,
                  padding:'6px 10px', maxHeight:120, overflowY:'auto',
                }}>
                  {logs.map((l, i) => <LogLine key={i} entry={l}/>)}
                </div>
              )}
              {status === 'done' && (
                <div style={{ display:'flex', gap:6, marginTop:6 }}>
                  <button onClick={() => { setStatus('idle'); setLogs([]); }} style={{
                    fontSize:10, padding:'3px 8px', borderRadius:4, border:'none',
                    cursor:'pointer', background:'rgba(52,211,153,.15)', color:'#34D399', fontWeight:700,
                  }}>🔄 Renvoyer</button>
                  <button onClick={resetSport} style={{
                    fontSize:10, padding:'3px 8px', borderRadius:4, border:'none',
                    cursor:'pointer', background:'var(--bg)', color:'var(--muted)',
                  }}>✕ Retirer</button>
                </div>
              )}
              {status === 'error' && (
                <div style={{ display:'flex', gap:6, marginTop:6 }}>
                  <button onClick={() => { setStatus('idle'); setLogs([]); }} style={{
                    fontSize:10, padding:'3px 8px', borderRadius:4, border:'none',
                    cursor:'pointer', background:'rgba(248,113,113,.15)', color:'#F87171', fontWeight:700,
                  }}>🔄 Réessayer</button>
                  <button onClick={resetSport} style={{
                    fontSize:10, padding:'3px 8px', borderRadius:4, border:'none',
                    cursor:'pointer', background:'var(--bg)', color:'var(--muted)',
                  }}>✕ Retirer</button>
                </div>
              )}
            </div>
          ))}
          {status === 'idle' && (
          <button
            onClick={() => fileRef.current.click()}
            style={{
              fontSize:11, padding:'5px 10px', borderRadius:6, border:'1px dashed var(--border)',
              background:'transparent', color:'var(--muted)', cursor:'pointer', alignSelf:'flex-start',
            }}>
            + Ajouter un autre fichier
          </button>
          )}
        </div>
      )}
    </div>
  );
}

function BonusCard({ type, onFileLoaded, initialFile }) {
  const fileRef = useRef();
  const initFile = initialFile ? { name: type.desc, content: initialFile } : null;
  const initPreview = initialFile ? initialFile.split('\n').filter(l => l.trim()).slice(0, 6) : [];
  const [file, setFile]     = useState(initFile);
  const [preview, setPreview] = useState(initPreview);
  const [taskId, setTaskId] = useState(null);
  const [logs, setLogs]     = useState([]);
  const [status, setStatus] = useState('idle'); // idle | running | done | error
  const [serverOk,  setServerOk]  = useState(null);

  // Notifier le parent si initialFile fourni
  React.useEffect(() => {
    if (initialFile) onFileLoaded?.(type.id, initialFile);
  }, []); // eslint-disable-line
  const [chromeCdp, setChromeCdp] = useState(null); // true=ok, false=non, null=check en cours
  const [chromeLaunching, setChromeLaunching] = useState(false);
  const pollRef = useRef(null);

  // Vérifier le serveur au montage
  useEffect(() => {
    fetch(`${SERVER}/status`).then(r => r.json()).then(() => setServerOk(true)).catch(() => setServerOk(false));
  }, []);

  const parseFile = (f) => {
    const reader = new FileReader();
    reader.onload = e => {
      const text = e.target.result;
      setFile({ name: f.name, content: text });
      // Preview : 5 premières lignes
      const lines = text.split('\n').filter(l => l.trim());
      setPreview(lines.slice(0, 6));
      onFileLoaded?.(type.id, text);
    };
    reader.readAsText(f);
  };

  const startPolling = (tid) => {
    pollRef.current = setInterval(async () => {
      try {
        const r = await fetch(`${SERVER}/task/${tid}`);
        const data = await r.json();
        setLogs([...data.logs]);
        if (data.done) {
          clearInterval(pollRef.current);
          setStatus(data.status === 'done' ? 'done' : 'error');
        }
      } catch {}
    }, 800);
  };

  const send = async () => {
    if (!file) return;
    setStatus('running');
    setLogs([]);
    try {
      let endpoint, body;
      if (type.id === 'casino') {
        endpoint = '/send/casino';
        body = { csv_content: file.content };
      } else if (type.id === 'cashback') {
        endpoint = '/send/cashback';
        body = { csv_content: file.content };
      } else {
        endpoint = '/send/sport';
        const label = type.id === 'sport100' ? '100' : '200';
        body = { csv_content: file.content, amount_label: label };
      }
      const r = await fetch(`${SERVER}${endpoint}`, {
        method: 'POST', headers: { 'Content-Type': 'application/json' },
        body: JSON.stringify(body),
      });
      const data = await r.json();
      setTaskId(data.task_id);
      startPolling(data.task_id);
    } catch(e) {
      setStatus('error');
      setLogs([{ ts: '--:--:--', msg: `Serveur local inaccessible : ${e.message}`, level: 'error' }]);
    }
  };

  const reset = () => {
    clearInterval(pollRef.current);
    setFile(null); setPreview([]); setTaskId(null);
    setLogs([]); setStatus('idle');
  };

  const statusColor = { idle:'var(--muted)', running:'#FBBF24', done:'#34D399', error:'#F87171' }[status];
  const statusLabel = { idle:'En attente', running:'En cours…', done:'✅ Envoyé', error:'❌ Erreur' }[status];
  const rowCount = preview.length > 1 ? preview.length - 1 : 0;

  return (
    <div style={{
      background: 'var(--bg2)', border: `1px solid var(--border)`,
      borderRadius: 12, padding: 18, display: 'flex', flexDirection: 'column', gap: 12,
      borderTop: `3px solid ${type.color}`,
    }}>
      {/* Header */}
      <div style={{ display:'flex', justifyContent:'space-between', alignItems:'center' }}>
        <div>
          <div style={{ fontSize: 14, fontWeight: 700, color: type.color }}>{type.label}</div>
          <div style={{ fontSize: 11, color: 'var(--muted)' }}>{type.desc}</div>
        </div>
        <div style={{ display:'flex', alignItems:'center', gap:8 }}>
          {serverOk === false && (
            <span style={{ fontSize:10, color:'#F87171', background:'rgba(248,113,113,.1)',
              padding:'2px 7px', borderRadius:4 }}>Serveur hors ligne</span>
          )}
          <span style={{ fontSize:11, fontWeight:600, color: statusColor }}>{statusLabel}</span>
        </div>
      </div>

      {/* Drop zone / fichier */}
      {!file ? (
        <div
          onClick={() => fileRef.current.click()}
          onDragOver={e => e.preventDefault()}
          onDrop={e => { e.preventDefault(); if(e.dataTransfer.files[0]) parseFile(e.dataTransfer.files[0]); }}
          style={{
            border: '2px dashed var(--border)', borderRadius: 8, padding: '20px 16px',
            textAlign: 'center', cursor: 'pointer', color: 'var(--muted)', fontSize: 12,
            transition: 'border-color .2s',
          }}
          onMouseEnter={e => e.currentTarget.style.borderColor = type.color}
          onMouseLeave={e => e.currentTarget.style.borderColor = 'var(--border)'}
        >
          <div style={{ fontSize: 22, marginBottom: 6 }}>📂</div>
          <div>Glisse le fichier ici ou <span style={{ color: type.color, fontWeight:600 }}>clique</span></div>
          <div style={{ fontSize: 10, marginTop: 4 }}>.csv uniquement</div>
          <input ref={fileRef} type="file" accept=".csv" hidden
            onChange={e => { if(e.target.files[0]) parseFile(e.target.files[0]); e.target.value=''; }}/>
        </div>
      ) : (
        <div style={{ background: 'var(--bg3)', borderRadius: 8, padding: '10px 12px' }}>
          <div style={{ display:'flex', justifyContent:'space-between', alignItems:'center', marginBottom:6 }}>
            <span style={{ fontSize:12, fontWeight:600, color:'var(--text)' }}>📄 {file.name}</span>
            <div style={{ display:'flex', gap:6 }}>
              <span style={{ fontSize:11, color:'var(--muted)' }}>{rowCount} joueurs</span>
              {status === 'idle' && (
                <button onClick={reset} style={{
                  fontSize:11, padding:'3px 9px', borderRadius:5, border:'1px solid #F87171',
                  background:'rgba(248,113,113,.1)', color:'#F87171', cursor:'pointer', fontWeight:600,
                }}>✕ Retirer</button>
              )}
            </div>
          </div>
          {/* Preview */}
          <div style={{ fontFamily:'monospace', fontSize:10, color:'var(--muted)' }}>
            {preview.map((line, i) => (
              <div key={i} style={{ color: i===0 ? type.color : 'var(--muted)', padding:'1px 0' }}>
                {line}
              </div>
            ))}
          </div>
        </div>
      )}

      {/* Logs */}
      {logs.length > 0 && (
        <div style={{
          background: 'var(--bg3)', borderRadius: 8, padding: '8px 12px',
          maxHeight: 160, overflowY: 'auto',
        }}>
          {logs.map((l, i) => <LogLine key={i} entry={l}/>)}
        </div>
      )}

      {/* Boutons */}
      <div style={{ display:'flex', gap:8 }}>
        {status === 'idle' && file && (
          <button onClick={send} disabled={!serverOk} style={{
            flex:1, padding:'8px 0', borderRadius:8, border:'none', cursor: serverOk ? 'pointer' : 'not-allowed',
            background: serverOk ? type.color : 'var(--bg3)',
            color: serverOk ? '#fff' : 'var(--muted)', fontWeight:700, fontSize:13,
          }}>
            🚀 Envoyer
          </button>
        )}
        {status === 'running' && (
          <div style={{
            flex:1, padding:'8px 0', borderRadius:8, background:'rgba(251,191,36,.1)',
            color:'#FBBF24', fontWeight:600, fontSize:13, textAlign:'center',
          }}>
            ⏳ Envoi en cours…
          </div>
        )}
        {status === 'done' && (
          <div style={{ display:'flex', gap:6, flex:1 }}>
            <button onClick={() => { setStatus('idle'); setLogs([]); setTaskId(null); }} style={{
              flex:1, padding:'8px 0', borderRadius:8, border:'none', cursor:'pointer',
              background:'rgba(52,211,153,.15)', color:'#34D399', fontWeight:700, fontSize:13,
            }}>
              🔄 Renvoyer
            </button>
            <button onClick={reset} style={{
              padding:'8px 12px', borderRadius:8, border:'none', cursor:'pointer',
              background:'var(--bg3)', color:'var(--muted)', fontWeight:600, fontSize:12,
            }}>
              ✕ Nouveau
            </button>
          </div>
        )}
        {status === 'error' && (
          <div style={{ display:'flex', gap:6, flex:1 }}>
            <button onClick={() => { setStatus('idle'); setLogs([]); setTaskId(null); }} style={{
              flex:1, padding:'8px 0', borderRadius:8, border:'none', cursor:'pointer',
              background:'rgba(248,113,113,.15)', color:'#F87171', fontWeight:700, fontSize:13,
            }}>
              🔄 Réessayer
            </button>
            <button onClick={reset} style={{
              padding:'8px 12px', borderRadius:8, border:'none', cursor:'pointer',
              background:'var(--bg3)', color:'var(--muted)', fontWeight:600, fontSize:12,
            }}>
              ✕ Nouveau
            </button>
          </div>
        )}
      </div>
    </div>
  );
}

export default function BonusAutoPage() {
  const [serverOk,  setServerOk]  = useState(null);
  const [chromeCdp, setChromeCdp] = useState(null); // true=ok, false=non, null=check en cours
  const [chromeLaunching, setChromeLaunching] = useState(false);
  const [loadedFiles, setLoadedFiles] = useState({});
  const [sportFiles,  setSportFiles]  = useState([]); // [{name, content, label}]
  const [batchStatus,   setBatchStatus]   = useState('idle');
  const [batchTaskId,   setBatchTaskId]   = useState(null);
  const [batchLogs,     setBatchLogs]     = useState([]);
  const [batchProgress, setBatchProgress] = useState(null);
  const batchPollRef = React.useRef(null);

  // Lire le transfert depuis ResultPage (sessionStorage + event)
  const applyTransfer = React.useCallback(() => {
    const transfer = sessionStorage.getItem('bonusTransfer');
    if (!transfer) return;
    try {
      const data = JSON.parse(transfer);
      sessionStorage.removeItem('bonusTransfer');
      if (data.casino_csv)   setLoadedFiles(prev => ({ ...prev, casino: data.casino_csv }));
      if (data.cashback_csv) setLoadedFiles(prev => ({ ...prev, cashback: data.cashback_csv }));
      if (data.sport_files?.length > 0) {
        setSportFiles(data.sport_files.map((sf, idx) => ({
          id: Date.now() + idx, name: sf.filename,
          content: sf.csv_content, label: sf.amount_label,
        })));
      }
    } catch(e) { console.warn('Transfer error:', e); }
  }, []); // eslint-disable-line

  useEffect(() => {
    applyTransfer(); // Au montage
    window.addEventListener('bonusTransfer', applyTransfer); // Depuis Results si déjà monté
    return () => window.removeEventListener('bonusTransfer', applyTransfer);
  }, [applyTransfer]);

  useEffect(() => {
    const check = () => fetch(`${SERVER}/status`)
      .then(r => r.json())
      .then(d => { setServerOk(true); setChromeCdp(d.chrome_cdp === true); })
      .catch(() => { setServerOk(false); setChromeCdp(false); });
    check();
    const interval = setInterval(check, 5000);
    return () => clearInterval(interval);
  }, []);

  const openChromeCdp = async () => {
    setChromeLaunching(true);
    try {
      const r = await fetch(`${SERVER}/open-chrome`, { method: 'POST' });
      const d = await r.json();
      if (d.ok) {
        setChromeCdp(true);
        alert('✅ ' + d.msg + '\n\nConnecte-toi au backoffice dans Chrome si pas déjà fait.');
      } else {
        alert('❌ ' + d.msg);
      }
    } catch(e) {
      alert('❌ Erreur : ' + e.message);
    }
    setChromeLaunching(false);
  };

  const handleFileLoaded = (typeId, content) => {
    setLoadedFiles(prev => ({ ...prev, [typeId]: content }));
  };

  const startBatchPolling = (tid) => {
    if (batchPollRef.current) clearInterval(batchPollRef.current);
    batchPollRef.current = setInterval(async () => {
      try {
        const [tr, sr] = await Promise.all([
          fetch(`${SERVER}/task/${tid}`).then(r => r.json()),
          fetch(`${SERVER}/sequence/${tid}`).then(r => r.json()),
        ]);
        setBatchLogs([...tr.logs]);
        setBatchProgress({ current: sr.current, total: sr.total, results: sr.results || [] });
        if (tr.done) {
          clearInterval(batchPollRef.current);
          setBatchStatus(tr.status);
        }
      } catch {}
    }, 1000);
  };

  const sendAll = async () => {
    if (batchPollRef.current) clearInterval(batchPollRef.current);
    setBatchStatus('running');
    setBatchProgress(null);
    setBatchLogs([{ ts: new Date().toLocaleTimeString(), msg: 'Démarrage envoi séquentiel…', level:'info' }]);
    const body = {};
    if (loadedFiles.casino)   body.casino_csv   = loadedFiles.casino;
    if (loadedFiles.cashback) body.cashback_csv = loadedFiles.cashback;
    body.sport_files = sportFiles.map(f => ({ csv_content: f.content, amount_label: f.label }));
    try {
      const r = await fetch(`${SERVER}/send/all`, {
        method:'POST', headers:{'Content-Type':'application/json'}, body: JSON.stringify(body),
      });
      const data = await r.json();
      setBatchTaskId(data.task_id);
      startBatchPolling(data.task_id);
    } catch(e) {
      setBatchLogs([{ ts:'--', msg: e.message, level:'error' }]);
      setBatchStatus('error');
    }
  };

  const pauseBatch = async () => {
    if (!batchTaskId) return;
    await fetch(`${SERVER}/pause/${batchTaskId}`, { method:'POST' });
    setBatchStatus('pausing');
  };

  const resumeBatch = async () => {
    if (!batchTaskId) return;
    setBatchStatus('running');
    await fetch(`${SERVER}/resume/${batchTaskId}`, { method:'POST' });
    startBatchPolling(batchTaskId);
  };

  const stopBatch = async () => {
    if (!batchTaskId) return;
    await fetch(`${SERVER}/stop/${batchTaskId}`, { method:'POST' });
    if (batchPollRef.current) clearInterval(batchPollRef.current);
    setBatchStatus('stopped');
  };

  const filesReady = Object.keys(loadedFiles).length + sportFiles.length;

  return (
    <div style={{ maxWidth: 1000, margin: '0 auto', padding: '20px 16px' }}>

      {/* Header + status serveur */}
      <div style={{ display:'flex', justifyContent:'space-between', alignItems:'flex-start', marginBottom:20 }}>
        <div>
          <div style={{ fontSize:20, fontWeight:700, color:'var(--text)', marginBottom:4 }}>
            🤖 Envoi automatique des bonus
          </div>
          <div style={{ fontSize:12, color:'var(--muted)' }}>
            Charge les fichiers générés par le Pipeline, puis lance l'envoi en un clic.
          </div>
        </div>
        <div style={{ display:'flex', gap:8, alignItems:'center' }}>
          {/* Statut serveur */}
          <div style={{
            display:'flex', alignItems:'center', gap:6, padding:'6px 12px',
            borderRadius:8, border:'1px solid var(--border)', background:'var(--bg2)',
          }}>
            <div style={{ width:7, height:7, borderRadius:'50%',
              background: serverOk === null ? '#FBBF24' : serverOk ? '#34D399' : '#F87171' }}/>
            <span style={{ fontSize:11, fontWeight:600, color:'var(--text)' }}>Serveur</span>
            <span style={{ fontSize:11, color:'var(--muted)' }}>
              {serverOk ? ':5001' : 'hors ligne'}
            </span>
          </div>
          {/* Statut Chrome CDP + bouton */}
          {serverOk && (
            <div style={{
              display:'flex', alignItems:'center', gap:6, padding:'6px 12px',
              borderRadius:8, border:'1px solid var(--border)', background:'var(--bg2)',
            }}>
              <div style={{ width:7, height:7, borderRadius:'50%',
                background: chromeCdp ? '#34D399' : '#F87171' }}/>
              <span style={{ fontSize:11, fontWeight:600, color:'var(--text)' }}>Chrome</span>
              <span style={{ fontSize:11, color:'var(--muted)' }}>
                {chromeCdp ? 'CDP actif' : 'non connecté'}
              </span>
              {!chromeCdp && (
                <button onClick={openChromeCdp} disabled={chromeLaunching} style={{
                  fontSize:10, padding:'3px 8px', borderRadius:5, border:'none',
                  cursor: chromeLaunching ? 'wait' : 'pointer',
                  background:'#60A5FA', color:'#fff', fontWeight:700,
                }}>
                  {chromeLaunching ? '⏳ Lancement…' : '🚀 Lancer Chrome'}
                </button>
              )}
            </div>
          )}
        </div>
      </div>

      {/* Instruction serveur hors ligne */}
      {serverOk === false && (
        <div style={{
          background:'rgba(248,113,113,.08)', border:'1px solid rgba(248,113,113,.3)',
          borderRadius:10, padding:'14px 16px', marginBottom:20, fontSize:12,
        }}>
          <div style={{ fontWeight:700, color:'#F87171', marginBottom:8 }}>
            ⚠️ Serveur local non détecté
          </div>
          <div style={{ color:'var(--muted)', lineHeight:1.8, marginBottom:12 }}>
            Le serveur local doit tourner pour que l'envoi fonctionne.<br/>
            <strong style={{ color:'var(--text)' }}>Option 1</strong> — Double-clique sur{' '}
            <code style={{ background:'var(--bg3)', padding:'2px 6px', borderRadius:4 }}>
              Lancer_Bonus_Server.command
            </code>{' '}
            (à placer sur ton Bureau)<br/>
            <strong style={{ color:'var(--text)' }}>Option 2</strong> — Terminal :{' '}
            <code style={{ background:'var(--bg3)', padding:'2px 6px', borderRadius:4 }}>
              python3 ~/Downloads/bonus_server.py
            </code>
          </div>
          <div style={{ display:'flex', gap:8, alignItems:'center', flexWrap:'wrap' }}>
            <button
              onClick={() => {
                // Essayer d'ouvrir Terminal directement via lien
                const cmd = 'python3 /Users/abelgaspardkollendame/Desktop/DAILY/cameroon-app/bonus_server.py';
                // Méthode 1 : lien terminal:// (macOS)
                const link = document.createElement('a');
                link.href = `terminal://open?command=${encodeURIComponent(cmd)}`;
                link.click();
                // Méthode 2 : copier dans le presse-papiers en fallback
                setTimeout(() => {
                  navigator.clipboard.writeText(cmd).then(() => {});
                }, 500);
              }}
              style={{
                padding:'7px 14px', borderRadius:7, border:'none', cursor:'pointer',
                background:'#F87171', color:'#fff', fontWeight:700, fontSize:11,
              }}>
              🖥️ Ouvrir Terminal
            </button>
            <button
              onClick={() => {
                const cmd = 'python3 /Users/abelgaspardkollendame/Desktop/DAILY/cameroon-app/bonus_server.py';
                navigator.clipboard.writeText(cmd);
                alert('✅ Commande copiée !\n\nColle dans Terminal (Cmd+V) puis Entrée.');
              }}
              style={{
                padding:'7px 14px', borderRadius:7, border:'none', cursor:'pointer',
                background:'var(--bg)', color:'var(--text)', fontWeight:700, fontSize:11,
                border:'1px solid var(--border)',
              }}>
              📋 Copier
            </button>
            <button
              onClick={() => window.location.reload()}
              style={{
                padding:'7px 14px', borderRadius:7, border:'1px solid var(--border)',
                cursor:'pointer', background:'var(--bg3)', color:'var(--text)',
                fontWeight:600, fontSize:11,
              }}>
              🔄 Vérifier à nouveau
            </button>
            <span style={{ fontSize:10, color:'var(--muted)' }}>
              Le serveur tourne sur localhost:5001
            </span>
          </div>
        </div>
      )}

      {/* Grille Casino + Cashback */}
      <div style={{
        display:'grid', gridTemplateColumns: 'repeat(auto-fit, minmax(280px, 1fr))',
        gap:14, marginBottom:14,
      }}>
        {BONUS_TYPES.map(type => (
          <BonusCard key={type.id} type={type} onFileLoaded={handleFileLoaded}
            initialFile={loadedFiles[type.id] || null}/>
        ))}
      </div>

      {/* Section Sport Bonus — multi-fichiers */}
      <SportBonusSection serverOk={serverOk} onFilesLoaded={(files) => setSportFiles(files)}
        initialFiles={sportFiles}/>

      {/* Bouton Tout envoyer */}
      {filesReady > 1 && (
        <div style={{
          background:'var(--bg2)', border:'1px solid var(--border)', borderRadius:12,
          padding:16, display:'flex', alignItems:'center', gap:14,
        }}>
          <div>
            <div style={{ fontSize:13, fontWeight:700, color:'var(--text)' }}>
              Tout envoyer en une fois
            </div>
            <div style={{ fontSize:11, color:'var(--muted)' }}>
              {filesReady} fichier{filesReady>1?'s':''} prêt{filesReady>1?'s':''} — chaque envoi s'ouvre dans Chrome
            </div>
          </div>
          <button
            onClick={sendAll}
            disabled={!serverOk || batchStatus === 'running'}
            style={{
              marginLeft:'auto', padding:'10px 24px', borderRadius:8, border:'none',
              cursor: serverOk && batchStatus !== 'running' ? 'pointer' : 'not-allowed',
              background: serverOk ? 'var(--accent)' : 'var(--bg3)',
              color: serverOk ? '#fff' : 'var(--muted)', fontWeight:700, fontSize:13,
            }}>
            {batchStatus === 'running' ? '⏳ En cours…' : '🚀 Tout envoyer'}
          </button>
          {/* Boutons Pause / Reprendre / Arrêter */}
          {(batchStatus === 'running' || batchStatus === 'pausing') && (
            <>
              <button onClick={pauseBatch} style={{
                padding:'10px 16px', borderRadius:8, border:'none', cursor:'pointer',
                background:'#FBBF24', color:'#000', fontWeight:700, fontSize:13,
              }}>⏸ Pause</button>
              <button onClick={stopBatch} style={{
                padding:'10px 16px', borderRadius:8, border:'none', cursor:'pointer',
                background:'#F87171', color:'#fff', fontWeight:700, fontSize:13,
              }}>⏹ Arrêter</button>
            </>
          )}
          {batchStatus === 'paused' && (
            <>
              <button onClick={resumeBatch} style={{
                padding:'10px 18px', borderRadius:8, border:'none', cursor:'pointer',
                background:'#34D399', color:'#fff', fontWeight:700, fontSize:13,
              }}>▶ Reprendre</button>
              <button onClick={stopBatch} style={{
                padding:'10px 16px', borderRadius:8, border:'none', cursor:'pointer',
                background:'#F87171', color:'#fff', fontWeight:700, fontSize:13,
              }}>⏹ Arrêter</button>
            </>
          )}
          {batchStatus === 'stopped' && (
            <span style={{ fontSize:12, color:'#F87171', fontWeight:600 }}>⏹ Arrêté</span>
          )}
        </div>
      )}

      {batchProgress && (
        <div style={{ marginTop:8, background:'var(--bg3)', borderRadius:8, padding:'10px 14px' }}>
          <div style={{ display:'flex', justifyContent:'space-between', fontSize:11, marginBottom:6 }}>
            <span style={{ color:'var(--text)', fontWeight:600 }}>
              Progression : {batchProgress.current}/{batchProgress.total}
            </span>
            <span style={{ color:'var(--muted)' }}>
              ✅ {batchProgress.results.filter(r=>r.status==='done').length} réussis
              {batchProgress.results.filter(r=>r.status==='error').length > 0 &&
                ` · ❌ ${batchProgress.results.filter(r=>r.status==='error').length} erreurs`}
            </span>
          </div>
          <div style={{ background:'var(--border)', borderRadius:4, height:6, overflow:'hidden' }}>
            <div style={{
              height:'100%', borderRadius:4, background:'var(--accent)',
              width: `${(batchProgress.current/batchProgress.total)*100}%`,
              transition:'width .3s',
            }}/>
          </div>
        </div>
      )}
      {batchLogs.length > 0 && (
        <div style={{ background:'var(--bg3)', borderRadius:8, padding:'8px 12px', marginTop:10,
          maxHeight:200, overflowY:'auto' }}>
          {batchLogs.map((l,i) => <LogLine key={i} entry={l}/>)}
        </div>
      )}

      {/* Guide */}
      <div style={{
        marginTop:24, background:'var(--bg2)', border:'1px solid var(--border)',
        borderRadius:12, padding:16,
      }}>
        <div style={{ fontSize:12, fontWeight:700, color:'var(--muted)', marginBottom:10,
          textTransform:'uppercase', letterSpacing:'.05em' }}>
          Comment ça marche
        </div>
        {[
          ['1', 'Lance le serveur', 'Terminal → python3 bonus_server.py — à faire une seule fois par session'],
          ['2', 'Charge les fichiers', 'Glisse les CSV générés par le Pipeline dans chaque carte'],
          ['3', 'Lance l\'envoi', 'Clique Envoyer sur chaque bonus ou Tout envoyer — Chrome s\'ouvre automatiquement'],
          ['4', 'Vérifie le log', 'Chaque étape est tracée en temps réel — les erreurs sont signalées'],
        ].map(([n, title, desc]) => (
          <div key={n} style={{ display:'flex', gap:12, marginBottom:10, alignItems:'flex-start' }}>
            <div style={{
              width:24, height:24, borderRadius:'50%', background:'var(--accent)',
              color:'#fff', fontSize:11, fontWeight:700, display:'flex',
              alignItems:'center', justifyContent:'center', flexShrink:0,
            }}>{n}</div>
            <div>
              <div style={{ fontSize:12, fontWeight:600, color:'var(--text)' }}>{title}</div>
              <div style={{ fontSize:11, color:'var(--muted)' }}>{desc}</div>
            </div>
          </div>
        ))}
      </div>
    </div>
  );
}
