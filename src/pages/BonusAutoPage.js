// src/pages/BonusAutoPage.js
// Envoi automatique des bonus via le serveur local Playwright
import React, { useState, useRef, useEffect } from 'react';

const SERVER = "http://localhost:5001";

const BONUS_TYPES = [
  { id: 'casino',    label: '🎰 Casino',    desc: 'casino_bonus_cm.csv',           color: '#818CF8' },
  { id: 'cashback',  label: '💸 Cashback',  desc: 'Daily_Cashback_Template.csv',   color: '#34D399' },
  { id: 'sport100',  label: '⚽ Sport 100', desc: 'Sport_Daily_100.csv',           color: '#F97316' },
  { id: 'sport200',  label: '⚽ Sport 200', desc: 'Sport_Daily_200.csv',           color: '#FBBF24' },
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

function BonusCard({ type, onFileLoaded }) {
  const fileRef = useRef();
  const [file, setFile]     = useState(null);
  const [preview, setPreview] = useState([]);
  const [taskId, setTaskId] = useState(null);
  const [logs, setLogs]     = useState([]);
  const [status, setStatus] = useState('idle'); // idle | running | done | error
  const [serverOk, setServerOk] = useState(null);
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
                  fontSize:10, padding:'2px 7px', borderRadius:4, border:'none',
                  background:'var(--bg)', color:'var(--muted)', cursor:'pointer',
                }}>✕ Changer</button>
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
        {(status === 'done' || status === 'error') && (
          <button onClick={reset} style={{
            flex:1, padding:'8px 0', borderRadius:8, border:'none', cursor:'pointer',
            background:'var(--bg3)', color:'var(--text)', fontWeight:600, fontSize:13,
          }}>
            🔄 Nouveau fichier
          </button>
        )}
      </div>
    </div>
  );
}

export default function BonusAutoPage() {
  const [serverOk, setServerOk] = useState(null);
  const [loadedFiles, setLoadedFiles] = useState({});
  const [batchStatus, setBatchStatus] = useState('idle');
  const [batchLogs, setBatchLogs] = useState([]);

  useEffect(() => {
    const check = () => fetch(`${SERVER}/status`)
      .then(r => r.json()).then(() => setServerOk(true)).catch(() => setServerOk(false));
    check();
    const interval = setInterval(check, 5000);
    return () => clearInterval(interval);
  }, []);

  const handleFileLoaded = (typeId, content) => {
    setLoadedFiles(prev => ({ ...prev, [typeId]: content }));
  };

  const sendAll = async () => {
    setBatchStatus('running');
    setBatchLogs([{ ts: new Date().toLocaleTimeString(), msg: 'Envoi de tous les bonus disponibles…', level:'info' }]);
    const body = {};
    if (loadedFiles.casino)    body.casino_csv    = loadedFiles.casino;
    if (loadedFiles.cashback)  body.cashback_csv  = loadedFiles.cashback;
    body.sport_files = [];
    if (loadedFiles.sport100)  body.sport_files.push({ csv_content: loadedFiles.sport100, amount_label: '100' });
    if (loadedFiles.sport200)  body.sport_files.push({ csv_content: loadedFiles.sport200, amount_label: '200' });

    try {
      const r = await fetch(`${SERVER}/send/all`, {
        method:'POST', headers:{'Content-Type':'application/json'}, body: JSON.stringify(body),
      });
      const data = await r.json();
      const taskIds = Object.values(data);
      setBatchLogs(prev => [...prev, { ts: new Date().toLocaleTimeString(), msg: `${taskIds.length} tâches lancées`, level:'success' }]);
      setBatchStatus('done');
    } catch(e) {
      setBatchLogs(prev => [...prev, { ts: '--', msg: e.message, level:'error' }]);
      setBatchStatus('error');
    }
  };

  const filesReady = Object.keys(loadedFiles).length;

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
        <div style={{
          display:'flex', alignItems:'center', gap:8, padding:'8px 14px',
          borderRadius:8, border:'1px solid var(--border)', background:'var(--bg2)',
        }}>
          <div style={{
            width:8, height:8, borderRadius:'50%',
            background: serverOk === null ? '#FBBF24' : serverOk ? '#34D399' : '#F87171',
          }}/>
          <span style={{ fontSize:12, fontWeight:600, color:'var(--text)' }}>
            Serveur local
          </span>
          <span style={{ fontSize:11, color:'var(--muted)' }}>
            {serverOk === null ? 'Vérification…' : serverOk ? 'Connecté · :5001' : 'Hors ligne'}
          </span>
          {serverOk === false && (
            <button onClick={() => window.location.reload()} style={{
              fontSize:10, padding:'2px 7px', borderRadius:4, border:'none',
              cursor:'pointer', background:'var(--bg)', color:'var(--muted)',
            }}>↻</button>
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
                // Copier la commande dans le presse-papiers
                navigator.clipboard.writeText('python3 ~/Downloads/bonus_server.py');
                alert('Commande copiée !\nColle-la dans un Terminal puis appuie sur Entrée.');
              }}
              style={{
                padding:'7px 14px', borderRadius:7, border:'none', cursor:'pointer',
                background:'#F87171', color:'#fff', fontWeight:700, fontSize:11,
              }}>
              📋 Copier la commande
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

      {/* Grille des 4 types de bonus */}
      <div style={{
        display:'grid', gridTemplateColumns: 'repeat(auto-fit, minmax(280px, 1fr))',
        gap:14, marginBottom:20,
      }}>
        {BONUS_TYPES.map(type => (
          <BonusCard key={type.id} type={type} onFileLoaded={handleFileLoaded}/>
        ))}
      </div>

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
        </div>
      )}

      {batchLogs.length > 0 && (
        <div style={{ background:'var(--bg3)', borderRadius:8, padding:'8px 12px', marginTop:10 }}>
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
