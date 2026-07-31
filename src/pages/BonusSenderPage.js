// src/pages/BonusSenderPage.js
// Automatisation envoi bonus Meridianbet BO
// Nécessite meridian_server.py en local sur http://localhost:5757

import React, { useState, useEffect, useRef } from 'react';
import * as XLSX from 'xlsx';
import Papa from 'papaparse';

const SERVER = "http://localhost:5757";
const POLL_INTERVAL = 3000; // vérifier le serveur toutes les 3s

function parseFile(file) {
  return new Promise((res) => {
    const ext = file.name.split('.').pop().toLowerCase();
    if (ext === 'csv') {
      Papa.parse(file, { header: true, skipEmptyLines: true, complete: r => res(r.data) });
    } else {
      const rd = new FileReader();
      rd.onload = e => {
        const wb = XLSX.read(e.target.result, { type: 'array' });
        res(XLSX.utils.sheet_to_json(wb.Sheets[wb.SheetNames[0]], { defval: '' }));
      };
      rd.readAsArrayBuffer(file);
    }
  });
}

export default function BonusSenderPage() {
  const [serverOnline, setServerOnline] = useState(false);
  const [serverStatus, setServerStatus] = useState(null);
  const [username, setUsername]         = useState('');
  const [password, setPassword]         = useState('');
  const [loggedIn, setLoggedIn]         = useState(false);
  const [loading, setLoading]           = useState('');
  const [message, setMessage]           = useState(null); // {type:'ok'|'err'|'warn', text}
  const [logs, setLogs]                 = useState([]);
  const [activeBonus, setActiveBonus]   = useState('sport');
  const [bonusData, setBonusData]       = useState({ sport: [], casino: [], deposit: [] });
  const [bonusFiles, setBonusFiles]     = useState({ sport: null, casino: null, deposit: null });

  const sportRef   = useRef();
  const casinoRef  = useRef();
  const depositRef = useRef();
  const refs = { sport: sportRef, casino: casinoRef, deposit: depositRef };

  // ── Polling statut serveur ──
  useEffect(() => {
    const check = async () => {
      try {
        const r = await fetch(`${SERVER}/status`, { signal: AbortSignal.timeout(2000) });
        const d = await r.json();
        setServerOnline(true);
        setServerStatus(d);
        setLoggedIn(d.logged_in);
        if (d.log) setLogs(d.log);
      } catch {
        setServerOnline(false);
        setServerStatus(null);
      }
    };
    check();
    const id = setInterval(check, POLL_INTERVAL);
    return () => clearInterval(id);
  }, []);

  const msg = (type, text) => setMessage({ type, text });

  // ── Login ──
  const handleLogin = async () => {
    if (!username || !password) return msg('err', 'Entre ton username et password');
    setLoading('login');
    msg(null, null);
    try {
      const r = await fetch(`${SERVER}/login`, {
        method: 'POST',
        headers: { 'Content-Type': 'application/json' },
        body: JSON.stringify({ username, password }),
      });
      const d = await r.json();
      if (d.ok) { setLoggedIn(true); msg('ok', 'Connecté au backoffice Meridianbet'); }
      else msg('err', d.error || 'Erreur de connexion');
    } catch (e) {
      msg('err', `Erreur réseau : ${e.message}`);
    }
    setLoading('');
  };

  // ── Chargement fichier bonus ──
  const handleFile = async (type, file) => {
    if (!file) return;
    const data = await parseFile(file);
    setBonusFiles(f => ({ ...f, [type]: file.name }));
    setBonusData(d => ({ ...d, [type]: data }));
    msg('ok', `${data.length} lignes chargées pour ${type.toUpperCase()}`);
  };

  // ── Envoi bonus ──
  const handleSend = async (type) => {
    const data = bonusData[type];
    if (!data?.length) return msg('err', `Charge d'abord un fichier ${type.toUpperCase()}`);
    if (!loggedIn) return msg('err', 'Connecte-toi d\'abord au backoffice');

    setLoading(type);
    msg(null, null);
    try {
      const r = await fetch(`${SERVER}/send-bonus`, {
        method: 'POST',
        headers: { 'Content-Type': 'application/json' },
        body: JSON.stringify({ type, data }),
      });
      const d = await r.json();
      if (d.ok) msg('ok', d.message || `Bonus ${type} envoyés avec succès`);
      else msg('err', d.error || 'Erreur lors de l\'envoi');
    } catch (e) {
      msg('err', `Erreur réseau : ${e.message}`);
    }
    setLoading('');
  };

  // ── Screenshot ──
  const handleScreenshot = async () => {
    window.open(`${SERVER}/screenshot`, '_blank');
  };

  // ── Close browser ──
  const handleClose = async () => {
    await fetch(`${SERVER}/close`, { method: 'POST' });
    setLoggedIn(false);
  };

  const bonusTypes = [
    { k: 'sport',   label: '⚽ Sport',   color: '#22C55E', desc: 'CSV : id, amount' },
    { k: 'casino',  label: '🎰 Casino',  color: '#A78BFA', desc: 'CSV : id, amount, currency' },
    { k: 'deposit', label: '💳 Deposit', color: '#60A5FA', desc: 'CSV : id, amount' },
  ];

  const statusColor = serverOnline ? '#22C55E' : '#F87171';
  const statusLabel = serverOnline ? '🟢 Serveur connecté' : '🔴 Serveur hors ligne';

  return (
    <div style={{ maxWidth: 800, margin: '0 auto', padding: '20px 16px' }}>

      {/* ── En-tête statut serveur ── */}
      <div style={{ display: 'flex', alignItems: 'center', justifyContent: 'space-between',
        background: 'var(--bg2)', border: `1px solid ${serverOnline ? 'rgba(34,197,94,.3)' : 'rgba(248,113,113,.3)'}`,
        borderRadius: 12, padding: '12px 16px', marginBottom: 20 }}>
        <div>
          <div style={{ fontSize: 13, fontWeight: 700, color: statusColor }}>{statusLabel}</div>
          <div style={{ fontSize: 11, color: 'var(--muted)', marginTop: 2 }}>
            {serverOnline
              ? `localhost:5757 · ${loggedIn ? '✅ Connecté au BO' : '⏳ Non connecté au BO'}`
              : 'Lance : python3 meridian_server.py'}
          </div>
        </div>
        {serverOnline && (
          <div style={{ display: 'flex', gap: 6 }}>
            <button onClick={handleScreenshot} style={{
              fontSize: 11, padding: '5px 10px', borderRadius: 6, border: '1px solid var(--border)',
              background: 'var(--bg3)', color: 'var(--text)', cursor: 'pointer' }}>
              📸 Capture
            </button>
            {loggedIn && (
              <button onClick={handleClose} style={{
                fontSize: 11, padding: '5px 10px', borderRadius: 6, border: 'none',
                background: '#F87171', color: '#fff', cursor: 'pointer', fontWeight: 600 }}>
                Fermer Chrome
              </button>
            )}
          </div>
        )}
      </div>

      {/* ── Message feedback ── */}
      {message?.text && (
        <div style={{
          padding: '10px 14px', borderRadius: 8, marginBottom: 16, fontSize: 12, fontWeight: 600,
          background: message.type === 'ok' ? 'rgba(34,197,94,.12)' : message.type === 'warn' ? 'rgba(251,191,36,.12)' : 'rgba(248,113,113,.12)',
          color: message.type === 'ok' ? '#22C55E' : message.type === 'warn' ? '#FBBF24' : '#F87171',
          border: `1px solid ${message.type === 'ok' ? 'rgba(34,197,94,.3)' : message.type === 'warn' ? 'rgba(251,191,36,.3)' : 'rgba(248,113,113,.3)'}`,
        }}>
          {message.text}
        </div>
      )}

      {/* ── Login ── */}
      {!loggedIn && serverOnline && (
        <div style={{ background: 'var(--bg2)', border: '1px solid var(--border)',
          borderRadius: 12, padding: 20, marginBottom: 20 }}>
          <div style={{ fontSize: 13, fontWeight: 700, color: 'var(--text)', marginBottom: 14 }}>
            🔐 Connexion Meridianbet Backoffice
          </div>
          <div style={{ display: 'grid', gridTemplateColumns: '1fr 1fr', gap: 10, marginBottom: 12 }}>
            <div>
              <label style={{ fontSize: 11, color: 'var(--muted)', display: 'block', marginBottom: 4 }}>Username</label>
              <input value={username} onChange={e => setUsername(e.target.value)}
                placeholder="Ton username BO" className="field-input"
                style={{ width: '100%', fontSize: 12, boxSizing: 'border-box' }} />
            </div>
            <div>
              <label style={{ fontSize: 11, color: 'var(--muted)', display: 'block', marginBottom: 4 }}>Password</label>
              <input type="password" value={password} onChange={e => setPassword(e.target.value)}
                onKeyDown={e => e.key === 'Enter' && handleLogin()}
                placeholder="Ton password BO" className="field-input"
                style={{ width: '100%', fontSize: 12, boxSizing: 'border-box' }} />
            </div>
          </div>
          <button onClick={handleLogin} disabled={loading === 'login'} style={{
            padding: '8px 20px', borderRadius: 8, border: 'none', cursor: 'pointer',
            background: '#4F46E5', color: '#fff', fontWeight: 700, fontSize: 12,
            opacity: loading === 'login' ? 0.7 : 1 }}>
            {loading === 'login' ? '⏳ Connexion...' : '🚀 Se connecter'}
          </button>
          <p style={{ fontSize: 11, color: 'var(--muted)', marginTop: 8 }}>
            Ton mot de passe ne quitte jamais ta machine — il va directement au serveur local.
          </p>
        </div>
      )}

      {/* ── Onglets bonus ── */}
      <div style={{ display: 'flex', gap: 6, marginBottom: 16 }}>
        {bonusTypes.map(bt => (
          <button key={bt.k} onClick={() => setActiveBonus(bt.k)} style={{
            padding: '7px 14px', borderRadius: 8, border: 'none', cursor: 'pointer', fontWeight: 600, fontSize: 12,
            background: activeBonus === bt.k ? bt.color : 'var(--bg3)',
            color: activeBonus === bt.k ? '#fff' : 'var(--muted)',
          }}>{bt.label}</button>
        ))}
      </div>

      {/* ── Section par type de bonus ── */}
      {bonusTypes.map(bt => activeBonus === bt.k && (
        <div key={bt.k} style={{ background: 'var(--bg2)', border: '1px solid var(--border)',
          borderRadius: 12, padding: 20, marginBottom: 16 }}>

          <div style={{ display: 'flex', justifyContent: 'space-between', alignItems: 'center', marginBottom: 16 }}>
            <div>
              <div style={{ fontSize: 14, fontWeight: 700, color: 'var(--text)' }}>{bt.label} Bonus</div>
              <div style={{ fontSize: 11, color: 'var(--muted)', marginTop: 2 }}>{bt.desc}</div>
            </div>
            {bonusData[bt.k].length > 0 && (
              <div style={{ fontSize: 12, fontWeight: 600, color: bt.color,
                background: `${bt.color}18`, padding: '4px 10px', borderRadius: 6 }}>
                {bonusData[bt.k].length} joueurs chargés
              </div>
            )}
          </div>

          {/* Zone de drop fichier */}
          <input ref={refs[bt.k]} type="file" accept=".csv,.xlsx" hidden
            onChange={e => handleFile(bt.k, e.target.files[0])} />
          <div onClick={() => refs[bt.k].current.click()}
            style={{ border: `2px dashed ${bonusData[bt.k].length ? bt.color : 'var(--border)'}`,
              borderRadius: 10, padding: '20px', textAlign: 'center', cursor: 'pointer',
              background: bonusData[bt.k].length ? `${bt.color}08` : 'transparent',
              marginBottom: 16, transition: 'all .2s' }}>
            {bonusFiles[bt.k]
              ? <div>
                  <div style={{ fontSize: 20, marginBottom: 4 }}>📄</div>
                  <div style={{ fontSize: 12, fontWeight: 600, color: bt.color }}>{bonusFiles[bt.k]}</div>
                  <div style={{ fontSize: 11, color: 'var(--muted)', marginTop: 2 }}>Clique pour remplacer</div>
                </div>
              : <div>
                  <div style={{ fontSize: 20, marginBottom: 4 }}>📂</div>
                  <div style={{ fontSize: 12, color: 'var(--muted)' }}>
                    Charge le fichier CSV {bt.k.toUpperCase()} exporté depuis MGM Retention
                  </div>
                  <div style={{ fontSize: 11, color: 'var(--muted)', marginTop: 4, opacity: 0.7 }}>{bt.desc}</div>
                </div>
            }
          </div>

          {/* Aperçu données */}
          {bonusData[bt.k].length > 0 && (
            <div style={{ background: 'var(--bg3)', borderRadius: 8, padding: '10px 12px', marginBottom: 16,
              maxHeight: 160, overflowY: 'auto', fontSize: 11, fontFamily: 'monospace' }}>
              <div style={{ color: 'var(--muted)', marginBottom: 6, fontFamily: 'sans-serif', fontWeight: 600 }}>
                Aperçu (5 premières lignes) :
              </div>
              {bonusData[bt.k].slice(0, 5).map((row, i) => (
                <div key={i} style={{ color: 'var(--text)', marginBottom: 2 }}>
                  {Object.entries(row).map(([k, v]) => `${k}: ${v}`).join(' | ')}
                </div>
              ))}
            </div>
          )}

          {/* Bouton envoi */}
          <button onClick={() => handleSend(bt.k)}
            disabled={!loggedIn || !bonusData[bt.k].length || loading === bt.k}
            style={{ width: '100%', padding: '12px', borderRadius: 8, border: 'none',
              cursor: loggedIn && bonusData[bt.k].length ? 'pointer' : 'not-allowed',
              background: loggedIn && bonusData[bt.k].length ? bt.color : 'var(--bg3)',
              color: loggedIn && bonusData[bt.k].length ? '#fff' : 'var(--muted)',
              fontWeight: 700, fontSize: 13, opacity: loading === bt.k ? 0.7 : 1 }}>
            {loading === bt.k
              ? `⏳ Envoi en cours...`
              : !loggedIn
              ? '🔐 Connecte-toi d\'abord'
              : !bonusData[bt.k].length
              ? '📂 Charge un fichier d\'abord'
              : `🚀 Envoyer ${bonusData[bt.k].length} bonus ${bt.label}`}
          </button>
        </div>
      ))}

      {/* ── Logs serveur ── */}
      {logs.length > 0 && (
        <div style={{ background: 'var(--bg2)', border: '1px solid var(--border)',
          borderRadius: 12, padding: 16 }}>
          <div style={{ fontSize: 12, fontWeight: 600, color: 'var(--muted)', marginBottom: 10,
            textTransform: 'uppercase', letterSpacing: '.05em' }}>
            Journal du serveur
          </div>
          <div style={{ maxHeight: 180, overflowY: 'auto' }}>
            {[...logs].reverse().map((l, i) => (
              <div key={i} style={{ display: 'flex', gap: 8, padding: '3px 0',
                borderBottom: i < logs.length - 1 ? '0.5px solid var(--border)' : 'none' }}>
                <span style={{ fontSize: 10, color: 'var(--muted)', fontFamily: 'monospace', minWidth: 50 }}>
                  {l.time}
                </span>
                <span style={{ fontSize: 11, color: l.level === 'ok' ? '#22C55E' : l.level === 'err' ? '#F87171' : l.level === 'warn' ? '#FBBF24' : 'var(--text)' }}>
                  {l.msg}
                </span>
              </div>
            ))}
          </div>
        </div>
      )}

      {/* ── Instructions si serveur offline ── */}
      {!serverOnline && (
        <div style={{ background: 'var(--bg2)', border: '1px solid var(--border)',
          borderRadius: 12, padding: 20, marginTop: 8 }}>
          <div style={{ fontSize: 13, fontWeight: 700, color: 'var(--text)', marginBottom: 12 }}>
            🚀 Comment démarrer le serveur
          </div>
          {[
            { step: '1', title: 'Installation (une seule fois)', code: 'pip install flask flask-cors playwright\nplaywright install chromium' },
            { step: '2', title: 'Lancement (chaque matin)', code: 'python3 meridian_server.py' },
            { step: '3', title: 'Vérification', code: '→ Cette page affichera "🟢 Serveur connecté"' },
          ].map(s => (
            <div key={s.step} style={{ marginBottom: 14 }}>
              <div style={{ fontSize: 11, fontWeight: 600, color: 'var(--muted)', marginBottom: 4 }}>
                Étape {s.step} — {s.title}
              </div>
              <pre style={{ background: 'var(--bg3)', padding: '8px 12px', borderRadius: 6,
                fontSize: 11, fontFamily: 'monospace', color: 'var(--text)', margin: 0,
                whiteSpace: 'pre-wrap' }}>{s.code}</pre>
            </div>
          ))}
        </div>
      )}
    </div>
  );
}
