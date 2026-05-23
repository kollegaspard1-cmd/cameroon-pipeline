// src/pages/HistoryPage.js
import React, { useEffect, useState } from 'react';
import { loadRuns } from '../lib/runHistory';

export default function HistoryPage() {
  const [runs,    setRuns]    = useState([]);
  const [loading, setLoading] = useState(true);
  const [error,   setError]   = useState('');

  useEffect(() => {
    loadRuns(30).then(setRuns).catch(err => setError(err.message)).finally(() => setLoading(false));
  }, []);

  if (loading) return (
    <div style={{ display:'flex', alignItems:'center', gap:10, color:'var(--muted)', fontSize:14 }}>
      <span style={{ animation:'spin 1s linear infinite', display:'inline-block' }}>⟳</span>
      Loading history from Firebase…
    </div>
  );

  if (error) return (
    <div style={{ padding:'12px 16px', background:'var(--danger-bg)', border:'1px solid var(--danger)', borderRadius:'var(--radius)', color:'var(--danger)', fontSize:13 }}>
      <b>Firebase error:</b> {error}<br/><br/>
      <span style={{ color:'var(--muted)' }}>Make sure you've replaced the placeholder config in <code>src/lib/firebase.js</code> with your real Firebase project credentials.</span>
    </div>
  );

  if (runs.length === 0) return (
    <div className="card" style={{ textAlign:'center', padding:'3rem', color:'var(--muted)' }}>
      <div style={{ fontSize:36, marginBottom:12 }}>🕐</div>
      <p style={{ fontSize:14 }}>No runs yet. Check the "Save to Firebase" option when running the pipeline.</p>
    </div>
  );

  function fmtDate(ts) {
    if (!ts) return '—';
    const d = ts.toDate?.() ?? new Date(ts);
    return d.toLocaleString();
  }

  return (
    <div>
      <div style={{ display:'flex', alignItems:'center', justifyContent:'space-between', marginBottom:'1rem' }}>
        <p className="section-label" style={{ margin:0 }}>Past runs</p>
        <span className="tag blue">{runs.length} runs</span>
      </div>
      <div className="card" style={{ padding:0, overflow:'hidden' }}>
        <table className="hist-table">
          <thead>
            <tr>
              <th>Date</th>
              <th>Input files</th>
              <th>Players</th>
              <th>Phones</th>
              <th>Casino</th>
              <th>Sport</th>
              <th>Cashback</th>
              <th>Aviator</th>
              <th>Files</th>
            </tr>
          </thead>
          <tbody>
            {runs.map(run => (
              <tr key={run.id}>
                <td style={{ whiteSpace:'nowrap', color:'var(--muted)', fontSize:12, fontFamily:'JetBrains Mono,monospace' }}>{fmtDate(run.createdAt)}</td>
                <td style={{ fontSize:12 }}>
                  <div style={{ fontWeight:500 }}>{run.fileNames?.excel ?? '—'}</div>
                  <div style={{ color:'var(--muted)' }}>{run.fileNames?.csv ?? ''}</div>
                </td>
                <td><b style={{ color:'var(--accent)', fontFamily:'JetBrains Mono,monospace' }}>{run.stats?.totalPlayers?.toLocaleString() ?? '—'}</b></td>
                <td style={{ fontFamily:'JetBrains Mono,monospace' }}>{run.stats?.phonesExtracted ?? '—'}</td>
                <td style={{ fontFamily:'JetBrains Mono,monospace' }}>{run.stats?.byCampaign?.casinoBonus ?? '—'}</td>
                <td style={{ fontFamily:'JetBrains Mono,monospace' }}>{run.stats?.byCampaign?.sportBonus ?? '—'}</td>
                <td style={{ fontFamily:'JetBrains Mono,monospace' }}>{run.stats?.byCampaign?.cashback ?? '—'}</td>
                <td style={{ fontFamily:'JetBrains Mono,monospace' }}>{run.stats?.byCampaign?.spribe ?? '—'}</td>
                <td><span className="tag green">{run.outputCount ?? '—'} files</span></td>
              </tr>
            ))}
          </tbody>
        </table>
      </div>
    </div>
  );
}
