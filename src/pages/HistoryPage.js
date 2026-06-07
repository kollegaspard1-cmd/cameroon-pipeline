// src/pages/HistoryPage.js
import React, { useEffect, useState } from 'react';
import { loadRuns } from '../lib/runHistory';

export default function HistoryPage() {
  const [runs,    setRuns]    = useState([]);
  const [loading, setLoading] = useState(true);
  const [error,   setError]   = useState('');

  useEffect(() => {
    loadRuns(30)
      .then(setRuns)
      .catch(err => setError(err.message))
      .finally(() => setLoading(false));
  }, []);

  if (loading) return <p style={{ color: 'var(--muted)', fontSize: '14px' }}>Loading history from Firebase…</p>;

  if (error) return (
    <div style={{ padding: '12px 16px', background: 'rgba(255,77,77,0.1)', border: '1px solid var(--danger)', borderRadius: '8px', color: 'var(--danger)', fontSize: '13px' }}>
      <b>Firebase error:</b> {error}
      <br/><br/>
      <span style={{ color: 'var(--muted)' }}>Make sure you've replaced the placeholder config in <code>src/lib/firebase.js</code> with your real Firebase project credentials.</span>
    </div>
  );

  if (runs.length === 0) return (
    <p style={{ color: 'var(--muted)', fontSize: '14px' }}>No runs yet. Check the "Save to Firebase" option when running the pipeline.</p>
  );

  function fmtDate(ts) {
    if (!ts) return '—';
    const d = ts.toDate?.() ?? new Date(ts);
    return d.toLocaleString();
  }

  return (
    <div>
      <p className="section-label">Past runs ({runs.length})</p>
      <div className="card" style={{ padding: 0, overflow: 'hidden' }}>
        <div style={{ overflowX: 'auto', WebkitOverflowScrolling: 'touch' }}>
        <table className="hist-table" style={{ minWidth: 700 }}>
          <thead>
            <tr>
              <th>Date</th>
              <th>Input files</th>
              <th>Players</th>
              <th>Phones</th>
              <th>Casino bonus</th>
              <th>Sport bonus</th>
              <th>Cashback</th>
              <th>Spribe</th>
              <th>Files</th>
            </tr>
          </thead>
          <tbody>
            {runs.map(run => (
              <tr key={run.id}>
                <td style={{ whiteSpace: 'nowrap', color: 'var(--muted)', fontSize: '12px' }}>{fmtDate(run.createdAt)}</td>
                <td style={{ fontSize: '12px' }}>
                  <div>{run.fileNames?.excel ?? '—'}</div>
                  <div style={{ color: 'var(--muted)' }}>{run.fileNames?.csv ?? ''}</div>
                </td>
                <td><b style={{ color: 'var(--accent)' }}>{run.stats?.totalPlayers?.toLocaleString() ?? '—'}</b></td>
                <td>{run.stats?.phonesExtracted ?? '—'}</td>
                <td>{run.stats?.byCampaign?.casinoBonus ?? '—'}</td>
                <td>{run.stats?.byCampaign?.sportBonus ?? '—'}</td>
                <td>{run.stats?.byCampaign?.cashback ?? '—'}</td>
                <td>{run.stats?.byCampaign?.spribe ?? '—'}</td>
                <td>
                  <span className="tag green">{run.outputCount ?? '—'} files</span>
                </td>
              </tr>
            ))}
          </tbody>
        </table>
        </div>
      </div>
    </div>
  );
}
