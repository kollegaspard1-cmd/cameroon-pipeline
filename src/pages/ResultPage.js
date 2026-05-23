// src/pages/ResultPage.js
import React, { useState } from 'react';
import { saveAs } from 'file-saver';
import JSZip from 'jszip';

export default function ResultPage({ result }) {
  const [downloading, setDownloading] = useState(false);

  if (!result) {
    return <p style={{ color: 'var(--muted)', fontSize: '14px' }}>No result yet — run the pipeline first.</p>;
  }

  const { stats, outputFiles, campaigns } = result;

  function downloadFile({ blob, filename }) {
    saveAs(blob, filename);
  }

  async function downloadAll() {
    setDownloading(true);
    const zip = new JSZip();
    for (const { blob, filename } of outputFiles) {
      zip.file(filename, blob);
    }
    const zipBlob = await zip.generateAsync({ type: 'blob' });
    saveAs(zipBlob, `cameroon_pipeline_${new Date().toISOString().slice(0,10)}.zip`);
    setDownloading(false);
  }

  const today = new Date();
  const dd = String(today.getDate()).padStart(2, '0');
  const mm = String(today.getMonth() + 1).padStart(2, '0');
  const yy = String(today.getFullYear()).slice(-2);
  const dateStr = `${dd}.${mm}.${yy}`;
  const finalName = `CM_Daily_${dateStr}.xlsx`;

  const campaignGroups = [
    {
      key: 'main',
      label: 'Main output',
      tag: 'blue',
      files: outputFiles.filter(f => f.filename === finalName || f.filename === 'Cameroon_YTD_phones.xlsx')
    },
    {
      key: 'casino',
      label: 'Casino campaigns',
      tag: 'amber',
      files: outputFiles.filter(f =>
        f.filename === 'casino_bonus_cm.csv' ||
        f.filename === 'Daily_Cashback_Template.csv' ||
        f.filename === 'spribe_freebets.csv'
      )
    },
    {
      key: 'sport',
      label: 'Sport bonus campaigns',
      tag: 'green',
      files: outputFiles.filter(f => f.filename.startsWith('CM_daily_'))
    }
  ];

  return (
    <div>
      {/* Stats */}
      <p className="section-label">Run summary</p>
      <div className="stat-grid">
        <div className="stat-card">
          <div className="stat-label">Total players</div>
          <div className="stat-value">{stats.totalPlayers.toLocaleString()}</div>
        </div>
        <div className="stat-card">
          <div className="stat-label">Phones</div>
          <div className="stat-value">{stats.phonesExtracted.toLocaleString()}</div>
        </div>
        <div className="stat-card">
          <div className="stat-label">Casino bonus</div>
          <div className="stat-value neutral">{stats.byCampaign.casinoBonus}</div>
        </div>
        <div className="stat-card">
          <div className="stat-label">Sport bonus</div>
          <div className="stat-value neutral">{stats.byCampaign.sportBonus}</div>
        </div>
        <div className="stat-card">
          <div className="stat-label">Cashback</div>
          <div className="stat-value neutral">{stats.byCampaign.cashback}</div>
        </div>
        <div className="stat-card">
          <div className="stat-label">Aviator freebets</div>
          <div className="stat-value neutral">{stats.byCampaign.spribe}</div>
        </div>
        <div className="stat-card">
          <div className="stat-label">Removed (Today&lt;0)</div>
          <div className="stat-value" style={{ color: stats.removedToday > 0 ? 'var(--danger)' : 'var(--text)', fontSize: '20px' }}>
            {stats.removedToday}
          </div>
        </div>
        <div className="stat-card">
          <div className="stat-label">Output files</div>
          <div className="stat-value neutral">{outputFiles.length}</div>
        </div>
      </div>

      {/* Download all */}
      <div style={{ marginBottom: '1.5rem', display: 'flex', gap: '8px' }}>
        <button className="btn primary" onClick={downloadAll} disabled={downloading}>
          {downloading ? '⟳ Zipping…' : '⬇ Download all as .zip'}
        </button>
      </div>

      {/* Campaign file groups */}
      {campaignGroups.map(group => (
        group.files.length > 0 && (
          <div key={group.key} style={{ marginBottom: '1.5rem' }}>
            <p className="section-label">
              {group.label}
              <span className={`tag ${group.tag}`} style={{ marginLeft: '8px' }}>{group.files.length} files</span>
            </p>
            <div className="file-list">
              {group.files.map(f => (
                <div key={f.filename} className="file-row">
                  <div>
                    <div className="file-row-name">{f.filename}</div>
                    <div className="file-row-meta">{getFileDescription(f.filename, campaigns)}</div>
                  </div>
                  <div className="file-row-actions">
                    <button className="btn sm" onClick={() => downloadFile(f)}>
                      ⬇ Download
                    </button>
                  </div>
                </div>
              ))}
            </div>
          </div>
        )
      ))}
    </div>
  );
}

function getFileDescription(filename, campaigns) {
  if (filename === 'Cameroon_final.xlsx') return 'MTD sheet — cleaned, segmented, with Today GGR';
  if (filename === 'Cameroon_YTD_phones.xlsx') return 'YTD phone numbers (+ stripped)';
  if (filename === 'casino_bonus_cm.csv') return 'Casino bonus players — account_id + amount + XAF';
  if (filename === 'Daily_Cashback_Template.csv') return 'Cash offer players — casino game, no specific preference';
  if (filename === 'spribe_freebets.csv') return 'Aviator freebet players — 5 freebets, 7-day validity';
  if (filename.startsWith('CM_daily_')) {
    const amt = filename.replace('CM_daily_', '').replace('.csv', '');
    return `Sport bonus — amount = ${Number(amt).toLocaleString()} XAF`;
  }
  return '';
}
