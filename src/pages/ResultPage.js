// src/pages/ResultPage.js
import React, { useState } from 'react';
import { saveAs } from 'file-saver';
import JSZip from 'jszip';

export default function ResultPage({ result }) {
  const [downloading, setDownloading] = useState(false);

  if (!result) return (
    <p style={{ color: 'var(--muted)', fontSize: 14 }}>No result yet — run the pipeline first.</p>
  );

  const { stats, outputFiles, campaigns } = result;

  const today = new Date();
  const dd = String(today.getDate()).padStart(2,'0');
  const mm = String(today.getMonth()+1).padStart(2,'0');
  const yy = String(today.getFullYear()).slice(-2);
  // Match any country prefix (CM_, TZ_, KE_, etc.)
  const finalName = outputFiles.find(f => f.filename.match(/^[A-Z]{2}_Daily_/))?.filename
    || `CM_Daily_${dd}.${mm}.${yy}.xlsx`;

  const campaignGroups = [
    { key:'main',   label:'Main output',          tag:'blue',
      files: outputFiles.filter(f => f.filename===finalName || f.filename==='Cameroon_YTD_phones.xlsx') },
    { key:'casino', label:'Casino campaigns',      tag:'amber',
      files: outputFiles.filter(f => ['casino_bonus_cm.csv','Daily_Cashback_Template.csv','spribe_freebets.csv'].includes(f.filename)) },
    { key:'sport',  label:'Sport bonus campaigns', tag:'green',
      files: outputFiles.filter(f => f.filename.startsWith('CM_daily_') || f.filename.startsWith('Sport_Daily_')) },
  ];

  async function downloadAll() {
    setDownloading(true);
    const zip = new JSZip();
    for (const { blob, filename } of outputFiles) zip.file(filename, blob);
    const zipBlob = await zip.generateAsync({ type:'blob' });
    saveAs(zipBlob, `cameroon_pipeline_${new Date().toISOString().slice(0,10)}.zip`);
    setDownloading(false);
  }

  const kpis = [
    { label:'Total Players', value: stats.totalPlayers.toLocaleString(), color:'var(--accent)' },
    { label:'Phones',        value: stats.phonesExtracted.toLocaleString(), color:'var(--sport)' },
    { label:'Casino Bonus',  value: stats.byCampaign.casinoBonus, color:'#D97706' },
    { label:'Sport Bonus',   value: stats.byCampaign.sportBonus, color:'var(--deposit)' },
    { label:'Cashback',      value: stats.byCampaign.cashback, color:'var(--text)' },
    { label:'Aviator FB',    value: stats.byCampaign.spribe, color:'var(--text)' },
    { label:'Output Files',  value: outputFiles.length, color:'var(--text)' },
    { label:'Removed (Today<0)', value: stats.removedToday, color: stats.removedToday>0?'var(--danger)':'var(--text)' },
  ];

  return (
    <div>
      {/* KPI Banner */}
      <div className="kpi-banner">
        {kpis.map(k => (
          <div key={k.label} className="kpi-item">
            <div className="kpi-label">{k.label}</div>
            <div className="kpi-value" style={{ color: k.color, fontSize: 22 }}>{k.value}</div>
          </div>
        ))}
      </div>

      {/* Download all */}
      <div style={{ marginBottom: '1.5rem' }}>
        <button className="btn success" onClick={downloadAll} disabled={downloading}
          style={{ fontSize: 14, padding: '10px 24px' }}>
          {downloading ? '⟳ Zipping…' : '⬇  Download all as .zip'}
        </button>
      </div>

      {/* File groups */}
      {campaignGroups.map(group => group.files.length > 0 && (
        <div key={group.key} style={{ marginBottom: '1.5rem' }}>
          <div style={{ display: 'flex', alignItems: 'center', gap: 8, marginBottom: 10 }}>
            <p className="section-label" style={{ margin: 0 }}>{group.label}</p>
            <span className={`tag ${group.tag}`}>{group.files.length} files</span>
          </div>
          <div className="file-list">
            {group.files.map(f => (
              <div key={f.filename} className="file-row">
                <div>
                  <div className="file-row-name">{f.filename}</div>
                  <div className="file-row-meta">{getFileDescription(f.filename, campaigns)}</div>
                </div>
                <button className="btn sm" onClick={() => saveAs(f.blob, f.filename)}>⬇ Download</button>
              </div>
            ))}
          </div>
        </div>
      ))}
    </div>
  );
}

function getFileDescription(filename, campaigns) {
  if (filename.includes('CM_Daily'))           return 'MTD sheet — cleaned, segmented, with Today GGR';
  if (filename === 'Cameroon_YTD_phones.xlsx') return 'YTD phone numbers (+ stripped)';
  if (filename === 'casino_bonus_cm.csv')      return 'Casino bonus players — account_id + amount + XAF';
  if (filename === 'Daily_Cashback_Template.csv') return 'Cash offer players — casino game, no specific preference';
  if (filename === 'spribe_freebets.csv')      return 'Aviator freebet players — 5 freebets, 7-day validity';
  if (filename.startsWith('CM_daily_') || filename.startsWith('Sport_Daily_')) {
    const amt = filename.replace('CM_daily_','').replace('Sport_Daily_','').replace('.csv','');
    return `Sport bonus — amount = ${Number(amt).toLocaleString()} XAF`;
  }
  return '';
}
