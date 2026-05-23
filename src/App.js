// src/App.js
import React, { useState } from 'react';
import UploadPage    from './pages/UploadPage';
import ResultPage    from './pages/ResultPage';
import HistoryPage   from './pages/HistoryPage';
import ObjectifsPage from './pages/ObjectifsPage';
import RapportsPage  from './pages/RapportsPage';
import './App.css';

const PAGES = [
  { id: 'upload',    label: 'Run Pipeline',  icon: '⚡', section: 'PIPELINE CM' },
  { id: 'result',    label: 'Results',       icon: '📊', section: null, requiresResult: true },
  { id: 'history',   label: 'History',       icon: '🕐', section: null },
  { id: 'objectifs', label: 'Objectifs',     icon: '🎯', section: 'CAMPRESJ' },
  { id: 'rapports',  label: 'Rapports',      icon: '📈', section: null },
];

const PAGE_TITLES = {
  upload:    { title: 'Run Pipeline',               sub: 'Cameroon Daily Campaign Engine' },
  result:    { title: 'Pipeline Results',            sub: 'Download output files' },
  history:   { title: 'Run History',                 sub: 'Firebase-stored past runs' },
  objectifs: { title: 'Objectifs CAMPRESJ',          sub: 'Générer les objectifs mensuels' },
  rapports:  { title: 'Rapports Réalisation',        sub: 'Réalisé vs Objectif par super' },
};

export default function App() {
  const [page, setPage]                   = useState('upload');
  const [pipelineResult, setPipelineResult] = useState(null);

  function handleResult(result) {
    setPipelineResult(result);
    setPage('result');
  }

  // Group pages by section
  let lastSection = null;

  return (
    <div className="app">
      {/* ── Sidebar ── */}
      <aside className="sidebar">
        <div className="sidebar-brand">
          <div className="brand-logo">
            <span className="brand-dot" />
            <span className="brand-name">Pipeline CM</span>
          </div>
          <div className="brand-sub">Cameroon Operations</div>
        </div>

        <div className="sidebar-section">
          {PAGES.map(p => {
            const showSection = p.section && p.section !== lastSection;
            if (showSection) lastSection = p.section;
            const disabled = p.requiresResult && !pipelineResult;
            return (
              <React.Fragment key={p.id}>
                {showSection && (
                  <div className="sidebar-section-label" style={{ marginTop: lastSection !== p.section ? '1.25rem' : 0 }}>
                    {p.section}
                  </div>
                )}
                <button
                  className={`nav-btn ${page === p.id ? 'active' : ''}`}
                  onClick={() => !disabled && setPage(p.id)}
                  disabled={disabled}
                >
                  <span className="nav-icon">{p.icon}</span>
                  {p.label}
                </button>
              </React.Fragment>
            );
          })}
        </div>
      </aside>

      {/* ── Main ── */}
      <div className="main-wrapper">
        <header className="topbar">
          <div style={{ display: 'flex', alignItems: 'center' }}>
            <span className="topbar-title">{PAGE_TITLES[page]?.title}</span>
            <span className="topbar-sub">— {PAGE_TITLES[page]?.sub}</span>
          </div>
        </header>

        <main className="main">
          {page === 'upload'    && <UploadPage    onResult={handleResult} />}
          {page === 'result'    && <ResultPage    result={pipelineResult} />}
          {page === 'history'   && <HistoryPage   />}
          {page === 'objectifs' && <ObjectifsPage />}
          {page === 'rapports'  && <RapportsPage  />}
        </main>
      </div>
    </div>
  );
}
