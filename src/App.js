// src/App.js
import React, { useState } from 'react';
import { AppProvider }          from './AppContext';
import UploadPage               from './pages/UploadPage';
import ResultPage               from './pages/ResultPage';
import HistoryPage              from './pages/HistoryPage';
import ObjectifsPage            from './pages/ObjectifsPage';
import RapportsPage             from './pages/RapportsPage';
import CMReportPage             from './pages/CMReportPage';
import MGMRetentionPage         from './pages/MGMRetentionPage';
import BetshopMarchesPage       from './pages/BetshopMarchesPage';
import BetshopMultiPeriodesPage from './pages/BetshopMultiPeriodesPage';
import './App.css';

const NAV = [
  {
    section: 'Pipeline CM',
    items: [
      { id: 'upload',      label: 'Daily Retention', icon: '⚡' },
      { id: 'result',      label: 'Results',         icon: '📊', requiresResult: true },
      { id: 'history',     label: 'History',          icon: '🕐' },
    ]
  },
  {
    section: 'Rapports',
    items: [
      { id: 'cmreport',    label: 'Daily Sport Report', icon: '📝' },
    ]
  },
  {
    section: 'CAMPRESJ',
    items: [
      { id: 'objectifs',   label: 'Objectifs',       icon: '🎯' },
      { id: 'rapports',    label: 'Réalisation',     icon: '📈' },
    ]
  },
  {
    section: 'Betshop',
    items: [
      { id: 'betshop-marches',    label: 'Par Marché',      icon: '🗺️' },
      { id: 'betshop-periodes',   label: 'Multi-Périodes',  icon: '📅' },
    ]
  },
  {
    section: 'MGM Retention',
    items: [
      { id: 'mgm',         label: 'Upload & Bonus',  icon: '📁' },
      { id: 'mgm-sport',   label: 'Sport',            icon: '⚽', sub: true },
      { id: 'mgm-casino',  label: 'Casino',           icon: '🎰', sub: true },
      { id: 'mgm-deposit', label: 'Deposit',          icon: '💳', sub: true },
    ]
  },
];

const TITLES = {
  upload:            { title: 'Run Pipeline',            sub: 'Cameroon Daily Campaign Engine' },
  result:            { title: 'Pipeline Results',         sub: 'Download output files' },
  history:           { title: 'Run History',              sub: 'Firebase-stored past runs' },
  cmreport:          { title: 'Daily Sport Report',       sub: 'Remplissage du template depuis les CSV' },
  objectifs:         { title: 'Objectifs CAMPRESJ',       sub: 'Générer les objectifs mensuels' },
  rapports:          { title: 'Rapports Réalisation',     sub: 'Réalisé vs Objectif par super' },
  'betshop-marches': { title: 'Betshop — Par Marché',    sub: 'Fichiers par marché + CAMPRESJ + Global' },
  'betshop-periodes':{ title: 'Betshop — Multi-Périodes', sub: 'Analyse Bimestrielle / Trimestrielle / Semestrielle / Annuelle' },
  mgm:               { title: 'MGM · Upload & Bonus',    sub: 'Upload data & bonus validator' },
  'mgm-sport':       { title: 'MGM · Sport',              sub: 'Players with 0 tickets this month' },
  'mgm-casino':      { title: 'MGM · Casino',             sub: 'Players with 0 casino payin this month' },
  'mgm-deposit':     { title: 'MGM · Deposit',            sub: 'Players with 0 deposit this month' },
};

export default function App() {
  const [page, setPage]                     = useState('upload');
  const [pipelineResult, setPipelineResult] = useState(null);
  const [mgmResults, setMgmResults]         = useState(null);

  function handleResult(result) { setPipelineResult(result); setPage('result'); }
  const mgmTab = page.startsWith('mgm-') ? page.replace('mgm-', '') : null;
  const mgmCounts = {
    sport:   mgmResults?.sport?.length   || 0,
    casino:  mgmResults?.casino?.length  || 0,
    deposit: mgmResults?.deposit?.length || 0,
  };

  return (
    <AppProvider>
      <div className="app">
        <aside className="sidebar">
          <div className="sidebar-brand">
            <div className="brand-logo"><span className="brand-dot" /><span className="brand-name">Pipeline CM</span></div>
            <div className="brand-sub">Cameroon Operations</div>
          </div>
          <nav style={{ padding:'0.5rem 0.75rem', flex:1, overflowY:'auto' }}>
            {NAV.map(group => (
              <div key={group.section} style={{ marginBottom:'0.1rem' }}>
                <div className="sidebar-section-label" style={{ marginTop:'1rem', marginBottom:'2px' }}>{group.section}</div>
                {group.items.map(item => {
                  const disabled = item.requiresResult && !pipelineResult;
                  const isMgmSub = item.sub;
                  const cat = item.id.replace('mgm-','');
                  const count = isMgmSub ? mgmCounts[cat] : null;
                  return (
                    <button key={item.id}
                      className={`nav-btn ${page === item.id ? 'active' : ''}`}
                      onClick={() => !disabled && setPage(item.id)}
                      disabled={disabled}
                      style={isMgmSub ? { paddingLeft:28 } : {}}>
                      <span className="nav-icon" style={isMgmSub?{fontSize:13}:{}}>{item.icon}</span>
                      <span style={isMgmSub?{fontSize:12}:{}}>{item.label}</span>
                      {count > 0 && <span className="nav-badge">{count.toLocaleString()}</span>}
                    </button>
                  );
                })}
              </div>
            ))}
          </nav>
        </aside>

        <div className="main-wrapper">
          <header className="topbar">
            <span className="topbar-title">{TITLES[page]?.title}</span>
            <span className="topbar-sub">— {TITLES[page]?.sub}</span>
          </header>
          <main className="main">
            {page === 'upload'           && <UploadPage       onResult={handleResult} />}
            {page === 'result'           && <ResultPage        result={pipelineResult} />}
            {page === 'history'          && <HistoryPage       />}
            {page === 'cmreport'         && <CMReportPage      />}
            {page === 'objectifs'        && <ObjectifsPage     onNavigate={setPage} />}
            {page === 'rapports'         && <RapportsPage      onNavigate={setPage} />}
            {page === 'betshop-marches'  && <BetshopMarchesPage onNavigate={setPage} />}
            {page === 'betshop-periodes' && <BetshopMultiPeriodesPage onNavigate={setPage} />}
            {(page === 'mgm' || page.startsWith('mgm-')) && (
              <MGMRetentionPage activeTab={mgmTab} onNavigate={setPage} results={mgmResults} onResults={setMgmResults} />
            )}
          </main>
        </div>
      </div>
    </AppProvider>
  );
}
