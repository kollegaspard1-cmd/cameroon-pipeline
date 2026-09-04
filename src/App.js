// src/App.js
import React, { useState } from 'react';
import { AppProvider, useAppContext } from './AppContext';
import { COUNTRIES }                  from './translations';
import UploadPage               from './pages/UploadPage';
import ResultPage               from './pages/ResultPage';
import HistoryPage              from './pages/HistoryPage';
import CampresjPage             from './pages/CampresjPage';
import CMReportPage             from './pages/CMReportPage';
import MGMRetentionPage         from './pages/MGMRetentionPage';
import BetshopMarchesPage       from './pages/BetshopMarchesPage';
import BetshopMultiPeriodesPage from './pages/BetshopMultiPeriodesPage';
import RetailReportPage         from './pages/RetailReportPage';
import WatchlistPage            from './pages/WatchlistPage';
import BonusAutoPage            from './pages/BonusAutoPage';
import './App.css';

const NAV = [
  {
    section: 'Pipeline CM',
    items: [
      { id: 'upload',      label: 'Daily Retention', icon: '⚡' },
      { id: 'bonus-auto',  label: 'Envoi Bonus',     icon: '🤖' },
      { id: 'result',      label: 'Results',         icon: '📊', requiresResult: true },
      { id: 'history',     label: 'History',         icon: '🕐' },
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
      { id: 'campresj',   label: 'Objectifs & Réalisation', icon: '🎯' },
    ]
  },
  {
    section: 'Betshop',
    items: [
      { id: 'betshop-marches',  label: 'Par Marché',     icon: '🗺️' },
      { id: 'betshop-periodes', label: 'Multi-Périodes', icon: '📅' },
    ]
  },
  {
    section: 'Retail',
    items: [
      { id: 'retail-report', label: 'Cashier & Deposits', icon: '📋' },
      { id: 'watchlist',     label: 'Watchlist',           icon: '⚠️' },
    ]
  },
  {
    section: 'Automatisation',
    items: [
      { id: 'bonus-sender', label: 'Bonus Sender', icon: '📦' },
    ]
  },
  {
    section: 'MGM Retention',
    items: [
      { id: 'mgm',         label: 'Upload & Bonus', icon: '📁' },
      { id: 'mgm-sport',   label: 'Sport',          icon: '⚽', sub: true },
      { id: 'mgm-casino',  label: 'Casino',         icon: '🎰', sub: true },
      { id: 'mgm-deposit', label: 'Deposit',        icon: '💳', sub: true },
    ]
  },
];

const TITLES = {
  'upload':            { title: 'Run Pipeline',             sub: 'Cameroon Daily Campaign Engine' },
  'bonus-auto':        { title: 'Bonus Auto',               sub: 'Envoi automatique Casino · Cashback · Sport' },
  'result':            { title: 'Pipeline Results',         sub: 'Download output files' },
  'history':           { title: 'Run History',              sub: 'Firebase-stored past runs' },
  'cmreport':          { title: 'Daily Sport Report',       sub: 'Remplissage du template depuis les CSV' },
  'campresj':          { title: 'CAMPRESJ',                 sub: 'Objectifs & Réalisation mensuelle' },
  'betshop-marches':   { title: 'Betshop — Par Marché',    sub: 'Fichiers par marché + CAMPRESJ + Global' },
  'betshop-periodes':  { title: 'Betshop — Multi-Périodes', sub: 'Analyse Bimestrielle / Trimestrielle / Semestrielle / Annuelle' },
  'retail-report':     { title: 'Retail Report',            sub: 'Cashier & Betshop · Deposits & Withdraws by Market' },
  'watchlist':         { title: 'Watchlist',                sub: 'Bet Performance MTD vs LM vs LY — Shops en retard' },
  'bonus-sender':      { title: 'Bonus Sender',             sub: 'Envoi automatisé des bonus' },
  'mgm':               { title: 'MGM · Upload & Bonus',    sub: 'Upload data & bonus validator' },
  'mgm-sport':         { title: 'MGM · Sport',             sub: 'Players with 0 tickets this month' },
  'mgm-casino':        { title: 'MGM · Casino',            sub: 'Players with 0 casino payin this month' },
  'mgm-deposit':       { title: 'MGM · Deposit',           sub: 'Players with 0 deposit this month' },
};

// ── Barre Langue / Pays (globale) ─────────────────────────────────────────
function LangCountryBar() {
  const { lang, setLang, country, setCountry } = useAppContext();
  return (
    <div style={{
      display: 'flex', gap: 6, alignItems: 'center',
      padding: '4px 10px',
      background: 'var(--bg2)',
      borderLeft: '1px solid var(--border)',
    }}>
      {['FR', 'EN'].map(l => (
        <button key={l} onClick={() => setLang(l)} style={{
          fontSize: 11, fontWeight: 700, padding: '2px 8px', borderRadius: 4,
          border: `1px solid ${lang === l ? 'var(--accent)' : 'var(--border)'}`,
          background: lang === l ? 'var(--accent)' : 'transparent',
          color: lang === l ? '#fff' : 'var(--muted)',
          cursor: 'pointer',
        }}>{l}</button>
      ))}
      <div style={{ width: 1, height: 14, background: 'var(--border)' }}/>
      <select
        value={country}
        onChange={e => setCountry(e.target.value)}
        style={{
          fontSize: 11, padding: '2px 6px', borderRadius: 4,
          border: '1px solid var(--border)', background: 'var(--bg3)',
          color: 'var(--text)', cursor: 'pointer',
        }}
      >
        {COUNTRIES.map(c => (
          <option key={c.code} value={c.code}>{c.flag} {c.name}</option>
        ))}
      </select>
    </div>
  );
}

// ── AppInner : écoute navigateTo ──────────────────────────────────────────
function AppInner({ setPage }) {
  const { navigateTo, setNavigateTo } = useAppContext();
  React.useEffect(() => {
    if (navigateTo) { setPage(navigateTo); setNavigateTo(null); }
  }, [navigateTo, setNavigateTo, setPage]);
  return null;
}

// ── App principal ─────────────────────────────────────────────────────────
export default function App() {
  const [page, setPage]                     = useState('upload');
  const [sidebarOpen, setSidebarOpen]       = useState(false);
  const [pipelineResult, setPipelineResult] = useState(null);
  const [mgmResults, setMgmResults]         = useState(null);

  function handleResult(result) { setPipelineResult(result); setPage('result'); setSidebarOpen(false); }
  const navigate = (p) => { setPage(p); setSidebarOpen(false); };
  const mgmTab = page.startsWith('mgm-') ? page.replace('mgm-', '') : null;
  const mgmCounts = {
    sport:   mgmResults?.sport?.length   || 0,
    casino:  mgmResults?.casino?.length  || 0,
    deposit: mgmResults?.deposit?.length || 0,
  };

  return (
    <AppProvider>
      <AppInner setPage={setPage}/>
      <div className="app">
        <div className={`sidebar-overlay ${sidebarOpen ? 'open' : ''}`}
          onClick={() => setSidebarOpen(false)} />

        <aside className={`sidebar ${sidebarOpen ? 'open' : ''}`}>
          <div className="sidebar-brand">
            <div className="brand-logo"><span className="brand-dot"/><span className="brand-name">Pipeline CM</span></div>
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
                      onClick={() => !disabled && navigate(item.id)}
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
            <button className="hamburger" onClick={() => setSidebarOpen(o => !o)} aria-label="Menu">☰</button>
            <span className="topbar-title">{TITLES[page]?.title}</span>
            <span className="topbar-sub">— {TITLES[page]?.sub}</span>
            <div style={{ marginLeft:'auto' }}>
              <LangCountryBar/>
            </div>
          </header>
          <main className="main">
            {page === 'upload'           && <UploadPage onResult={handleResult}/>}
            {page === 'bonus-auto'       && <BonusAutoPage/>}
            {page === 'result'           && <ResultPage result={pipelineResult}/>}
            {page === 'history'          && <HistoryPage/>}
            {page === 'cmreport'         && <CMReportPage/>}
            {page === 'campresj'         && <CampresjPage/>}
            {page === 'retail-report'    && <RetailReportPage/>}
            {page === 'watchlist'        && <WatchlistPage/>}
            {page === 'betshop-marches'  && <BetshopMarchesPage onNavigate={setPage}/>}
            {page === 'betshop-periodes' && <BetshopMultiPeriodesPage onNavigate={setPage}/>}
            {(page === 'mgm' || page.startsWith('mgm-')) && (
              <MGMRetentionPage activeTab={mgmTab} onNavigate={setPage} results={mgmResults} onResults={setMgmResults}/>
            )}
          </main>
        </div>
      </div>
    </AppProvider>
  );
}
