// src/App.js
import React, { useState } from 'react';
import UploadPage from './pages/UploadPage';
import ResultPage from './pages/ResultPage';
import HistoryPage from './pages/HistoryPage';
import './App.css';

export default function App() {
  const [page, setPage] = useState('upload');  // 'upload' | 'result' | 'history'
  const [pipelineResult, setPipelineResult] = useState(null);

  function handleResult(result) {
    setPipelineResult(result);
    setPage('result');
  }

  return (
    <div className="app">
      <header className="app-header">
        <div className="header-inner">
          <div className="brand">
            <span className="brand-dot" />
            <span className="brand-name">Pipeline CM</span>
            <span className="brand-sub">Cameroon Daily Campaign Engine</span>
          </div>
          <nav className="nav">
            <button
              className={`nav-btn ${page === 'upload' ? 'active' : ''}`}
              onClick={() => setPage('upload')}
            >
              <svg width="16" height="16" viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2"><path d="M21 15v4a2 2 0 01-2 2H5a2 2 0 01-2-2v-4"/><polyline points="17 8 12 3 7 8"/><line x1="12" y1="3" x2="12" y2="15"/></svg>
              Run Pipeline
            </button>
            <button
              className={`nav-btn ${page === 'result' ? 'active' : ''}`}
              onClick={() => setPage('result')}
              disabled={!pipelineResult}
            >
              <svg width="16" height="16" viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2"><rect x="3" y="3" width="18" height="18" rx="2"/><path d="M9 9h6M9 12h6M9 15h4"/></svg>
              Results
            </button>
            <button
              className={`nav-btn ${page === 'history' ? 'active' : ''}`}
              onClick={() => setPage('history')}
            >
              <svg width="16" height="16" viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2"><circle cx="12" cy="12" r="10"/><polyline points="12 6 12 12 16 14"/></svg>
              History
            </button>
          </nav>
        </div>
      </header>

      <main className="main">
        {page === 'upload'  && <UploadPage  onResult={handleResult} />}
        {page === 'result'  && <ResultPage  result={pipelineResult} />}
        {page === 'history' && <HistoryPage />}
      </main>
    </div>
  );
}
