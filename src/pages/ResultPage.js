// src/pages/ResultPage.js
import React, { useState, useMemo } from 'react';
import { useAppContext } from '../AppContext';
import { saveAs } from 'file-saver';
import JSZip from 'jszip';

// ── Pill component ───────────────────────────────────────────────────────────
const PILL_STYLES = {
  'casino':       { bg:'#2d1f5e', color:'#b59fff' },
  'sport bonus':  { bg:'#1d3a6e', color:'#7aadff' },
  'MIX':          { bg:'#1d4a3a', color:'#5dcea0' },
  'Cash':         { bg:'#3d2a10', color:'#f5a623' },
  'Aviator':      { bg:'#3d1a10', color:'#ff7a5a' },
  'Casino bonus': { bg:'#2d1f5e', color:'#b59fff' },
};
function Pill({ label }) {
  if (!label) return <span style={{ color:'var(--muted)' }}>—</span>;
  const s = PILL_STYLES[label] || { bg:'var(--bg3)', color:'var(--muted)' };
  return (
    <span style={{
      display:'inline-block', fontSize:11, padding:'3px 9px',
      borderRadius:999, fontWeight:600, background:s.bg, color:s.color,
    }}>{label}</span>
  );
}

// ── Collapsible section ───────────────────────────────────────────────────────
function Section({ icon, title, count, countColor='var(--accent)', filename, children }) {
  const [open, setOpen] = useState(true);
  return (
    <div style={{
      background:'var(--bg2)', border:'1px solid var(--border)',
      borderRadius:10, marginBottom:12, overflow:'hidden',
    }}>
      <div
        onClick={() => setOpen(o => !o)}
        style={{
          display:'flex', alignItems:'center', gap:10,
          padding:'12px 16px', cursor:'pointer',
          borderBottom: open ? '1px solid var(--border)' : 'none',
          userSelect:'none',
        }}
      >
        <span style={{ fontSize:18 }}>{icon}</span>
        <div style={{ flex:1 }}>
          <div style={{ fontSize:14, fontWeight:700, color:'var(--text)' }}>{title}</div>
          {filename && <div style={{ fontSize:11, color:'var(--muted)', marginTop:1 }}>{filename}</div>}
        </div>
        <span style={{
          fontSize:13, fontWeight:700, color:countColor,
          background:'var(--bg3)', borderRadius:6, padding:'2px 10px',
        }}>{count}</span>
        <span style={{ color:'var(--muted)', fontSize:12, marginLeft:4 }}>
          {open ? '▲' : '▼'}
        </span>
      </div>
      {open && <div style={{ padding:'0 16px 14px' }}>{children}</div>}
    </div>
  );
}

// ── Mini table ────────────────────────────────────────────────────────────────
function MiniTable({ rows, cols, maxRows=8 }) {
  const [showAll, setShowAll] = useState(false);
  const displayed = showAll ? rows : rows.slice(0, maxRows);
  if (!rows.length) return <div style={{ fontSize:12, color:'var(--muted)', padding:'10px 0' }}>No data</div>;
  return (
    <>
      <table style={{ width:'100%', borderCollapse:'collapse', fontSize:12, marginTop:10 }}>
        <thead>
          <tr>
            {cols.map(c => (
              <th key={c.key} style={{
                textAlign: c.right ? 'right' : 'left',
                padding:'6px 8px', color:'var(--muted)', fontWeight:600,
                borderBottom:'1px solid var(--border)', whiteSpace:'nowrap', fontSize:11,
              }}>{c.label}</th>
            ))}
          </tr>
        </thead>
        <tbody>
          {displayed.map((row, i) => (
            <tr key={i} style={{ borderBottom:'0.5px solid var(--border)' }}>
              {cols.map(c => (
                <td key={c.key} style={{
                  padding:'5px 8px',
                  textAlign: c.right ? 'right' : 'left',
                  color:'var(--text)',
                  fontFamily: c.mono ? 'monospace' : 'inherit',
                  fontSize: 12,
                }}>
                  {c.pill
                    ? <Pill label={row[c.key] || ''} />
                    : c.right
                      ? (row[c.key] != null ? Number(row[c.key]).toLocaleString('fr-FR') : '—')
                      : (row[c.key] ?? '—')
                  }
                </td>
              ))}
            </tr>
          ))}
        </tbody>
      </table>
      {rows.length > maxRows && (
        <button
          onClick={() => setShowAll(o => !o)}
          style={{
            marginTop:8, fontSize:11, color:'var(--accent)', background:'none',
            border:'none', cursor:'pointer', padding:0,
          }}>
          {showAll ? '▲ Show less' : `▼ Show all ${rows.length} rows`}
        </button>
      )}
    </>
  );
}

// ── MTD columns ───────────────────────────────────────────────────────────────
const MTD_COLS = [
  { key:'account_id',    label:'Account ID',    mono:true },
  { key:'firstname',     label:'Name' },
  { key:'Yesterday ggr', label:'Yesterday GGR', right:true },
  { key:'MTD total ggr', label:'MTD total GGR', right:true },
  { key:'YTD',           label:'YTD',           right:true },
  { key:'Today',         label:'Today',         right:true },
  { key:'game',          label:'Game',          pill:true },
  { key:'offer',         label:'Offer',         pill:true },
  { key:'5%',            label:'5%',            right:true },
  { key:'amount',        label:'Amount',        right:true },
  { key:'Mixsport',      label:'Mixsport',      right:true },
];

const SEGMENT_FILTERS = ['All', 'casino', 'sport bonus', 'MIX'];

// ── Main component ────────────────────────────────────────────────────────────
export default function ResultPage({ result }) {
  const [downloading,  setDownloading]  = useState(false);
  const [search,       setSearch]       = useState('');
  const [segFilter,    setSegFilter]    = useState('All');
  const { setNavigateTo } = useAppContext();
  const [activeTab,    setActiveTab]    = useState('campaigns');

  // Hooks AVANT early return
  const allMtd = useMemo(() => result?.mtdRows || [], [result]);
  const filteredMtd = useMemo(() => {
    let rows = allMtd;
    if (segFilter !== 'All') rows = rows.filter(r => r.game === segFilter);
    if (search.trim()) {
      const q = search.toLowerCase();
      rows = rows.filter(r =>
        String(r.account_id||'').includes(q) ||
        (r.firstname||'').toLowerCase().includes(q) ||
        (r.game||'').toLowerCase().includes(q) ||
        (r.offer||'').toLowerCase().includes(q)
      );
    }
    return rows;
  }, [allMtd, segFilter, search]);

  if (!result) return (
    <p style={{ color:'var(--muted)', fontSize:14 }}>
      No result yet — run the pipeline first.
    </p>
  );

  const { stats, outputFiles, campaigns, mtdRows, countryConfig } = result;
  const currency = countryConfig?.currency || 'XAF';
  const prefix   = countryConfig?.prefix   || 'CM';

  const today = new Date();
  const dd = String(today.getDate()).padStart(2,'0');
  const mm = String(today.getMonth()+1).padStart(2,'0');
  const yy = String(today.getFullYear()).slice(-2);
  const dailyName = outputFiles.find(f => f.filename.match(/^[A-Z]{2}_Daily_/))?.filename
    || `${prefix}_Daily_${dd}.${mm}.${yy}.xlsx`;

  // Sport groups
  const sportFiles = outputFiles
    .filter(f => f.filename.startsWith('Sport_Daily_') || f.filename.startsWith('CM_daily_'))
    .map(f => {
      const amt = f.filename.replace(/^(Sport_Daily_|CM_daily_)/,'').replace('.csv','');
      const count = campaigns?.cm_daily?.[amt]?.length ?? 0;
      return { filename:f.filename, blob:f.blob, amt:Number(amt), count };
    })
    .sort((a,b) => a.amt - b.amt);

  // Casino bonus rows for table
  const cbRows  = campaigns?.casino_bonus_cm || [];
  const cashRows = campaigns?.daily_cashback || [];
  const spribeRows = campaigns?.spribe_freebets || [];
  const phoneRows  = result?.phones || [];

  // Available MTD cols (only those present in data)
  const availCols = MTD_COLS.filter(c => allMtd.length > 0 && c.key in allMtd[0]);

  async function downloadAll() {
    setDownloading(true);
    const zip = new JSZip();
    for (const { blob, filename } of outputFiles) zip.file(filename, blob);
    const zipBlob = await zip.generateAsync({ type:'blob' });
    saveAs(zipBlob, `pipeline_${new Date().toISOString().slice(0,10)}.zip`);
    setDownloading(false);
  }

  // Transférer les fichiers vers Bonus Auto via AppContext
  function transferToBonus() {
    const toCSV = (rows, cols) =>
      [cols.join(','), ...rows.map(r => cols.map(k => r[k] ?? '').join(','))].join('\n');

    const bonusData = {};

    // Casino
    if (cbRows.length > 0) {
      bonusData.casino_csv = toCSV(cbRows, ['account_id','amount','currency']);
      bonusData.casino_filename = 'casino_bonus_cm.csv';
    }

    // Cashback
    if (cashRows.length > 0) {
      const cashMapped = cashRows.map(r => ({
        account_id: r.account_id,
        amount: r[' amount '] ?? r.amount ?? '',
        currency_iso3: r.currency_iso3 ?? 'XAF',
      }));
      bonusData.cashback_csv = toCSV(cashMapped, ['account_id','amount','currency_iso3']);
      bonusData.cashback_filename = 'Daily_Cashback_Template.csv';
    }

    // Sport files (blob → text)
    bonusData.sport_files = [];
    const loadSportFiles = async () => {
      for (const sf of sportFiles) {
        const text = await sf.blob.text();
        bonusData.sport_files.push({
          filename: sf.filename,
          csv_content: text,
          amount_label: String(sf.amt),
        });
      }
      // Stocker dans sessionStorage pour que BonusAutoPage le lise
      sessionStorage.setItem('bonusTransfer', JSON.stringify(bonusData));
      setNavigateTo('bonus-auto');
      // Dispatcher l'event après un délai pour laisser BonusAutoPage se monter
      setTimeout(() => window.dispatchEvent(new Event('bonusTransfer')), 300);
    };
    loadSportFiles();
  }

  return (
    <div>

      {/* ── KPI Banner ── */}
      <div style={{
        display:'grid', gridTemplateColumns:'repeat(5,1fr)',
        gap:10, marginBottom:20,
      }}>
        {[
          { label:'JOUEURS',          value:stats.totalPlayers,           color:'var(--text)' },
          { label:'BONUS DE CASINO',  value:stats.byCampaign.casinoBonus, color:'#3B82F6' },
          { label:'BONUS SPORT',      value:stats.byCampaign.sportBonus,  color:'#22C55E' },
          { label:'REMBOURSEMENT',    value:stats.byCampaign.cashback,    color:'#F59E0B' },
          { label:'SPRIBE / AVIATEUR',value:stats.byCampaign.spribe,      color:'#A78BFA' },
        ].map(k => (
          <div key={k.label} style={{
            background:'var(--bg2)', border:'1px solid var(--border)',
            borderRadius:10, padding:'14px 16px',
          }}>
            <div style={{ fontSize:26, fontWeight:700, color:k.color }}>
              {k.value.toLocaleString()}
            </div>
            <div style={{ fontSize:10, color:'var(--muted)', marginTop:4, fontWeight:600, letterSpacing:'.06em' }}>
              {k.label}
            </div>
          </div>
        ))}
      </div>

      {/* ── Tabs ── */}
      <div style={{ display:'flex', gap:6, marginBottom:16, flexWrap:'wrap', alignItems:'center' }}>
        {[
          { id:'campaigns', label:`🎯 Campagnes` },
          { id:'sport',     label:`⚽ Bonus Sport (${sportFiles.length})` },
          { id:'downloads', label:`⬇ Téléchargements (${outputFiles.length})` },
          { id:'mtd',       label:`📋 Tableau MTD (${allMtd.length})` },
        ].map(t => (
          <button key={t.id}
            onClick={() => setActiveTab(t.id)}
            style={{
              fontSize:12, padding:'7px 14px', borderRadius:7, border:'none',
              cursor:'pointer', fontWeight:600,
              background: activeTab===t.id ? 'var(--accent)' : 'var(--bg3)',
              color: activeTab===t.id ? '#fff' : 'var(--muted)',
            }}>
            {t.label}
          </button>
        ))}
        {/* Bouton transfert vers Bonus Auto */}
        {(cbRows.length > 0 || cashRows.length > 0 || sportFiles.length > 0) && (
          <button onClick={transferToBonus} style={{
            marginLeft:'auto', fontSize:12, padding:'7px 16px', borderRadius:7,
            border:'none', cursor:'pointer', fontWeight:700,
            background:'linear-gradient(135deg,#818CF8,#34D399)',
            color:'#fff', display:'flex', alignItems:'center', gap:6,
          }}>
            🤖 → Envoi Bonus
          </button>
        )}
      </div>

      {/* ── Campagnes tab ── */}
      {activeTab === 'campaigns' && (
        <div style={{ display:'grid', gridTemplateColumns:'1fr 1fr', gap:12 }}>

          <Section icon="🎰" title="Bonus de casino"
            filename="casino_bonus_[préfixe].csv"
            count={cbRows.length} countColor="#3B82F6">
            <MiniTable rows={cbRows} cols={[
              { key:'account_id', label:'ID du compte', mono:true },
              { key:'amount',     label:'Montant', right:true },
              { key:'currency',   label:'Devise' },
            ]}/>
          </Section>

          <Section icon="💸" title="Remboursement quotidien"
            filename="Modèle_de_remboursement_quotidien.csv"
            count={cashRows.length} countColor="#F59E0B">
            <MiniTable rows={cashRows.map(r => ({
              account_id: r.account_id,
              amount: r[' amount '],
              currency: r.currency_iso3,
            }))} cols={[
              { key:'account_id', label:'ID du compte', mono:true },
              { key:'amount',     label:'Montant', right:true },
              { key:'currency',   label:'Devise' },
            ]}/>
          </Section>

          <Section icon="✈️" title="Spribe / Aviateur"
            filename="spribe_freebets.csv"
            count={spribeRows.length} countColor="#A78BFA">
            <MiniTable rows={spribeRows} cols={[
              { key:'opPlayerId',             label:'ID du joueur',      mono:true },
              { key:'ticketAmount_amount',    label:'Montant du billet', right:true },
              { key:'num_of_free_ticketAmounts', label:'Paris gratuits', right:true },
              { key:'game',                   label:'Jeu' },
              { key:'start_date',             label:'Commencer' },
            ]}/>
          </Section>

          <Section icon="📞" title="Extraction de données téléphoniques"
            filename={`${prefix}_YTD_phones.xlsx`}
            count={phoneRows.length} countColor="#22C55E">
            <MiniTable rows={phoneRows} cols={[
              { key:'account_id', label:'ID du compte', mono:true },
              { key:'phone',      label:'Téléphone',    mono:true },
            ]}/>
          </Section>

        </div>
      )}

      {/* ── Sport tab ── */}
      {activeTab === 'sport' && (
        <div>
          <p className="section-label" style={{ marginBottom:12 }}>
            BONUS SPORT QUOTIDIEN — CLASSÉS PAR MONTANT
          </p>
          <div style={{
            display:'grid',
            gridTemplateColumns:'repeat(auto-fill, minmax(160px, 1fr))',
            gap:10, marginBottom:20,
          }}>
            {sportFiles.map(g => (
              <div key={g.filename}
                onClick={() => saveAs(g.blob, g.filename)}
                style={{
                  background:'var(--bg2)', border:'1px solid var(--border)',
                  borderRadius:10, padding:'14px 16px', cursor:'pointer',
                  transition:'border-color .15s',
                }}
                onMouseEnter={e => e.currentTarget.style.borderColor='var(--accent)'}
                onMouseLeave={e => e.currentTarget.style.borderColor='var(--border)'}
              >
                <div style={{ fontSize:10, fontWeight:700, color:'var(--muted)', letterSpacing:'.06em', marginBottom:4 }}>
                  MONTANT DE LA PRIME
                </div>
                <div style={{ fontSize:20, fontWeight:700, color:'var(--text)' }}>
                  {g.amt.toLocaleString('fr-FR')}
                  <span style={{ fontSize:12, color:'var(--muted)', marginLeft:4 }}>{currency}</span>
                </div>
                <div style={{ fontSize:13, color:'var(--muted)', marginTop:4 }}>
                  {g.count} joueurs
                </div>
                <div style={{ fontSize:10, color:'var(--muted)', marginTop:6, fontFamily:'monospace' }}>
                  {g.filename}
                </div>
              </div>
            ))}
          </div>
        </div>
      )}

      {/* ── Downloads tab ── */}
      {activeTab === 'downloads' && (
        <div>
          <p className="section-label" style={{ marginBottom:10 }}>TÉLÉCHARGEMENTS</p>
          <div style={{
            background:'var(--bg2)', border:'1px solid var(--border)',
            borderRadius:10, padding:'16px',
          }}>
            <div style={{ display:'flex', flexWrap:'wrap', gap:8 }}>
              <button
                onClick={downloadAll}
                disabled={downloading}
                style={{
                  fontSize:12, padding:'7px 16px', borderRadius:7,
                  border:'none', cursor:'pointer', fontWeight:700,
                  background:'var(--accent)', color:'#fff',
                }}>
                {downloading ? '⟳ Compression…' : '📦 Tout télécharger'}
              </button>
              {outputFiles.map(f => (
                <button key={f.filename}
                  onClick={() => saveAs(f.blob, f.filename)}
                  style={{
                    fontSize:12, padding:'7px 14px', borderRadius:7,
                    border:'1px solid var(--border)', cursor:'pointer',
                    background:'var(--bg3)', color:'var(--text)', fontWeight:500,
                  }}>
                  {getIcon(f.filename)} {f.filename}
                </button>
              ))}
            </div>
          </div>
        </div>
      )}

      {/* ── MTD table tab ── */}
      {activeTab === 'mtd' && (
        <div>
          <div style={{ display:'flex', alignItems:'center', gap:10, marginBottom:12, flexWrap:'wrap' }}>
            <div style={{ fontSize:16, fontWeight:700, color:'var(--text)', display:'flex', alignItems:'center', gap:8 }}>
              📊 Tableau principal MTD
            </div>
            <input
              value={search} onChange={e => setSearch(e.target.value)}
              placeholder="Rechercher un compte, un jeu…"
              style={{
                flex:1, minWidth:200, padding:'7px 12px', fontSize:12,
                background:'var(--bg3)', border:'1px solid var(--border)',
                borderRadius:7, color:'var(--text)', outline:'none',
              }}
            />
            <span style={{ fontSize:12, color:'var(--muted)', whiteSpace:'nowrap' }}>
              {filteredMtd.length.toLocaleString()} joueurs
            </span>
          </div>

          {/* Segment filter */}
          <div style={{ display:'flex', gap:4, marginBottom:12 }}>
            {SEGMENT_FILTERS.map(f => (
              <button key={f}
                onClick={() => setSegFilter(f)}
                style={{
                  fontSize:11, padding:'5px 14px', borderRadius:6,
                  border:'none', cursor:'pointer', fontWeight:600,
                  background: segFilter===f ? 'var(--accent)' : 'var(--bg3)',
                  color: segFilter===f ? '#fff' : 'var(--muted)',
                }}>
                {f}
              </button>
            ))}
          </div>

          <div style={{ overflowX:'auto', overflowY:'auto', maxHeight:520, background:'var(--bg2)', border:'1px solid var(--border)', borderRadius:10 }}>
            <table style={{ width:'100%', borderCollapse:'collapse', fontSize:12 }}>
              <thead>
                <tr>
                  {availCols.map(c => (
                    <th key={c.key} style={{
                      textAlign: c.right ? 'right' : 'left',
                      padding:'8px 12px', color:'var(--muted)', fontWeight:600,
                      borderBottom:'1px solid var(--border)', whiteSpace:'nowrap',
                      fontSize:11, letterSpacing:'.04em',
                    }}>{c.label}</th>
                  ))}
                </tr>
              </thead>
              <tbody>
                {filteredMtd.slice(0,500).map((row, i) => (
                  <tr key={i} style={{
                    borderBottom:'0.5px solid var(--border)',
                    background: i%2===0 ? 'transparent' : 'rgba(255,255,255,.02)',
                  }}>
                    {availCols.map(c => (
                      <td key={c.key} style={{
                        padding:'6px 12px',
                        textAlign: c.right ? 'right' : 'left',
                        fontFamily: c.mono ? 'monospace' : 'inherit',
                        color:'var(--text)', fontSize:12,
                      }}>
                        {c.pill
                          ? <Pill label={row[c.key] || ''} />
                          : c.right
                            ? (row[c.key] != null && row[c.key] !== ''
                                ? Number(row[c.key]).toLocaleString('fr-FR')
                                : '—')
                            : (row[c.key] ?? '—')
                        }
                      </td>
                    ))}
                  </tr>
                ))}
              </tbody>
            </table>
          </div>

          {filteredMtd.length > 500 && (
            <div style={{ fontSize:11, color:'var(--muted)', marginTop:6, textAlign:'right' }}>
              Affichage des 500 premiers sur {filteredMtd.length.toLocaleString()} joueurs
            </div>
          )}
        </div>
      )}

    </div>
  );
}

function getIcon(filename) {
  if (filename.match(/^[A-Z]{2}_Daily_/))        return '📋';
  if (filename.includes('_YTD_phones'))           return '📞';
  if (filename.includes('casino_bonus'))          return '🎰';
  if (filename.includes('Cashback') || filename.includes('remboursement')) return '💸';
  if (filename.includes('spribe'))                return '✈️';
  if (filename.startsWith('Sport_Daily_'))        return '⚽';
  return '📄';
}
