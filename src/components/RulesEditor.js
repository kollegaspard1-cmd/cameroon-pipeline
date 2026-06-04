// src/components/RulesEditor.js
// UI component for editing, importing, saving rounding rules

import React, { useState, useRef } from 'react';
import XLSX from 'xlsx-js-style';
import Papa from 'papaparse';
import { saveRulesForCountry, deleteRulesForCountry, DEFAULT_RULES } from '../lib/roundingRules';

const OPERATORS = [
  { value: 'lt',  label: '< (inférieur à)' },
  { value: 'gte', label: '≥ (supérieur ou égal à)' },
  { value: 'lte', label: '≤ (inférieur ou égal à)' },
  { value: 'gt',  label: '> (supérieur à)' },
];

const T = {
  FR: {
    title:        'Regles arrondi',
    subtitle:     'Comment les montants bonus sont arrondis',
    condition:    'Condition',
    threshold:    'Seuil',
    step:         'Pas arrondi',
    label:        'Libellé',
    add:          '+ Ajouter une règle',
    save:         '💾 Sauvegarder',
    reset:        '↺ Réinitialiser',
    delete:       '🗑 Supprimer config',
    import:       '📥 Importer JSON',
    export:       '📤 Exporter JSON',
    saving:       'Sauvegarde…',
    saved:        '✅ Sauvegardé !',
    error:        '❌ Erreur',
    preview:      'Aperçu',
    previewTest:  'Valeur test',
    result:       'Résultat',
    deleteConfirm:'Réinitialiser aux valeurs par défaut ?',
    noRules:      'Aucune règle — ajoutez-en une',
    firebaseNote: 'Les règles sont sauvegardées par pays sur Firebase',
  },
  EN: {
    title:        'Rounding Rules',
    subtitle:     'Defines how bonus amounts are rounded',
    condition:    'Condition',
    threshold:    'Threshold',
    step:         'Rounding step',
    label:        'Label',
    add:          '+ Add rule',
    save:         '💾 Save',
    reset:        '↺ Reset',
    delete:       '🗑 Delete config',
    import:       '📥 Import JSON',
    export:       '📤 Export JSON',
    saving:       'Saving…',
    saved:        '✅ Saved!',
    error:        '❌ Error',
    preview:      'Preview',
    previewTest:  'Test value',
    result:       'Result',
    deleteConfirm:'Reset to default values?',
    noRules:      'No rules — add one',
    firebaseNote: 'Rules are saved per country on Firebase',
  }
};

export default function RulesEditor({ rules, onChange, countryCode, lang = 'FR', onClose }) {
  const [saving, setSaving]   = useState('');
  const [testVal, setTestVal] = useState(750);
  const importRef = useRef();

  const t = T[lang] || T.FR;

  const addRule = () => {
    const newRule = {
      id: `r${Date.now()}`,
      label: '',
      threshold: 1000,
      step: 100,
      operator: 'lt',
    };
    onChange([...rules, newRule]);
  };

  const updateRule = (id, field, value) => {
    onChange(rules.map(r => r.id === id ? { ...r, [field]: field === 'step' || field === 'threshold' ? Number(value) : value } : r));
  };

  const removeRule = (id) => onChange(rules.filter(r => r.id !== id));

  const previewResult = (val) => {
    if (!val || isNaN(val)) return '—';
    for (const rule of [...rules].sort((a,b) => {
      if (a.operator==='lt' && b.operator==='gte') return -1;
      return a.threshold - b.threshold;
    })) {
      const match = rule.operator==='lt'  ? val < rule.threshold
                  : rule.operator==='gte' ? val >= rule.threshold
                  : rule.operator==='lte' ? val <= rule.threshold
                  : rule.operator==='gt'  ? val > rule.threshold : false;
      if (match) return Math.floor(val / rule.step) * rule.step;
    }
    return Math.floor(val / 100) * 100;
  };

  const saveToFirebase = async () => {
    setSaving('saving');
    try {
      await saveRulesForCountry(countryCode, rules);
      setSaving('saved');
      setTimeout(() => setSaving(''), 2000);
    } catch(e) {
      setSaving('error');
      setTimeout(() => setSaving(''), 3000);
    }
  };

  const resetToDefault = async () => {
    if (!window.confirm(t.deleteConfirm)) return;
    setSaving('saving');
    try {
      await deleteRulesForCountry(countryCode);
      onChange([...DEFAULT_RULES]);
      setSaving('saved');
      setTimeout(() => setSaving(''), 2000);
    } catch(e) {
      setSaving('error');
      setTimeout(() => setSaving(''), 3000);
    }
  };

  const COND_LABELS = {
    lt:  lang==='FR'?'inférieur à (<)':'less than (<)',
    gte: lang==='FR'?'supérieur ou égal à (≥)':'greater or equal (≥)',
    lte: lang==='FR'?'inférieur ou égal à (≤)':'less or equal (≤)',
    gt:  lang==='FR'?'supérieur à (>)':'greater than (>)',
  };

  const rulesToRows = () => rules.map(r => ({
    [lang==='FR'?'Libellé':'Label']: r.label,
    [lang==='FR'?'Condition':'Condition']: COND_LABELS[r.operator] || r.operator,
    [lang==='FR'?'Seuil':'Threshold']: r.threshold,
    [lang==='FR'?'Pas arrondi':'Rounding step']: r.step,
  }));

  const exportCSV = () => {
    const rows = rulesToRows();
    const csv = [
      Object.keys(rows[0]).join(','),
      ...rows.map(r => Object.values(r).join(','))
    ].join('\n');
    const blob = new Blob([csv], { type: 'text/csv;charset=utf-8;' });
    const a = document.createElement('a');
    a.href = URL.createObjectURL(blob);
    a.download = `regles_arrondi_${countryCode}.csv`;
    a.click();
  };

  const exportXLSX_rules = () => {
    const rows = rulesToRows();
    const ws = XLSX.utils.json_to_sheet(rows);
    // Style headers
    const headers = Object.keys(rows[0]);
    headers.forEach((h, i) => {
      const addr = XLSX.utils.encode_cell({r:0, c:i});
      if (ws[addr]) ws[addr].s = {
        font: { bold:true, color:{rgb:'FFFFFF'} },
        fill: { patternType:'solid', fgColor:{rgb:'1F3864'} },
        alignment: { horizontal:'center' }
      };
    });
    ws['!cols'] = [{ wch:20 },{ wch:28 },{ wch:12 },{ wch:16 }];
    const wb = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(wb, ws, lang==='FR'?'Règles arrondi':'Rounding Rules');
    const buf = XLSX.write(wb, { bookType:'xlsx', type:'array' });
    const blob = new Blob([buf], { type:'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet' });
    const a = document.createElement('a');
    a.href = URL.createObjectURL(blob);
    a.download = `regles_arrondi_${countryCode}.xlsx`;
    a.click();
  };

  const downloadTemplate = () => {
    const templateRows = [
      { [lang==='FR'?'Libellé':'Label']: lang==='FR'?'< 1 000':'< 1,000', [lang==='FR'?'Condition':'Condition']: COND_LABELS['lt'], [lang==='FR'?'Seuil':'Threshold']: 1000, [lang==='FR'?'Pas arrondi':'Rounding step']: 100 },
      { [lang==='FR'?'Libellé':'Label']: lang==='FR'?'≥ 1 000':'≥ 1,000', [lang==='FR'?'Condition':'Condition']: COND_LABELS['gte'], [lang==='FR'?'Seuil':'Threshold']: 1000, [lang==='FR'?'Pas arrondi':'Rounding step']: 500 },
    ];
    const ws = XLSX.utils.json_to_sheet(templateRows);
    const headers = Object.keys(templateRows[0]);
    headers.forEach((h, i) => {
      const addr = XLSX.utils.encode_cell({r:0, c:i});
      if (ws[addr]) ws[addr].s = {
        font: { bold:true, color:{rgb:'FFFFFF'} },
        fill: { patternType:'solid', fgColor:{rgb:'375623'} },
        alignment: { horizontal:'center' }
      };
    });
    ws['!cols'] = [{ wch:20 },{ wch:28 },{ wch:12 },{ wch:16 }];
    const wb = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(wb, ws, 'Template');
    const buf = XLSX.write(wb, { bookType:'xlsx', type:'array' });
    const blob = new Blob([buf], { type:'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet' });
    const a = document.createElement('a');
    a.href = URL.createObjectURL(blob);
    a.download = `template_regles_arrondi.xlsx`;
    a.click();
  };

  const parseCondition = (str) => {
    const s = (str||'').toLowerCase();
    if (s.includes('<') && !s.includes('=') && !s.includes('>')) return 'lt';
    if (s.includes('≥') || s.includes('>=') || (s.includes('>') && s.includes('='))) return 'gte';
    if (s.includes('≤') || s.includes('<=') || (s.includes('<') && s.includes('='))) return 'lte';
    if (s.includes('>') && !s.includes('=') && !s.includes('<')) return 'gt';
    if (s === 'lt') return 'lt';
    if (s === 'gte') return 'gte';
    if (s === 'lte') return 'lte';
    if (s === 'gt') return 'gt';
    return 'lt';
  };

  const importFile = (e) => {
    const file = e.target.files[0];
    if (!file) return;
    const ext = file.name.split('.').pop().toLowerCase();
    if (ext === 'csv') {
      Papa.parse(file, {
        header: true, skipEmptyLines: true,
        complete: ({ data }) => {
          const parsed = data.map((row, i) => {
            const vals = Object.values(row);
            const keys = Object.keys(row);
            const labelKey = keys.find(k => k.toLowerCase().includes('lib') || k.toLowerCase().includes('label'));
            const condKey  = keys.find(k => k.toLowerCase().includes('cond'));
            const seuilKey = keys.find(k => k.toLowerCase().includes('seuil') || k.toLowerCase().includes('thresh'));
            const stepKey  = keys.find(k => k.toLowerCase().includes('pas') || k.toLowerCase().includes('step') || k.toLowerCase().includes('arrondi'));
            return {
              id: `r${Date.now()}_${i}`,
              label:     row[labelKey] || vals[0] || '',
              operator:  parseCondition(row[condKey] || vals[1] || 'lt'),
              threshold: parseFloat(row[seuilKey] || vals[2]) || 1000,
              step:      parseFloat(row[stepKey]  || vals[3]) || 100,
            };
          });
          onChange(parsed);
        }
      });
    } else {
      const reader = new FileReader();
      reader.onload = (ev) => {
        const wb = XLSX.read(ev.target.result, { type:'array' });
        const ws = wb.Sheets[wb.SheetNames[0]];
        const data = XLSX.utils.sheet_to_json(ws, { defval:'' });
        const parsed = data.map((row, i) => {
          const keys = Object.keys(row);
          const labelKey = keys.find(k => k.toLowerCase().includes('lib') || k.toLowerCase().includes('label'));
          const condKey  = keys.find(k => k.toLowerCase().includes('cond'));
          const seuilKey = keys.find(k => k.toLowerCase().includes('seuil') || k.toLowerCase().includes('thresh'));
          const stepKey  = keys.find(k => k.toLowerCase().includes('pas') || k.toLowerCase().includes('step') || k.toLowerCase().includes('arrondi'));
          const vals = Object.values(row);
          return {
            id: `r${Date.now()}_${i}`,
            label:     String(row[labelKey] || vals[0] || ''),
            operator:  parseCondition(String(row[condKey] || vals[1] || 'lt')),
            threshold: parseFloat(String(row[seuilKey] || vals[2])) || 1000,
            step:      parseFloat(String(row[stepKey]  || vals[3])) || 100,
          };
        }).filter(r => r.threshold && r.step);
        onChange(parsed);
      };
      reader.readAsArrayBuffer(file);
    }
    e.target.value = '';
  };

  const btnBase = { padding:'7px 14px', borderRadius:6, fontSize:12, fontWeight:600, cursor:'pointer', border:'1px solid var(--border)', background:'var(--bg2)', color:'var(--text2)' };

  return (
    <div style={{ position:'fixed', inset:0, background:'rgba(0,0,0,0.4)', zIndex:1000, display:'flex', alignItems:'center', justifyContent:'center' }}>
      <div style={{ background:'var(--bg2)', borderRadius:12, padding:'1.5rem', width:'100%', maxWidth:680, maxHeight:'90vh', overflowY:'auto', boxShadow:'0 8px 32px rgba(0,0,0,0.2)' }}>

        {/* Header */}
        <div style={{ display:'flex', alignItems:'flex-start', justifyContent:'space-between', marginBottom:'1rem' }}>
          <div>
            <h3 style={{ fontSize:16, fontWeight:700, color:'var(--text)', margin:0 }}>{t.title} — {countryCode}</h3>
            <p style={{ fontSize:12, color:'var(--muted)', margin:'4px 0 0' }}>{t.subtitle}</p>
          </div>
          <button onClick={onClose} style={{ ...btnBase, padding:'4px 10px', fontSize:16, lineHeight:1 }}>✕</button>
        </div>

        {/* Rules list */}
        {rules.length === 0 ? (
          <div style={{ textAlign:'center', padding:'2rem', color:'var(--muted)', fontSize:13 }}>{t.noRules}</div>
        ) : (
          <div style={{ display:'flex', flexDirection:'column', gap:'0.6rem', marginBottom:'1rem' }}>
            {rules.map((rule, i) => (
              <div key={rule.id} style={{ display:'grid', gridTemplateColumns:'1fr 120px 80px 120px 36px', gap:8, alignItems:'center',
                padding:'10px 12px', background:'var(--bg3)', borderRadius:8, border:'1px solid var(--border)' }}>
                <div>
                  <div style={{ fontSize:10, color:'var(--muted)', fontWeight:600, marginBottom:3 }}>{t.label}</div>
                  <input value={rule.label} onChange={e => updateRule(rule.id,'label',e.target.value)}
                    placeholder={`Règle ${i+1}`} className="field-input" style={{ width:'100%', fontSize:12 }} />
                </div>
                <div>
                  <div style={{ fontSize:10, color:'var(--muted)', fontWeight:600, marginBottom:3 }}>{t.condition}</div>
                  <select value={rule.operator} onChange={e => updateRule(rule.id,'operator',e.target.value)}
                    className="field-input" style={{ width:'100%', fontSize:11, cursor:'pointer' }}>
                    {OPERATORS.map(op => <option key={op.value} value={op.value}>{op.label}</option>)}
                  </select>
                </div>
                <div>
                  <div style={{ fontSize:10, color:'var(--muted)', fontWeight:600, marginBottom:3 }}>{t.threshold}</div>
                  <input type="number" value={rule.threshold} onChange={e => updateRule(rule.id,'threshold',e.target.value)}
                    className="field-input" style={{ width:'100%', fontSize:12 }} />
                </div>
                <div>
                  <div style={{ fontSize:10, color:'var(--muted)', fontWeight:600, marginBottom:3 }}>{t.step}</div>
                  <input type="number" value={rule.step} onChange={e => updateRule(rule.id,'step',e.target.value)}
                    className="field-input" style={{ width:'100%', fontSize:12 }} />
                </div>
                <button onClick={() => removeRule(rule.id)}
                  style={{ ...btnBase, padding:'6px 8px', color:'var(--danger)', borderColor:'rgba(239,68,68,0.3)', marginTop:16 }}>
                  🗑
                </button>
              </div>
            ))}
          </div>
        )}

        {/* Preview */}
        <div style={{ background:'var(--accent-dim)', borderRadius:8, padding:'10px 14px', marginBottom:'1rem', display:'flex', alignItems:'center', gap:'1rem' }}>
          <span style={{ fontSize:12, fontWeight:600, color:'var(--accent)' }}>{t.preview} :</span>
          <input type="number" value={testVal} onChange={e => setTestVal(Number(e.target.value))}
            className="field-input" style={{ width:100, fontSize:12 }} />
          <span style={{ fontSize:12, color:'var(--muted)' }}>→</span>
          <span style={{ fontSize:14, fontWeight:700, color:'var(--accent)', fontFamily:'JetBrains Mono, monospace' }}>
            {previewResult(testVal).toLocaleString()}
          </span>
          <span style={{ fontSize:11, color:'var(--muted)', marginLeft:4 }}>
            ({rules.find(r => {
              const v = testVal;
              return r.operator==='lt' ? v<r.threshold : r.operator==='gte' ? v>=r.threshold : r.operator==='lte' ? v<=r.threshold : v>r.threshold;
            })?.label || 'défaut'})
          </span>
        </div>

        {/* Actions */}
        <div style={{ display:'flex', gap:8, flexWrap:'wrap', alignItems:'center' }}>
          <button onClick={addRule} style={{ ...btnBase, color:'var(--accent)', borderColor:'var(--accent)', background:'var(--accent-dim)' }}>
            {t.add}
          </button>
          <button onClick={saveToFirebase}
            style={{ ...btnBase, background:'var(--accent)', color:'#fff', borderColor:'var(--accent)' }}>
            {saving === 'saving' ? t.saving : saving === 'saved' ? t.saved : saving === 'error' ? t.error : t.save}
          </button>
          <button onClick={resetToDefault} style={{ ...btnBase }}>
            {t.reset}
          </button>
          <button onClick={exportCSV} style={{ ...btnBase }}>
            📊 {lang==='FR'?'Exporter CSV':'Export CSV'}
          </button>
          <button onClick={exportXLSX_rules} style={{ ...btnBase }}>
            📗 {lang==='FR'?'Exporter XLSX':'Export XLSX'}
          </button>
          <button onClick={downloadTemplate} style={{ ...btnBase, color:'var(--deposit)', borderColor:'var(--deposit)' }}>
            📥 {lang==='FR'?'Template XLSX':'XLSX Template'}
          </button>
          <button onClick={() => importRef.current.click()} style={{ ...btnBase }}>
            📂 {lang==='FR'?'Importer CSV/XLSX':'Import CSV/XLSX'}
          </button>
          <input ref={importRef} type="file" accept=".csv,.xlsx,.xls" hidden onChange={importFile} />
          <span style={{ marginLeft:'auto', fontSize:10, color:'var(--muted2)', fontStyle:'italic' }}>
            {t.firebaseNote}
          </span>
        </div>
      </div>
    </div>
  );
}
