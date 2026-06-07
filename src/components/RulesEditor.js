// src/components/RulesEditor.js
import React, { useState, useRef } from 'react';
import XLSX from 'xlsx-js-style';
import { saveRulesForCountry, deleteRulesForCountry, DEFAULT_RULES } from '../lib/roundingRules';

const OPERATORS_FR = [
  { value: 'lt',      label: '< (inferieur a)' },
  { value: 'gte',     label: '>= (superieur ou egal a)' },
  { value: 'lte',     label: '<= (inferieur ou egal a)' },
  { value: 'gt',      label: '> (superieur a)' },
  { value: 'between', label: '<-> Entre X et Y' },
];
const OPERATORS_EN = [
  { value: 'lt',      label: '< (less than)' },
  { value: 'gte',     label: '>= (greater or equal)' },
  { value: 'lte',     label: '<= (less or equal)' },
  { value: 'gt',      label: '> (greater than)' },
  { value: 'between', label: '<-> Between X and Y' },
];

const T = {
  FR: {
    title:        'Regles arrondi',
    subtitle:     'Comment les montants bonus sont arrondis',
    condition:    'Condition',
    threshold:    'Seuil',
    threshold2:   'Seuil Max',
    from:         'De (>=)',
    to:           'A (<)',
    step:         'Pas arrondi',
    label:        'Libelle',
    add:          '+ Ajouter une regle',
    save:         'Sauvegarder',
    reset:        'Reinitialiser',
    exportXlsx:   'Exporter XLSX',
    template:     'Template XLSX',
    importXlsx:   'Importer XLSX',
    saving:       'Sauvegarde...',
    saved:        'Sauvegarde !',
    error:        'Erreur',
    preview:      'Apercu',
    deleteConfirm:'Reinitialiser aux valeurs par defaut ?',
    noRules:      'Aucune regle - ajoutez-en une',
    firebaseNote: 'Les regles sont sauvegardees par pays sur Firebase',
    delete:       'Supprimer',
  },
  EN: {
    title:        'Rounding Rules',
    subtitle:     'Defines how bonus amounts are rounded',
    condition:    'Condition',
    threshold:    'Threshold',
    threshold2:   'Max Threshold',
    from:         'From (>=)',
    to:           'To (<)',
    step:         'Rounding step',
    label:        'Label',
    add:          '+ Add rule',
    save:         'Save',
    reset:        'Reset',
    exportXlsx:   'Export XLSX',
    template:     'XLSX Template',
    importXlsx:   'Import XLSX',
    saving:       'Saving...',
    saved:        'Saved!',
    error:        'Error',
    preview:      'Preview',
    deleteConfirm:'Reset to default values?',
    noRules:      'No rules - add one',
    firebaseNote: 'Rules are saved per country on Firebase',
    delete:       'Delete',
  }
};

export default function RulesEditor({ rules, onChange, countryCode, lang = 'FR', onClose }) {
  const [saving, setSaving]   = useState('');
  const [testVal, setTestVal] = useState(750);
  const importRef = useRef();

  const t = T[lang] || T.EN;
  const OPERATORS = lang === 'EN' ? OPERATORS_EN : OPERATORS_FR;

  const addRule = () => {
    onChange([...rules, {
      id: `r${Date.now()}`,
      label: '', threshold: 1000, threshold2: 5000, step: 100, operator: 'lt',
    }]);
  };

  const updateRule = (id, field, value) => {
    onChange(rules.map(r => r.id === id
      ? { ...r, [field]: ['step','threshold','threshold2'].includes(field) ? Number(value) : value }
      : r));
  };

  const removeRule = (id) => onChange(rules.filter(r => r.id !== id));

  const matchRule = (rule, val) => {
    if (rule.operator === 'between') return val >= rule.threshold && val < (rule.threshold2 || rule.threshold + 1000);
    if (rule.operator === 'lt')  return val < rule.threshold;
    if (rule.operator === 'gte') return val >= rule.threshold;
    if (rule.operator === 'lte') return val <= rule.threshold;
    if (rule.operator === 'gt')  return val > rule.threshold;
    return false;
  };

  const previewResult = (val) => {
    if (!val || isNaN(val)) return '—';
    const sorted = [...rules].sort((a,b) => a.threshold - b.threshold);
    for (const rule of sorted) {
      if (matchRule(rule, val)) return Math.floor(val / rule.step) * rule.step;
    }
    return Math.floor(val / 100) * 100;
  };

  const matchedRule = rules.find(r => matchRule(r, testVal));

  const saveToFirebase = async () => {
    setSaving('saving');
    try {
      await saveRulesForCountry(countryCode, rules);
      setSaving('saved');
      setTimeout(() => setSaving(''), 2000);
    } catch(e) { setSaving('error'); setTimeout(() => setSaving(''), 3000); }
  };

  const resetToDefault = async () => {
    if (!window.confirm(t.deleteConfirm)) return;
    setSaving('saving');
    try {
      await deleteRulesForCountry(countryCode);
      onChange([...DEFAULT_RULES]);
      setSaving('saved');
      setTimeout(() => setSaving(''), 2000);
    } catch(e) { setSaving('error'); setTimeout(() => setSaving(''), 3000); }
  };

  const COND_LABELS = {
    lt:      lang==='EN'?'less than (<)':'inferieur a (<)',
    gte:     lang==='EN'?'greater or equal (>=)':'superieur ou egal a (>=)',
    lte:     lang==='EN'?'less or equal (<=)':'inferieur ou egal a (<=)',
    gt:      lang==='EN'?'greater than (>)':'superieur a (>)',
    between: lang==='EN'?'between X and Y':'entre X et Y',
  };

  const rulesToRows = () => rules.map(r => ({
    [t.label]:     r.label,
    [t.condition]: COND_LABELS[r.operator] || r.operator,
    [t.threshold]: r.threshold,
    [t.threshold2]: r.operator==='between' ? (r.threshold2||'') : '',
    [t.step]:      r.step,
  }));

  const exportXLSX_rules = () => {
    const rows = rulesToRows();
    const ws = XLSX.utils.json_to_sheet(rows);
    Object.keys(rows[0]).forEach((h, i) => {
      const addr = XLSX.utils.encode_cell({r:0, c:i});
      if (ws[addr]) ws[addr].s = {
        font: { bold:true, color:{rgb:'FFFFFF'} },
        fill: { patternType:'solid', fgColor:{rgb:'1F3864'} },
        alignment: { horizontal:'center' }
      };
    });
    ws['!cols'] = [{wch:20},{wch:26},{wch:12},{wch:16},{wch:16}];
    const wb = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(wb, ws, t.title.slice(0,31));
    const buf = XLSX.write(wb, { bookType:'xlsx', type:'array' });
    const blob = new Blob([buf], { type:'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet' });
    const a = document.createElement('a');
    a.href = URL.createObjectURL(blob);
    a.download = `regles_arrondi_${countryCode}.xlsx`;
    a.click();
  };

  const downloadTemplate = () => {
    const templateRows = [
      { [t.label]:'< 1 000',  [t.condition]:COND_LABELS['lt'],  [t.threshold]:1000, [t.threshold2]:'', [t.step]:100 },
      { [t.label]:'>= 1 000', [t.condition]:COND_LABELS['gte'], [t.threshold]:1000, [t.threshold2]:'', [t.step]:500 },
    ];
    const ws = XLSX.utils.json_to_sheet(templateRows);
    Object.keys(templateRows[0]).forEach((h, i) => {
      const addr = XLSX.utils.encode_cell({r:0, c:i});
      if (ws[addr]) ws[addr].s = {
        font: { bold:true, color:{rgb:'FFFFFF'} },
        fill: { patternType:'solid', fgColor:{rgb:'375623'} },
        alignment: { horizontal:'center' }
      };
    });
    ws['!cols'] = [{wch:20},{wch:26},{wch:12},{wch:16},{wch:16}];
    const wb = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(wb, ws, 'Template');
    const buf = XLSX.write(wb, { bookType:'xlsx', type:'array' });
    const blob = new Blob([buf], { type:'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet' });
    const a = document.createElement('a');
    a.href = URL.createObjectURL(blob);
    a.download = 'template_regles_arrondi.xlsx';
    a.click();
  };

  const parseCondition = (str) => {
    const s = (str||'').toLowerCase();
    if (s.includes('between') || s.includes('entre')) return 'between';
    if (s.includes('<') && !s.includes('=') && !s.includes('>')) return 'lt';
    if (s.includes('>=') || s.includes('>= ') || s.includes('superieur ou')) return 'gte';
    if (s.includes('<=') || s.includes('<= ') || s.includes('inferieur ou')) return 'lte';
    if (s.includes('>') && !s.includes('=')) return 'gt';
    if (['lt','gte','lte','gt','between'].includes(s)) return s;
    return 'lt';
  };

  const importFile = (e) => {
    const file = e.target.files[0];
    if (!file) return;
    const reader = new FileReader();
    reader.onload = (ev) => {
      const wb = XLSX.read(ev.target.result, { type:'array' });
      const ws = wb.Sheets[wb.SheetNames[0]];
      const data = XLSX.utils.sheet_to_json(ws, { defval:'' });
      const parsed = data.map((row, i) => {
        const keys = Object.keys(row);
        const vals = Object.values(row);
        const labelKey  = keys.find(k => k.toLowerCase().includes('lib') || k.toLowerCase().includes('label'));
        const condKey   = keys.find(k => k.toLowerCase().includes('cond'));
        const seuilKey  = keys.find(k => (k.toLowerCase().includes('seuil') || k.toLowerCase().includes('thresh')) && !k.toLowerCase().includes('max'));
        const seuil2Key = keys.find(k => k.toLowerCase().includes('max') || k.toLowerCase().includes('seuil2') || k.toLowerCase().includes('threshold2'));
        const stepKey   = keys.find(k => k.toLowerCase().includes('pas') || k.toLowerCase().includes('step') || k.toLowerCase().includes('arrondi'));
        return {
          id: `r${Date.now()}_${i}`,
          label:      String(row[labelKey]  || vals[0] || ''),
          operator:   parseCondition(String(row[condKey] || vals[1] || 'lt')),
          threshold:  parseFloat(String(row[seuilKey]  || vals[2])) || 1000,
          threshold2: seuil2Key && row[seuil2Key] ? parseFloat(String(row[seuil2Key])) : 5000,
          step:       parseFloat(String(row[stepKey]   || vals[4])) || 100,
        };
      }).filter(r => r.threshold && r.step);
      onChange(parsed);
    };
    reader.readAsArrayBuffer(file);
    e.target.value = '';
  };

  const btnBase = {
    padding:'7px 14px', borderRadius:6, fontSize:12, fontWeight:600,
    cursor:'pointer', border:'1px solid var(--border)',
    background:'var(--bg2)', color:'var(--text2)',
    display:'flex', alignItems:'center', gap:6,
  };

  return (
    <div style={{ position:'fixed', inset:0, background:'rgba(0,0,0,0.4)', zIndex:1000, display:'flex', alignItems:'center', justifyContent:'center', padding:'1rem' }}>
      <div style={{ background:'var(--bg2)', borderRadius:12, padding:'1.5rem', width:'100%', maxWidth:720, maxHeight:'90vh', overflowY:'auto', boxShadow:'0 8px 32px rgba(0,0,0,0.2)' }}>

        {/* Header */}
        <div style={{ display:'flex', alignItems:'flex-start', justifyContent:'space-between', marginBottom:'1rem' }}>
          <div>
            <h3 style={{ fontSize:16, fontWeight:700, color:'var(--text)', margin:0 }}>{t.title} — {countryCode}</h3>
            <p style={{ fontSize:12, color:'var(--muted)', margin:'4px 0 0' }}>{t.subtitle}</p>
          </div>
          <button onClick={onClose} style={{ ...btnBase, padding:'4px 10px', fontSize:16 }}>✕</button>
        </div>

        {/* Rules list */}
        {rules.length === 0 ? (
          <div style={{ textAlign:'center', padding:'2rem', color:'var(--muted)', fontSize:13 }}>{t.noRules}</div>
        ) : (
          <div style={{ display:'flex', flexDirection:'column', gap:'0.6rem', marginBottom:'1rem' }}>
            {rules.map((rule, i) => (
              <div key={rule.id} style={{
                display:'grid',
                gridTemplateColumns: rule.operator==='between'
                  ? '1fr 130px 80px 80px 110px 36px'
                  : '1fr 130px 90px 110px 36px',
                gap:8, alignItems:'end',
                padding:'10px 12px', background:'var(--bg3)',
                borderRadius:8, border:'1px solid var(--border)'
              }}>
                {/* Label */}
                <div>
                  <div style={{ fontSize:10, color:'var(--muted)', fontWeight:600, marginBottom:3 }}>{t.label}</div>
                  <input value={rule.label} onChange={e => updateRule(rule.id,'label',e.target.value)}
                    placeholder={`${t.label} ${i+1}`} className="field-input" style={{ width:'100%', fontSize:12 }} />
                </div>
                {/* Condition */}
                <div>
                  <div style={{ fontSize:10, color:'var(--muted)', fontWeight:600, marginBottom:3 }}>{t.condition}</div>
                  <select value={rule.operator} onChange={e => updateRule(rule.id,'operator',e.target.value)}
                    className="field-input" style={{ width:'100%', fontSize:11, cursor:'pointer' }}>
                    {OPERATORS.map(op => <option key={op.value} value={op.value}>{op.label}</option>)}
                  </select>
                </div>
                {/* Threshold */}
                <div>
                  <div style={{ fontSize:10, color:'var(--muted)', fontWeight:600, marginBottom:3 }}>
                    {rule.operator==='between' ? t.from : t.threshold}
                  </div>
                  <input type="number" value={rule.threshold} onChange={e => updateRule(rule.id,'threshold',e.target.value)}
                    className="field-input" style={{ width:'100%', fontSize:12 }} />
                </div>
                {/* Threshold2 (between only) */}
                {rule.operator==='between' && (
                  <div>
                    <div style={{ fontSize:10, color:'var(--muted)', fontWeight:600, marginBottom:3 }}>{t.to}</div>
                    <input type="number" value={rule.threshold2||''} onChange={e => updateRule(rule.id,'threshold2',e.target.value)}
                      className="field-input" style={{ width:'100%', fontSize:12 }} />
                  </div>
                )}
                {/* Step */}
                <div>
                  <div style={{ fontSize:10, color:'var(--muted)', fontWeight:600, marginBottom:3 }}>{t.step}</div>
                  <input type="number" value={rule.step} onChange={e => updateRule(rule.id,'step',e.target.value)}
                    className="field-input" style={{ width:'100%', fontSize:12 }} />
                </div>
                {/* Delete */}
                <button onClick={() => removeRule(rule.id)}
                  style={{ ...btnBase, padding:'6px 8px', color:'var(--danger)', borderColor:'rgba(239,68,68,0.3)', justifyContent:'center' }}>
                  🗑
                </button>
              </div>
            ))}
          </div>
        )}

        {/* Preview */}
        <div style={{ background:'var(--accent-dim)', borderRadius:8, padding:'10px 14px', marginBottom:'1rem', display:'flex', alignItems:'center', gap:'1rem', flexWrap:'wrap' }}>
          <span style={{ fontSize:12, fontWeight:600, color:'var(--accent)' }}>{t.preview} :</span>
          <input type="number" value={testVal} onChange={e => setTestVal(Number(e.target.value))}
            className="field-input" style={{ width:100, fontSize:12 }} />
          <span style={{ fontSize:12, color:'var(--muted)' }}>→</span>
          <span style={{ fontSize:14, fontWeight:700, color:'var(--accent)', fontFamily:'JetBrains Mono, monospace' }}>
            {previewResult(testVal).toLocaleString()}
          </span>
          {matchedRule && (
            <span style={{ fontSize:11, color:'var(--muted)' }}>({matchedRule.label || t.label})</span>
          )}
        </div>

        {/* Actions */}
        <div style={{ display:'flex', gap:8, flexWrap:'wrap', alignItems:'center' }}>
          <button onClick={addRule} style={{ ...btnBase, color:'var(--accent)', borderColor:'var(--accent)', background:'var(--accent-dim)' }}>
            {t.add}
          </button>
          <button onClick={saveToFirebase}
            style={{ ...btnBase, background:'var(--accent)', color:'#fff', borderColor:'var(--accent)' }}>
            💾 {saving==='saving'?t.saving:saving==='saved'?t.saved:saving==='error'?t.error:t.save}
          </button>
          <button onClick={resetToDefault} style={{ ...btnBase }}>
            ↺ {t.reset}
          </button>
          <button onClick={exportXLSX_rules} style={{ ...btnBase }}>
            📗 {t.exportXlsx}
          </button>
          <button onClick={downloadTemplate} style={{ ...btnBase, color:'var(--deposit)', borderColor:'var(--deposit)' }}>
            📥 {t.template}
          </button>
          <button onClick={() => importRef.current.click()} style={{ ...btnBase }}>
            📂 {t.importXlsx}
          </button>
          <input ref={importRef} type="file" accept=".xlsx,.xls" hidden onChange={importFile} />
          <span style={{ marginLeft:'auto', fontSize:10, color:'var(--muted2)', fontStyle:'italic' }}>
            {t.firebaseNote}
          </span>
        </div>
      </div>
    </div>
  );
}
