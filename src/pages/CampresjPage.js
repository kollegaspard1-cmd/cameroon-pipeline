// src/pages/CampresjPage.js
// Fusion ObjectifsPage + RapportsPage — cycle mensuel unifié
import React, { useState, useRef } from 'react';
import Papa from 'papaparse';
import XLSX from 'xlsx-js-style';
import { saveAs } from 'file-saver';
import JSZip from 'jszip';
import { useAppContext } from '../AppContext';
import { saveRecap, loadRecapSaved, deleteRecap, saveListe, loadListe, deleteListe } from '../lib/savedFiles';

export default function CampresjPage() {
  // ── RECAP partagé ────────────────────────────────
  const [recapSaved,   setRecapSaved]   = useState(null);
  const [recapMeta,    setRecapMeta]    = useState(null);
  const [recapStatus,  setRecapStatus]  = useState('init');
  const [superList,    setSuperList]    = useState(null);

  // ── Phase 1 — Objectifs ──────────────────────────
  const [csvFile,      setCsvFile]      = useState(null);
  const [recapFileObj] = useState(null);
  const [recapFileRap] = useState(null);
  const [listFile,     setListFile]     = useState(null);
  const [listData,     setListData]     = useState(null);
  const [listeSaved,   setListeSaved]   = useState(null);
  const [listeStatus,  setListeStatus]  = useState('init');
  const [listeMeta,    setListeMeta]    = useState(null);
  const [selectedProp, setSelectedProp] = useState('');
  const [propResult,   setPropResult]   = useState(null);
  const [selectedSupObj, setSelectedSupObj] = useState('');
  const [supResultObj, setSupResultObj] = useState(null);
  const [mois,         setMois]         = useState('MAI');
  const [annee,        setAnnee]        = useState(2026);
  const [taux,         setTaux]         = useState(15);
  const [supers,       setSupers]       = useState('YAOUNDE,DOUALA,OUEST');
  const [runningObj,   setRunningObj]   = useState(false);
  const [logsObj,      setLogsObj]      = useState([]);
  const [resultObj,    setResultObj]    = useState(null);
  const csvRef   = useRef();
  const listRef  = useRef();
  const listeUpRef = useRef();

  // ── Phase 2 — Réalisation ────────────────────────
  const [reportFile,   setReportFile]   = useState(null);
  const [selectedSupRap, setSelectedSupRap] = useState('');
  const [supResultRap, setSupResultRap] = useState(null);
  const [moisRap,      setMoisRap]      = useState('');
  const [runningRap,   setRunningRap]   = useState(false);
  const [logsRap,      setLogsRap]      = useState([]);
  const [resultsRap,   setResultsRap]   = useState(null);
  const [supersRap,    setSupersRap]    = useState([]);
  const reportRef = useRef();

  const { setObjectifsFiles } = useAppContext();

  // ── Firebase load on mount ───────────────────────
  React.useEffect(() => {
    const run = async () => {
      setRecapStatus('loading');
      const recapResult = await loadRecapSaved();
      if (recapResult) {
        setRecapSaved(recapResult.zoneData);
        setRecapMeta(recapResult.meta);
        buildSuperList(recapResult.zoneData);
        setRecapStatus('loaded');
      } else setRecapStatus('empty');

      setListeStatus('loading');
      const listeResult = await loadListe();
      if (listeResult) {
        setListeSaved(listeResult.propMap);
        setListData(listeResult.propMap);
        setListeMeta(listeResult.meta);
        setListeStatus('loaded');
      } else setListeStatus('empty');
    };
    run();
  }, []); // eslint-disable-line

  // ── Sync moisRap depuis mois quand vide ──────────
  React.useEffect(() => {
    if (!moisRap && mois && annee) setMoisRap(`${mois.toUpperCase()} ${annee}`);
  }, [mois, annee]); // eslint-disable-line

  const logObj = (msg) => setLogsObj(p => [...p, msg]);
  const logRap = (msg) => setLogsRap(p => [...p, msg]);

  // ── Firebase RECAP helpers ───────────────────────
  const saveRecapToFirebase = async (zoneData, filename) => {
    setRecapStatus('saving');
    try {
      const meta = await saveRecap(zoneData, filename);
      setRecapSaved(zoneData);
      setRecapMeta({ ...meta, source: 'Firebase' });
      setRecapStatus('saved');
      setTimeout(() => setRecapStatus('loaded'), 2000);
    } catch(e) { setRecapStatus('error'); }
  };

  const deleteRecapFromFirebase = async () => {
    if (!window.confirm('Supprimer le RECAP sauvegardé ?')) return;
    await deleteRecap();
    setRecapSaved(null); setRecapMeta(null); setSuperList(null); setRecapStatus('empty');
  };

  const saveListeToFirebase = async (propMap, filename) => {
    setListeStatus('saving');
    try {
      const meta = await saveListe(propMap, filename);
      setListeSaved(propMap); setListeMeta(meta); setListeStatus('saved');
      setTimeout(() => setListeStatus('loaded'), 2000);
    } catch(e) { setListeStatus('error'); }
  };

  const deleteListeFromStorage = async () => {
    if (!window.confirm('Supprimer la liste sauvegardée ?')) return;
    await deleteListe();
    setListeSaved(null); setListData(null); setListeMeta(null); setListeStatus('empty');
  };

  // ── Normalisation ────────────────────────────────
  const normalize = (text) => String(text).replace(/\s+/g,' ').trim().toLowerCase();

  // ════════════════════════════════════════════════
  // LOADERS COMMUNS
  // ════════════════════════════════════════════════

  const loadRecapFile = (file) => new Promise((res, rej) => {
    const reader = new FileReader();
    reader.onload = (e) => {
      try {
        const wb = XLSX.read(e.target.result, { type:'array' });
        const zoneData = {};
        wb.SheetNames.forEach(sheet => {
          const ws   = wb.Sheets[sheet];
          const rows = XLSX.utils.sheet_to_json(ws, { header:1, defval:'' });
          const data = [];
          // Trouver la ligne d'en-tête dynamiquement
          let dataStart = 2;
          for (let i = 0; i < Math.min(rows.length, 6); i++) {
            const vals = rows[i].map(v => String(v||'').trim().toUpperCase());
            if (vals.includes('ZONES') || vals.includes('ZONE')) { dataStart = i + 1; break; }
          }
          for (let i = dataStart; i < rows.length; i++) {
            const [zone, sup, sd, st] = rows[i];
            const z = String(zone||'').trim(), s = String(sup||'').trim();
            const d = String(sd||'').trim(),   t = String(st||'').trim();
            if (z && s && d && t && s !== 'SUPER EN CHARGE' && z !== 'ZONES' && z !== 'ZONE') {
              data.push([z, s, d, t]);
            }
          }
          zoneData[sheet] = data;
        });
        res(zoneData);
      } catch(err) { rej(err); }
    };
    reader.onerror = rej;
    reader.readAsArrayBuffer(file);
  });

  // loadRecap pour RapportsPage (format objet {ZONES,SUPER,SALLES,NOMS_TECH})
  const loadRecapRapports = (file) => new Promise((res, rej) => {
    const reader = new FileReader();
    reader.onload = e => {
      try {
        const wb = XLSX.read(e.target.result, { type: 'array' });
        const result = {};
        wb.SheetNames.forEach(sheet => {
          const ws = wb.Sheets[sheet];
          const raw = XLSX.utils.sheet_to_json(ws, { header: 1, defval: '' });
          let hr = null;
          for (let i = 0; i < raw.length; i++) {
            const vs = raw[i].map(v => String(v).trim().toUpperCase());
            if (vs.includes('ZONES') && vs.includes('SUPER EN CHARGE')) { hr = i; break; }
          }
          if (hr === null) return;
          const headers = raw[hr].map(v => String(v).trim().toUpperCase());
          const ci = {
            ZONES:  headers.findIndex(h => h === 'ZONES' || h === 'ZONE'),
            SUPER:  headers.findIndex(h => h.includes('SUPER EN CHARGE')),
            SALLES: headers.findIndex(h => h === 'SALLES' || h === 'SALLE'),
            TECH:   headers.findIndex(h => h.includes('NOMS TECHNIQUE') || h.includes('NOM TECHNIQUE')),
          };
          if (Object.values(ci).some(v => v < 0)) return;
          const rows = []; let lastZone = '', lastSuper = '';
          for (let i = hr + 1; i < raw.length; i++) {
            const row = raw[i];
            const zone  = String(row[ci.ZONES]  || '').trim() || lastZone;
            const sup   = String(row[ci.SUPER]  || '').trim().replace(/\s+/g,' ') || lastSuper;
            const salle = String(row[ci.SALLES] || '').trim();
            const tech  = String(row[ci.TECH]   || '').trim();
            if (!salle && !tech) continue;
            if (zone.toUpperCase() === 'ZONES' || zone.toUpperCase() === 'NAN') continue;
            if (zone) lastZone = zone;
            if (sup)  lastSuper = sup;
            if (salle) {
              let k = tech.toUpperCase().replace(/ /g,'');
              if (k === 'ZOETELECAMTEL') k = 'ZOATELECAMTEL';
              rows.push({ ZONES: lastZone, SUPER: lastSuper, SALLES: salle, NOMS_TECH: tech, _k: k });
            }
          }
          result[sheet] = rows;
        });
        res(result);
      } catch(err) { rej(err); }
    };
    reader.onerror = rej;
    reader.readAsArrayBuffer(file);
  });

  const loadCSV = (file) => new Promise((res, rej) => {
    Papa.parse(file, {
      header:true, skipEmptyLines:true,
      complete: ({ data }) => {
        const lookup = {};
        data.forEach(row => {
          const name  = row['Betshop name'] || '';
          const parts = name.split('.');
          const salle = parts.length >= 3 ? parts[parts.length-1].trim() : name.trim();
          lookup[normalize(salle)] = row;
        });
        res(lookup);
      },
      error: rej
    });
  });

  const loadListFile = (file) => new Promise((res, rej) => {
    const reader = new FileReader();
    reader.onload = (e) => {
      try {
        const wb   = XLSX.read(e.target.result, { type:'array' });
        const ws   = wb.Sheets[wb.SheetNames[0]];
        const rows = XLSX.utils.sheet_to_json(ws, { header:1, defval:'' });
        const propMap = {};
        const cityMap = {}; // shop → city
        for (let i = 1; i < rows.length; i++) {
          const shop = String(rows[i][2] || '').trim();
          const city = String(rows[i][3] || '').trim(); // col D = City
          const prop = String(rows[i][5] || '').trim();
          if (shop && prop) {
            if (!propMap[prop]) propMap[prop] = [];
            propMap[prop].push(shop);
          }
          if (shop && city) cityMap[shop] = city;
        }
        setListData(propMap);
        res({ propMap, cityMap });
      } catch(err) { rej(err); }
    };
    reader.onerror = rej;
    reader.readAsArrayBuffer(file);
  });

  const loadBetshopReport = (file) => new Promise((res, rej) => {
    const reader = new FileReader();
    reader.onload = e => {
      try {
        const wb   = XLSX.read(e.target.result, {type:'array'});
        const ws   = wb.Sheets['Sheet1'] || wb.Sheets[wb.SheetNames[0]];
        const raw  = XLSX.utils.sheet_to_json(ws, {header:1, defval:''});
        const hdrs = raw[0].map(v => String(v).trim());
        const si   = hdrs.findIndex(h => h === 'Shop');
        const t    = hdrs.findIndex(h => h === 'Bet tickets MTD');
        const tl   = hdrs.findIndex(h => h === 'Bet tickets LM');
        const ty   = hdrs.findIndex(h => h === 'Bet tickets LY');
        const t15l = hdrs.findIndex(h => h === 'Bet Ticket LM +15%');
        const t15y = hdrs.findIndex(h => h.includes('LY +15'));
        const p    = hdrs.findIndex(h => h === 'Payin Bet MTD');
        const g    = hdrs.findIndex(h => h === 'GGR Bet MTD');
        const gl   = hdrs.findIndex(h => h === 'GGR Bet LM');
        const gy   = hdrs.findIndex(h => h === 'GGR Bet LY');
        if (si < 0) throw new Error('Colonne Shop introuvable');
        const rows = [];
        for (let i = 1; i < raw.length; i++) {
          const r = raw[i];
          const shop = String(r[si]||'').trim();
          if (!shop) continue;
          rows.push({
            _k: shop.toUpperCase().replace(/ /g,''),
            Shop: shop,
            Tickets_MTD: Number(r[t])||0,
            Payin_MTD:   Number(r[p])||0,
            GGR_MTD:     Number(r[g])||0,
            'Bet tickets MTD':    Number(r[t])||0,
            'Bet tickets LM':     Number(r[tl])||0,
            'Bet tickets LY':     Number(r[ty])||0,
            'Bet Ticket LM +15%': Number(r[t15l])||0,
            'Bet ticket LY +15 %':Number(r[t15y])||0,
            'GGR Bet MTD': Number(r[g])||0,
            'GGR Bet LM':  Number(r[gl])||0,
            'GGR Bet LY':  Number(r[gy])||0,
          });
        }
        res(rows);
      } catch(err) { rej(err); }
    };
    reader.onerror = rej;
    reader.readAsArrayBuffer(file);
  });

  // ── buildSuperList : accepte tableaux ET objets ──
  const buildSuperList = (recapData) => {
    const superMap = {};
    Object.entries(recapData).forEach(([ville, rows]) => {
      rows.forEach(r => {
        let key;
        if (Array.isArray(r)) {
          const [z, sup, sd, st] = r;
          key = (sup||'').trim().replace(/\s+/g,' ');
          if (!key) return;
          if (!superMap[key]) superMap[key] = { ville, rows: [] };
          superMap[key].rows.push({ ZONES:z, SUPER:key, SALLES:sd, NOMS_TECH:st,
            _k: (st||'').toUpperCase().replace(/ /g,'') });
        } else {
          key = (r.SUPER||r.sup||'').trim().replace(/\s+/g,' ');
          if (!key) return;
          const rn = { ZONES:r.ZONES||r.ZONE||r.z||'', SUPER:key,
            SALLES:r.SALLES||r.SALLE||r.sd||'', NOMS_TECH:r.NOMS_TECH||r.NOM_TECH||r.st||'',
            _k:r._k||(r.NOMS_TECH||r.NOM_TECH||r.st||'').toUpperCase().replace(/ /g,'') };
          if (!superMap[key]) superMap[key] = { ville, rows: [] };
          superMap[key].rows.push(rn);
        }
      });
    });
    setSuperList(superMap);
    return superMap;
  };

  // ════════════════════════════════════════════════
  // PHASE 1 — STYLES & HELPERS OBJECTIFS
  // ════════════════════════════════════════════════
  const SO = {
    HEADER: { patternType:'solid', fgColor:{ rgb:'1F3864' } },
    GREEN:  { patternType:'solid', fgColor:{ rgb:'375623' } },
    RED:    { patternType:'solid', fgColor:{ rgb:'7B2C2C' } },
    BLUE:   { patternType:'solid', fgColor:{ rgb:'1F4E79' } },
    ALT:    { patternType:'solid', fgColor:{ rgb:'EBF3FB' } },
    WHITE:  { patternType:'solid', fgColor:{ rgb:'FFFFFF' } },
    TOTAL:  { patternType:'solid', fgColor:{ rgb:'D9D9D9' } },
    ZONE:   { patternType:'solid', fgColor:{ rgb:'2E75B6' } },
  };
  const bdrO = () => { const t={style:'thin',color:{rgb:'BFBFBF'}}; return {top:t,bottom:t,left:t,right:t}; };
  const fntO = (bold,sz,rgb) => ({name:'Calibri',bold:!!bold,sz:sz||10,color:{rgb:rgb||'000000'}});
  const alnO = (h,wrap) => ({horizontal:h||'left',vertical:'center',wrapText:!!wrap});
  const scO  = (ws,r,c,v,style) => { const a=XLSX.utils.encode_cell({r:r-1,c:c-1}); ws[a]={v:v==null?'':v,t:typeof v==='number'?'n':'s',s:style}; };
  const scnO = (ws,r,c,v,style,fmt) => { const a=XLSX.utils.encode_cell({r:r-1,c:c-1}); ws[a]={v:typeof v==='number'?v:0,t:'n',s:style,z:fmt}; };
  const mgO  = (ws,r1,c1,r2,c2) => { if(!ws['!merges'])ws['!merges']=[]; ws['!merges'].push({s:{r:r1-1,c:c1-1},e:{r:r2-1,c:c2-1}}); };

  const buildSheet = (rowsData, csvLookup, titleText, moisUp, anneeN, hasZone=true, hasSuper=true, tauxVal=15) => {
    const ws = {};
    const anneePrec = anneeN - 1;
    const colDefs = [];
    if (hasZone)  colDefs.push({ h:'ZONE', fill:SO.HEADER, w:14 });
    if (hasSuper) colDefs.push({ h:'SUPER EN CHARGE', fill:SO.HEADER, w:22 });
    colDefs.push(
      { h:'SALLE', fill:SO.HEADER, w:28 },
      { h:'NOM TECHNIQUE', fill:SO.HEADER, w:18 },
      { h:`Payin ${moisUp} ${anneePrec} (FCFA)`, fill:SO.GREEN, w:22 },
      { h:`Payin Objectif (${moisUp} ${anneeN}) +${tauxVal}%`, fill:SO.HEADER, w:26 },
      { h:`GGR ${moisUp} ${anneePrec} (FCFA)`, fill:SO.RED, w:20 },
      { h:`Tickets ${moisUp} ${anneePrec}`, fill:SO.BLUE, w:16 },
      { h:`Tickets Objectif (${moisUp} ${anneeN}) +${tauxVal}%`, fill:SO.HEADER, w:24 },
    );
    const ncols = colDefs.length;
    const pCol  = (hasZone?1:0) + (hasSuper?1:0) + 3;
    scO(ws,2,1,titleText,{font:fntO(true,11,'FFFFFF'),fill:SO.HEADER,alignment:alnO('left'),border:bdrO()});
    mgO(ws,2,1,2,ncols);
    ws['!rows'] = [{},{hpt:20},{},{hpt:42}];
    colDefs.forEach((cd,i) => {
      scO(ws,4,i+1,cd.h,{font:fntO(true,10,'FFFFFF'),fill:cd.fill,alignment:alnO('center',true),border:bdrO()});
    });
    let dataRow=5, prevGroup=null;
    const totals=[0,0,0,0,0], missing=[];
    let rowIdx=0;
    rowsData.forEach(([zone,sup,sd,st]) => {
      const groupKey = hasSuper ? `${zone}|${sup}` : zone;
      if (groupKey !== prevGroup) {
        const label = hasSuper&&hasZone?`── ${zone}  ·  ${sup}`:hasZone?`── ${zone}`:`── ${sup}`;
        scO(ws,dataRow,1,label,{font:fntO(true,9,'FFFFFF'),fill:SO.ZONE,border:bdrO()});
        for(let c=2;c<=ncols;c++) scO(ws,dataRow,c,'',{font:fntO(false,9,'FFFFFF'),fill:SO.ZONE,border:bdrO()});
        if(!ws['!rows'][dataRow-1]) ws['!rows'][dataRow-1]={};
        ws['!rows'][dataRow-1]={hpt:14};
        dataRow++; prevGroup=groupKey;
      }
      const rowFill = rowIdx%2===0 ? SO.ALT : SO.WHITE;
      let col=1;
      if(hasZone)  { scO(ws,dataRow,col,zone,{font:fntO(false,9),fill:rowFill,border:bdrO(),alignment:alnO('left')}); col++; }
      if(hasSuper) { scO(ws,dataRow,col,sup, {font:fntO(false,9),fill:rowFill,border:bdrO(),alignment:alnO('left')}); col++; }
      scO(ws,dataRow,col,sd,{font:fntO(false,9),fill:rowFill,border:bdrO(),alignment:alnO('left')}); col++;
      scO(ws,dataRow,col,st,{font:fntO(false,9),fill:rowFill,border:bdrO(),alignment:alnO('left')}); col++;
      const csvRow = csvLookup[normalize(st)];
      if (csvRow) {
        const payin=parseInt(csvRow['Payin money'])||0, tickets=parseInt(csvRow['Payin count'])||0;
        const gross=parseFloat(csvRow['Gross'])||0;
        const payinObj=Math.round(payin*(1+tauxVal/100)), ticketsObj=Math.round(tickets*(1+tauxVal/100));
        totals[0]+=payin;totals[1]+=payinObj;totals[2]+=gross;totals[3]+=tickets;totals[4]+=ticketsObj;
        scnO(ws,dataRow,col,payin,     {font:fntO(false,9),fill:rowFill,border:bdrO(),alignment:alnO('right')},'#,##0'); col++;
        scnO(ws,dataRow,col,payinObj,  {font:fntO(true,9), fill:SO.ALT, border:bdrO(),alignment:alnO('right')},'#,##0'); col++;
        scnO(ws,dataRow,col,gross,     {font:fntO(false,9),fill:rowFill,border:bdrO(),alignment:alnO('right')},'#,##0.00'); col++;
        scnO(ws,dataRow,col,tickets,   {font:fntO(false,9),fill:rowFill,border:bdrO(),alignment:alnO('right')},'#,##0'); col++;
        scnO(ws,dataRow,col,ticketsObj,{font:fntO(true,9), fill:SO.ALT, border:bdrO(),alignment:alnO('right')},'#,##0');
      } else {
        missing.push(st);
        for(let c=col;c<=ncols;c++) scO(ws,dataRow,c,'',{font:fntO(false,9),fill:rowFill,border:bdrO()});
      }
      dataRow++;rowIdx++;
    });
    const totalR=dataRow;
    scO(ws,totalR,1,`TOTAL  (${rowsData.length} salles)`,{font:fntO(true,10),fill:SO.TOTAL,border:bdrO(),alignment:alnO('left')});
    if(pCol>1) mgO(ws,totalR,1,totalR,pCol-1);
    for(let c=2;c<pCol;c++) scO(ws,totalR,c,'',{font:fntO(true,10),fill:SO.TOTAL,border:bdrO()});
    [[pCol,totals[0],'#,##0'],[pCol+1,totals[1],'#,##0'],[pCol+2,Math.round(totals[2]*100)/100,'#,##0.00'],[pCol+3,totals[3],'#,##0'],[pCol+4,totals[4],'#,##0']].forEach(([c,v,fmt]) => {
      scnO(ws,totalR,c,v,{font:fntO(true,10),fill:SO.TOTAL,border:bdrO(),alignment:alnO('right')},fmt);
    });
    ws['!cols'] = colDefs.map(cd=>({wch:cd.w}));
    ws['!ref']  = XLSX.utils.encode_range({s:{r:0,c:0},e:{r:totalR-1,c:ncols-1}});
    return {ws,missing};
  };

  // ════════════════════════════════════════════════
  // PHASE 1 — PIPELINES OBJECTIFS
  // ════════════════════════════════════════════════
  const runPipelineObj = async () => {
    setRunningObj(true); setLogsObj([]); setResultObj(null);
    try {
      const moisUp=mois.toUpperCase().trim(), anneeN=parseInt(annee), tauxVal=parseInt(taux);
      const prefix=`${moisUp}${anneeN}`;
      const zonesSup=supers.split(',').map(s=>s.trim().toUpperCase()).filter(Boolean);
      logObj('📂 Chargement des données...');
      let zoneData;
      if (recapFileObj) {
        zoneData = await loadRecapFile(recapFileObj);
        logObj(`  ✓ RECAP : ${Object.keys(zoneData).length} zones`);
      } else if (recapSaved) {
        zoneData = recapSaved;
        logObj('  ✓ RECAP chargé depuis Firebase');
      } else throw new Error('Veuillez uploader le RECAP des salles par zone');
      const csvLookup = await loadCSV(csvFile);
      logObj(`  ✓ CSV : ${Object.keys(csvLookup).length} salles`);
      const allZones=Object.keys(zoneData), allRows=allZones.flatMap(z=>zoneData[z]);
      logObj(`\n📋 ${allZones.length} zones · ${allRows.length} salles`);
      buildSuperList(zoneData);
      const zip=new JSZip(), allMissing=[];

      logObj('\n🌍 Génération GLOBAL...');
      const {ws:wsG,missing:mG} = buildSheet(allRows,csvLookup,
        `OBJECTIFS CAMPRESJ — ${moisUp} ${anneeN}  ·  GLOBAL  ·  ${allRows.length} salles`,
        moisUp,anneeN,true,true,tauxVal);
      allMissing.push(...mG);
      const wbG=XLSX.utils.book_new();
      XLSX.utils.book_append_sheet(wbG,wsG,`GLOBAL ${moisUp} ${anneeN}`.slice(0,31));
      zip.file(`${prefix}_GLOBAL.xlsx`,XLSX.write(wbG,{bookType:'xlsx',type:'array'}));
      logObj(`  ✅ ${prefix}_GLOBAL.xlsx`);

      logObj('\n🗺️  Génération par zone...');
      for (const zoneName of allZones) {
        const rows=zoneData[zoneName];
        const {ws,missing}=buildSheet(rows,csvLookup,
          `OBJECTIFS CAMPRESJ — ${moisUp} ${anneeN}  ·  ${zoneName}  ·  ${rows.length} salles`,
          moisUp,anneeN,true,true,tauxVal);
        allMissing.push(...missing);
        const wb=XLSX.utils.book_new();
        XLSX.utils.book_append_sheet(wb,ws,zoneName.slice(0,31));
        zip.file(`${prefix}_ZONE_${zoneName}.xlsx`,XLSX.write(wb,{bookType:'xlsx',type:'array'}));
        logObj(`  ✅ ${prefix}_ZONE_${zoneName}.xlsx`);
      }

      logObj(`\n👤 Génération par super (${zonesSup.join(',')})...`);
      for (const zoneName of zonesSup) {
        if (!zoneData[zoneName]) { logObj(`  ⚠️ Zone '${zoneName}' introuvable`); continue; }
        const superMap={};
        zoneData[zoneName].forEach(([z,sup,sd,st]) => {
          if(!superMap[sup]) superMap[sup]=[];
          superMap[sup].push([z,sup,sd,st]);
        });
        for (const [sup,rows] of Object.entries(superMap)) {
          const {ws,missing}=buildSheet(rows,csvLookup,
            `OBJECTIFS ${moisUp} ${anneeN}  ·  ${zoneName}  ·  ${sup}  ·  ${rows.length} salles`,
            moisUp,anneeN,true,false,tauxVal);
          allMissing.push(...missing);
          const wb=XLSX.utils.book_new();
          XLSX.utils.book_append_sheet(wb,ws,sup.slice(0,31));
          const safe=sup.replace(/ /g,'_').replace(/\//g,'-');
          zip.file(`${prefix}_${zoneName}_${safe}.xlsx`,XLSX.write(wb,{bookType:'xlsx',type:'array'}));
          logObj(`  ✅ ${prefix}_${zoneName}_${safe}.xlsx`);
        }
      }

      const blob=await zip.generateAsync({type:'blob'});
      const uniqueMissing=[...new Set(allMissing)];
      setResultObj({blob,filename:`Objectifs_${moisUp}${anneeN}.zip`,missing:uniqueMissing});
      setObjectifsFiles(prev=>[...prev.filter(f=>f.mois!==`${moisUp} ${anneeN}`),
        {blob,filename:`Objectifs_${moisUp}${anneeN}.zip`,mois:`${moisUp} ${anneeN}`,type:'global'}]);
      // Pré-remplir le mois de Réalisation
      setMoisRap(`${moisUp} ${anneeN}`);
      logObj(`\n✅ Terminé !${uniqueMissing.length?` ⚠️ ${uniqueMissing.length} salle(s) sans données CSV`:''}`);
    } catch(err) { logObj(`❌ Erreur : ${err.message}`); }
    setRunningObj(false);
  };

  const runSuperPipelineObj = async () => {
    if (!selectedSupObj || !superList || !csvFile) return;
    setSupResultObj(null);
    try {
      const moisUp=mois.toUpperCase().trim(), anneeN=parseInt(annee), tauxVal=parseInt(taux);
      const rows=superList[selectedSupObj].rows.map(r=>[r.ZONES,r.SUPER,r.SALLES,r.NOMS_TECH]);
      logObj(`\n👤 Objectifs Super : ${selectedSupObj}...`);
      const csvLookup=await loadCSV(csvFile);
      const {ws,missing}=buildSheet(rows,csvLookup,
        `OBJECTIFS ${moisUp} ${anneeN}  ·  ${selectedSupObj}  ·  ${rows.length} salles`,
        moisUp,anneeN,true,false,tauxVal);
      const wb=XLSX.utils.book_new();
      XLSX.utils.book_append_sheet(wb,ws,selectedSupObj.slice(0,31));
      const filename=`${moisUp}${anneeN}_SUPER_${selectedSupObj.replace(/ /g,'_').replace(/[/]/g,'-')}.xlsx`;
      const blob=new Blob([XLSX.write(wb,{bookType:'xlsx',type:'array'})],
        {type:'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'});
      setSupResultObj({blob,filename,count:rows.length,missing});
      logObj(`  ✅ ${filename}`);
    } catch(err) { logObj(`❌ Erreur : ${err.message}`); }
  };

  const runPropPipeline = async () => {
    if (!selectedProp || !listData || !csvFile) return;
    setPropResult(null);
    try {
      const moisUp=mois.toUpperCase().trim(), anneeN=parseInt(annee), tauxVal=parseInt(taux);
      const propShops=new Set(listData[selectedProp].map(s=>normalize(s)));
      logObj(`\n🏪 Objectifs ${selectedProp}...`);
      const csvLookup=await loadCSV(csvFile);
      let propRows=[];
      if (recapFileObj || recapSaved) {
        const zd=recapFileObj?await loadRecapFile(recapFileObj):recapSaved;
        propRows=Object.values(zd).flat().filter(([z,sup,sd,st])=>
          propShops.has(normalize(st))||propShops.has(normalize(sd)));
        logObj(`  ${propRows.length} salles dans le RECAP`);
      } else {
        listData[selectedProp].forEach(s=>propRows.push(['—','—',s,s]));
      }
      if (!propRows.length) { logObj('  ⚠️ Aucune salle trouvée'); return; }
      const {ws,missing}=buildSheet(propRows,csvLookup,
        `OBJECTIFS ${selectedProp} — ${moisUp} ${anneeN}  ·  ${propRows.length} salles`,
        moisUp,anneeN,!!recapFileObj||!!recapSaved,!!recapFileObj||!!recapSaved,tauxVal);
      const wb=XLSX.utils.book_new();
      XLSX.utils.book_append_sheet(wb,ws,selectedProp.slice(0,31));
      const filename=`${moisUp}${anneeN}_${selectedProp.replace(/ /g,'_')}.xlsx`;
      const blob=new Blob([XLSX.write(wb,{bookType:'xlsx',type:'array'})],
        {type:'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'});
      setPropResult({blob,filename,count:propRows.length,missing});
      logObj(`  ✅ ${filename}`);
    } catch(err) { logObj(`❌ Erreur : ${err.message}`); }
  };

  // ════════════════════════════════════════════════
  // PHASE 2 — STYLES & HELPERS RÉALISATION
  // ════════════════════════════════════════════════
  const C = {
    DARK:'1F3864',BLUE:'1F4E79',MID:'2E5F8A',GREY:'404040',
    GREEN:'375623',RED:'7B2C2C',GOLD:'BF8F00',TTOTAL:'2E4B7A',
    OK:'E2EFDA',OKF:'375623',KO:'FCE4D6',KOF:'C55A11',
    OK15:'FFF2CC',OK15F:'7F6000',ALT:'EBF3FB',KPI:'F2F2F2',
  };
  const fillR  = (rgb)    => ({patternType:'solid',fgColor:{rgb}});
  const bdrR   = ()       => {const t={style:'thin',color:{rgb:'BFBFBF'}};return{top:t,bottom:t,left:t,right:t};};
  const fntR   = (bold=false,sz=10,rgb='000000') => ({name:'Arial',bold,sz,color:{rgb}});
  const alnR   = (h='center',v='center',wrap=false) => ({horizontal:h,vertical:v,wrapText:wrap});
  const scR    = (ws,r,c,v,bold=false,sz=10,fc='000000',bg=null,h='left') => {
    const addr=XLSX.utils.encode_cell({r:r-1,c:c-1});
    ws[addr]={v,t:typeof v==='number'?'n':'s',s:{font:fntR(bold,sz,fc),fill:bg?fillR(bg):undefined,alignment:alnR(h),border:bdrR()}};
  };

  const buildMaster = (recap, betRows) => {
    const betMap={}; betRows.forEach(r=>{betMap[r._k]=r;});
    const out={};
    Object.entries(recap).forEach(([ville,rows]) => {
      out[ville]=rows.map(r=>{
        const rn=Array.isArray(r)
          ?{ZONES:r[0]||'',SUPER:r[1]||'',SALLES:r[2]||'',NOMS_TECH:r[3]||'',_k:(r[3]||r[2]||'').toUpperCase().replace(/ /g,'')}
          :{ZONES:r.ZONES||r.ZONE||r.z||'',SUPER:r.SUPER||r.sup||'',SALLES:r.SALLES||r.SALLE||r.sd||'',NOMS_TECH:r.NOMS_TECH||r.NOM_TECH||r.st||'',_k:r._k||(r.NOMS_TECH||r.NOM_TECH||r.st||'').toUpperCase().replace(/ /g,'')};
        const b=betMap[rn._k]||{};
        const mtd=b.Tickets_MTD||0;
        const obj=b['Bet ticket LY +15 %']||b['Bet Ticket LY +15%']||b['Bet Ticket LM +15%']||0;
        const diff=mtd-obj, pct=obj>0?(diff/obj)*100:null;
        let statut='❌ Non atteint';
        if(pct!==null&&pct>=15) statut='🏆 Obj+15%';
        else if(pct!==null&&pct>=0) statut='✅ Atteint';
        return {...rn,Tickets_MTD:mtd,Payin_MTD:b.Payin_MTD||0,GGR_MTD:b.GGR_MTD||0,Tickets_Obj:obj,Diff:diff,Pct:pct,Statut:statut};
      });
    });
    return out;
  };

  const makeSupers = (dfV) => {
    const supMap={};
    dfV.forEach(r=>{
      const supKey=(r.SUPER||'').trim().replace(/\s+/g,' ');
      if(!supMap[supKey]) supMap[supKey]=[];
      supMap[supKey].push({...r,SUPER:supKey});
    });
    return Object.entries(supMap).map(([sup,rows])=>{
      const zones=[...new Set(rows.map(r=>r.ZONES))];
      const att=rows.filter(r=>r.Statut!=='❌ Non atteint').length;
      const natt=rows.filter(r=>r.Statut==='❌ Non atteint').length;
      return {
        zone:zones.join(' / '),super:sup,nb:rows.length,att,natt,
        mtd:rows.reduce((s,r)=>s+r.Tickets_MTD,0),
        obj:rows.reduce((s,r)=>s+r.Tickets_Obj,0),
        ggr:rows.reduce((s,r)=>s+r.GGR_MTD,0),
        rows:rows.map(r=>({zone:r.ZONES,salle:r.SALLES,tech:r.NOMS_TECH,mtd:r.Tickets_MTD,obj:r.Tickets_Obj,diff:r.Diff,pct:r.Pct,ggr:r.GGR_MTD,statut:r.Statut}))
      };
    }).sort((a,b)=>{
      const pa=a.obj>0?(a.mtd-a.obj)/a.obj*100:0;
      const pb=b.obj>0?(b.mtd-b.obj)/b.obj*100:0;
      return pb-pa;
    });
  };

  const writeClassement = (supers,moisLabel) => {
    const ws={},merges=[];let R=1;
    scR(ws,R,1,`CLASSEMENT SUPERS — ${moisLabel}`,true,13,'FFFFFF',C.DARK,'center');
    merges.push({s:{r:R-1,c:0},e:{r:R-1,c:7}});R++;
    ['Rang','Super','Zone','Salles','Atteints','Non att.','Tickets MTD','% vs Obj'].forEach((h,i)=>scR(ws,R,i+1,h,true,10,'FFFFFF',C.BLUE,'center'));
    R++;
    supers.forEach((sd,idx)=>{
      const pct=sd.obj>0?(sd.mtd-sd.obj)/sd.obj*100:0;
      const bg=pct>=15?C.OK15:(pct>=0?C.OK:C.KO);
      const fc=pct>=15?C.OK15F:(pct>=0?C.OKF:C.KOF);
      const medal=['🥇','🥈','🥉'][idx]||'';
      [[1,`${medal} ${idx+1}`,true,10,'000000',bg,'center'],[2,sd.super,false,10,'000000',bg,'left'],
       [3,sd.zone,false,9,'555555',bg,'left'],[4,sd.nb,false,10,'000000',bg,'center'],
       [5,sd.att,false,10,C.OKF,bg,'center'],[6,sd.natt,false,10,C.KOF,bg,'center'],
       [7,sd.mtd,true,10,'000000',bg,'right'],[8,`${pct>=0?'+':''}${pct.toFixed(1)}%`,true,10,fc,bg,'center'],
      ].forEach(([c,v,b,s,fc2,bg2,h])=>scR(ws,R,c,v,b,s,fc2,bg2,h));
      R++;
    });
    ws['!cols']=[6,28,20,8,8,8,14,12].map(w=>({wch:w}));
    ws['!merges']=merges;
    ws['!ref']=XLSX.utils.encode_range({s:{r:0,c:0},e:{r:R-1,c:7}});
    return ws;
  };

  const writeDetail = (master,moisLabel) => {
    const ws={};let R=1;
    scR(ws,R,1,`DÉTAIL PAR ZONE & SUPER — ${moisLabel}`,true,12,'FFFFFF',C.DARK,'center');
    ws['!merges']=[{s:{r:0,c:0},e:{r:0,c:7}}];R++;
    ['Zone','Super','Salle','Nom technique','Tickets MTD','Objectif','Diff','% vs Obj'].forEach((h,i)=>scR(ws,R,i+1,h,true,10,'FFFFFF',C.BLUE,'center'));
    R++;
    Object.entries(master).forEach(([ville,rows])=>{
      scR(ws,R,1,`── ${ville}`,true,10,'FFFFFF',C.MID,'left');
      for(let c=2;c<=8;c++) scR(ws,R,c,'',false,10,'FFFFFF',C.MID,'left');R++;
      rows.forEach(r=>{
        const bg=r.Pct===null?C.KPI:(r.Pct>=15?C.OK15:(r.Pct>=0?C.OK:C.KO));
        const dfc=(r.Diff||0)>=0?C.OKF:C.KOF;
        const pfc=(r.Pct===null||r.Pct<0)?C.KOF:C.OKF;
        scR(ws,R,1,r.ZONES,false,9,'555555',bg,'left');
        scR(ws,R,2,r.SUPER||'',false,9,'000000',bg,'left');
        scR(ws,R,3,r.SALLES||'',false,9,'000000',bg,'left');
        scR(ws,R,4,r.NOMS_TECH,false,9,'000000',bg,'left');
        scR(ws,R,5,r.Tickets_MTD,true,10,'000000',bg,'right');
        scR(ws,R,6,r.Tickets_Obj,false,10,'000000',bg,'right');
        scR(ws,R,7,r.Diff,true,10,dfc,bg,'right');
        scR(ws,R,8,r.Pct!==null?parseFloat(r.Pct.toFixed(2)):'',true,10,pfc,bg,'right');
        if(ws[XLSX.utils.encode_cell({r:R-1,c:7})]) ws[XLSX.utils.encode_cell({r:R-1,c:7})].z='0.00"%"';
        R++;
      });
    });
    ws['!cols']=[14,22,26,24,14,14,10,12].map(w=>({wch:w}));
    ws['!ref']=XLSX.utils.encode_range({s:{r:0,c:0},e:{r:R-1,c:7}});
    return ws;
  };

  const writeSuperSheet = (sd,moisLabel,rang,total) => {
    const ws={},merges=[];let R=1;
    const pct=sd.obj>0?(sd.mtd-sd.obj)/sd.obj*100:0;
    scR(ws,R,1,`SUPER : ${sd.super}  ·  ${sd.zone}  ·  ${moisLabel}  —  Rang ${rang}/${total}`,true,13,'FFFFFF',C.BLUE,'center');
    merges.push({s:{r:0,c:0},e:{r:0,c:7}});R++;
    const kpis=[
      ['Salles en charge',`${sd.nb} salles`,'Objectifs atteints',`${sd.att} / ${sd.nb}`],
      ['Non atteints',`${sd.natt} / ${sd.nb}`,'Tickets réalisés (MTD)',sd.mtd.toLocaleString('fr-FR')],
      ['Tickets objectif',sd.obj.toLocaleString('fr-FR'),'Réalisation vs objectif',`${pct>=0?'▲ ':'▼ '}${pct>0?'+':''}${pct.toFixed(1)}%`],
    ];
    kpis.forEach(([la,lv,ra,rv])=>{
      for(let c=1;c<=8;c++) scR(ws,R,c,'',false,10,'000000',C.KPI,'left');
      scR(ws,R,1,la,true,10);merges.push({s:{r:R-1,c:0},e:{r:R-1,c:1}});
      scR(ws,R,3,lv,true,11,'000000',C.KPI,'center');merges.push({s:{r:R-1,c:2},e:{r:R-1,c:3}});
      scR(ws,R,5,ra,true,10);merges.push({s:{r:R-1,c:4},e:{r:R-1,c:5}});
      const vbg=rv.includes('%')?(pct>=0?C.OK:C.KO):'DEEAF1';
      const vfc=pct>=0?C.OKF:C.KOF;
      scR(ws,R,7,rv,true,11,vfc,vbg,'center');merges.push({s:{r:R-1,c:6},e:{r:R-1,c:7}});R++;
    });
    ['Salle (terrain)','Nom technique','Tickets MTD','Objectif (+15%)','Diff','% vs Obj','GGR MTD','Statut'].forEach((h,i)=>{
      scR(ws,R,i+1,h,true,9,'FFFFFF',[C.BLUE,C.BLUE,C.BLUE,C.MID,C.MID,C.MID,C.RED,C.GREY][i],'center');
    });R++;
    const sorted=[...sd.rows].sort((a,b)=>{
      const aNo=a.mtd===0&&a.obj===0,bNo=b.mtd===0&&b.obj===0;
      if(aNo!==bNo) return aNo?1:-1;
      return (b.pct||0)-(a.pct||0);
    });
    sorted.forEach(r=>{
      const noData=r.mtd===0&&r.obj===0;
      const bg=noData?C.KPI:(r.statut==='🏆 Obj+15%'?C.OK15:(r.statut==='✅ Atteint'?C.OK:C.KO));
      const dfc=(r.diff||0)>=0?C.OKF:C.KOF;
      const pfc=(!r.pct||r.pct<0)?C.KOF:C.OKF;
      scR(ws,R,1,r.salle,false,10,'000000',bg,'left');
      scR(ws,R,2,r.tech, false,10,'000000',bg,'left');
      if(noData){
        for(let c=3;c<=7;c++){const a=XLSX.utils.encode_cell({r:R-1,c:c-1});ws[a]={v:'—',t:'s',s:{font:fntR(false,10,'888888'),fill:fillR(bg),border:bdrR(),alignment:alnR('center')}};}
        scR(ws,R,8,'Données absentes',false,9,'888888',C.KPI,'center');
      } else {
        scR(ws,R,3,r.mtd,true,10,'000000',bg,'right');
        scR(ws,R,4,r.obj,false,10,'000000',bg,'right');
        scR(ws,R,5,r.diff,true,10,dfc,bg,'right');
        if(r.pct!==null){scR(ws,R,6,parseFloat(r.pct.toFixed(2)),true,10,pfc,bg,'right');if(ws[XLSX.utils.encode_cell({r:R-1,c:5})]) ws[XLSX.utils.encode_cell({r:R-1,c:5})].z='0.00"%"';}
        scR(ws,R,7,Math.round(r.ggr),false,10,'000000',bg,'right');
        scR(ws,R,8,r.statut,true,9,r.statut==='🏆 Obj+15%'?C.OK15F:(r.statut==='✅ Atteint'?C.OKF:C.KOF),bg,'center');
      }R++;
    });
    scR(ws,R,1,`TOTAL — ${sd.super}`,true,10,'FFFFFF',C.BLUE,'left');
    merges.push({s:{r:R-1,c:0},e:{r:R-1,c:1}});
    scR(ws,R,3,sd.mtd,true,10,'FFFFFF',C.BLUE,'right');
    scR(ws,R,4,sd.obj,true,10,'FFFFFF',C.BLUE,'right');
    scR(ws,R,5,sd.mtd-sd.obj,true,10,'FFFFFF',C.BLUE,'right');
    scR(ws,R,6,`${pct>0?'+':''}${pct.toFixed(2)}%`,true,10,'FFFFFF',C.BLUE,'center');
    scR(ws,R,7,Math.round(sd.ggr),true,10,'FFFFFF',C.BLUE,'right');
    ws['!cols']=[26,24,14,16,12,10,14,14].map(w=>({wch:w}));
    ws['!merges']=merges;
    ws['!ref']=XLSX.utils.encode_range({s:{r:0,c:0},e:{r:R-1,c:7}});
    return ws;
  };

  const buildWb = (supers,masterSub,moisLabel) => {
    const wb=XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(wb,writeClassement(supers,moisLabel),'① Classement Supers');
    XLSX.utils.book_append_sheet(wb,writeDetail(masterSub,moisLabel),'② Détail par Zone & Super');
    const medals={1:'🥇',2:'🥈',3:'🥉'};
    const usedNames=new Set(['① Classement Supers','② Détail par Zone & Super']);
    supers.forEach((sd,i)=>{
      const em=medals[i+1]||'  ';
      let name=`${em} ${sd.super.trim()}`.slice(0,31),finalName=name,counter=2;
      while(usedNames.has(finalName)){finalName=`${name.slice(0,28)} ${counter}`.slice(0,31);counter++;}
      usedNames.add(finalName);
      XLSX.utils.book_append_sheet(wb,writeSuperSheet(sd,moisLabel,i+1,supers.length),finalName);
    });
    return wb;
  };

  // ════════════════════════════════════════════════
  // PHASE 2 — PIPELINES RÉALISATION
  // ════════════════════════════════════════════════
  const getRecapRap = async () => {
    if (recapFileRap) return loadRecapRapports(recapFileRap);
    if (recapSaved)   return recapSaved;
    throw new Error('Veuillez uploader le RECAP des salles par zone');
  };

  const runSuperPipelineRap = async () => {
    if (!selectedSupRap || !superList || !reportFile) return;
    setSupResultRap(null);
    try {
      const moisUp=moisRap.toUpperCase();
      logRap(`\n👤 Réalisation Super : ${selectedSupRap}...`);
      const [betRows, recap] = await Promise.all([loadBetshopReport(reportFile), getRecapRap()]);
      const supData=superList[selectedSupRap];
      const master=buildMaster(recap,betRows);
      const vRows=master[supData.ville]||[];
      const supRows=vRows.filter(r=>(r.SUPER||'').trim().replace(/\s+/g,' ')===selectedSupRap);
      if(!supRows.length){logRap(`  ⚠️ Aucune salle pour ${selectedSupRap}`);return;}
      logRap(`  ${supRows.length} salles`);
      const supers=makeSupers(supRows);
      const wb=buildWb(supers,{[supData.ville]:supRows},moisUp);
      const safe=selectedSupRap.replace(/ /g,'_').replace(/[/]/g,'-');
      const filename=`ANALYSE_SUPER_${safe}_${moisUp.replace(/ /g,'_')}.xlsx`;
      const blob=new Blob([XLSX.write(wb,{bookType:'xlsx',type:'array'})],
        {type:'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'});
      setSupResultRap({blob,filename});
      logRap(`  ✅ ${filename}`);
    } catch(err){logRap(`❌ Erreur : ${err.message}`);}
  };

  const runAllSupersPipeline = async () => {
    if (!reportFile || (!recapFileRap && !recapSaved)) return;
    setRunningRap(true); setLogsRap([]); setSupersRap([]);
    try {
      const moisUp = moisRap.toUpperCase();
      logRap('📂 Chargement...');
      const recap  = await getRecapRap();
      const betRows = await loadBetshopReport(reportFile);
      logRap(`  ✓ Betshop : ${betRows.length} salles`);
      const master = buildMaster(recap, betRows);
      const villes = Object.keys(recap);
      const superFiles = [];

      for (const ville of villes) {
        const supMap = {};
        (master[ville]||[]).forEach(r => {
          const sup = (r.SUPER||'').trim().replace(/\s+/g,' ');
          if (!sup) return;
          if (!supMap[sup]) supMap[sup] = [];
          supMap[sup].push(r);
        });
        for (const [sup, rows] of Object.entries(supMap)) {
          const supersLocal = makeSupers(rows);
          const wb  = buildWb(supersLocal, {[ville]: rows}, moisUp);
          const safe = sup.replace(/ /g,'_').replace(/[/]/g,'-');
          const filename = `ANALYSE_SUPER_${safe}_${moisUp.replace(/ /g,'_')}.xlsx`;
          const blob = new Blob(
            [XLSX.write(wb, {bookType:'xlsx', type:'array'})],
            {type:'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'}
          );
          // eslint-disable-next-line no-unused-vars
          superFiles.push({filename, blob, super: sup, ville, nb: rows.length,
            att: supersLocal.reduce((s,x)=>s+x.att,0),
            natt: supersLocal.reduce((s,x)=>s+x.natt,0)});
          logRap(`    ✅ ${filename}`);
        }
      }

      setSupersRap(superFiles);
      logRap(`\n✅ ${superFiles.length} rapports supers générés`);
    } catch(err) { logRap(`❌ Erreur : ${err.message}`); }
    setRunningRap(false);
  };

  const runPipelineRap = async () => {
    setRunningRap(true);setLogsRap([]);setResultsRap(null);
    try {
      const moisUp=moisRap.toUpperCase();
      logRap('📂 Chargement...');
      const recap=await getRecapRap();
      if(recapSaved&&!recapFileRap) logRap('  ✓ RECAP depuis Firebase');
      const betRows=await loadBetshopReport(reportFile);
      logRap(`  ✓ Betshop : ${betRows.length} salles`);
      buildSuperList(recap);
      const villes=Object.keys(recap);
      logRap(`  Villes : ${villes.join(', ')}`);
      const master=buildMaster(recap,betRows);
      const zip=new JSZip();

      for (const ville of villes) {
        logRap(`  Génération ${ville}...`);
        const supers=makeSupers(master[ville]);
        const wb=buildWb(supers,{[ville]:master[ville]},moisUp);
        const fname=`ANALYSE_${ville.toUpperCase().replace(/ /g,'_')}_${moisUp.replace(/ /g,'_')}.xlsx`;
        zip.file(fname,XLSX.write(wb,{bookType:'xlsx',type:'array'}));
        logRap(`    ✅ ${fname}`);
      }

      logRap('  Génération GLOBAL...');
      const allSupers=villes.flatMap(v=>makeSupers(master[v]));
      allSupers.sort((a,b)=>{const pa=a.obj>0?(a.mtd-a.obj)/a.obj*100:0,pb=b.obj>0?(b.mtd-b.obj)/b.obj*100:0;return pb-pa;});
      const wbG=buildWb(allSupers,master,moisUp);
      const fnameG=`ANALYSE_GLOBAL_${moisUp.replace(/ /g,'_')}.xlsx`;
      zip.file(fnameG,XLSX.write(wbG,{bookType:'xlsx',type:'array'}));
      logRap(`    ✅ ${fnameG}`);

      // ── Fichiers individuels par super ──────────
      const superFiles = [];
      for (const ville of villes) {
        const supMap = {};
        (master[ville]||[]).forEach(r => {
          const sup = (r.SUPER||'').trim().replace(/\s+/g,' ');
          if (!sup) return;
          if (!supMap[sup]) supMap[sup] = [];
          supMap[sup].push(r);
        });
        for (const [sup, rows] of Object.entries(supMap)) {
          const supersLocal = makeSupers(rows);
          const wb = buildWb(supersLocal, {[ville]: rows}, moisUp);
          const safe = sup.replace(/ /g,'_').replace(/[/]/g,'-');
          const filename = `ANALYSE_SUPER_${safe}_${moisUp.replace(/ /g,'_')}.xlsx`;
          const blob2 = new Blob(
            [XLSX.write(wb, {bookType:'xlsx', type:'array'})],
            {type:'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'}
          );
          // eslint-disable-next-line no-unused-vars
          const sd = supersLocal[0] || {};
          superFiles.push({filename, blob: blob2, super: sup, ville, nb: rows.length, att: sd.att||0, natt: sd.natt||0});
          logRap(`    ✅ ${filename}`);
        }
      }
      setSupersRap(superFiles);

      const blob=await zip.generateAsync({type:'blob'});
      setResultsRap({blob,filename:`Rapports_${moisUp.replace(/ /g,'_')}.zip`});
      logRap(`\n✅ Terminé — ${villes.length+1} fichiers zones + ${superFiles.length} supers`);
    } catch(err){logRap(`❌ Erreur : ${err.message}`);}
    setRunningRap(false);
  };

  // ════════════════════════════════════════════════
  // UI HELPERS
  // ════════════════════════════════════════════════
  const statusTag = (status, loaded) => {
    if(status==='loading') return <span style={{fontSize:12,color:'var(--muted)'}}>⟳ Chargement...</span>;
    if(status==='saving')  return <span style={{fontSize:12,color:'var(--accent)'}}>⟳ Sauvegarde...</span>;
    if(status==='saved')   return <span style={{fontSize:12,color:'var(--deposit)'}}>✅ Sauvegardé !</span>;
    if(status==='error')   return <span style={{fontSize:12,color:'var(--danger)'}}>❌ Erreur Firebase</span>;
    if(status==='loaded'&&loaded) return <span style={{fontSize:12,color:'var(--deposit)',fontWeight:600}}>✅ Chargé depuis Firebase</span>;
    return null;
  };

  const phaseHeader = (num, label, sub, color) => (
    <div style={{display:'flex',alignItems:'center',gap:10,marginBottom:'1rem'}}>
      <div style={{width:28,height:28,borderRadius:'50%',background:color,display:'flex',alignItems:'center',
        justifyContent:'center',fontSize:13,fontWeight:700,color:'#fff',flexShrink:0}}>
        {num}
      </div>
      <div>
        <div style={{fontSize:15,fontWeight:700,color:'var(--text)'}}>{label}</div>
        <div style={{fontSize:12,color:'var(--muted)'}}>{sub}</div>
      </div>
    </div>
  );

  // ════════════════════════════════════════════════
  // RENDER
  // ════════════════════════════════════════════════
  return (
    <div>

      {/* ── RECAP PARTAGÉ ── */}
      <div className="card" style={{marginBottom:'1.25rem',borderLeft:'3px solid var(--accent)',borderRadius:8}}>
        <div style={{display:'flex',alignItems:'center',justifyContent:'space-between',marginBottom:'0.5rem'}}>
          <div>
            <p className="section-label" style={{margin:0}}>RECAP DES SALLES PAR ZONE — PARTAGÉ</p>
            {recapMeta && (
              <div style={{fontSize:11,color:'var(--muted)',marginTop:3}}>
                📁 {recapMeta.filename} · {recapMeta.nbSalles} salles · sauvegardé le {recapMeta.savedAt}
              </div>
            )}
          </div>
          <div style={{display:'flex',gap:8,alignItems:'center'}}>
            {statusTag(recapStatus, recapSaved)}
            <label htmlFor="recap-upload-input"
              style={{padding:'6px 14px',fontSize:12,fontWeight:600,cursor:'pointer',
                border:'1px solid var(--border2)',borderRadius:6,background:'var(--bg3)',color:'var(--text2)',
                display:'inline-flex',alignItems:'center'}}>
              {recapSaved ? '🔄 Mettre à jour' : '📂 Uploader'}
            </label>
            <input id="recap-upload-input" type="file" accept=".xlsx" style={{display:'none'}} onChange={async e=>{
              const f=e.target.files[0]; if(!f) return;
              const zd=await loadRecapFile(f);
              buildSuperList(zd);
              await saveRecapToFirebase(zd, f.name);
              e.target.value='';
            }}/>
            {recapSaved && (
              <button onClick={deleteRecapFromFirebase}
                style={{padding:'6px 12px',fontSize:12,fontWeight:600,cursor:'pointer',
                  border:'1px solid rgba(239,68,68,0.3)',borderRadius:6,background:'var(--danger-bg)',color:'var(--danger)'}}>
                🗑 Supprimer
              </button>
            )}

          </div>
        </div>
        {recapStatus==='empty'&&!recapSaved && (
          <div style={{padding:'10px 14px',background:'var(--casino-bg)',borderRadius:6,fontSize:12,color:'var(--casino)',border:'1px solid rgba(245,158,11,0.2)'}}>
            ⚠️ Aucun RECAP sauvegardé — uploadez le fichier pour le sauvegarder dans l'application
          </div>
        )}
        {recapSaved && (
          <div style={{display:'flex',flexWrap:'wrap',gap:6,marginTop:6}}>
            {Object.entries(recapSaved).map(([zone,rows])=>(
              <span key={zone} style={{fontSize:11,padding:'2px 10px',borderRadius:100,
                background:'var(--accent-dim)',color:'var(--accent)',fontWeight:600}}>
                {zone} · {rows.length} salles
              </span>
            ))}
          </div>
        )}
      </div>

      {/* ════════════════════════════════════════════
          PHASE 1 — OBJECTIFS
      ═════════════════════════════════════════════ */}
      <div className="card" style={{marginBottom:'1.25rem',borderTop:'3px solid #7F77DD'}}>
        {phaseHeader(1,'Objectifs du mois','Génération en début de mois — LY +15%','#7F77DD')}

        {/* LISTE sauvegardée */}
        <div style={{background:'var(--bg2)',borderRadius:6,padding:'10px 14px',marginBottom:'1rem'}}>
          <div style={{display:'flex',alignItems:'center',justifyContent:'space-between'}}>
            <div>
              <span style={{fontSize:12,fontWeight:600,color:'var(--text2)'}}>LISTE DES SALLES PAR PROPRIÉTAIRE</span>
              {listeMeta && <span style={{fontSize:11,color:'var(--muted)',marginLeft:8}}>· {listeMeta.nbProps} propriétaires · {listeMeta.savedAt}</span>}
            </div>
            <div style={{display:'flex',gap:6,alignItems:'center'}}>
              {statusTag(listeStatus,listeSaved)}
              <button onClick={()=>listeUpRef.current.click()}
                style={{padding:'4px 10px',fontSize:11,cursor:'pointer',border:'1px solid var(--border2)',borderRadius:5,background:'var(--bg3)',color:'var(--text2)'}}>
                {listeSaved?'🔄 Mettre à jour':'📂 Uploader'}
              </button>
              {listeSaved&&<button onClick={deleteListeFromStorage}
                style={{padding:'4px 8px',fontSize:11,cursor:'pointer',border:'1px solid rgba(239,68,68,0.3)',borderRadius:5,background:'var(--danger-bg)',color:'var(--danger)'}}>
                🗑
              </button>}
              <input ref={listeUpRef} type="file" accept=".xlsx,.xls" hidden onChange={async e=>{
                const f=e.target.files[0];if(!f)return;
                const {propMap:pm, cityMap:cm}=await loadListFile(f);
                await saveListeToFirebase(pm,f.name,cm);
                e.target.value='';
              }}/>
            </div>
          </div>
          {listeSaved&&(
            <div style={{display:'flex',flexWrap:'wrap',gap:5,marginTop:6}}>
              {Object.entries(listData||{}).sort(([a],[b])=>a.localeCompare(b)).map(([prop,shops])=>(
                <span key={prop} style={{fontSize:10,padding:'1px 8px',borderRadius:100,
                  background:'var(--deposit-bg)',color:'var(--deposit)',fontWeight:600}}>
                  {prop} · {shops.length}
                </span>
              ))}
            </div>
          )}
        </div>

        {/* Fichiers source */}
        <p className="section-label">Fichiers source</p>
        <div style={{marginBottom:'1rem'}}>
          <div onClick={()=>csvRef.current.click()}
            onDragOver={e=>e.preventDefault()}
            onDrop={e=>{e.preventDefault();setCsvFile(e.dataTransfer.files[0]);}}
            className={`drop-zone ${csvFile?'has-file':''}`}>
            <input ref={csvRef} type="file" accept=".csv" hidden onChange={e=>setCsvFile(e.target.files[0])}/>
            <div className="drop-icon">📋</div>
            <div className="drop-label">shift_report_summary.csv</div>
            <div className="drop-sub">Betshop name · Payin money · Payin count · Gross</div>
            {csvFile&&<div className="file-chip">✓ {csvFile.name}</div>}
          </div>
        </div>

        {/* Paramètres */}
        <div style={{display:'flex',gap:'1.5rem',flexWrap:'wrap',marginBottom:'1rem'}}>
          {[
            {label:'Mois (ex: MAI)',val:mois,set:setMois,w:120,type:'text'},
            {label:'Année',val:annee,set:e=>setAnnee(Number(e)),w:100,type:'number'},
            {label:'Taux objectif (%)',val:taux,set:e=>setTaux(Number(e)),w:100,type:'number'},
          ].map(({label,val,set,w,type})=>(
            <label key={label} style={{display:'flex',flexDirection:'column',gap:4}}>
              <span style={{fontSize:12,color:'var(--muted)',fontWeight:600}}>{label}</span>
              <input type={type} value={val} onChange={e=>set(e.target.value)} className="field-input" style={{width:w}}/>
            </label>
          ))}
          <label style={{display:'flex',flexDirection:'column',gap:4,flex:1}}>
            <span style={{fontSize:12,color:'var(--muted)',fontWeight:600}}>Zones avec fichiers par super</span>
            <input value={supers} onChange={e=>setSupers(e.target.value)} className="field-input"/>
          </label>
        </div>

        {/* Fichier liste optionnel */}
        <p className="section-label">Fichier optionnel — Objectifs par propriétaire</p>
        <div onClick={()=>listRef.current.click()} onDragOver={e=>e.preventDefault()}
          onDrop={e=>{e.preventDefault();const f=e.dataTransfer.files[0];setListFile(f);loadListFile(f).then(({propMap,cityMap})=>saveListeToFirebase(propMap,f.name,cityMap));}}
          className={`drop-zone ${listFile?'has-file':''}`} style={{marginBottom:'1rem'}}>
          <input ref={listRef} type="file" accept=".xlsx,.xls" hidden
            onChange={e=>{const f=e.target.files[0];setListFile(f);loadListFile(f).then(({propMap,cityMap})=>saveListeToFirebase(propMap,f.name,cityMap));}}/>
          <div className="drop-icon">🏪</div>
          <div className="drop-label">LISTE_DES_SALLES_CAMEROON.xlsx</div>
          <div className="drop-sub">Market · Shop · City · Address · Proprietaire</div>
          {listFile&&<div className="file-chip">✓ {listFile.name}</div>}
        </div>

        {/* Sélecteurs super / propriétaire */}
        {superList&&(
          <div style={{display:'grid',gridTemplateColumns:'1fr 1fr',gap:'1rem',marginBottom:'1rem'}}>
            {/* Super */}
            <div style={{background:'var(--bg2)',borderRadius:6,padding:'12px'}}>
              <p className="section-label" style={{marginBottom:8}}>Objectifs par Super en charge</p>
              <div style={{display:'flex',gap:8,alignItems:'flex-end'}}>
                <select value={selectedSupObj} onChange={e=>setSelectedSupObj(e.target.value)}
                  className="field-input" style={{flex:1,cursor:'pointer'}}>
                  <option value="">-- Choisir un super --</option>
                  {Object.entries(superList).sort(([a],[b])=>a.localeCompare(b)).map(([sup,data])=>(
                    <option key={sup} value={sup}>{sup} ({data.rows.length} salles)</option>
                  ))}
                </select>
                <button onClick={runSuperPipelineObj}
                  disabled={!selectedSupObj||!csvFile}
                  className="btn primary" style={{whiteSpace:'nowrap',padding:'8px 14px'}}>
                  ▶ Objectifs super
                </button>
              </div>
              {supResultObj&&(
                <button onClick={()=>saveAs(supResultObj.blob,supResultObj.filename)}
                  className="btn success" style={{marginTop:8,fontSize:12}}>
                  ⬇ {supResultObj.filename}
                </button>
              )}
            </div>
            {/* Propriétaire */}
            {listData&&(
              <div style={{background:'var(--bg2)',borderRadius:6,padding:'12px'}}>
                <p className="section-label" style={{marginBottom:8}}>Objectifs par Propriétaire</p>
                <div style={{display:'flex',gap:8,alignItems:'flex-end'}}>
                  <select value={selectedProp} onChange={e=>setSelectedProp(e.target.value)}
                    className="field-input" style={{flex:1,cursor:'pointer'}}>
                    <option value="">-- Choisir --</option>
                    {Object.entries(listData).sort(([a],[b])=>a.localeCompare(b)).map(([prop,shops])=>(
                      <option key={prop} value={prop}>{prop} ({shops.length})</option>
                    ))}
                  </select>
                  <button onClick={runPropPipeline}
                    disabled={!selectedProp||!csvFile}
                    className="btn primary" style={{whiteSpace:'nowrap',padding:'8px 14px'}}>
                    ▶ Objectifs prop.
                  </button>
                </div>
                {propResult&&(
                  <button onClick={()=>saveAs(propResult.blob,propResult.filename)}
                    className="btn success" style={{marginTop:8,fontSize:12}}>
                    ⬇ {propResult.filename}
                  </button>
                )}
              </div>
            )}
          </div>
        )}

        {/* Bouton principal */}
        <button onClick={runPipelineObj}
          disabled={!csvFile||(!recapFileObj&&!recapSaved)||runningObj}
          className="btn primary" style={{fontSize:14,padding:'11px 28px',marginBottom:'1rem'}}>
          {runningObj?'⟳ Génération...':'▶  Générer tous les objectifs'}
        </button>

        {logsObj.length>0&&(
          <div className="log-console" style={{marginBottom:'1rem'}}>
            {logsObj.map((l,i)=><span key={i} className="log-line" style={{display:'block'}}>{l}</span>)}
          </div>
        )}

        {resultObj&&(
          <div style={{display:'flex',alignItems:'center',gap:'1rem',flexWrap:'wrap'}}>
            <button onClick={()=>saveAs(resultObj.blob,resultObj.filename)}
              className="btn success" style={{fontSize:13}}>
              ⬇  Télécharger {resultObj.filename}
            </button>
            {resultObj.missing.length>0&&(
              <span style={{fontSize:12,color:'var(--casino)'}}>
                ⚠️ {resultObj.missing.length} salle(s) sans données CSV
              </span>
            )}
          </div>
        )}
      </div>

      {/* Connecteur visuel */}
      <div style={{textAlign:'center',margin:'0.5rem 0',color:'var(--muted)',fontSize:12,display:'flex',alignItems:'center',justifyContent:'center',gap:8}}>
        <div style={{flex:1,height:1,background:'var(--border)'}}/>
        <span>↓ RECAP transmis automatiquement · mois pré-rempli</span>
        <div style={{flex:1,height:1,background:'var(--border)'}}/>
      </div>

      {/* ════════════════════════════════════════════
          PHASE 2 — RÉALISATION
      ═════════════════════════════════════════════ */}
      <div className="card" style={{marginBottom:'1.25rem',borderTop:'3px solid #1D9E75'}}>
        {phaseHeader(2,'Réalisation vs Objectifs','Rapport analytique en cours / fin de mois — objectif = LY+15% du rapport','#1D9E75')}

        {/* Note */}
        <div style={{padding:'8px 12px',background:'var(--accent-dim)',borderRadius:6,fontSize:12,color:'var(--accent)',marginBottom:'1rem',fontWeight:500}}>
          ✓ RECAP chargé depuis Firebase · Objectif = colonne "Bet ticket LY +15 %" — aucun fichier objectifs à uploader
        </div>

        {/* Rapport betshop */}
        <p className="section-label">Rapport Betshop brut</p>
        <div onClick={()=>reportRef.current.click()} onDragOver={e=>e.preventDefault()}
          onDrop={e=>{e.preventDefault();setReportFile(e.dataTransfer.files[0]);}}
          className={`drop-zone ${reportFile?'has-file':''}`} style={{marginBottom:'1rem'}}>
          <input ref={reportRef} type="file" accept=".xlsx,.xls" hidden
            onChange={e=>setReportFile(e.target.files[0])}/>
          <div className="drop-icon">📋</div>
          <div className="drop-label">Betshop_report_Cameroon_YYYY-MM.xlsx</div>
          <div className="drop-sub">Colonnes : Market · Shop · Bet tickets MTD · LM · LY · LY+15% · Payin · GGR</div>
          {reportFile&&<div className="file-chip">✓ {reportFile.name}</div>}
        </div>



        {/* Super individuel */}
        {superList&&(
          <div style={{background:'var(--bg2)',borderRadius:6,padding:'12px',marginBottom:'1rem'}}>
            <p className="section-label" style={{marginBottom:8}}>Réalisation individuelle — Super en charge</p>
            <div style={{display:'flex',gap:'1rem',alignItems:'flex-end',flexWrap:'wrap'}}>
              <select value={selectedSupRap} onChange={e=>setSelectedSupRap(e.target.value)}
                className="field-input" style={{flex:1,cursor:'pointer'}}>
                <option value="">-- Choisir un super --</option>
                {Object.entries(superList).sort(([a],[b])=>a.localeCompare(b)).map(([sup,data])=>(
                  <option key={sup} value={sup}>{sup} ({data.rows.length} salles — {data.ville})</option>
                ))}
              </select>
              <button onClick={runSuperPipelineRap}
                disabled={!selectedSupRap||!reportFile||(!recapFileRap&&!recapSaved)}
                className="btn primary" style={{padding:'9px 20px',whiteSpace:'nowrap'}}>
                ▶  Réalisation {selectedSupRap||'super'}
              </button>
            </div>
            {supResultRap&&(
              <button onClick={()=>saveAs(supResultRap.blob,supResultRap.filename)}
                className="btn success" style={{marginTop:8,fontSize:13}}>
                ⬇  Télécharger {supResultRap.filename}
              </button>
            )}
          </div>
        )}

        {/* Mois pré-rempli */}
        <div style={{marginBottom:'1.5rem'}}>
          <label style={{display:'flex',flexDirection:'column',gap:4,fontSize:13,color:'var(--muted)',maxWidth:220}}>
            Mois
            <input value={moisRap} onChange={e=>setMoisRap(e.target.value)} className="field-input"/>
          </label>
          {moisRap&&mois&&moisRap.toUpperCase().includes(mois.toUpperCase())&&(
            <div style={{fontSize:11,color:'var(--accent)',marginTop:4}}>
              🔗 Pré-rempli depuis la Phase 1
            </div>
          )}
        </div>

        <div style={{display:'flex',gap:'1rem',flexWrap:'wrap',marginBottom:'1.5rem'}}>
          <button onClick={runPipelineRap}
            disabled={!reportFile||(!recapFileRap&&!recapSaved)||runningRap}
            className="btn primary" style={{fontSize:14,padding:'11px 28px'}}>
            {runningRap?'⏳ Génération...':'▶ Générer les rapports (zones + global)'}
          </button>
          <button onClick={runAllSupersPipeline}
            disabled={!reportFile||(!recapFileRap&&!recapSaved)||runningRap}
            className="btn primary" style={{fontSize:14,padding:'11px 28px',background:'var(--accent)'}}>
            {runningRap?'⏳ Génération...':'👤 Générer tous les rapports supers'}
          </button>
        </div>

        {logsRap.length>0&&(
          <div className="log-console" style={{marginBottom:'1rem'}}>
            {logsRap.map((l,i)=><span key={i} className="log-line" style={{display:'block'}}>{l}</span>)}
          </div>
        )}

        {resultsRap&&(
          <div style={{marginTop:'0.5rem'}}>
            <button onClick={()=>saveAs(resultsRap.blob,resultsRap.filename)}
              className="btn success" style={{fontSize:14,padding:'10px 24px'}}>
              ⬇ Télécharger tout — {resultsRap.filename}
            </button>
          </div>
        )}

        {/* ── Résultats individuels par super ── */}
        {supersRap.length>0&&(
          <div style={{marginTop:'1.5rem'}}>
            <p className="section-label">Rapports par Super — {supersRap.length} fichiers</p>
            <div style={{display:'flex',flexDirection:'column',gap:6}}>
              {supersRap.map((f,i)=>(
                <div key={i} style={{display:'flex',alignItems:'center',justifyContent:'space-between',
                  padding:'10px 14px',background:'var(--bg2)',borderRadius:6,
                  border:'1px solid var(--border)'}}>
                  <div>
                    <div style={{fontSize:13,fontWeight:600,color:'var(--text)'}}>{f.super}</div>
                    <div style={{fontSize:11,color:'var(--muted)',marginTop:2}}>
                      {f.ville} · {f.nb} salles ·{' '}
                      <span style={{color:'var(--deposit)'}}>✅ {f.att} atteints</span>
                      {' · '}
                      <span style={{color:'var(--danger)'}}>❌ {f.natt} non atteints</span>
                    </div>
                  </div>
                  <button onClick={()=>saveAs(f.blob,f.filename)}
                    className="btn success"
                    style={{fontSize:12,padding:'6px 14px',whiteSpace:'nowrap'}}>
                    ⬇ Download
                  </button>
                </div>
              ))}
            </div>
          </div>
        )}
      </div>

    </div>
  );
}
