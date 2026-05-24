// src/pages/BetshopMarchesPage.js
import React, { useState, useRef } from 'react';
import * as XLSX from 'xlsx';
import { saveAs } from 'file-saver';
import JSZip from 'jszip';

// ── Palette ──────────────────────────────────────
const C = {
  dark_blue:'1F3864', mid_blue:'2E75B6', green_med:'375623', green_bg:'E2EFDA',
  green_fg:'375623', red_dark:'5C0A0A', red_med:'7B2C2C', red_bg:'FCE4D6',
  red_fg:'C55A11', gold_bg:'FFF2CC', gold_fg:'7F6000', gold_dark:'BF8F00',
  grey:'F2F2F2', white:'FFFFFF',
};
const FMT_NB='#,##0', FMT_DIF='#,##0;[Red]-#,##0', FMT_PCT='0.0"%"';

function bdr() {
  const t = { style:'thin', color:{ rgb:'BFBFBF' } };
  return { top:t, bottom:t, left:t, right:t };
}
function fill(rgb) { return { patternType:'solid', fgColor:{ rgb } }; }
function fnt(bold=false, sz=10, rgb='000000', italic=false) {
  return { name:'Arial', bold, sz, color:{ rgb }, italic };
}
function aln(h='center', v='center', wrap=false) {
  return { horizontal:h, vertical:v, wrapText:wrap };
}
function sc(ws, r, c, v, style) {
  const addr = XLSX.utils.encode_cell({ r:r-1, c:c-1 });
  ws[addr] = { v, t: typeof v === 'number' ? 'n' : 's', s: style };
}
function scn(ws, r, c, v, style, fmt) {
  const addr = XLSX.utils.encode_cell({ r:r-1, c:c-1 });
  ws[addr] = { v, t:'n', s: style, z: fmt };
}
function merge(ws, r1, c1, r2, c2) {
  if (!ws['!merges']) ws['!merges'] = [];
  ws['!merges'].push({ s:{ r:r1-1, c:c1-1 }, e:{ r:r2-1, c:c2-1 } });
}

// ── Cell style builders ───────────────────────────
function hStyle(bg, fg='FFFFFF', sz=10) {
  return { font:fnt(true,sz,fg), fill:fill(bg), border:bdr(), alignment:aln('center','center',true) };
}
function dStyle(bg='FFFFFF', bold=false, fg='000000', h='right') {
  return { font:fnt(bold,10,fg), fill:fill(bg), border:bdr(), alignment:aln(h,'center') };
}
function titleStyle(bg) {
  return { font:fnt(true,13,'FFFFFF'), fill:fill(bg), alignment:aln('center','center') };
}
function sectionStyle(bg) {
  return { font:fnt(true,11,'FFFFFF'), fill:fill(bg), border:{ top:{style:'medium',color:{rgb:'404040'}}, bottom:{style:'medium',color:{rgb:'404040'}}, left:{style:'medium',color:{rgb:'404040'}}, right:{style:'medium',color:{rgb:'404040'}} }, alignment:aln('center','center') };
}

// ── Prep dataframe ────────────────────────────────
function prepDf(rows) {
  return rows.map((row, i) => {
    const n = (k) => parseFloat(row[k]) || 0;
    const mtd  = n('Bet tickets MTD');
    const lm   = n('Bet tickets LM');
    const ly   = n('Bet tickets LY');
    const lm15 = n('Bet Ticket LM +15%');
    const ly15 = n('Bet ticket LY +15 %');
    const pmtd = n('Payin Bet MTD'), plm = n('Payin Bet LM'), ply = n('Payin Bet LY');
    const gmtd = n('GGR Bet MTD'),  glm = n('GGR Bet LM');
    return {
      ...row,
      Shop: row['Shop'] || row['Betshop name'] || '',
      'Bet tickets MTD': mtd, 'Bet tickets LM': lm, 'Bet tickets LY': ly,
      'Bet Ticket LM +15%': lm15, 'Bet ticket LY +15 %': ly15,
      'Payin Bet MTD': pmtd, 'Payin Bet LM': plm, 'Payin Bet LY': ply,
      'GGR Bet MTD': gmtd, 'GGR Bet LM': glm,
      Diff_T_LM: mtd - lm, Diff_T_LY: mtd - ly,
      Diff_T_LM15: mtd - lm15, Diff_T_LY15: mtd - ly15,
      Pct_T_LM: lm !== 0 ? (mtd - lm) / lm * 100 : null,
      Pct_T_LY: ly !== 0 ? (mtd - ly) / ly * 100 : null,
      Diff_P_LM: pmtd - plm, Diff_P_LY: pmtd - ply,
      Pct_P_LM: plm !== 0 ? (pmtd - plm) / plm * 100 : null,
      Diff_G_LM: gmtd - glm,
      Pct_G_LM: glm !== 0 ? (gmtd - glm) / Math.abs(glm) * 100 : null,
      Statut: mtd === 0 ? '⛔ INACTIF' : '✅ Actif',
    };
  });
}

// ── Sheet: Données ────────────────────────────────
function buildOngData(wb, sheetName, df, titre, mois) {
  const ws = {};
  // Title
  const titleR = 2;
  for (let c = 1; c <= 21; c++) {
    const addr = XLSX.utils.encode_cell({ r: titleR-1, c: c-1 });
    ws[addr] = { v: c === 1 ? `${titre} | ${mois}  ·  ${df.length} salles` : '', t:'s', s: titleStyle(C.dark_blue) };
  }
  merge(ws, titleR, 1, titleR, 21);

  // Group headers row 4
  const ghRow = 4;
  [
    [1,1,'SALLE',C.dark_blue], [2,10,'🎫 TICKETS','1F4E79'],
    [11,16,'💰 PAYIN (FCFA)',C.green_med], [17,21,'📈 GGR (FCFA)',C.red_med]
  ].forEach(([cs,ce,val,bg]) => {
    sc(ws, ghRow, cs, val, hStyle(bg, 'FFFFFF', 10));
    if (cs !== ce) merge(ws, ghRow, cs, ghRow, ce);
  });

  // Sub-headers row 5
  const shRow = 5;
  const subHdrs = [
    [1,'Salle',C.dark_blue],[2,'MTD','1F4E79'],[3,'LM','1F4E79'],[4,'LY','1F4E79'],
    [5,'Diff LM','2E5F8A'],[6,'% LM','2E5F8A'],[7,'Diff LY','2E5F8A'],[8,'% LY','2E5F8A'],
    [9,'LM+15%','1F4E79'],[10,'Diff LM+15%','2E5F8A'],
    [11,'MTD',C.green_med],[12,'LM',C.green_med],[13,'LY',C.green_med],
    [14,'Diff LM','4A7C3F'],[15,'% LM','4A7C3F'],[16,'Diff LY','4A7C3F'],
    [17,'MTD',C.red_med],[18,'LM',C.red_med],[19,'LY',C.red_med],
    [20,'Diff LM','A04040'],[21,'% LM','A04040'],
  ];
  subHdrs.forEach(([col, lbl, bg]) => sc(ws, shRow, col, lbl, hStyle(bg, 'FFFFFF', 9)));

  // Data rows
  df.forEach((row, i) => {
    const dr = shRow + 1 + i;
    const inactive = row.Statut === '⛔ INACTIF';
    const bg = inactive ? 'EFEFEF' : (i % 2 === 0 ? 'EBF3FB' : C.white);
    const nameStyle = { font: fnt(false, 10, inactive ? '7F7F7F' : '000000', inactive), fill: fill(bg), border: bdr(), alignment: aln('left','center') };
    sc(ws, dr, 1, row.Shop, nameStyle);

    const dStyle2 = (v, good, inv=false) => {
      const isGood = inv ? v < 0 : v > 0;
      const fgColor = (v === null || v === undefined) ? '000000' : (isGood ? C.green_fg : C.red_fg);
      return { font: fnt(true, 10, fgColor), fill: fill(bg), border: bdr(), alignment: aln('right','center') };
    };
    const plain = { font: fnt(false,10,'000000'), fill: fill(bg), border: bdr(), alignment: aln('right','center') };

    scn(ws,dr,2,row['Bet tickets MTD'],plain,FMT_NB);
    scn(ws,dr,3,row['Bet tickets LM'],plain,FMT_NB);
    scn(ws,dr,4,row['Bet tickets LY'],plain,FMT_NB);
    scn(ws,dr,5,row.Diff_T_LM, dStyle2(row.Diff_T_LM, true), FMT_DIF);
    scn(ws,dr,6,row.Pct_T_LM||0, dStyle2(row.Pct_T_LM, true), FMT_PCT);
    scn(ws,dr,7,row.Diff_T_LY, dStyle2(row.Diff_T_LY, true), FMT_DIF);
    scn(ws,dr,8,row.Pct_T_LY||0, dStyle2(row.Pct_T_LY, true), FMT_PCT);
    scn(ws,dr,9,row['Bet Ticket LM +15%'],plain,FMT_NB);
    scn(ws,dr,10,row.Diff_T_LM15, dStyle2(row.Diff_T_LM15, true), FMT_DIF);
    scn(ws,dr,11,row['Payin Bet MTD'],plain,FMT_NB);
    scn(ws,dr,12,row['Payin Bet LM'],plain,FMT_NB);
    scn(ws,dr,13,row['Payin Bet LY'],plain,FMT_NB);
    scn(ws,dr,14,row.Diff_P_LM, dStyle2(row.Diff_P_LM, true), FMT_DIF);
    scn(ws,dr,15,row.Pct_P_LM||0, dStyle2(row.Pct_P_LM, true), FMT_PCT);
    scn(ws,dr,16,row.Diff_P_LY, dStyle2(row.Diff_P_LY, true), FMT_DIF);
    scn(ws,dr,17,row['GGR Bet MTD'],plain,FMT_NB);
    scn(ws,dr,18,row['GGR Bet LM'],plain,FMT_NB);
    scn(ws,dr,19,row['GGR Bet LY'],plain,FMT_NB);
    scn(ws,dr,20,row.Diff_G_LM, dStyle2(row.Diff_G_LM, true), FMT_DIF);
    scn(ws,dr,21,row.Pct_G_LM||0, dStyle2(row.Pct_G_LM, true), FMT_PCT);
  });

  // Total row
  const tr = shRow + 1 + df.length;
  const sum = (k) => df.reduce((s, r) => s + (parseFloat(r[k]) || 0), 0);
  const tStyle = { font: fnt(true,10,'FFFFFF'), fill: fill(C.dark_blue), border: bdr(), alignment: aln('right','center') };
  sc(ws, tr, 1, `TOTAL (${df.length})`, { ...tStyle, alignment: aln('left','center') });
  scn(ws,tr,2,sum('Bet tickets MTD'),tStyle,FMT_NB);
  scn(ws,tr,3,sum('Bet tickets LM'),tStyle,FMT_NB);
  scn(ws,tr,4,sum('Bet tickets LY'),tStyle,FMT_NB);
  scn(ws,tr,5,sum('Diff_T_LM'),tStyle,FMT_DIF);
  scn(ws,tr,11,sum('Payin Bet MTD'),tStyle,FMT_NB);
  scn(ws,tr,14,sum('Diff_P_LM'),tStyle,FMT_DIF);
  scn(ws,tr,17,sum('GGR Bet MTD'),tStyle,FMT_NB);
  scn(ws,tr,20,sum('Diff_G_LM'),tStyle,FMT_DIF);

  ws['!cols'] = [32,14,14,14,13,10,13,10,14,14,14,14,14,13,10,13,14,14,14,13,10].map(w => ({wch:w}));
  ws['!ref'] = XLSX.utils.encode_range({ s:{r:0,c:0}, e:{r:tr-1,c:20} });
  XLSX.utils.book_append_sheet(wb, ws, sheetName.slice(0,31));
}

// ── Sheet: Résumé ─────────────────────────────────
function buildOngResume(wb, sheetName, df, titre, mois) {
  const ws = {};
  const sum = (k) => df.reduce((s, r) => s + (parseFloat(r[k]) || 0), 0);
  const nb_actif = df.filter(r => r.Statut === '✅ Actif').length;
  const prog = df.filter(r => r.Diff_T_LM > 0).length;
  const recul = df.filter(r => r.Diff_T_LM < 0).length;

  // Title
  sc(ws, 2, 1, `${titre} — RÉSUMÉ | ${mois}`, titleStyle(C.dark_blue));
  merge(ws, 2, 1, 2, 6);

  // KPIs
  sc(ws, 4, 1, '📊 INDICATEURS CLÉS', sectionStyle(C.mid_blue));
  merge(ws, 4, 1, 4, 6);

  const hdrs = ['Indicateur','Valeur MTD','Valeur LM','Variation','Var %','Tendance'];
  hdrs.forEach((h, i) => sc(ws, 5, i+1, h, hStyle(C.mid_blue)));

  const kpis = [
    ['Salles totales', df.length, null],
    ['Salles actives', nb_actif, null],
    ['En progression vs LM', prog, null],
    ['En recul vs LM', recul, null],
    ['Tickets MTD', sum('Bet tickets MTD'), sum('Bet tickets LM')],
    ['Payin MTD (FCFA)', sum('Payin Bet MTD'), sum('Payin Bet LM')],
    ['GGR MTD (FCFA)', sum('GGR Bet MTD'), sum('GGR Bet LM')],
  ];

  kpis.forEach(([label, val, lm], i) => {
    const r = 6 + i;
    const bg = i % 2 === 0 ? C.grey : C.white;
    sc(ws, r, 1, label, dStyle(bg, true, '000000', 'left'));
    scn(ws, r, 2, val, dStyle(bg), FMT_NB);
    if (lm !== null) {
      scn(ws, r, 3, lm, dStyle(bg), FMT_NB);
      const diff = val - lm;
      const pct = lm !== 0 ? diff / lm * 100 : 0;
      const isGood = diff >= 0;
      const varStyle = { font: fnt(true,10,isGood?C.green_fg:C.red_fg), fill: fill(isGood?C.green_bg:C.red_bg), border: bdr(), alignment: aln('right','center') };
      scn(ws, r, 4, diff, varStyle, FMT_DIF);
      scn(ws, r, 5, pct, varStyle, FMT_PCT);
      sc(ws, r, 6, `${diff>=0?'▲':'▼'} ${pct>0?'+':''}${pct.toFixed(1)}%`, { ...varStyle, alignment: aln('center','center') });
    } else {
      [3,4,5,6].forEach(c => sc(ws, r, c, '', dStyle(bg)));
    }
  });

  // Top/Bottom 10
  const sorted_prog = [...df].sort((a,b) => b.Diff_T_LM - a.Diff_T_LM).slice(0,10);
  const sorted_recul = [...df].sort((a,b) => a.Diff_T_LM - b.Diff_T_LM).slice(0,10);
  const nb = Math.min(10, df.length);
  let r = 14;
  sc(ws, r, 1, `🏆 TOP ${nb} PROGRESSIONS`, sectionStyle(C.green_med));
  merge(ws, r, 1, r, 3);
  sc(ws, r, 4, `⚠️ TOP ${nb} RECULS`, sectionStyle(C.red_med));
  merge(ws, r, 4, r, 6);
  r++;
  ['Salle','Tickets MTD','Diff LM','Salle','Tickets MTD','Diff LM'].forEach((h,i) => {
    sc(ws, r, i+1, h, hStyle(i < 3 ? '4A7C3F' : 'A04040', 'FFFFFF', 9));
  });
  for (let i = 0; i < nb; i++) {
    r++;
    const rp = sorted_prog[i], rr = sorted_recul[i];
    const bgp = i % 2 === 0 ? C.green_bg : C.white;
    const bgr = i % 2 === 0 ? C.red_bg : C.white;
    sc(ws, r, 1, rp?.Shop||'', dStyle(bgp,false,'000000','left'));
    scn(ws, r, 2, rp?.['Bet tickets MTD']||0, dStyle(bgp), FMT_NB);
    scn(ws, r, 3, rp?.Diff_T_LM||0, { font:fnt(true,10,C.green_fg), fill:fill(bgp), border:bdr(), alignment:aln('right','center') }, FMT_DIF);
    sc(ws, r, 4, rr?.Shop||'', dStyle(bgr,false,'000000','left'));
    scn(ws, r, 5, rr?.['Bet tickets MTD']||0, dStyle(bgr), FMT_NB);
    scn(ws, r, 6, rr?.Diff_T_LM||0, { font:fnt(true,10,C.red_fg), fill:fill(bgr), border:bdr(), alignment:aln('right','center') }, FMT_DIF);
  }

  ws['!cols'] = [32,18,18,16,16,18].map(w => ({wch:w}));
  ws['!ref'] = XLSX.utils.encode_range({ s:{r:0,c:0}, e:{r:r,c:5} });
  XLSX.utils.book_append_sheet(wb, ws, sheetName.slice(0,31));
}

// ── Sheet: TAPER+ ─────────────────────────────────
function buildTaperPlus(wb, sheetName, df, seuilPlus, titre, mois) {
  const ws = {};
  sc(ws, 2, 2, `TAPER +  ·  ${titre}  ·  ${mois}`, titleStyle(C.dark_blue));
  merge(ws, 2, 2, 2, 8);
  sc(ws, 3, 2, `⚙️ Seuil minimum : ${seuilPlus.toLocaleString()} tickets MTD`, { font:fnt(false,9,C.gold_fg), fill:fill(C.gold_bg), border:bdr(), alignment:aln('left','center') });
  merge(ws, 3, 2, 3, 8);

  const dp = df.filter(r => r['Bet tickets MTD'] >= seuilPlus);
  const blocs = [
    { titre:'Plus que LM', emoji:'✅', rows: dp.filter(r=>r.Diff_T_LM>0).sort((a,b)=>b.Diff_T_LM-a.Diff_T_LM), bg:C.green_med, bgh:'4A7C3F', bge:C.green_bg, ref:'LM', rk:'Bet tickets LM', dk:'Diff_T_LM', pk:'Pct_T_LM' },
    { titre:'Plus que LY', emoji:'📅', rows: dp.filter(r=>r.Diff_T_LY>0).sort((a,b)=>b.Diff_T_LY-a.Diff_T_LY), bg:'1F6B38', bgh:'2E8B57', bge:'EBF5EB', ref:'LY', rk:'Bet tickets LY', dk:'Diff_T_LY', pk:'Pct_T_LY' },
    { titre:'Dépassé LM+15%', emoji:'🏆', rows: dp.filter(r=>r.Diff_T_LM15>0).sort((a,b)=>b.Diff_T_LM15-a.Diff_T_LM15), bg:C.gold_dark, bgh:'A07800', bge:C.gold_bg, ref:'LM+15%', rk:'Bet Ticket LM +15%', dk:'Diff_T_LM15', pk:'Pct_T_LM' },
  ];

  let row = 5;
  blocs.forEach(bloc => {
    sc(ws, row, 2, `${bloc.emoji}  ${bloc.titre}  —  ${bloc.rows.length} salle${bloc.rows.length!==1?'s':''}`, sectionStyle(bloc.bg));
    merge(ws, row, 2, row, 8); row++;
    sc(ws, row, 2, `Seuil ${seuilPlus.toLocaleString()} tickets min`, { font:fnt(false,9,'FFFFFF',true), fill:fill(bloc.bgh), border:bdr(), alignment:aln('left','center') });
    merge(ws, row, 2, row, 8); row++;
    ['#','Salle','Tickets MTD',bloc.ref,'Diff','%','GGR MTD'].forEach((h,i) => sc(ws, row, i+2, h, hStyle(bloc.bgh,'FFFFFF',9)));
    row++;
    bloc.rows.forEach((r, i) => {
      const bg = i%2===0 ? bloc.bge : C.white;
      const isGood = true;
      sc(ws, row, 2, i+1, { font:fnt(true,10,'FFFFFF'), fill:fill(bloc.bg), border:bdr(), alignment:aln('center','center') });
      sc(ws, row, 3, r.Shop, dStyle(bg,false,'000000','left'));
      scn(ws, row, 4, r['Bet tickets MTD'], dStyle(bg), FMT_NB);
      scn(ws, row, 5, r[bloc.rk], dStyle(bg), FMT_NB);
      scn(ws, row, 6, r[bloc.dk], { font:fnt(true,10,C.green_fg), fill:fill(bg), border:bdr(), alignment:aln('right','center') }, FMT_DIF);
      scn(ws, row, 7, r[bloc.pk]||0, { font:fnt(true,10,C.green_fg), fill:fill(bg), border:bdr(), alignment:aln('right','center') }, FMT_PCT);
      scn(ws, row, 8, r['GGR Bet MTD'], dStyle(bg), FMT_NB);
      row++;
    });
    if (bloc.rows.length > 0) {
      const tStyle2 = { font:fnt(true,10,'FFFFFF'), fill:fill(bloc.bg), border:bdr(), alignment:aln('right','center') };
      sc(ws, row, 2, `TOTAL (${bloc.rows.length})`, { ...tStyle2, alignment:aln('left','center') });
      scn(ws, row, 4, bloc.rows.reduce((s,r)=>s+r['Bet tickets MTD'],0), tStyle2, FMT_NB);
      scn(ws, row, 6, bloc.rows.reduce((s,r)=>s+r[bloc.dk],0), tStyle2, FMT_DIF);
      scn(ws, row, 8, bloc.rows.reduce((s,r)=>s+r['GGR Bet MTD'],0), tStyle2, FMT_NB);
      row++;
    }
    row++;
  });

  ws['!cols'] = [3,5,30,13,13,14,10,14].map(w => ({wch:w}));
  ws['!ref'] = XLSX.utils.encode_range({ s:{r:0,c:0}, e:{r:row,c:8} });
  XLSX.utils.book_append_sheet(wb, ws, sheetName.slice(0,31));
}

// ── Sheet: TAPER- ─────────────────────────────────
function buildTaperMoins(wb, sheetName, df, seuilMoins, titre, mois) {
  const ws = {};
  sc(ws, 2, 2, `TAPER −  ·  ${titre}  ·  ${mois}`, titleStyle(C.dark_blue));
  merge(ws, 2, 2, 2, 8);

  const dm = df.filter(r => r['Bet tickets MTD'] >= seuilMoins);
  const blocs = [
    { titre:'Moins que LM', emoji:'⚠️', rows: dm.filter(r=>r.Diff_T_LM<0).sort((a,b)=>a.Diff_T_LM-b.Diff_T_LM), bg:C.red_med, bgh:'A04040', bge:C.red_bg, dk:'Diff_T_LM', pk:'Pct_T_LM' },
    { titre:'Moins que LY', emoji:'📉', rows: dm.filter(r=>r.Diff_T_LY<0).sort((a,b)=>a.Diff_T_LY-b.Diff_T_LY), bg:'8B1A1A', bgh:'B03030', bge:'FCE4D6', dk:'Diff_T_LY', pk:'Pct_T_LY' },
    { titre:"N'a pas atteint LM+15%", emoji:'🔴', rows: dm.filter(r=>r.Diff_T_LM15<0).sort((a,b)=>a.Diff_T_LM15-b.Diff_T_LM15), bg:C.red_dark, bgh:'8B1A1A', bge:'FCE4D6', dk:'Diff_T_LM15', pk:'Pct_T_LM' },
  ];

  let row = 5;
  blocs.forEach(bloc => {
    sc(ws, row, 2, `${bloc.emoji}  ${bloc.titre}  —  ${bloc.rows.length} salle${bloc.rows.length!==1?'s':''}`, sectionStyle(bloc.bg));
    merge(ws, row, 2, row, 8); row++;
    sc(ws, row, 2, `Seuil ${seuilMoins.toLocaleString()} tickets min`, { font:fnt(false,9,'FFFFFF',true), fill:fill(bloc.bgh), border:bdr(), alignment:aln('left','center') });
    merge(ws, row, 2, row, 8); row++;
    ['#','Salle','Tickets MTD','Référence','Diff','%','GGR MTD'].forEach((h,i) => sc(ws, row, i+2, h, hStyle(bloc.bgh,'FFFFFF',9)));
    row++;
    bloc.rows.forEach((r, i) => {
      const bg = i%2===0 ? bloc.bge : C.white;
      sc(ws, row, 2, i+1, { font:fnt(true,10,'FFFFFF'), fill:fill(bloc.bg), border:bdr(), alignment:aln('center','center') });
      sc(ws, row, 3, r.Shop, dStyle(bg,false,'000000','left'));
      scn(ws, row, 4, r['Bet tickets MTD'], dStyle(bg), FMT_NB);
      scn(ws, row, 5, 0, dStyle(bg), FMT_NB);
      scn(ws, row, 6, r[bloc.dk], { font:fnt(true,10,C.red_fg), fill:fill(bg), border:bdr(), alignment:aln('right','center') }, FMT_DIF);
      scn(ws, row, 7, r[bloc.pk]||0, { font:fnt(true,10,C.red_fg), fill:fill(bg), border:bdr(), alignment:aln('right','center') }, FMT_PCT);
      scn(ws, row, 8, r['GGR Bet MTD'], dStyle(bg), FMT_NB);
      row++;
    });
    row++;
  });

  ws['!cols'] = [3,5,30,13,13,14,10,14].map(w => ({wch:w}));
  ws['!ref'] = XLSX.utils.encode_range({ s:{r:0,c:0}, e:{r:row,c:8} });
  XLSX.utils.book_append_sheet(wb, ws, sheetName.slice(0,31));
}

// ── Build workbook for one market ─────────────────
function buildMarketWb(df, label, mois, seuilPlus, seuilMoins) {
  const wb = XLSX.utils.book_new();
  buildOngData(wb, `① ${label.slice(0,25)}`, df, label, mois);
  buildOngResume(wb, '② Résumé', df, label, mois);
  buildTaperPlus(wb, '③ TAPER+', df, seuilPlus, label, mois);
  buildTaperMoins(wb, '④ TAPER−', df, seuilMoins, label, mois);
  return wb;
}

// ── Main Component ────────────────────────────────
export default function BetshopMarchesPage() {
  const [evalFile, setEvalFile]   = useState(null);
  const [reportFile, setReportFile] = useState(null);
  const [mois, setMois]           = useState('AVRIL 2026');
  const [seuilPlus, setSeuilPlus] = useState(500);
  const [seuilMoins, setSeuilMoins] = useState(2000);
  const [marches, setMarches]     = useState('Balnda,Janvier,Bunker,Hippolyte,LUniversel,Franck');
  const [uniGroupes, setUniGroupes] = useState('UNI DG,UNI SF,UNI TA');
  const [running, setRunning]     = useState(false);
  const [logs, setLogs]           = useState([]);
  const [result, setResult]       = useState(null);
  const evalRef   = useRef();
  const reportRef = useRef();

  const log = (msg) => setLogs(p => [...p, msg]);

  const readXlsx = (file, sheet) => new Promise((res, rej) => {
    const reader = new FileReader();
    reader.onload = e => {
      try {
        const wb = XLSX.read(e.target.result, { type:'array' });
        const wsName = sheet || wb.SheetNames[0];
        const ws = wb.Sheets[wsName] || wb.Sheets[wb.SheetNames[0]];
        res(XLSX.utils.sheet_to_json(ws, { defval: 0 }));
      } catch(err) { rej(err); }
    };
    reader.onerror = rej;
    reader.readAsArrayBuffer(file);
  });

  const runPipeline = async () => {
    setRunning(true); setLogs([]); setResult(null);
    try {
      log('📂 Chargement des fichiers...');
      const [evalRows, reportRows] = await Promise.all([
        readXlsx(evalFile, 'CAMPRESJ'),
        readXlsx(reportFile, 'Sheet1'),
      ]);
      const campresj_list = new Set(evalRows.map(r => r['Shop'] || r['Betshop name'] || ''));
      log(`  ✓ Eval CAMPRESJ : ${campresj_list.size} salles`);
      log(`  ✓ Report : ${reportRows.length} salles`);

      const df_all = prepDf(reportRows);
      const moisStr = mois.toUpperCase();
      const zip = new JSZip();

      // Global
      log('🌍 Génération Global...');
      const df_camp_all = df_all.filter(r => campresj_list.has(r.Shop));
      const wbGlobal = XLSX.utils.book_new();
      buildOngData(wbGlobal, '① Global', df_all, 'Global', moisStr);
      buildOngData(wbGlobal, '② CAMPRESJ', df_camp_all, 'CAMPRESJ', moisStr);
      buildOngResume(wbGlobal, '③ Résumé Global', df_all, 'Global', moisStr);
      buildTaperPlus(wbGlobal, '④ TAPER+ Global', df_all, seuilPlus, 'Global', moisStr);
      buildTaperMoins(wbGlobal, '⑤ TAPER− Global', df_all, seuilMoins, 'Global', moisStr);
      zip.file(`Betshop_Global_${moisStr.replace(/ /g,'_')}.xlsx`, XLSX.write(wbGlobal, { bookType:'xlsx', type:'array' }));
      log(`  ✅ Global (${df_all.length} salles)`);

      // CAMPRESJ
      log('📋 Génération CAMPRESJ...');
      const wbCamp = buildMarketWb(df_camp_all, 'CAMPRESJ', moisStr, seuilPlus, seuilMoins);
      zip.file(`Betshop_CAMPRESJ_${moisStr.replace(/ /g,'_')}.xlsx`, XLSX.write(wbCamp, { bookType:'xlsx', type:'array' }));
      log(`  ✅ CAMPRESJ (${df_camp_all.length} salles)`);

      // Marchés standards
      const marchesList = marches.split(',').map(m => m.trim()).filter(Boolean);
      for (const market of marchesList) {
        const df_m = df_all.filter(r => (r['Market']||'') === market);
        if (df_m.length === 0) { log(`  ⚠️ ${market} : 0 salles trouvées`); continue; }
        const wb_m = buildMarketWb(prepDf(df_m), market, moisStr, seuilPlus, seuilMoins);
        zip.file(`Betshop_${market.replace(/ /g,'_')}_${moisStr.replace(/ /g,'_')}.xlsx`, XLSX.write(wb_m, { bookType:'xlsx', type:'array' }));
        log(`  ✅ ${market} (${df_m.length} salles)`);
      }

      // Sous-groupes UNI
      const uniList = uniGroupes.split(',').map(u => u.trim()).filter(Boolean);
      const df_uni = df_all.filter(r => (r['Market']||'') === 'UNI 1l');
      for (const uni of uniList) {
        const df_u = prepDf(df_uni.filter(r => r.Shop.startsWith(uni)));
        if (df_u.length === 0) { log(`  ⚠️ ${uni} : 0 salles trouvées`); continue; }
        const wb_u = buildMarketWb(df_u, uni, moisStr, seuilPlus, seuilMoins);
        zip.file(`Betshop_${uni.replace(/ /g,'_')}_${moisStr.replace(/ /g,'_')}.xlsx`, XLSX.write(wb_u, { bookType:'xlsx', type:'array' }));
        log(`  ✅ ${uni} (${df_u.length} salles)`);
      }

      const blob = await zip.generateAsync({ type:'blob' });
      const filename = `Betshop_Marches_${moisStr.replace(/ /g,'_')}.zip`;
      setResult({ blob, filename });
      log(`\n✅ Terminé !`);
    } catch(err) {
      log(`❌ Erreur : ${err.message}`);
    }
    setRunning(false);
  };

  return (
    <div>
      {/* Uploads */}
      <p className="section-label">Fichiers source</p>
      <div style={{ display:'grid', gridTemplateColumns:'1fr 1fr', gap:'1rem', marginBottom:'1.25rem' }}>
        {[
          { label:'Betshop Cameroon Evaluation.xlsx', sub:'Onglet CAMPRESJ — liste des salles CAMPRESJ', ref:evalRef, setter:setEvalFile, file:evalFile, icon:'📊' },
          { label:'Betshop report Cameroon.xlsx', sub:'Onglet Sheet1 — données complètes par salle', ref:reportRef, setter:setReportFile, file:reportFile, icon:'📋' },
        ].map(({ label, sub, ref, setter, file, icon }) => (
          <div key={label} className={`drop-zone ${file ? 'has-file' : ''}`}
            onClick={() => ref.current.click()}
            onDragOver={e => e.preventDefault()}
            onDrop={e => { e.preventDefault(); setter(e.target.files?.[0] || e.dataTransfer.files[0]); }}>
            <input ref={ref} type="file" accept=".xlsx,.xls" hidden onChange={e => setter(e.target.files[0])} />
            <div className="drop-icon">{icon}</div>
            <div className="drop-label">{label}</div>
            <div className="drop-sub">{sub}</div>
            {file && <div className="file-chip">✓ {file.name}</div>}
          </div>
        ))}
      </div>

      {/* Paramètres */}
      <div className="card" style={{ marginBottom:'1.25rem' }}>
        <p className="section-label">Paramètres</p>
        <div style={{ display:'grid', gridTemplateColumns:'1fr 1fr 1fr', gap:'1rem', marginBottom:'1rem' }}>
          <label style={{ display:'flex', flexDirection:'column', gap:5 }}>
            <span style={{ fontSize:12, color:'var(--muted)', fontWeight:600 }}>Mois</span>
            <input value={mois} onChange={e => setMois(e.target.value)} className="field-input" />
          </label>
          <label style={{ display:'flex', flexDirection:'column', gap:5 }}>
            <span style={{ fontSize:12, color:'var(--muted)', fontWeight:600 }}>Seuil TAPER+ (tickets min)</span>
            <input type="number" value={seuilPlus} onChange={e => setSeuilPlus(Number(e.target.value))} className="field-input" />
          </label>
          <label style={{ display:'flex', flexDirection:'column', gap:5 }}>
            <span style={{ fontSize:12, color:'var(--muted)', fontWeight:600 }}>Seuil TAPER− (tickets min)</span>
            <input type="number" value={seuilMoins} onChange={e => setSeuilMoins(Number(e.target.value))} className="field-input" />
          </label>
        </div>
        <div style={{ display:'grid', gridTemplateColumns:'1fr 1fr', gap:'1rem' }}>
          <label style={{ display:'flex', flexDirection:'column', gap:5 }}>
            <span style={{ fontSize:12, color:'var(--muted)', fontWeight:600 }}>Marchés standard (séparés par virgule)</span>
            <input value={marches} onChange={e => setMarches(e.target.value)} className="field-input" style={{ width:'100%' }} />
          </label>
          <label style={{ display:'flex', flexDirection:'column', gap:5 }}>
            <span style={{ fontSize:12, color:'var(--muted)', fontWeight:600 }}>Sous-groupes UNI (séparés par virgule)</span>
            <input value={uniGroupes} onChange={e => setUniGroupes(e.target.value)} className="field-input" style={{ width:'100%' }} />
          </label>
        </div>
      </div>

      <button className="btn primary" disabled={!evalFile || !reportFile || running}
        onClick={runPipeline} style={{ fontSize:14, padding:'11px 28px', marginBottom:'1.5rem' }}>
        {running ? '⟳ Génération...' : '▶  Générer les fichiers par marché'}
      </button>

      {/* Logs */}
      {logs.length > 0 && (
        <>
          <p className="section-label">Console</p>
          <div className="log-console" style={{ marginBottom:'1.5rem' }}>
            {logs.map((l, i) => <span key={i} className="log-line" style={{ display:'block' }}>{l}</span>)}
          </div>
        </>
      )}

      {/* Result */}
      {result && (
        <div className="card">
          <button className="btn success" onClick={() => saveAs(result.blob, result.filename)}
            style={{ fontSize:14, padding:'10px 24px' }}>
            ⬇  Télécharger {result.filename}
          </button>
        </div>
      )}
    </div>
  );
}
