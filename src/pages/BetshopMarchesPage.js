// src/pages/BetshopMarchesPage.js
import React, { useState, useRef } from 'react';
import XLSX from 'xlsx-js-style';
import { saveAs } from 'file-saver';
import JSZip from 'jszip';
import { useAppContext } from '../AppContext';

// ── Palette ──────────────────────────────────────
const C = {
  dark_blue:'1F3864', mid_blue:'2E75B6', green_med:'375623', green_bg:'E2EFDA',
  green_fg:'375623', red_dark:'5C0A0A', red_med:'7B2C2C', red_bg:'FCE4D6',
  red_fg:'C55A11', gold_bg:'FFF2CC', gold_fg:'7F6000', gold_dark:'BF8F00',
  grey:'F2F2F2', white:'FFFFFF',
};
const FMT_NB='#,##0', FMT_DIF='#,##0;[Red]-#,##0', FMT_PCT='0.0"%"';
const SEUIL_ALERTE_PCT = 20;

function bdr(s='thin') {
  const col = s==='thin' ? 'BFBFBF' : '404040';
  const t = { style:s, color:{ rgb:col } };
  return { top:t, bottom:t, left:t, right:t };
}
function fill(rgb) { return { patternType:'solid', fgColor:{ rgb } }; }
function fnt(bold=false,sz=10,rgb='000000',italic=false) {
  return { name:'Arial', bold, sz, color:{ rgb }, italic };
}
function aln(h='center',v='center',wrap=false) {
  return { horizontal:h, vertical:v, wrapText:wrap };
}
function sc(ws,r,c,v,style) {
  const addr = XLSX.utils.encode_cell({ r:r-1, c:c-1 });
  ws[addr] = { v:v===null||v===undefined?'':v, t:typeof v==='number'?'n':'s', s:style };
}
function scn(ws,r,c,v,style,fmt) {
  const addr = XLSX.utils.encode_cell({ r:r-1, c:c-1 });
  ws[addr] = { v:typeof v==='number'?v:0, t:'n', s:style, z:fmt };
}
function mg(ws,r1,c1,r2,c2) {
  if (!ws['!merges']) ws['!merges'] = [];
  ws['!merges'].push({ s:{r:r1-1,c:c1-1}, e:{r:r2-1,c:c2-1} });
}
function hS(bg,fg='FFFFFF',sz=10) {
  return { font:fnt(true,sz,fg), fill:fill(bg), border:bdr(), alignment:aln('center','center',true) };
}
function dS(bg='FFFFFF',bold=false,fg='000000',h='right') {
  return { font:fnt(bold,10,fg), fill:fill(bg), border:bdr(), alignment:aln(h,'center') };
}
function tS(bg,sz=12) {
  return { font:fnt(true,sz,'FFFFFF'), fill:fill(bg), border:bdr('medium'), alignment:aln('center','center') };
}
function sS(bg) {
  return { font:fnt(true,11,'FFFFFF'), fill:fill(bg), border:bdr('medium'), alignment:aln('left','center') };
}
function mS(bg,fg,italic=false,sz=9) {
  return { font:fnt(false,sz,fg,italic), fill:fill(bg), border:bdr(), alignment:aln('left','center',true) };
}

function colorDelta(v, inv=false) {
  const good = inv ? v < 0 : v > 0;
  return good ? C.green_fg : C.red_fg;
}

// ── Prep dataframe ────────────────────────────────
function prepDf(rows) {
  return rows.map((row) => {
    const n=(k)=>parseFloat(row[k])||0;
    const mtd=n('Bet tickets MTD'), lm=n('Bet tickets LM'), ly=n('Bet tickets LY');
    const lm15=n('Bet Ticket LM +15%'), ly15=n('Bet ticket LY +15 %');
    const pmtd=n('Payin Bet MTD'), plm=n('Payin Bet LM'), ply=n('Payin Bet LY');
    const gmtd=n('GGR Bet MTD'), glm=n('GGR Bet LM');
    return {
      ...row,
      Shop: row['Shop']||row['Betshop name']||'',
      'Bet tickets MTD':mtd,'Bet tickets LM':lm,'Bet tickets LY':ly,
      'Bet Ticket LM +15%':lm15,'Bet ticket LY +15 %':ly15,
      'Payin Bet MTD':pmtd,'Payin Bet LM':plm,'Payin Bet LY':ply,
      'GGR Bet MTD':gmtd,'GGR Bet LM':glm,
      Diff_T_LM:mtd-lm, Diff_T_LY:mtd-ly,
      Diff_T_LM15:mtd-lm15, Diff_T_LY15:mtd-ly15,
      Pct_T_LM:lm!==0?(mtd-lm)/lm*100:null,
      Pct_T_LY:ly!==0?(mtd-ly)/ly*100:null,
      Diff_P_LM:pmtd-plm, Diff_P_LY:pmtd-ply,
      Pct_P_LM:plm!==0?(pmtd-plm)/plm*100:null,
      Diff_G_LM:gmtd-glm,
      Pct_G_LM:glm!==0?(gmtd-glm)/Math.abs(glm)*100:null,
      Statut:mtd===0?'⛔ INACTIF':'✅ Actif',
    };
  });
}

// ── Sheet: Données ────────────────────────────────
function buildOngData(wb, sheetName, df, titre, mois) {
  const ws={};
  const nr=df.length;
  sc(ws,2,1,`${titre} | ${mois}  ·  ${df.length} salles`,tS(C.dark_blue,13));
  mg(ws,2,1,2,21);

  const ghRow=4;
  [[1,1,'SALLE',C.dark_blue],[2,10,'🎫 TICKETS','1F4E79'],[11,16,'💰 PAYIN',C.green_med],[17,21,'📈 GGR',C.red_med]].forEach(([cs,ce,val,bg])=>{
    sc(ws,ghRow,cs,val,hS(bg,'FFFFFF',10));
    if(cs!==ce) mg(ws,ghRow,cs,ghRow,ce);
  });

  const subH=['Salle','MTD','LM','LY','Diff LM','% LM','Diff LY','% LY','LM+15%','Diff LM+15%','MTD','LM','LY','Diff LM','% LM','Diff LY','MTD','LM','LY','Diff LM','% LM'];
  const subBg=['1F3864','1F4E79','1F4E79','1F4E79','2E5F8A','2E5F8A','2E5F8A','2E5F8A','1F4E79','2E5F8A',C.green_med,C.green_med,C.green_med,'4A7C3F','4A7C3F','4A7C3F',C.red_med,C.red_med,C.red_med,'A04040','A04040'];
  subH.forEach((h,i)=>sc(ws,5,i+1,h,hS(subBg[i],'FFFFFF',9)));

  df.forEach((row,i)=>{
    const dr=6+i, bg=row.Statut==='⛔ INACTIF'?'EFEFEF':i%2===0?'EBF3FB':C.white;
    const fgN=row.Statut==='⛔ INACTIF'?'7F7F7F':'000000';
    sc(ws,dr,1,row.Shop,{font:fnt(false,10,fgN),fill:fill(bg),border:bdr(),alignment:aln('left','center')});
    const vals=[
      [2,row['Bet tickets MTD'],FMT_NB,false],
      [3,row['Bet tickets LM'],FMT_NB,false],
      [4,row['Bet tickets LY'],FMT_NB,false],
      [5,row.Diff_T_LM,FMT_DIF,true],
      [6,row.Pct_T_LM||0,FMT_PCT,true],
      [7,row.Diff_T_LY,FMT_DIF,true],
      [8,row.Pct_T_LY||0,FMT_PCT,true],
      [9,row['Bet Ticket LM +15%'],FMT_NB,false],
      [10,row.Diff_T_LM15,FMT_DIF,true],
      [11,row['Payin Bet MTD'],FMT_NB,false],
      [12,row['Payin Bet LM'],FMT_NB,false],
      [13,row['Payin Bet LY'],FMT_NB,false],
      [14,row.Diff_P_LM,FMT_DIF,true],
      [15,row.Pct_P_LM||0,FMT_PCT,true],
      [16,row.Diff_P_LY,FMT_DIF,true],
      [17,row['GGR Bet MTD'],FMT_NB,false],
      [18,row['GGR Bet LM'],FMT_NB,false],
      [19,row['GGR Bet LY']||0,FMT_NB,false],
      [20,row.Diff_G_LM,FMT_DIF,true],
      [21,row.Pct_G_LM||0,FMT_PCT,true],
    ];
    vals.forEach(([c,v,fmt,delta])=>{
      const fg=delta?(v>=0?C.green_fg:C.red_fg):'000000';
      scn(ws,dr,c,v,{font:fnt(delta,10,fg),fill:fill(bg),border:bdr(),alignment:aln('right','center')},fmt);
    });
  });

  // Total
  const tr=6+nr;
  const sum=k=>df.reduce((s,r)=>s+(parseFloat(r[k])||0),0);
  const tSt={font:fnt(true,10,'FFFFFF'),fill:fill(C.dark_blue),border:bdr(),alignment:aln('right','center')};
  sc(ws,tr,1,`TOTAL (${nr})`,{...tSt,alignment:aln('left','center')});
  [[2,sum('Bet tickets MTD'),FMT_NB],[3,sum('Bet tickets LM'),FMT_NB],[4,sum('Bet tickets LY'),FMT_NB],
   [5,sum('Diff_T_LM'),FMT_DIF],[11,sum('Payin Bet MTD'),FMT_NB],[14,sum('Diff_P_LM'),FMT_DIF],
   [17,sum('GGR Bet MTD'),FMT_NB],[20,sum('Diff_G_LM'),FMT_DIF]].forEach(([c,v,fmt])=>{
    scn(ws,tr,c,v,tSt,fmt);
  });

  ws['!cols']=[32,14,14,14,13,10,13,10,14,14,14,14,14,13,10,13,14,14,14,13,10].map(w=>({wch:w}));
  ws['!ref']=XLSX.utils.encode_range({s:{r:0,c:0},e:{r:tr-1,c:20}});
  XLSX.utils.book_append_sheet(wb,ws,sheetName.slice(0,31));
}

// ── Sheet: Résumé ─────────────────────────────────
function buildOngResume(wb, sheetName, df, titre, mois) {
  const ws={};
  const sum=k=>df.reduce((s,r)=>s+(parseFloat(r[k])||0),0);
  const nb_actif=df.filter(r=>r.Statut==='✅ Actif').length;
  const prog=df.filter(r=>r.Diff_T_LM>0).length;
  const recul=df.filter(r=>r.Diff_T_LM<0).length;

  sc(ws,2,1,`${titre} — RÉSUMÉ | ${mois}`,tS(C.dark_blue,13));
  mg(ws,2,1,2,6);
  sc(ws,4,1,'📊 INDICATEURS CLÉS',sS(C.mid_blue));
  mg(ws,4,1,4,6);

  const hdrs=['Indicateur','Valeur MTD','Valeur LM','Variation','Var %','Tendance'];
  hdrs.forEach((h,i)=>sc(ws,5,i+1,h,hS(C.mid_blue)));

  const kpis=[
    ['Salles totales',df.length,null],
    ['Salles actives',nb_actif,null],
    ['En progression vs LM',prog,null],
    ['En recul vs LM',recul,null],
    ['Tickets MTD',sum('Bet tickets MTD'),sum('Bet tickets LM')],
    ['Payin MTD (FCFA)',sum('Payin Bet MTD'),sum('Payin Bet LM')],
    ['GGR MTD (FCFA)',sum('GGR Bet MTD'),sum('GGR Bet LM')],
  ];
  kpis.forEach(([label,val,lm],i)=>{
    const r=6+i, bg=i%2===0?C.grey:C.white;
    sc(ws,r,1,label,dS(bg,true,'000000','left'));
    scn(ws,r,2,val,dS(bg),FMT_NB);
    if(lm!==null){
      const diff=val-lm, pct=lm!==0?diff/lm*100:0, isGood=diff>=0;
      const vS={font:fnt(true,10,isGood?C.green_fg:C.red_fg),fill:fill(isGood?C.green_bg:C.red_bg),border:bdr(),alignment:aln('right','center')};
      scn(ws,r,3,lm,dS(bg),FMT_NB);
      scn(ws,r,4,diff,vS,FMT_DIF);
      scn(ws,r,5,pct,vS,FMT_PCT);
      sc(ws,r,6,`${diff>=0?'▲':'▼'} ${pct>0?'+':''}${pct.toFixed(1)}%`,{...vS,alignment:aln('center','center')});
    } else { [3,4,5,6].forEach(c=>sc(ws,r,c,'',dS(bg))); }
  });

  // Top/Bottom
  const sorted_prog=[...df].sort((a,b)=>b.Diff_T_LM-a.Diff_T_LM).slice(0,10);
  const sorted_recul=[...df].sort((a,b)=>a.Diff_T_LM-b.Diff_T_LM).slice(0,10);
  const nb=Math.min(10,df.length);
  let r=14;
  sc(ws,r,1,`🏆 TOP ${nb} PROGRESSIONS`,sS(C.green_med)); mg(ws,r,1,r,3);
  sc(ws,r,4,`⚠️ TOP ${nb} RECULS`,sS(C.red_med)); mg(ws,r,4,r,6); r++;
  ['Salle','Tickets MTD','Diff LM','Salle','Tickets MTD','Diff LM'].forEach((h,i)=>sc(ws,r,i+1,h,hS(i<3?'4A7C3F':'A04040','FFFFFF',9)));
  for(let i=0;i<nb;i++){
    r++; const rp=sorted_prog[i], rr=sorted_recul[i];
    const bgp=i%2===0?C.green_bg:C.white, bgr=i%2===0?C.red_bg:C.white;
    sc(ws,r,1,rp?.Shop||'',dS(bgp,false,'000000','left'));
    scn(ws,r,2,rp?.['Bet tickets MTD']||0,dS(bgp),FMT_NB);
    scn(ws,r,3,rp?.Diff_T_LM||0,{font:fnt(true,10,C.green_fg),fill:fill(bgp),border:bdr(),alignment:aln('right','center')},FMT_DIF);
    sc(ws,r,4,rr?.Shop||'',dS(bgr,false,'000000','left'));
    scn(ws,r,5,rr?.['Bet tickets MTD']||0,dS(bgr),FMT_NB);
    scn(ws,r,6,rr?.Diff_T_LM||0,{font:fnt(true,10,C.red_fg),fill:fill(bgr),border:bdr(),alignment:aln('right','center')},FMT_DIF);
  }

  ws['!cols']=[32,18,18,16,16,18].map(w=>({wch:w}));
  ws['!ref']=XLSX.utils.encode_range({s:{r:0,c:0},e:{r:r,c:5}});
  XLSX.utils.book_append_sheet(wb,ws,sheetName.slice(0,31));
}

// ── Sheet: Analyse automatique ────────────────────
function buildOngAnalyse(wb, sheetName, df, label, mois) {
  const ws={};
  const sum=k=>df.reduce((s,r)=>s+(parseFloat(r[k])||0),0);
  const nb_total=df.length;
  const nb_actif=df.filter(r=>r.Statut==='✅ Actif').length;
  const nb_inactif=nb_total-nb_actif;
  const df_actif=df.filter(r=>r.Statut==='✅ Actif');
  const t_mtd=sum('Bet tickets MTD'), t_lm=sum('Bet tickets LM'), t_ly=sum('Bet tickets LY');
  const diff_lm=t_mtd-t_lm, diff_ly=t_mtd-t_ly;
  const pct_lm=t_lm!==0?diff_lm/t_lm*100:0, pct_ly=t_ly!==0?diff_ly/t_ly*100:0;
  const ggr_mtd=sum('GGR Bet MTD'), ggr_lm=sum('GGR Bet LM');
  const diff_ggr=ggr_mtd-ggr_lm, pct_ggr=ggr_lm!==0?diff_ggr/Math.abs(ggr_lm)*100:0;
  const nb_prog=df.filter(r=>r.Diff_T_LM>0).length;
  const nb_recul=df.filter(r=>r.Diff_T_LM<0).length;
  const top3=[...df_actif].sort((a,b)=>b.Diff_T_LM-a.Diff_T_LM).slice(0,3);
  const bot3=[...df_actif].sort((a,b)=>a.Diff_T_LM-b.Diff_T_LM).slice(0,3);
  const alerte_lm=df_actif.filter(r=>(r.Pct_T_LM||0)<-SEUIL_ALERTE_PCT).sort((a,b)=>a.Pct_T_LM-b.Pct_T_LM);
  const alerte_ly=df_actif.filter(r=>(r.Pct_T_LY||0)<-SEUIL_ALERTE_PCT).sort((a,b)=>a.Pct_T_LY-b.Pct_T_LY);

  const arrow=v=>v>=0?'▲':'▼';
  const signe=v=>v>=0?`+${v.toFixed(1)}%`:`${v.toFixed(1)}%`;
  const qual=v=>v>=15?'excellente performance':v>=5?'bonne progression':v>=0?'légère progression':v>=-5?'légère baisse':v>=-15?'baisse notable':'forte baisse préoccupante';

  sc(ws,2,1,`ANALYSE AUTOMATIQUE  ·  ${label}  ·  ${mois}`,tS(C.dark_blue,13));
  mg(ws,2,1,2,7);
  sc(ws,3,1,`⚙️  Seuil alerte baisse : >${SEUIL_ALERTE_PCT}% vs LM ou LY  |  Généré automatiquement`,
    mS(C.gold_bg,C.gold_fg,true,9));
  mg(ws,3,1,3,7);

  let row=5;
  const writeSection=(titre, bgTitle, items)=>{
    row++;
    sc(ws,row,2,titre,sS(bgTitle)); mg(ws,row,2,row,7); row++;
    items.forEach(([emoji,texte,detail,bgLig,bold])=>{
      sc(ws,row,2,emoji,{font:fnt(false,12),fill:fill(bgLig),border:bdr(),alignment:aln('center','center')});
      sc(ws,row,3,texte,{font:fnt(bold,10,'000000'),fill:fill(bgLig),border:bdr(),alignment:aln('left','center',true)});
      mg(ws,row,3,row,5);
      sc(ws,row,6,`➜  ${detail}`,{font:fnt(false,9,'2E2E2E',true),fill:fill(bgLig),border:bdr(),alignment:aln('left','center',true)});
      mg(ws,row,6,row,7);
      row++;
    });
  };

  const bg_perf=pct_lm>=0?C.dark_blue:C.red_med;
  writeSection('📊  PERFORMANCE GLOBALE DU MARCHÉ', bg_perf, [
    [pct_lm>=0?'📈':'📉',
     `Le marché ${label} affiche une ${qual(pct_lm)} sur les tickets : ${t_mtd.toLocaleString()} MTD vs ${t_lm.toLocaleString()} LM (${arrow(pct_lm)} ${signe(pct_lm)}).`,
     pct_lm>=0?'Maintenir la dynamique et identifier les leviers de croissance.':'Analyser les causes de baisse et renforcer l\'animation commerciale.',
     pct_lm>=0?C.green_bg:C.red_bg, true],
    ['📅',
     `Vs année précédente : ${t_mtd.toLocaleString()} MTD vs ${t_ly.toLocaleString()} LY (${arrow(pct_ly)} ${signe(pct_ly)}) — tendance annuelle ${pct_ly>=0?'favorable':'défavorable'}.`,
     pct_ly>=0?'La tendance annuelle est solide — capitaliser.':'Tendance annuelle défavorable — un plan de redressement est recommandé.',
     pct_ly>=0?C.green_bg:C.red_bg, false],
    ['💰',
     `GGR : ${ggr_mtd.toLocaleString()} FCFA MTD vs ${ggr_lm.toLocaleString()} FCFA LM (${arrow(pct_ggr)} ${signe(pct_ggr)}).`,
     diff_ggr>=0?'GGR en hausse — surveiller le ratio GGR/Payin.':'GGR en baisse — revoir les conditions de jeu.',
     diff_ggr>=0?C.green_bg:C.red_bg, false],
    ['🏪',
     `${nb_actif} salle${nb_actif>1?'s':''} active${nb_actif>1?'s':''} sur ${nb_total} — ${nb_prog} en progression, ${nb_recul} en recul vs LM.${nb_inactif>0?` ${nb_inactif} inactive${nb_inactif>1?'s':''}.`:''}`,
     nb_inactif>0?`Investiguer les ${nb_inactif} salle${nb_inactif>1?'s':''} inactive${nb_inactif>1?'s':''}.`:'Toutes les salles sont actives ce mois.',
     C.grey, false],
  ]);

  writeSection('🏆  SALLES EN TÊTE — Meilleures progressions vs LM', C.green_med,
    top3.map((r,i)=>['',
      `${'🥇🥈🥉'[i]}  ${r.Shop} — ${r['Bet tickets MTD'].toLocaleString()} tickets MTD (${arrow(r.Diff_T_LM)} ${r.Diff_T_LM>0?'+':''}${Math.round(r.Diff_T_LM).toLocaleString()}, ${signe(r.Pct_T_LM||0)} vs LM)`,
      r.Pct_T_LM>=15?'Performance excellente — à valoriser comme modèle.':'Bonne progression — encourager et maintenir l\'effort.',
      [C.gold_bg,C.green_bg,'EBF5EB'][i], i===0])
  );

  writeSection('⚠️  SALLES EN QUEUE — Plus fortes baisses vs LM', C.red_med,
    bot3.map((r,i)=>['',
      `${'🔴🟠🟡'[i]}  ${r.Shop} — ${r['Bet tickets MTD'].toLocaleString()} tickets MTD (${arrow(r.Diff_T_LM)} ${r.Diff_T_LM>0?'+':''}${Math.round(r.Diff_T_LM).toLocaleString()}, ${signe(r.Pct_T_LM||0)} vs LM)`,
      Math.abs(r.Pct_T_LM||0)>SEUIL_ALERTE_PCT?`Baisse de ${Math.abs(r.Pct_T_LM||0).toFixed(1)}% — visite terrain et plan d'action immédiat requis.`:`Baisse de ${Math.abs(r.Pct_T_LM||0).toFixed(1)}% — à surveiller.`,
      [C.red_bg,'FCE4D6',C.gold_bg][i], i===0])
  );

  const alertItems=[];
  if(pct_lm<0) alertItems.push(['🚨',`ALERTE MARCHÉ : ${label} est globalement en baisse de ${Math.abs(pct_lm).toFixed(1)}% vs LM (${diff_lm>0?'+':''}${Math.round(diff_lm).toLocaleString()} tickets).`,'Action prioritaire : réunion de pilotage et revue des salles sous-performantes.','FFC7CE',true]);
  alerte_lm.forEach(r=>alertItems.push(['📛',`${r.Shop} : baisse de ${Math.abs(r.Pct_T_LM||0).toFixed(1)}% vs LM (${Math.round(r.Diff_T_LM)>0?'+':''}${Math.round(r.Diff_T_LM).toLocaleString()} tickets — MTD : ${r['Bet tickets MTD'].toLocaleString()}).`,'Visite terrain urgente — identifier la cause.',C.red_bg,false]));
  const shopsLm=new Set(alerte_lm.map(r=>r.Shop));
  alerte_ly.filter(r=>!shopsLm.has(r.Shop)).forEach(r=>alertItems.push(['⚡',`${r.Shop} : baisse de ${Math.abs(r.Pct_T_LY||0).toFixed(1)}% vs LY — MTD : ${r['Bet tickets MTD'].toLocaleString()}.`,'Tendance annuelle défavorable — revoir la stratégie.',C.gold_bg,false]));
  if(!alertItems.length) alertItems.push(['✅',`Aucune alerte critique ce mois-ci sur le marché ${label}.`,'Situation sous contrôle — maintenir le suivi mensuel.',C.green_bg,false]);
  writeSection(`🚨  ALERTES — Baisses > ${SEUIL_ALERTE_PCT}% vs LM ou LY`,'7B2C2C',alertItems);

  ws['!cols']=[4,36,55,18,18,14,14].map(w=>({wch:w}));
  ws['!ref']=XLSX.utils.encode_range({s:{r:0,c:0},e:{r:row,c:6}});
  XLSX.utils.book_append_sheet(wb,ws,sheetName.slice(0,31));
}

// ── Sheet: TAPER+ ─────────────────────────────────
function buildTaperPlus(wb, sheetName, df, seuil, titre, mois) {
  const ws={};
  sc(ws,2,2,`TAPER +  ·  ${titre}  ·  ${mois}`,tS(C.dark_blue,12)); mg(ws,2,2,2,8);
  sc(ws,3,2,`⚙️ Seuil minimum : ${seuil.toLocaleString()} tickets MTD`,mS(C.gold_bg,C.gold_fg,false,9)); mg(ws,3,2,3,8);

  const dp=df.filter(r=>r['Bet tickets MTD']>=seuil);
  const blocs=[
    {titre:'Plus que LM',emoji:'✅',rows:dp.filter(r=>r.Diff_T_LM>0).sort((a,b)=>b.Diff_T_LM-a.Diff_T_LM),bg:C.green_med,bgh:'4A7C3F',bge:C.green_bg,ref:'LM',rk:'Bet tickets LM',dk:'Diff_T_LM',pk:'Pct_T_LM'},
    {titre:'Plus que LY',emoji:'📅',rows:dp.filter(r=>r.Diff_T_LY>0).sort((a,b)=>b.Diff_T_LY-a.Diff_T_LY),bg:'1F6B38',bgh:'2E8B57',bge:'EBF5EB',ref:'LY',rk:'Bet tickets LY',dk:'Diff_T_LY',pk:'Pct_T_LY'},
    {titre:'Dépassé LM+15%',emoji:'🏆',rows:dp.filter(r=>r.Diff_T_LM15>0).sort((a,b)=>b.Diff_T_LM15-a.Diff_T_LM15),bg:C.gold_dark,bgh:'A07800',bge:C.gold_bg,ref:'LM+15%',rk:'Bet Ticket LM +15%',dk:'Diff_T_LM15',pk:'Pct_T_LM'},
  ];

  let row=5;
  blocs.forEach(bloc=>{
    sc(ws,row,2,`${bloc.emoji}  ${bloc.titre}  —  ${bloc.rows.length} salle${bloc.rows.length!==1?'s':''}`,sS(bloc.bg)); mg(ws,row,2,row,8); row++;
    sc(ws,row,2,`Seuil ${seuil.toLocaleString()} tickets min`,mS(bloc.bgh,'FFFFFF',true,9)); mg(ws,row,2,row,8); row++;
    ['#','Salle','Tickets MTD',bloc.ref,'Diff','%','GGR MTD'].forEach((h,i)=>sc(ws,row,i+2,h,hS(bloc.bgh,'FFFFFF',9))); row++;
    bloc.rows.forEach((r,i)=>{
      const bg=i%2===0?bloc.bge:C.white;
      sc(ws,row,2,i+1,{font:fnt(true,10,'FFFFFF'),fill:fill(bloc.bg),border:bdr(),alignment:aln('center','center')});
      sc(ws,row,3,r.Shop,dS(bg,false,'000000','left'));
      scn(ws,row,4,r['Bet tickets MTD'],dS(bg),FMT_NB);
      scn(ws,row,5,r[bloc.rk],dS(bg),FMT_NB);
      scn(ws,row,6,r[bloc.dk],{font:fnt(true,10,C.green_fg),fill:fill(bg),border:bdr(),alignment:aln('right','center')},FMT_DIF);
      scn(ws,row,7,r[bloc.pk]||0,{font:fnt(true,10,C.green_fg),fill:fill(bg),border:bdr(),alignment:aln('right','center')},FMT_PCT);
      scn(ws,row,8,r['GGR Bet MTD'],dS(bg),FMT_NB);
      row++;
    });
    if(bloc.rows.length>0){
      const tSt2={font:fnt(true,10,'FFFFFF'),fill:fill(bloc.bg),border:bdr(),alignment:aln('right','center')};
      sc(ws,row,2,`TOTAL (${bloc.rows.length})`,{...tSt2,alignment:aln('left','center')});
      scn(ws,row,4,bloc.rows.reduce((s,r)=>s+r['Bet tickets MTD'],0),tSt2,FMT_NB);
      scn(ws,row,6,bloc.rows.reduce((s,r)=>s+r[bloc.dk],0),tSt2,FMT_DIF);
      scn(ws,row,8,bloc.rows.reduce((s,r)=>s+r['GGR Bet MTD'],0),tSt2,FMT_NB);
      row++;
    }
    row++;
  });

  ws['!cols']=[3,5,30,13,13,14,10,14].map(w=>({wch:w}));
  ws['!ref']=XLSX.utils.encode_range({s:{r:0,c:0},e:{r:row,c:8}});
  XLSX.utils.book_append_sheet(wb,ws,sheetName.slice(0,31));
}

// ── Sheet: TAPER- ─────────────────────────────────
function buildTaperMoins(wb, sheetName, df, seuil, titre, mois) {
  const ws={};
  sc(ws,2,2,`TAPER −  ·  ${titre}  ·  ${mois}`,tS(C.dark_blue,12)); mg(ws,2,2,2,8);

  const dm=df.filter(r=>r['Bet tickets MTD']>=seuil);
  const blocs=[
    {titre:'Moins que LM',emoji:'⚠️',rows:dm.filter(r=>r.Diff_T_LM<0).sort((a,b)=>a.Diff_T_LM-b.Diff_T_LM),bg:C.red_med,bgh:'A04040',bge:C.red_bg,dk:'Diff_T_LM',pk:'Pct_T_LM'},
    {titre:'Moins que LY',emoji:'📉',rows:dm.filter(r=>r.Diff_T_LY<0).sort((a,b)=>a.Diff_T_LY-b.Diff_T_LY),bg:'8B1A1A',bgh:'B03030',bge:'FCE4D6',dk:'Diff_T_LY',pk:'Pct_T_LY'},
    {titre:"N'a pas atteint LM+15%",emoji:'🔴',rows:dm.filter(r=>r.Diff_T_LM15<0).sort((a,b)=>a.Diff_T_LM15-b.Diff_T_LM15),bg:C.red_dark,bgh:'8B1A1A',bge:'FCE4D6',dk:'Diff_T_LM15',pk:'Pct_T_LM'},
  ];

  let row=5;
  blocs.forEach(bloc=>{
    sc(ws,row,2,`${bloc.emoji}  ${bloc.titre}  —  ${bloc.rows.length} salle${bloc.rows.length!==1?'s':''}`,sS(bloc.bg)); mg(ws,row,2,row,8); row++;
    sc(ws,row,2,`Seuil ${seuil.toLocaleString()} tickets min`,mS(bloc.bgh,'FFFFFF',true,9)); mg(ws,row,2,row,8); row++;
    ['#','Salle','Tickets MTD','Référence','Diff','%','GGR MTD'].forEach((h,i)=>sc(ws,row,i+2,h,hS(bloc.bgh,'FFFFFF',9))); row++;
    bloc.rows.forEach((r,i)=>{
      const bg=i%2===0?bloc.bge:C.white;
      sc(ws,row,2,i+1,{font:fnt(true,10,'FFFFFF'),fill:fill(bloc.bg),border:bdr(),alignment:aln('center','center')});
      sc(ws,row,3,r.Shop,dS(bg,false,'000000','left'));
      scn(ws,row,4,r['Bet tickets MTD'],dS(bg),FMT_NB);
      scn(ws,row,5,0,dS(bg),FMT_NB);
      scn(ws,row,6,r[bloc.dk],{font:fnt(true,10,C.red_fg),fill:fill(bg),border:bdr(),alignment:aln('right','center')},FMT_DIF);
      scn(ws,row,7,r[bloc.pk]||0,{font:fnt(true,10,C.red_fg),fill:fill(bg),border:bdr(),alignment:aln('right','center')},FMT_PCT);
      scn(ws,row,8,r['GGR Bet MTD'],dS(bg),FMT_NB);
      row++;
    });
    row++;
  });

  ws['!cols']=[3,5,30,13,13,14,10,14].map(w=>({wch:w}));
  ws['!ref']=XLSX.utils.encode_range({s:{r:0,c:0},e:{r:row,c:8}});
  XLSX.utils.book_append_sheet(wb,ws,sheetName.slice(0,31));
}

// ── Sheet: Intersections ──────────────────────────
function buildIntersections(wb, sheetName, df, seuilP, seuilM, titre, mois) {
  const ws={};
  sc(ws,2,2,`INTERSECTIONS LM ∩ LY  ·  ${titre}  ·  ${mois}`,tS(C.dark_blue,12)); mg(ws,2,2,2,10);
  sc(ws,3,2,`⚙️ Seuil TAPER+ : ${seuilP.toLocaleString()}  |  Seuil TAPER− : ${seuilM.toLocaleString()}  |  ℹ️ Salles présentes dans les DEUX critères simultanément`,mS(C.gold_bg,C.gold_fg,true,9)); mg(ws,3,2,3,10);

  const dp=df.filter(r=>r['Bet tickets MTD']>=seuilP);
  const dm=df.filter(r=>r['Bet tickets MTD']>=seuilM);
  const sp_lm=new Set(dp.filter(r=>r.Diff_T_LM>0).map(r=>r.Shop));
  const sp_ly=new Set(dp.filter(r=>r.Diff_T_LY>0).map(r=>r.Shop));
  const sp_lm15=new Set(dp.filter(r=>r.Diff_T_LM15>0).map(r=>r.Shop));
  const sp_ly15=new Set(dp.filter(r=>r.Diff_T_LY15>0).map(r=>r.Shop));
  const sm_lm=new Set(dm.filter(r=>r.Diff_T_LM<0).map(r=>r.Shop));
  const sm_ly=new Set(dm.filter(r=>r.Diff_T_LY<0).map(r=>r.Shop));
  const sm_lm15=new Set(dm.filter(r=>r.Diff_T_LM15<0).map(r=>r.Shop));
  const sm_ly15=new Set(dm.filter(r=>r.Diff_T_LY15<0).map(r=>r.Shop));

  const inter=(base,s1,s2,sortK,asc=false)=>{
    const result=base.filter(r=>s1.has(r.Shop)&&s2.has(r.Shop));
    return asc?result.sort((a,b)=>a[sortK]-b[sortK]):result.sort((a,b)=>b[sortK]-a[sortK]);
  };

  let row=5;
  sc(ws,row,2,'▲  BLOCS POSITIFS',sS(C.green_med)); mg(ws,row,2,row,10); row+=2;

  const ibloc=(titre2,emoji,rows2,bgt,bgh,bge,note,inv=false)=>{
    sc(ws,row,2,`${emoji}  ${titre2}  —  ${rows2.length} salle${rows2.length!==1?'s':''}`,sS(bgt)); mg(ws,row,2,row,10); row++;
    sc(ws,row,2,note,mS(bgh,'FFFFFF',true,9)); mg(ws,row,2,row,10); row++;
    if(rows2.length===0){
      sc(ws,row,2,'Aucune salle ne satisfait les deux critères.',{font:fnt(false,10,'7F7F7F',true),fill:fill(C.grey),border:bdr(),alignment:aln('center','center')});
      mg(ws,row,2,row,10); row+=2; return;
    }
    // Headers
    [[2,''],[3,''],[4,''],[5,'▶ vs LM'],[8,'▶ vs LY']].forEach(([c,v])=>{if(v){sc(ws,row,c,v,hS(bgh,'FFFFFF',9)); mg(ws,row,c,row,c+2);}});
    row++;
    ['#','Salle','MTD','LM','Diff LM','% LM','LY','Diff LY','% LY'].forEach((h,i)=>sc(ws,row,i+2,h,hS(bgh,'FFFFFF',9))); row++;
    rows2.forEach((r,i)=>{
      const bg=i%2===0?bge:C.white;
      sc(ws,row,2,i+1,{font:fnt(true,10,'FFFFFF'),fill:fill(bgt),border:bdr(),alignment:aln('center','center')});
      sc(ws,row,3,r.Shop,dS(bg,false,'000000','left'));
      scn(ws,row,4,r['Bet tickets MTD'],dS(bg),FMT_NB);
      scn(ws,row,5,r['Bet tickets LM'],dS(bg),FMT_NB);
      const fgLM=colorDelta(r.Diff_T_LM,inv);
      scn(ws,row,6,r.Diff_T_LM,{font:fnt(true,10,fgLM),fill:fill(bg),border:bdr(),alignment:aln('right','center')},FMT_DIF);
      scn(ws,row,7,r.Pct_T_LM||0,{font:fnt(true,10,fgLM),fill:fill(bg),border:bdr(),alignment:aln('right','center')},FMT_PCT);
      scn(ws,row,8,r['Bet tickets LY'],dS(bg),FMT_NB);
      const fgLY=colorDelta(r.Diff_T_LY,inv);
      scn(ws,row,9,r.Diff_T_LY,{font:fnt(true,10,fgLY),fill:fill(bg),border:bdr(),alignment:aln('right','center')},FMT_DIF);
      scn(ws,row,10,r.Pct_T_LY||0,{font:fnt(true,10,fgLY),fill:fill(bg),border:bdr(),alignment:aln('right','center')},FMT_PCT);
      row++;
    });
    // Total
    const tSt2={font:fnt(true,10,'FFFFFF'),fill:fill(bgt),border:bdr(),alignment:aln('right','center')};
    sc(ws,row,2,`TOTAL (${rows2.length})`,{...tSt2,alignment:aln('left','center')});
    scn(ws,row,4,rows2.reduce((s,r)=>s+r['Bet tickets MTD'],0),tSt2,FMT_NB);
    scn(ws,row,5,rows2.reduce((s,r)=>s+r['Bet tickets LM'],0),tSt2,FMT_NB);
    scn(ws,row,6,rows2.reduce((s,r)=>s+r.Diff_T_LM,0),tSt2,FMT_DIF);
    scn(ws,row,8,rows2.reduce((s,r)=>s+r['Bet tickets LY'],0),tSt2,FMT_NB);
    scn(ws,row,9,rows2.reduce((s,r)=>s+r.Diff_T_LY,0),tSt2,FMT_DIF);
    row+=2;
  };

  ibloc('BLOC 1 — Progression vs LM ET vs LY','✅',inter(dp,sp_lm,sp_ly,'Diff_T_LM'),C.green_med,'4A7C3F','EBF5EB',`MTD > LM  ET  MTD > LY  |  seuil ${seuilP.toLocaleString()}`);
  ibloc('BLOC 2 — Dépassement LM+15% ET LY+15%','🏆',inter(dp,sp_lm15,sp_ly15,'Diff_T_LM15'),C.gold_dark,'A07800',C.gold_bg,`MTD > LM+15%  ET  MTD > LY+15%  |  seuil ${seuilP.toLocaleString()}`);
  sc(ws,row,2,'▼  BLOCS NÉGATIFS',sS(C.red_med)); mg(ws,row,2,row,10); row+=2;
  ibloc('BLOC 3 — Recul vs LM ET vs LY','⚠️',inter(dm,sm_lm,sm_ly,'Diff_T_LM',true),C.red_med,'A04040',C.red_bg,`MTD < LM  ET  MTD < LY  |  seuil ${seuilM.toLocaleString()}`,true);
  ibloc("BLOC 4 — N'a atteint ni LM+15% ni LY+15%",'🔴',inter(dm,sm_lm15,sm_ly15,'Diff_T_LM15',true),C.red_dark,'8B1A1A','FCE4D6',`MTD < LM+15%  ET  MTD < LY+15%  |  seuil ${seuilM.toLocaleString()}`,true);

  ws['!cols']=[3,5,30,13,13,14,10,13,14,10].map(w=>({wch:w}));
  ws['!ref']=XLSX.utils.encode_range({s:{r:0,c:0},e:{r:row,c:9}});
  XLSX.utils.book_append_sheet(wb,ws,sheetName.slice(0,31));
}

// ── Build file (marche mode = 6 onglets) ──────────
function buildFileMarche(df, label, mois, seuilP, seuilM) {
  const wb=XLSX.utils.book_new();
  buildOngData(wb,`① ${label}`.slice(0,31),df,label,mois);
  buildOngResume(wb,'② Résumé',df,label,mois);
  buildOngAnalyse(wb,'③ Analyse',df,label,mois);
  buildTaperPlus(wb,'④ TAPER+ — Marché',df,seuilP,label,mois);
  buildTaperMoins(wb,'⑤ TAPER− — Marché',df,seuilM,label,mois);
  buildIntersections(wb,'⑥ Intersections',df,seuilP,seuilM,label,mois);
  return wb;
}

// ── Build file (complet mode = 11 onglets) ─────────
function buildFileComplet(dfAll, dfCamp, label, mois, seuilP, seuilM) {
  const wb=XLSX.utils.book_new();
  buildOngData(wb,'① Global',dfAll,`${label} — Toutes salles`,mois);
  buildOngData(wb,'② CAMPRESJ',dfCamp,`CAMPRESJ — ${label}`,mois);
  buildOngResume(wb,'③ CAMPRESJ — Résumé',dfCamp,label,mois);
  buildOngAnalyse(wb,'④ Analyse — Global',dfAll,label,mois);
  buildOngAnalyse(wb,'⑤ Analyse — CAMPRESJ',dfCamp,'CAMPRESJ',mois);
  buildTaperPlus(wb,'⑥ TAPER+ — Global',dfAll,seuilP,label,mois);
  buildTaperPlus(wb,'⑦ TAPER+ — CAMPRESJ',dfCamp,seuilP,'CAMPRESJ',mois);
  buildTaperMoins(wb,'⑧ TAPER− — Global',dfAll,seuilM,label,mois);
  buildTaperMoins(wb,'⑨ TAPER− — CAMPRESJ',dfCamp,seuilM,'CAMPRESJ',mois);
  buildIntersections(wb,'⑩ Intersections — Global',dfAll,seuilP,seuilM,label,mois);
  buildIntersections(wb,'⑪ Intersections — CAMPRESJ',dfCamp,seuilP,seuilM,'CAMPRESJ',mois);
  return wb;
}

// ── Main Component ────────────────────────────────
export default function BetshopMarchesPage({ onNavigate }) {
  const [evalFile,   setEvalFile]   = useState(null);
  const [reportFile, setReportFile] = useState(null);
  const [mois,       setMois]       = useState('AVRIL 2026');
  const [seuilPlus,  setSeuilPlus]  = useState(500);
  const [seuilMoins, setSeuilMoins] = useState(2000);
  const [marches,    setMarches]    = useState('Balnda,Janvier,Bunker,Hippolyte,LUniversel,Franck');
  const [uniGroupes, setUniGroupes] = useState('UNI DG,UNI SF,UNI TA');
  const [running,    setRunning]    = useState(false);
  const [logs,       setLogs]       = useState([]);
  const [result,     setResult]     = useState(null);
  // Single market
  const [marketList,    setMarketList]    = useState(null);
  const [selectedMarket,setSelectedMarket]= useState('');
  const [marketResult,  setMarketResult]  = useState(null);

  const { setBetshopForMulti } = useAppContext();
  const evalRef   = useRef();
  const reportRef = useRef();
  const log = (msg) => setLogs(p => [...p, msg]);

  const readXlsx = (file, sheet) => new Promise((res,rej) => {
    const reader=new FileReader();
    reader.onload=e=>{
      try {
        const wb=XLSX.read(e.target.result,{type:'array'});
        const wsName=sheet||wb.SheetNames[0];
        const ws=wb.Sheets[wsName]||wb.Sheets[wb.SheetNames[0]];
        res(XLSX.utils.sheet_to_json(ws,{defval:0}));
      } catch(err){rej(err);}
    };
    reader.onerror=rej;
    reader.readAsArrayBuffer(file);
  });

  const buildMarketList = (evalRows, reportRows) => {
    const campSet=new Set(evalRows.map(r=>r['Shop']||r['Betshop name']||''));
    const markets=new Set(reportRows.map(r=>String(r['Market']||'').trim()).filter(Boolean));
    const list={};
    markets.forEach(m=>{ list[m]=reportRows.filter(r=>String(r['Market']||'').trim()===m); });
    // Add UNI subgroups
    const uniRows=reportRows.filter(r=>String(r['Market']||'').trim()==='UNI 1l');
    ['UNI DG','UNI SF','UNI TA'].forEach(u=>{
      const rows=uniRows.filter(r=>String(r['Shop']||'').startsWith(u));
      if(rows.length>0) list[u]=rows;
    });
    list['CAMPRESJ']=reportRows.filter(r=>campSet.has(r['Shop']||r['Betshop name']||''));
    list['Global']=reportRows;
    setMarketList(list);
    return {list, campSet};
  };

  // Run single market pipeline
  const runMarketPipeline = async () => {
    if(!selectedMarket||!marketList||!evalFile||!reportFile) return;
    setMarketResult(null);
    try {
      const moisUp=mois.toUpperCase();
      log(`\n📁 Génération ${selectedMarket}...`);
      const [evalRows, reportRows] = await Promise.all([
        readXlsx(evalFile,'CAMPRESJ'),
        readXlsx(reportFile,'Sheet1'),
      ]);
      const campSet=new Set(evalRows.map(r=>r['Shop']||r['Betshop name']||''));
      const rows=marketList[selectedMarket];
      const df=prepDf(rows);

      let wb;
      if(selectedMarket==='Global'){
        const dfCamp=prepDf(reportRows.filter(r=>campSet.has(r['Shop']||r['Betshop name']||'')));
        wb=buildFileComplet(df,dfCamp,'Global',moisUp,seuilPlus,seuilMoins);
      } else if(selectedMarket==='CAMPRESJ'){
        wb=buildFileMarche(df,'CAMPRESJ',moisUp,seuilPlus,seuilMoins);
      } else {
        wb=buildFileMarche(df,selectedMarket,moisUp,seuilPlus,seuilMoins);
      }

      const safe=selectedMarket.replace(/ /g,'_');
      const filename=`Betshop_${safe}_${moisUp.replace(/ /g,'_')}.xlsx`;
      const blob=new Blob([XLSX.write(wb,{bookType:'xlsx',type:'array'})],
        {type:'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'});
      setMarketResult({blob,filename,count:df.length});
      log(`  ✅ ${filename} (${df.length} salles)`);
    } catch(err){log(`❌ Erreur : ${err.message}`);}
  };

  const runPipeline = async () => {
    setRunning(true); setLogs([]); setResult(null);
    try {
      log('📂 Chargement des fichiers...');
      const [evalRows, reportRows] = await Promise.all([
        readXlsx(evalFile,'CAMPRESJ'),
        readXlsx(reportFile,'Sheet1'),
      ]);
      buildMarketList(evalRows,reportRows);
      const campSet2=new Set(evalRows.map(r=>r['Shop']||r['Betshop name']||''));
      log(`  ✓ Eval CAMPRESJ : ${campSet2.size} salles`);
      log(`  ✓ Report : ${reportRows.length} salles`);

      const df_all=prepDf(reportRows);
      const moisUp=mois.toUpperCase();
      const moisSafe=moisUp.replace(/ /g,'_');
      const zip=new JSZip();

      // Global (11 onglets)
      log('🌍 Génération Global...');
      const df_camp_all=prepDf(reportRows.filter(r=>campSet2.has(r['Shop']||r['Betshop name']||'')));
      const wbGlobal=buildFileComplet(df_all,df_camp_all,'Global',moisUp,seuilPlus,seuilMoins);
      zip.file(`Betshop_Global_${moisSafe}.xlsx`,XLSX.write(wbGlobal,{bookType:'xlsx',type:'array'}));
      log(`  ✅ Global (${df_all.length} salles) — 11 onglets`);

      // CAMPRESJ (6 onglets)
      log('📋 Génération CAMPRESJ...');
      const wbCamp=buildFileMarche(df_camp_all,'CAMPRESJ',moisUp,seuilPlus,seuilMoins);
      zip.file(`Betshop_CAMPRESJ_${moisSafe}.xlsx`,XLSX.write(wbCamp,{bookType:'xlsx',type:'array'}));
      log(`  ✅ CAMPRESJ (${df_camp_all.length} salles) — 6 onglets`);

      // Marchés standards (6 onglets chacun)
      log('🗺️  Génération par marché...');
      const marchesList=marches.split(',').map(m=>m.trim()).filter(Boolean);
      for(const market of marchesList){
        const df_m=prepDf(reportRows.filter(r=>String(r['Market']||'').trim()===market));
        if(df_m.length===0){log(`  ⚠️ ${market} : 0 salles`);continue;}
        const wb_m=buildFileMarche(df_m,market,moisUp,seuilPlus,seuilMoins);
        zip.file(`Betshop_${market.replace(/ /g,'_')}_${moisSafe}.xlsx`,XLSX.write(wb_m,{bookType:'xlsx',type:'array'}));
        log(`  ✅ ${market} (${df_m.length} salles)`);
      }

      // Sous-groupes UNI (6 onglets chacun)
      log('🏢 Génération UNI...');
      const df_uni_all=prepDf(reportRows.filter(r=>String(r['Market']||'').trim()==='UNI 1l'));
      const uniList=uniGroupes.split(',').map(u=>u.trim()).filter(Boolean);
      for(const uni of uniList){
        const df_u=df_uni_all.filter(r=>String(r.Shop||'').startsWith(uni));
        if(df_u.length===0){log(`  ⚠️ ${uni} : 0 salles`);continue;}
        const wb_u=buildFileMarche(df_u,uni,moisUp,seuilPlus,seuilMoins);
        zip.file(`Betshop_${uni.replace(/ /g,'_')}_${moisSafe}.xlsx`,XLSX.write(wb_u,{bookType:'xlsx',type:'array'}));
        log(`  ✅ ${uni} (${df_u.length} salles)`);
      }

      const blob=await zip.generateAsync({type:'blob'});
      setResult({blob,filename:`Betshop_Marches_${moisSafe}.zip`});
      // Store CAMPRESJ file for jumelage with Multi-Périodes
      const campresj_data = XLSX.write(buildFileMarche(df_camp_all,'CAMPRESJ',moisUp,seuilPlus,seuilMoins),{bookType:'xlsx',type:'array'});
      const campresj_blob = new Blob([campresj_data],{type:'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'});
      const campresj_file = new File([campresj_blob], `Betshop_CAMPRESJ_${moisSafe}.xlsx`, {type:'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'});
      setBetshopForMulti({ mois: moisUp, file: campresj_file });
      log(`\n✅ Terminé !`);
    } catch(err){log(`❌ Erreur : ${err.message}`);}
    setRunning(false);
  };

  return (
    <div>
      {/* Uploads */}
      <p className="section-label">Fichiers source</p>
      <div style={{display:'grid',gridTemplateColumns:'1fr 1fr',gap:'1rem',marginBottom:'1.25rem'}}>
        {[
          {label:'Betshop Cameroon Evaluation.xlsx',sub:'Onglet CAMPRESJ',ref:evalRef,setter:setEvalFile,file:evalFile,icon:'📊'},
          {label:'Betshop report Cameroon.xlsx',sub:'Onglet Sheet1 — données complètes',ref:reportRef,setter:setReportFile,file:reportFile,icon:'📋'},
        ].map(({label,sub,ref,setter,file,icon})=>(
          <div key={label} className={`drop-zone ${file?'has-file':''}`}
            onClick={()=>ref.current.click()}
            onDragOver={e=>e.preventDefault()}
            onDrop={e=>{e.preventDefault();const f=e.dataTransfer.files[0];setter(f);}}>
            <input ref={ref} type="file" accept=".xlsx,.xls" hidden onChange={async e=>{
              const f=e.target.files[0]; setter(f);
              if(evalFile||reportFile){
                try{
                  const [ev,rep]=await Promise.all([
                    readXlsx(evalFile||f,'CAMPRESJ'),
                    readXlsx(reportFile||f,'Sheet1'),
                  ]);
                  buildMarketList(ev,rep);
                }catch(err){}
              }
            }}/>
            <div className="drop-icon">{icon}</div>
            <div className="drop-label">{label}</div>
            <div className="drop-sub">{sub}</div>
            {file&&<div className="file-chip">✓ {file.name}</div>}
          </div>
        ))}
      </div>

      {/* Paramètres */}
      <div className="card" style={{marginBottom:'1.25rem'}}>
        <p className="section-label">Paramètres</p>
        <div style={{display:'grid',gridTemplateColumns:'1fr 1fr 1fr',gap:'1rem',marginBottom:'1rem'}}>
          <label style={{display:'flex',flexDirection:'column',gap:5}}>
            <span style={{fontSize:12,color:'var(--muted)',fontWeight:600}}>Mois</span>
            <input value={mois} onChange={e=>setMois(e.target.value)} className="field-input"/>
          </label>
          <label style={{display:'flex',flexDirection:'column',gap:5}}>
            <span style={{fontSize:12,color:'var(--muted)',fontWeight:600}}>Seuil TAPER+ (tickets min)</span>
            <input type="number" value={seuilPlus} onChange={e=>setSeuilPlus(Number(e.target.value))} className="field-input"/>
          </label>
          <label style={{display:'flex',flexDirection:'column',gap:5}}>
            <span style={{fontSize:12,color:'var(--muted)',fontWeight:600}}>Seuil TAPER− (tickets min)</span>
            <input type="number" value={seuilMoins} onChange={e=>setSeuilMoins(Number(e.target.value))} className="field-input"/>
          </label>
        </div>
        <div style={{display:'grid',gridTemplateColumns:'1fr 1fr',gap:'1rem'}}>
          <label style={{display:'flex',flexDirection:'column',gap:5}}>
            <span style={{fontSize:12,color:'var(--muted)',fontWeight:600}}>Marchés standard (séparés par virgule)</span>
            <input value={marches} onChange={e=>setMarches(e.target.value)} className="field-input" style={{width:'100%'}}/>
          </label>
          <label style={{display:'flex',flexDirection:'column',gap:5}}>
            <span style={{fontSize:12,color:'var(--muted)',fontWeight:600}}>Sous-groupes UNI (séparés par virgule)</span>
            <input value={uniGroupes} onChange={e=>setUniGroupes(e.target.value)} className="field-input" style={{width:'100%'}}/>
          </label>
        </div>
      </div>

      {/* Sélecteur marché individuel */}
      {marketList && (
        <div className="card" style={{marginBottom:'1.25rem'}}>
          <p className="section-label">Générer le fichier d'un seul marché</p>
          <div style={{display:'flex',gap:'1rem',alignItems:'flex-end',flexWrap:'wrap'}}>
            <label style={{display:'flex',flexDirection:'column',gap:5,flex:1}}>
              <span style={{fontSize:12,color:'var(--muted)',fontWeight:600}}>Marché</span>
              <select value={selectedMarket} onChange={e=>setSelectedMarket(e.target.value)}
                className="field-input" style={{cursor:'pointer'}}>
                <option value="">-- Choisir un marché --</option>
                {Object.entries(marketList).sort(([a],[b])=>a.localeCompare(b)).map(([m,rows])=>(
                  <option key={m} value={m}>{m} ({rows.length} salles)</option>
                ))}
              </select>
            </label>
            <button onClick={runMarketPipeline}
              disabled={!selectedMarket||!evalFile||!reportFile}
              className="btn primary" style={{padding:'9px 20px',whiteSpace:'nowrap'}}>
              ▶  Générer {selectedMarket||'marché'}
            </button>
          </div>
          {marketResult&&(
            <div style={{marginTop:'1rem',display:'flex',alignItems:'center',gap:'1rem'}}>
              <button onClick={()=>saveAs(marketResult.blob,marketResult.filename)}
                className="btn success" style={{fontSize:13}}>
                ⬇  Télécharger {marketResult.filename}
              </button>
              <span style={{fontSize:12,color:'var(--muted)'}}>{marketResult.count} salles</span>
            </div>
          )}
        </div>
      )}

      <button className="btn primary" disabled={!evalFile||!reportFile||running}
        onClick={runPipeline} style={{fontSize:14,padding:'11px 28px',marginBottom:'1.5rem'}}>
        {running?'⟳ Génération...':'▶  Générer tous les marchés'}
      </button>

      {logs.length>0&&(
        <>
          <p className="section-label">Console</p>
          <div className="log-console" style={{marginBottom:'1.5rem'}}>
            {logs.map((l,i)=><span key={i} className="log-line" style={{display:'block'}}>{l}</span>)}
          </div>
        </>
      )}

      {result&&(
        <div className="card">
          <button onClick={()=>saveAs(result.blob,result.filename)}
            className="btn success" style={{fontSize:14,padding:'10px 24px'}}>
            ⬇  Télécharger {result.filename}
          </button>
          {onNavigate && (
            <button onClick={() => onNavigate('betshop-periodes')}
              style={{ marginLeft:'1rem', padding:'10px 20px', fontSize:13, background:'var(--accent-dim)', color:'var(--accent)', border:'1px solid var(--accent)', borderRadius:'var(--radius)', cursor:'pointer', fontWeight:600 }}>
              📅 Analyser en Multi-Périodes →
            </button>
          )}
        </div>
      )}
    </div>
  );
}
