// src/lib/pipeline.js
// Full port of cameroon_pipeline.py → runs in-browser with SheetJS + PapaParse

import * as XLSX from 'xlsx';
import Papa from 'papaparse';
import { format, addDays } from 'date-fns';

// ─── Helpers ────────────────────────────────────────────────────────────────

const floorMultiple = (val, m) => Math.floor(val / m) * m;

function roundSport(val) {
  if (val == null || isNaN(val)) return null;
  return val < 1000 ? floorMultiple(val, 100) : floorMultiple(val, 500);
}

function roundMix(val) {
  if (val == null || isNaN(val)) return null;
  return val < 1000 ? floorMultiple(val, 100) : floorMultiple(val, 500);
}

// ─── File readers ────────────────────────────────────────────────────────────

export function readExcelSheets(arrayBuffer) {
  const wb = XLSX.read(arrayBuffer, { type: 'array' });
  const sheets = {};
  for (const name of wb.SheetNames) {
    sheets[name] = XLSX.utils.sheet_to_json(wb.Sheets[name], { defval: null });
  }
  return sheets;
}

export function readCSV(text) {
  const result = Papa.parse(text, { header: true, skipEmptyLines: true, dynamicTyping: true });
  return result.data;
}

// ─── Drop columns helper ─────────────────────────────────────────────────────

function dropCols(rows, cols) {
  return rows.map(r => {
    const out = { ...r };
    for (const c of cols) delete out[c];
    return out;
  });
}

function renameCols(rows, map) {
  return rows.map(r => {
    const out = {};
    for (const [k, v] of Object.entries(r)) out[map[k] ?? k] = v;
    return out;
  });
}

// ─── Main pipeline ───────────────────────────────────────────────────────────

export function runPipeline(sheets, csvRows, onProgress, threshold = 2000, maxSport = 30000) {

  onProgress?.('Loading MTD sheet…', 5);

  // ── Step 1: MTD ───────────────────────────────────────────────────────────
  let mtd = sheets['MTD'] ?? [];

  const DROP_MTD = [
    'companyname','currency','playerTypeRisk','lastname',
    'betting_bonus_money','cash_tickets','bonus_tickets','bonus_ticket_payin',
    'cash_ticket_payin','no_win_bonus_tickets','ticket_bonus_converted',
    'casino_bonus_money','casino_bonus_payin','played_casinogametypes',
    'most_played_casino_game_type','casino_bonus_ggr','casino_standard_payin',
    'casino_slot_payin','casino_crash_payin','casino_live_payin',
    'casino_bonus_to_cash','cash_back_casino','freebet_superheli',
    'freespin_aviator','all_free_spins_conv','deposit','withdraw',
    'last_bet_cash_activity','last_casino_cash_activity','last_bet_bonus_activity',
    'last_casino_bonus_activity','last_casino_activity_days','last_bet_activity_days'
  ];
  mtd = dropCols(mtd, DROP_MTD);
  mtd = renameCols(mtd, { 'total ggr': 'MTD total ggr' });

  // Yesterday ggr
  mtd = mtd.map(r => ({
    ...r,
    'Yesterday ggr': (r.ticket_yesterday_ggr ?? 0) + (r.casino_yesterday_ggr ?? 0)
  }));

  // Sort desc, filter >= 1999
  mtd.sort((a, b) => b['Yesterday ggr'] - a['Yesterday ggr']);
  mtd = mtd.filter(r => r['Yesterday ggr'] >= threshold);

  mtd = dropCols(mtd, ['ticket_ggr', 'casino_ggr']);

  onProgress?.('Loading YTD sheet…', 15);

  // ── Step 1: YTD ───────────────────────────────────────────────────────────
  let ytd = sheets['YTD'] ?? [];

  // Extract phones
  const phones = ytd
    .filter(r => r.phone != null && String(r.phone).trim() !== '')
    .map(r => ({
      account_id: r.account_id,
      phone: String(r.phone).replace('+', '').trim()
    }));

  const DROP_YTD = [
    'companyname','currency','playerTypeRisk','phone','firstname','lastname',
    'betting_bonus_money','cash_tickets','bonus_tickets','bonus_ticket_payin',
    'cash_ticket_payin','ticket_ggr','ticket_yesterday_ggr','no_win_bonus_tickets',
    'ticket_bonus_converted','casino_bonus_money','casino_bonus_payin',
    'played_casinogametypes','most_played_casino_game_type','casino_bonus_ggr',
    'casino_standard_payin','casino_slot_payin','casino_crash_payin','casino_live_payin',
    'casino_bonus_to_cash','cash_back_casino','freebet_superheli','freespin_aviator',
    'all_free_spins_conv','casino_ggr','casino_yesterday_ggr','deposit','withdraw',
    'last_bet_cash_activity','last_casino_cash_activity','last_bet_bonus_activity',
    'last_casino_bonus_activity','last_casino_activity_days','last_bet_activity_days'
  ];
  ytd = dropCols(ytd, DROP_YTD);

  // ── Step 1: Last 3 days ───────────────────────────────────────────────────
  let l3 = sheets['Last 3 days'] ?? [];
  const DROP_L3 = [
    'companyname','currency','playerTypeRisk','phone','firstname','lastname',
    'betting_bonus_money','cash_tickets','bonus_tickets','bonus_ticket_payin',
    'cash_ticket_payin','ticket_ggr','ticket_yesterday_ggr','no_win_bonus_tickets',
    'ticket_bonus_converted','casino_bonus_money','casino_bonus_payin','played_casinogametypes'
  ];
  l3 = dropCols(l3, DROP_L3);

  onProgress?.('Calculating segments…', 30);

  // ── Step 2: Merge YTD into MTD ────────────────────────────────────────────
  const ytdDedup = {};
  for (const r of ytd) {
    if (!(r.account_id in ytdDedup)) ytdDedup[r.account_id] = r['total ggr'] ?? 0;
  }
  mtd = mtd.map(r => ({ ...r, YTD: ytdDedup[r.account_id] ?? 0 }));

  // game column
  mtd = mtd.map(r => {
    const yggr = r['Yesterday ggr'] || 0;
    const ratio = yggr !== 0 ? (r.casino_yesterday_ggr ?? 0) / yggr : null;
    return { ...r, game: ratio != null && ratio >= 0.40 ? 'casino' : 'sport bonus' };
  });

  // 5% column
  mtd = mtd.map(r => ({ ...r, '5%': Math.round(r['Yesterday ggr'] * 0.05 * 10000) / 10000 }));

  // offer, freebets, amount, Mixsport columns
  mtd = mtd.map(r => ({ ...r, offer: '', freebets: '', amount: null, Mixsport: null }));

  // Remove Category == Negative
  mtd = mtd.filter(r => r.Category !== 'Negative');

  // Remove YTD or MTD total ggr < 0
  mtd = mtd.filter(r => (r.YTD ?? 0) >= 0 && (r['MTD total ggr'] ?? 0) >= 0);

  // Build l3 map: account_id → most_played_casino_game_type
  const l3Map = {};
  for (const r of l3) {
    if (!(r.account_id in l3Map)) l3Map[r.account_id] = r.most_played_casino_game_type;
  }

  function assignOffer(row) {
    if (row.game !== 'casino') return '';
    const val = l3Map[row.account_id];
    if (val == null || val === '') return 'Cash';
    if (val === 'CRASH_GAMES') return 'Aviator';
    if (val === 'SLOTS') return 'Casino bonus';
    return 'Cash';
  }

  mtd = mtd.map(r => ({ ...r, offer: assignOffer(r) }));

  // Other casino categories → Cash
  mtd = mtd.map(r => {
    if (r.game === 'casino' && !['Aviator', 'Casino bonus', 'Cash', ''].includes(r.offer)) {
      return { ...r, offer: 'Cash' };
    }
    return r;
  });

  // casino + ticket_yesterday_ggr != 0 + offer != Cash → MIX
  mtd = mtd.map(r => {
    if (r.game === 'casino' && (r.ticket_yesterday_ggr ?? 0) !== 0 && r.offer !== 'Cash') {
      return { ...r, game: 'MIX' };
    }
    return r;
  });

  // MIX + 5% < 350 → casino
  mtd = mtd.map(r => {
    if (r.game === 'MIX' && r['5%'] < 350) return { ...r, game: 'casino' };
    return r;
  });

  // casino + Aviator + 5% < 250 → Casino bonus
  mtd = mtd.map(r => {
    if (r.game === 'casino' && r.offer === 'Aviator' && r['5%'] < 250) {
      return { ...r, offer: 'Casino bonus' };
    }
    return r;
  });

  onProgress?.('Computing amounts & freebets…', 50);

  // ── Amount & freebets calculations ────────────────────────────────────────
  mtd = mtd.map(r => {
    let { offer, game, '5%': pct, amount, freebets, Mixsport } = r;

    if (offer === 'Cash') {
      amount = floorMultiple(pct, 50);
    }
    if (game === 'casino' && offer === 'Casino bonus') {
      amount = floorMultiple(pct, 50);
    }
    if (game === 'casino' && offer === 'Aviator') {
      freebets = '5*';
      amount = floorMultiple(pct / 5, 5);
    }
    if (game === 'sport bonus') {
      amount = Math.min(roundSport(pct), maxSport);
    }
    if (game === 'MIX' && offer === 'Casino bonus') {
      const half = pct / 2;
      amount = floorMultiple(half, 50);
      Mixsport = Math.min(roundMix(half), maxSport);
    }
    if (game === 'MIX' && offer === 'Aviator') {
      freebets = '5*';
      const half = pct / 2;
      amount = floorMultiple(half / 5, 5);
      Mixsport = Math.min(roundMix(half), maxSport);
    }

    return { ...r, offer, game, amount, freebets, Mixsport };
  });

  onProgress?.('Merging CSV transactions…', 65);

  // ── Step 3: Today column from CSV ─────────────────────────────────────────
  const csvMap = {};
  for (const r of csvRows) {
    const id = r['Account ID'];
    if (id != null && !(id in csvMap)) csvMap[id] = r['Total Gross'] ?? 0;
  }

  const beforeCount = mtd.length;
  mtd = mtd.map(r => ({ ...r, Today: csvMap[r.account_id] ?? 0 }));
  mtd = mtd.filter(r => (r.Today ?? 0) >= 0);
  const removedToday = beforeCount - mtd.length;

  onProgress?.('Generating campaign files…', 80);

  // ── Step 4: Campaign files ────────────────────────────────────────────────
  const today = new Date();
  const todayStr = format(today, 'yyyy-MM-dd');
  const endStr = format(addDays(today, 7), 'yyyy-MM-dd');


  const SENTINEL_ID = 1326274;

  // 1. casino_bonus_cm.csv
  const cbRows = mtd
    .filter(r => ['casino', 'MIX'].includes(r.game) && r.offer === 'Casino bonus' && r.account_id !== SENTINEL_ID)
    .map(r => ({ account_id: r.account_id, amount: r.amount, currency: 'XAF' }));

  // 2. CM_daily_XXX.csv (grouped by amount)
  const dailyRows = [
    ...mtd.filter(r => r.game === 'sport bonus' && r.account_id !== SENTINEL_ID).map(r => ({ id: r.account_id, amount: r.amount })),
    ...mtd.filter(r => r.game === 'MIX' && r.account_id !== SENTINEL_ID).map(r => ({ id: r.account_id, amount: r.Mixsport }))
  ].filter(r => r.amount != null && r.amount > 0);

  const dailyGroups = {};
  for (const r of dailyRows) {
    const key = String(r.amount);
    if (!dailyGroups[key]) dailyGroups[key] = [];
    dailyGroups[key].push(r);
  }

  // 3. Daily_Cashback_Template.csv
  const cashRows = mtd
    .filter(r => r.game === 'casino' && r.offer === 'Cash' && r.account_id !== SENTINEL_ID)
    .map(r => ({ account_id: r.account_id, ' amount ': r.amount, currency_iso3: 'XAF' }));

  // 4. spribe_freebets.csv
  const spribeRows = mtd
    .filter(r => r.offer === 'Aviator' && r.account_id !== SENTINEL_ID)
    .map(r => ({
      opPlayerId: r.account_id,
      currency: 'XAF',
      ticketAmount_amount: r.amount,
      num_of_free_ticketAmounts: 5,
      start_date: todayStr,
      end_date: endStr,
      game: 'AVIATOR'
    }));

  onProgress?.('Pipeline complete!', 100);

  return {
    mtd,
    phones,
    campaigns: {
      casino_bonus_cm: cbRows,
      cm_daily: dailyGroups,
      daily_cashback: cashRows,
      spribe_freebets: spribeRows
    },
    stats: {
      totalPlayers: mtd.length,
      removedToday,
      phonesExtracted: phones.length,
      byCampaign: {
        casinoBonus: cbRows.length,
        sportBonus: Object.values(dailyGroups).reduce((s, g) => s + g.length, 0),
        cashback: cashRows.length,
        spribe: spribeRows.length
      }
    }
  };
}

// ─── Export helpers ──────────────────────────────────────────────────────────

export function exportCSV(rows, filename) {
  const csv = Papa.unparse(rows);
  const blob = new Blob([csv], { type: 'text/csv;charset=utf-8;' });
  return { blob, filename };
}

export function exportXLSX(rows, sheetName, filename) {
  const ws = XLSX.utils.json_to_sheet(rows);
  const wb = XLSX.utils.book_new();
  XLSX.utils.book_append_sheet(wb, ws, sheetName);
  const buf = XLSX.write(wb, { bookType: 'xlsx', type: 'array' });
  const blob = new Blob([buf], { type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet' });
  return { blob, filename };
}

export function buildAllExports(result) {
  const files = [];

  // CM_Daily_DD.MM.YY.xlsx (MTD sheet)
  const d = new Date();
  const dd = String(d.getDate()).padStart(2, '0');
  const mm = String(d.getMonth() + 1).padStart(2, '0');
  const yy = String(d.getFullYear()).slice(-2);
  const finalName = `CM_Daily_${dd}.${mm}.${yy}.xlsx`;
  files.push(exportXLSX(result.mtd, 'MTD', finalName));

  // Cameroon_YTD_phones.xlsx
  files.push(exportXLSX(result.phones, 'Phones', 'Cameroon_YTD_phones.xlsx'));

  // casino_bonus_cm.csv
  files.push(exportCSV(result.campaigns.casino_bonus_cm, 'casino_bonus_cm.csv'));

  // Daily cashback
  files.push(exportCSV(result.campaigns.daily_cashback, 'Daily_Cashback_Template.csv'));

  // spribe_freebets.csv
  files.push(exportCSV(result.campaigns.spribe_freebets, 'spribe_freebets.csv'));

  // CM_daily_XXX.csv files
  for (const [amt, rows] of Object.entries(result.campaigns.cm_daily)) {
    files.push(exportCSV(rows, `CM_daily_${amt}.csv`));
  }

  return files;
}
