/************************************************************
 * Trendy_FBE.gs — trendy + stock + pokrycie + zamówienie (T) + rotacja (U)
 * Wejście  : 'Sprzedaż_okresowa_FBE'
 * Wyjście  : 'Trendy_FBE'
 * Stock    : karta 'ADS_FBE' (kolumna „Stock") – opcjonalna
 * 
 * ZMIANA: Usunięto kolumnę "Deficyt_w_miesiącu_1" (była ostatnią kolumną)
 ************************************************************/

// ===== Nazwy arkuszy =====
const IN_SHEET_TRENDS  = 'Sprzedaż_okresowa_FBE';
const OUT_SHEET_TRENDS = 'Trendy_FBE';

// ===== Docelowe kolumny na wyniki pomocnicze =====
const COL_REORDER_PROPOSAL  = 26; // Z (nie będzie resetowana przy przeliczeniu)
const COL_ROTATION_LEFTOVER = 27; // AA (nie będzie resetowana przy przeliczeniu)

// ===== Kolorowanie tła dla stocku =====
const HIGHLIGHT_STOCK_WITH_INTRANSIT = true;
const INTRANSIT_BG_COLOR = '#e8f1fb'; // delikatny niebieski

// ===== Parametry trendów =====
const TREND_WINDOW = 6, MA3_WINDOW=3, MA6_WINDOW=6;
const EXCLUDE_CURRENT = false, NORM_THRESHOLD = 0.15;
const PROPORTIONAL_MIN_DAY=10, PROPORTIONAL_CAP=1.6, PROPORTIONAL_MIN_UNITS=0;

// ——— Nagłówki stałe (placeholders) i format końcowy ———
const HEADER_BASE_PROPOSAL = 'Propozycja_zamówienia';
const HEADER_BASE_LEFTOVER = 'Sztuki_pozostałe_po_[…M]';

function ensurePersistentHelperColumns_(sh) {
  if (!sh) return;
  ensureColumns_(sh, COL_ROTATION_LEFTOVER);
  const lc = sh.getLastColumn();
  const hdr = sh.getRange(1,1,1,lc).getDisplayValues()[0];
  const curV = String(hdr[COL_REORDER_PROPOSAL-1] || '').trim();
  if (!curV) sh.getRange(1, COL_REORDER_PROPOSAL).setValue(HEADER_BASE_PROPOSAL);
  const curW = String(hdr[COL_ROTATION_LEFTOVER-1] || '').trim();
  if (!curW) sh.getRange(1, COL_ROTATION_LEFTOVER).setValue(HEADER_BASE_LEFTOVER);
}

function setColumnHeaderAndClear_(sh, col1, newHeader) {
  const lr = Math.max(2, sh.getLastRow());
  sh.getRange(1, col1).setValue(newHeader);
  if (lr > 1) {
    const rng = sh.getRange(2, col1, lr-1, 1);
    try { rng.clearContent(); } catch(_){}
    try { rng.clearNote(); } catch(_){}
    try { rng.clearDataValidations(); } catch(_){}
  }
}

/* ============================================================
 *   GŁÓWNA TABELA: liczenie trendów FBE + stock + pokrycie
 *   ZMIANA: Usunięto kolumnę Deficyt_w_miesiącu_1
 * ============================================================ */
function computeTrendsFromSalesPeriods(mult1, mult2, mult3) {
  // Domyślnie 0% jeśli nie podano
  mult1 = Number(mult1) || 0;
  mult2 = Number(mult2) || 0;
  mult3 = Number(mult3) || 0;

  if (typeof requireActivePartner_ === 'function') requireActivePartner_();

  const ss   = SpreadsheetApp.getActive();
  const inSh = ss.getSheetByName(IN_SHEET_TRENDS);
  if (!inSh) throw new Error('Brak arkusza "' + IN_SHEET_TRENDS + '".');

  const { stockMap, priceNetMap, performanceMap } = readAdsFbeMaps_();
  const VAT_RO = 0.21;
  const inTransitMap = readInTransitMap_();
  const productInfoMap = readPartnerProductInfoMap_();

  const lr = inSh.getLastRow(), lc = inSh.getLastColumn();
  if (lr < 2) throw new Error('Brak danych do trendów.');

  const vals = inSh.getRange(1,1,lr,lc).getValues();
  const hdrDisp = inSh.getRange(1,1,1,lc).getDisplayValues()[0];
  const monthCols = getMonthColumns_(hdrDisp);
  if (!monthCols.length) throw new Error('Nie znaleziono kolumn miesięcznych (MM-YYYY).');

  const today = new Date(), nowY=today.getFullYear(), nowM=today.getMonth()+1;
  const lastIdx = monthCols.length-1;
  let useLast = EXCLUDE_CURRENT ? lastIdx-1 : lastIdx;
  const lastIsCurrent = (monthCols[lastIdx].y===nowY && monthCols[lastIdx].m===nowM);

  let scaleCurrent=false, scaleFactor=1;
  if (lastIsCurrent) {
    const day=today.getDate();
    if (day < PROPORTIONAL_MIN_DAY) useLast = lastIdx-1;
    else {
      const proporcja = daysInMonth_(nowY,nowM)/Math.max(1,day);
      scaleFactor = Math.min(PROPORTIONAL_CAP, proporcja);
      scaleCurrent = true;
    }
  }
  if (useLast < 0) throw new Error('Za mało miesięcy po odjęciu bieżącego.');

  // ===== ZMIANA: Dynamiczne nagłówki dla prognoz =====
  const formatMultiplier = (m) => {
    if (m === 0) return '';
    const sign = m > 0 ? '+' : '';
    return ` [${sign}${m}%]`;
  };

  const out = [[
    'part_number_key','ID','SKU','EAN',
    '3 miesiące wstecz','Sprzedaż dwa miesiące wstecz','Sprzedaż ubiegły miesiąc','Sprzedaż w obecnym miesiącu',
    'Trend_M0/M-1_%','Trend_M-1/M-2_%','Trend_M-2/M-3_%',
    'Suma_3M','Średnia_3M','Suma_6M','Średnia_6M','Slope_6M','Klasyfikacja',
    'Performance',
    `Prognoza_nast_M${formatMultiplier(mult1)}`,
    `Prognoza_+2M${formatMultiplier(mult2)}`,
    `Prognoza_+3M${formatMultiplier(mult3)}`,
    'Stock','Cena_netto','Cena_brutto','Pokrycie_magazynu'
  ]];

  for (let r=1; r<vals.length; r++) {
    const key = String(vals[r][0]||'').trim(); if (!key) continue;

    const info = lookupProductInfo_(productInfoMap, key);

    const series = monthCols.map(mc => toNum_(vals[r][mc.colIdx0]));

    if (lastIsCurrent && scaleCurrent) {
      const lastVal0 = series[series.length-1] || 0;
      series[series.length-1] = Math.max(PROPORTIONAL_MIN_UNITS, Math.round(lastVal0*scaleFactor));
    }

    const salesCurrent = series[useLast] || 0;
    const salesPrev1 = (useLast > 0 ? series[useLast-1] || 0 : 0);
    const salesPrev2 = (useLast > 1 ? series[useLast-2] || 0 : 0);
    const salesPrev3 = (useLast > 2 ? series[useLast-3] || 0 : 0);

    const trendM0M1 = (salesPrev1>0 ? ((salesCurrent-salesPrev1)/salesPrev1)*100 : (salesCurrent>0 ? 100 : ''));
    const trendM1M2 = (salesPrev2>0 ? ((salesPrev1-salesPrev2)/salesPrev2)*100 : (salesPrev1>0 ? 100 : ''));
    const trendM2M3 = (salesPrev3>0 ? ((salesPrev2-salesPrev3)/salesPrev3)*100 : (salesPrev2>0 ? 100 : ''));

    const sum3 = sumTail_(series, useLast, MA3_WINDOW), avg3 = avgFromSum_(sum3, Math.min(MA3_WINDOW, useLast+1));
    const sum6 = sumTail_(series, useLast, MA6_WINDOW), avg6 = avgFromSum_(sum6, Math.min(MA6_WINDOW, useLast+1));
    const slope = linreg_(tailWindow_(series, useLast, MA6_WINDOW)).slope;

    const cls = (slope > 0.5 ? '▲ rośnie' : (slope < -0.5 ? '▼ spada' : '▶ stabilny'));

    const performance = performanceMap.has(key) ? performanceMap.get(key) : '';

    // ===== ZMIANA: Obliczanie prognoz bazowych, potem mnożenie =====
    const fc1_base = Math.max(0, Math.round(avg3 + 1 * slope));
    const fc2_base = Math.max(0, Math.round(avg3 + 2 * slope));
    const fc3_base = Math.max(0, Math.round(avg3 + 3 * slope));

    const fc1 = Math.round(fc1_base * (1 + mult1 / 100));
    const fc2 = Math.round(fc2_base * (1 + mult2 / 100));
    const fc3 = Math.round(fc3_base * (1 + mult3 / 100));

    const stockADS   = Number(stockMap.get(key) || 0);
    const inTransit  = Number(inTransitMap.get(key) || 0);
    const stockTotal = stockADS + inTransit;
    let remaining = stockTotal, coverMonths = 0;

    if (remaining >= fc1) { remaining -= fc1; coverMonths = 1; }
    else { coverMonths = (fc1 > 0 ? 0 : 1); }

    if (coverMonths >= 1) {
      if (remaining >= fc2) { remaining -= fc2; coverMonths = 2; }
      else if (fc2 === 0)   { coverMonths = 2; }
    }
    if (coverMonths >= 2) {
      if (remaining >= fc3) { remaining -= fc3; coverMonths = 3; }
      else if (fc3 === 0)   { coverMonths = 3; }
    }

    let coverageMsg = '';
    if (coverMonths >= 3)       coverageMsg = 'wystarczy na 3 miesiące lub więcej';
    else if (coverMonths === 2) coverageMsg = 'wystarczy na pokrycie sprzedaży na dwa miesiące';
    else if (coverMonths === 1) coverageMsg = 'wystarczy na jeden miesiąc';
    else                        coverageMsg = 'nie wystarczy na pokrycie sprzedaży w tym miesiącu';

    const net = priceNetMap.has(key) ? Number(priceNetMap.get(key)) : '';
    const gross = (net === '' ? '' : round_(net * (1 + VAT_RO), 2));

    // ZMIANA: usunięto deficitM1 i coverMonths (duplikat info z coverageMsg)
    out.push([
      key,
      info ? info.id  : '',
      info ? info.sku : '',
      info ? info.ean : '',
      (useLast>2 ? salesPrev3 : ''),
      (useLast>1 ? salesPrev2 : ''),
      (useLast>0 ? salesPrev1 : ''),
      salesCurrent,
      isNum_(trendM0M1) ? round_(trendM0M1, 1) : trendM0M1,
      isNum_(trendM1M2) ? round_(trendM1M2, 1) : trendM1M2,
      isNum_(trendM2M3) ? round_(trendM2M3, 1) : trendM2M3,
      sum3, round_(avg3,2), sum6, round_(avg6,2), round_(slope,3),
      cls,
      performance,
      fc1, fc2, fc3,
      stockTotal,
      net, gross,
      coverageMsg
      // USUNIĘTO: coverMonths (duplikat), deficitM1
    ]);
  }

  let outSh = ss.getSheetByName(OUT_SHEET_TRENDS) || ss.insertSheet(OUT_SHEET_TRENDS);

  ensureColumns_(outSh, COL_REORDER_PROPOSAL);

  const baseCols = out[0].length;
  const lastRowExisting = Math.max(2, outSh.getLastRow(), out.length);

  outSh.getRange(1, 1, lastRowExisting, baseCols).clearContent();
  outSh.getRange(1, 1, out.length, baseCols).setValues(out);
  outSh.setFrozenRows(1);

  // Adnotacje dla Stock
  try {
    const hdrNow = outSh.getRange(1, 1, 1, outSh.getLastColumn()).getDisplayValues()[0];
    const cStock1 = colIndexByHeaderSoft_FBE(hdrNow, 'Stock');
    if (cStock1) {
      const nRows = out.length - 1;
      if (nRows > 0) {
        const rngStockData = outSh.getRange(2, cStock1, nRows, 1);

        try { rngStockData.clearNote(); } catch(_){}
        try {
          const clearBgs = Array.from({length: nRows}, () => ['']);
          rngStockData.setBackgrounds(clearBgs);
        } catch(_){}

        const notes = [];
        const bgs   = [];
        const keysCol1 = colIndexByHeaderSoft_FBE(hdrNow, 'part_number_key');
        let keys = [];
        if (keysCol1) {
          keys = outSh.getRange(2, keysCol1, nRows, 1).getDisplayValues().map(r => String(r[0]||'').trim());
        }

        const stockValsNow = rngStockData.getValues().map(r => Number(r[0]||0));

        for (let i = 0; i < nRows; i++) {
          const key = keys[i] || '';
          const ads = Number(stockMap.get(key) || 0);
          const it  = Number(inTransitMap.get(key) || 0);
          if (it > 0) {
            const totalShown = stockValsNow[i];
            notes.push([`${totalShown} = ${ads} (ADS_FBE) + ${it} (w drodze)`]);
            bgs.push([HIGHLIGHT_STOCK_WITH_INTRANSIT ? INTRANSIT_BG_COLOR : '']);
          } else {
            notes.push(['']);
            bgs.push(['']);
          }
        }

        try { rngStockData.setNotes(notes); } catch(_){}
        try { rngStockData.setBackgrounds(bgs); } catch(_){}
      }
    }
  } catch(e) {
    Logger.log('Stock annotation failed: ' + (e && (e.stack || e.message) || e));
  }

  // KRYTYCZNE: formatowanie warunkowe dla pokrycia
  try { 
    ensureCoverageFormatting_(outSh); 
    SpreadsheetApp.flush();
  } catch (e) {
    Logger.log('ensureCoverageFormatting_ failed: ' + (e && (e.stack || e.message) || e));
  }

  try { flagLow3mOnTrendy_(outSh); } catch (e) { Logger.log('flagLow3mOnTrendy_ skipped: ' + e); }
  
  SpreadsheetApp.flush();
}

/** Oznacz w Trendy_FBE PNK o Suma_3M ≤ 1 lekko czerwonym tłem (tylko kolumna PNK). */
function flagLow3mOnTrendy_(sh) {
  if (!sh) return;
  const vals = sh.getDataRange().getValues();
  if (vals.length < 2) return;

  const headers = vals[0].map(String);

  const pickCol1 = (cands) => {
    for (const cand of cands) {
      const idx1 = colIndexByHeaderSoft_FBE(headers, cand);
      if (idx1 > 0) return idx1;
    }
    return 0;
  };

  const idxPNK1   = pickCol1(['part_number_key','PNK','part_number','SKU','sku']);
  const idxSuma31 = pickCol1(['Suma_3M','Suma 3M','SUMA_3M']);

  if (!idxPNK1 || !idxSuma31) return;

  const n = vals.length - 1;
  const bgs = Array.from({ length: n }, () => ['']);

  for (let r = 1; r < vals.length; r++) {
    const s3 = Number(vals[r][idxSuma31 - 1]);
    if (!isNaN(s3) && s3 <= 1) {
      bgs[r - 1][0] = '#ffdddd';
    }
  }

  sh.getRange(2, idxPNK1, n, 1).setBackgrounds(bgs);
}

/* ============================================================
 *   HELPER: stock + ceny + wydajność z „ADS_FBE"
 * ============================================================ */
function readAdsFbeMaps_() {
  const ss = SpreadsheetApp.getActive();
  const sh = ss.getSheetByName('ADS_FBE');
  const stockMap = new Map();
  const priceNetMap = new Map();
  const performanceMap = new Map();
  
  if (!sh) return { stockMap, priceNetMap, performanceMap };

  const lr = sh.getLastRow(), lc = sh.getLastColumn();
  if (lr < 2) return { stockMap, priceNetMap, performanceMap };

  const readLc = Math.max(17, lc);
  const width = Math.max(1, Math.min(readLc, sh.getMaxColumns()));
  const vals = sh.getRange(1, 1, lr, width).getValues();
  const header = vals[0].map(h => String(h || '').trim().toLowerCase());

  const idxKey = findAny_(header, ['part number key (pnk)','pnk','part_number_key','part number key']);
  if (idxKey === -1) return { stockMap, priceNetMap, performanceMap };
  
  const idxStock = findAny_(header, ['stock','stoc']);
  const colNetPrice0 = 15; // kolumna P (0-based: 15)
  const colPerformance0 = 16; // kolumna Q (0-based: 16)
  const hasPrice = colNetPrice0 < width;
  const hasPerformance = colPerformance0 < width;

  for (let r = 1; r < vals.length; r++) {
    const key = String(vals[r][idxKey] || '').trim();
    if (!key) continue;
    
    if (idxStock !== -1) {
      const num = Number(String(vals[r][idxStock] ?? '').replace(',', '.'));
      stockMap.set(key, isFinite(num) ? num : 0);
    }

    if (hasPrice) {
      const raw = vals[r][colNetPrice0];
      const num = Number(String(raw ?? '').replace(',', '.'));
      if (isFinite(num)) priceNetMap.set(key, num);
    }

    if (hasPerformance) {
      const rawPerf = vals[r][colPerformance0];
      const numPerf = Number(String(rawPerf ?? '').replace(',', '.'));
      if (isFinite(numPerf)) {
        performanceMap.set(key, numPerf);
      } else {
        const strPerf = String(rawPerf || '').trim();
        if (strPerf) performanceMap.set(key, strPerf);
      }
    }
  }
  
  return { stockMap, priceNetMap, performanceMap };
}

function readStockMap_FBE_() {
  return readAdsFbeMaps_().stockMap;
}

function readNetPriceMap_FBE_() {
  return readAdsFbeMaps_().priceNetMap;
}

/**
 * Czyta „Dostawa w drodze": PNK w kolumnie F (6), ilość w kolumnie Q (17).
 * Zwraca Map<PNK, suma_qty>
 */
function readInTransitMap_() {
  const ss = SpreadsheetApp.getActive();
  const sh = ss.getSheetByName('Dostawa w drodze');
  const map = new Map();
  if (!sh) return map;

  const lr = sh.getLastRow(), lc = Math.max(17, sh.getLastColumn());
  if (lr < 2) return map;

  const vals = sh.getRange(2, 1, lr - 1, lc).getDisplayValues();

  const IDX_PNK = 5;
  const IDX_QTY = 16;

  for (let i = 0; i < vals.length; i++) {
    const pnk = String(vals[i][IDX_PNK] || '').trim();
    if (!pnk) continue;
    const raw = String(vals[i][IDX_QTY] ?? '').replace(',', '.');
    const qty = Number(raw);
    if (!isFinite(qty) || qty <= 0) continue;

    map.set(pnk, (map.get(pnk) || 0) + qty);
  }
  return map;
}

/** Czyta arkusz "Podstawowe Informacje" z pliku partnera. */
function readPartnerProductInfoMap_() {
  const map = new Map();
  if (typeof getSS_ProdRoFbe_ !== 'function') return map;

  let ss;
  try {
    ss = getSS_ProdRoFbe_();
  } catch (e) {
    Logger.log('readPartnerProductInfoMap_: ' + (e && (e.stack || e.message) || e));
    ss = null;
  }
  if (!ss) return map;

  let sh;
  try {
    sh = ss.getSheetByName('Podstawowe Informacje');
  } catch (e) {
    Logger.log('readPartnerProductInfoMap_/sheet: ' + (e && (e.stack || e.message) || e));
    sh = null;
  }
  if (!sh) return map;

  const lr = sh.getLastRow();
  const lc = sh.getLastColumn();
  if (lr < 2 || lc < 1) return map;

  const width = Math.max(3, Math.min(sh.getMaxColumns ? sh.getMaxColumns() : lc, lc));
  const values = sh.getRange(1, 1, lr, width).getDisplayValues();
  if (!values || values.length < 2) return map;

  const header = values[0].map(h => String(h || '').trim());
  const headerLower = header.map(h => h.toLowerCase());

  let idxId  = findAny_(headerLower, ['id','product id','product_id','productid']);
  if (idxId === -1 && header.length >= 1) idxId = 0;
  let idxSku = findAny_(headerLower, ['sku','seller sku','seller_sku','sellersku','sku produktu','sku sprzedawcy']);
  if (idxSku === -1 && header.length >= 2) idxSku = 1;
  let idxEan = findAny_(headerLower, ['ean','ean13','ean_13','kod ean','kod_ean','barcode','bar code','ean code']);
  if (idxEan === -1 && header.length >= 3) idxEan = 2;
  const idxPnk = findAny_(headerLower, ['part number key (pnk)','part number key','part_number_key','pnk','part number']);

  for (let r = 1; r < values.length; r++) {
    const row = values[r];
    const id = (idxId !== -1) ? String(row[idxId] || '').trim() : '';
    const sku = (idxSku !== -1) ? String(row[idxSku] || '').trim() : '';
    const ean = (idxEan !== -1) ? String(row[idxEan] || '').trim() : '';

    if (!id && !sku && !ean) continue;

    const info = { id, sku, ean };
    const register = (val) => {
      const norm = normalizeKeyForMap_(val);
      if (!norm || map.has(norm)) return;
      map.set(norm, info);
    };

    if (idxPnk !== -1) register(row[idxPnk]);
    register(sku);
    register(id);
  }

  return map;
}

function normalizeKeyForMap_(value) {
  const s = String(value || '').trim();
  return s ? s.toUpperCase() : '';
}

function lookupProductInfo_(map, key) {
  if (!map || typeof map.get !== 'function') return null;
  const norm = normalizeKeyForMap_(key);
  if (!norm) return null;
  return map.get(norm) || null;
}

/**
 * Czyta limity wysyłki do FBE z arkusza ADS_KZ.
 */
function readAdsKzLimitMap_(){
  const map = new Map();
  const ss = SpreadsheetApp.getActive();
  const sh = ss.getSheetByName('ADS_KZ');
  if (!sh) return map;

  const lr = sh.getLastRow();
  if (lr < 2) return map;

  const vals = sh.getRange(2, 1, lr-1, 14).getValues();
  for (let i=0;i<vals.length;i++){
    const row = vals[i];
    const pnk = String(row[3] || '').trim();
    if (!pnk) continue;
    const raw = row[13];
    const num = Number(String(raw ?? '').replace(',','.'));
    if (isFinite(num)) map.set(pnk, Math.max(0, num));
  }
  return map;
}

function buildRemainThisMonthMap_(inSh) {
  const map = new Map();
  const lr = inSh.getLastRow(), lc = inSh.getLastColumn();
  if (lr < 2) return map;

  const hdr = inSh.getRange(1,1,1,lc).getDisplayValues()[0];
  const monthCols = getMonthColumns_(hdr);
  const today = new Date(), nowY = today.getFullYear(), nowM = today.getMonth()+1;
  const lastIdx = monthCols.length - 1;
  if (lastIdx < 0 || !(monthCols[lastIdx].y === nowY && monthCols[lastIdx].m === nowM)) return map;

  const day    = today.getDate();
  const daysIn = daysInMonth_(nowY, nowM);
  const colM2D = monthCols[lastIdx].colIdx0 + 1;

  const vals = inSh.getRange(2, 1, lr-1, Math.max(1, colM2D)).getValues();
  for (let i=0; i<vals.length; i++){
    const key = String(vals[i][0]||'').trim();
    if (!key) continue;
    const m2d = Number(vals[i][colM2D-1] || 0);
    const fullEst = Math.round(m2d * (daysIn / Math.max(1, day)));
    const remain  = Math.max(0, fullEst - m2d);
    map.set(key, remain);
  }
  return map;
}

/* ============================================================
 *   PROPOZYCJA ZAMÓWIENIA → NADPISUJEMY KOLUMNĘ Z (była AA)
 * ============================================================ */
function applyReorderProposalFBE(targetM){
  const ss   = SpreadsheetApp.getActive();
  const out  = ss.getSheetByName(OUT_SHEET_TRENDS);
  const inSh = ss.getSheetByName(IN_SHEET_TRENDS);
  if (!out)  throw new Error('Brak arkusza "'+OUT_SHEET_TRENDS+'". Najpierw policz trendy.');
  if (!inSh) throw new Error('Brak arkusza "'+IN_SHEET_TRENDS+'".');

  const labelDyn = `${HEADER_BASE_PROPOSAL}[${String(targetM).replace('.', ',')}M]`;
  ensureColumns_(out, COL_REORDER_PROPOSAL);
  setColumnHeaderAndClear_(out, COL_REORDER_PROPOSAL, labelDyn);

  const lrO  = out.getLastRow();
  const lcO0 = out.getLastColumn();
  if (lrO < 2) return;

  const hdrO = out.getRange(1,1,1,lcO0).getDisplayValues()[0];

  // Funkcja pomocnicza do szukania kolumny po prefiksie (obsługuje nagłówki z mnożnikami)
  const findColByPrefix = (headers, prefix) => {
    for (let i = 0; i < headers.length; i++) {
      const h = String(headers[i] || '').trim();
      if (h === prefix || h.startsWith(prefix + ' [') || h.startsWith(prefix + '[')) {
        return i + 1; // zwróć indeks 1-based
      }
    }
    return 0;
  };

  const cKey   = colIndexByHeaderSoft_FBE(hdrO, 'part_number_key');
  const cF1    = findColByPrefix(hdrO, 'Prognoza_nast_M');
  const cF2    = findColByPrefix(hdrO, 'Prognoza_+2M');
  const cF3    = findColByPrefix(hdrO, 'Prognoza_+3M');
  const cStock = colIndexByHeaderSoft_FBE(hdrO, 'Stock');
  const cSlope = colIndexByHeaderSoft_FBE(hdrO, 'Slope_6M');

  if (!(cKey && cF1 && cF2 && cF3 && cStock)) {
    Logger.log('DEBUG kolumny w Trendy_FBE:');
    Logger.log('Nagłówki: ' + JSON.stringify(hdrO));
    Logger.log('cKey=' + cKey + ', cF1=' + cF1 + ', cF2=' + cF2 + ', cF3=' + cF3 + ', cStock=' + cStock);
    throw new Error('Brakuje kolumn: part_number_key / Prognoza_* / Stock w "'+OUT_SHEET_TRENDS+'". Sprawdź logi.');
  }

  const remainMap = buildRemainThisMonthMap_(inSh);
  const limitMap  = readAdsKzLimitMap_();

  const body = out.getRange(
    2, 1, lrO-1, Math.max(lcO0, cStock, cF3, cSlope || 0)
  ).getValues();
  const keys = out.getRange(2, cKey, lrO-1, 1).getDisplayValues().map(r=>String(r[0]||'').trim());

  const result = [];
  const notes  = [];
  const M = Math.max(0, Number(String(targetM).replace(',','.')) || 0);
  const Mfull = Math.floor(M);
  const Mfrac = M - Mfull;

  for (let i=0; i<body.length; i++){
    const row  = body[i];
    const key  = keys[i];
    if (!key) { result.push(['']); notes.push(['']); continue; }

    const f1    = Math.max(0, Number(row[cF1-1]||0));
    const f2    = Math.max(0, Number(row[cF2-1]||0));
    const f3    = Math.max(0, Number(row[cF3-1]||0));
    const slope = (cSlope ? Number(row[cSlope-1]||0) : 0);

    let stock = Math.max(0, Number(row[cStock-1]||0));
    if (remainMap.has(key)) stock = Math.max(0, stock - remainMap.get(key));

    let need = 0;

    if (Mfull >= 1) need += f1;
    if (Mfull >= 2) need += f2;
    if (Mfull >= 3) need += f3;

    if (Mfull > 3) {
      let prev = f3;
      for (let k = 4; k <= Mfull; k++){
        prev = Math.max(0, Math.round(prev + slope));
        need += prev;
      }
    }

    if (Mfrac > 0) {
      let nextMonthForecast;
      if      (Mfull === 0) nextMonthForecast = f1;
      else if (Mfull === 1) nextMonthForecast = f2;
      else if (Mfull === 2) nextMonthForecast = f3;
      else                  nextMonthForecast = Math.max(0, Math.round(f3 + slope * (Mfull - 2)));
      need += nextMonthForecast * Mfrac;
    }

    const rawQty = Math.max(0, Math.ceil(need - stock));

    let finalQty = rawQty;
    let note = '';
    if (limitMap.has(key)) {
      const lim = Math.max(0, Number(limitMap.get(key)) || 0);
      if (isFinite(lim) && rawQty > lim) {
        finalQty = lim;
        note = 'maks ' + String(lim);
      }
    }

    result.push([finalQty]);
    notes.push([note]);
  }

  const rng = out.getRange(2, COL_REORDER_PROPOSAL, result.length, 1);
  rng.setValues(result);
  try { rng.setNotes(notes); } catch (_){}
}

/* ============================================================
 *   UTILITIES (lokalne)
 * ============================================================ */

function ensureColumns_(sh, colMin){
  const lc = sh.getLastColumn();
  if (lc < colMin) sh.insertColumnsAfter(lc, colMin - lc);
}

function neutralizeRangeFormatting_(rng) {
  try {
    rng.clearFormat();
    rng.clearDataValidations();
    rng.setFontColor('#000000');
    rng.setBackground(null);
    rng.setNumberFormat('@');
  } catch (_) {}
}

function toNum_(v){ const n=Number(String(v??'').replace(',','.')); return isFinite(n)?n:0; }
function isNum_(v){ return typeof v==='number' && isFinite(v); }
function round_(x,d){ const p=Math.pow(10,d|0); return Math.round((x+Number.EPSILON)*p)/p; }
function mean_(arr){ if(!arr||!arr.length) return 0; let s=0,c=0; for(const x of arr){ const n=toNum_(x); s+=n; c++; } return c? s/c : 0; }

function daysInMonth_(y,m){ return new Date(y,m,0).getDate(); }

function getMonthColumns_(hdrDispRow){
  const cols = [];
  for (let c=0; c<hdrDispRow.length; c++){
    const parsed = parseHeaderMonth_(hdrDispRow[c]);
    if (parsed) cols.push({colIdx0:c, ...parsed});
  }
  return cols;
}

function monthKey_(y, m){
  return `${y}-${String(m).padStart(2,'0')}`;
}

function parseHeaderMonth_(h){
  if (h instanceof Date && !isNaN(h)) {
    const y=h.getFullYear(), m=h.getMonth()+1;
    return { label:`${String(m).padStart(2,'0')}-${y}`, y, m };
  }
  let s = String(h||'').trim(); if (!s) return null;
  s = s.replace(/[\u2010-\u2015\u2212\uFE58\uFE63\uFF0D]/g,'-')
       .replace(/[\/.]/g,'-')
       .replace(/\s+/g,'').replace(/-+/g,'-');
  let mm, yyyy, M = s.match(/^(\d{2})-(\d{4})$/);
  if (M) { mm=Number(M[1]); yyyy=Number(M[2]); }
  else { M = s.match(/^(\d{4})-(\d{2})$/); if (M) { yyyy=Number(M[1]); mm=Number(M[2]); } else return null; }
  if (!(mm>=1 && mm<=12)) return null;
  return { label:`${String(mm).padStart(2,'0')}-${yyyy}`, y:yyyy, m:mm };
}
function sumTail_(arr, endIdx, win){ let s=0; for (let i=Math.max(0,endIdx-win+1); i<=endIdx; i++) s+=toNum_(arr[i]); return s; }
function avgFromSum_(sum, cnt){ return cnt>0 ? sum/cnt : 0; }
function tailWindow_(arr,endIdx,win){ const st=Math.max(0,endIdx-win+1); return arr.slice(st,endIdx+1).map(toNum_); }
function linreg_(y){
  const n=y.length; let sx=0, sy=0, sxy=0, sxx=0;
  for (let i=0;i<n;i++){ const xi=i, yi=toNum_(y[i]); sx+=xi; sy+=yi; sxy+=xi*yi; sxx+=xi*xi; }
  const denom = n*sxx - sx*sx; if (!denom) return { slope:0, intercept:mean_(y) };
  const slope = (n*sxy - sx*sy) / denom; const intercept=(sy - slope*sx) / n; return { slope, intercept };
}
function colIndexByHeaderSoft_FBE(hdrRow, exactHeader) {
  if (!Array.isArray(hdrRow)) return 0;
  
  const pozycja = hdrRow.findIndex(function(h) {
    return String(h || '').trim() === String(exactHeader || '').trim();
  });
  
  return (pozycja >= 0) ? (pozycja + 1) : 0;
}
function findAny_(lowerHdrRow, candidatesLower){
  for (const name of candidatesLower){
    const i = lowerHdrRow.indexOf(String(name).toLowerCase());
    if (i !== -1) return i;
  }
  return -1;
}
function ensureCoverageFormatting_(sh) {
  if (!sh) return;
  const lastRow = Math.max(2, sh.getLastRow());
  const hdr = sh.getRange(1, 1, 1, Math.max(1, sh.getLastColumn())).getDisplayValues()[0];

  const cCover = colIndexByHeaderSoft_FBE(hdr, 'Pokrycie_magazynu');
  if (!cCover) return;

  const rng = sh.getRange(2, cCover, lastRow - 1, 1);

  const rules = sh.getConditionalFormatRules();
  let keep = rules.filter(rule => !rule.getRanges().some(r =>
    r.getColumn() <= cCover && cCover < r.getColumn() + r.getNumColumns()
  ));

  const cQ = colIndexByHeaderSoft_FBE(hdr, 'Cena_netto');
  const cR = colIndexByHeaderSoft_FBE(hdr, 'Cena_brutto');
  keep = keep.filter(rule => !rule.getRanges().some(r => {
    const c1 = r.getColumn(), c2 = c1 + r.getNumColumns() - 1;
    const touchesQ = cQ && c1 <= cQ && cQ <= c2;
    const touchesR = cR && c1 <= cR && cR <= c2;
    return touchesQ || touchesR;
  }));

  const red = SpreadsheetApp.newConditionalFormatRule()
    .whenTextEqualTo('nie wystarczy na pokrycie sprzedaży w tym miesiącu')
    .setBackground('#fde7e9')
    .setRanges([rng]).build();

  const green1 = SpreadsheetApp.newConditionalFormatRule()
    .whenTextEqualTo('wystarczy na jeden miesiąc')
    .setBackground('#e6f4ea')
    .setRanges([rng]).build();

  const green2 = SpreadsheetApp.newConditionalFormatRule()
    .whenTextEqualTo('wystarczy na pokrycie sprzedaży na dwa miesiące')
    .setBackground('#e6f4ea')
    .setRanges([rng]).build();

  const green3 = SpreadsheetApp.newConditionalFormatRule()
    .whenTextEqualTo('wystarczy na 3 miesiące lub więcej')
    .setBackground('#e6f4ea')
    .setRanges([rng]).build();

  sh.setConditionalFormatRules(keep.concat([red, green1, green2, green3]));
}

/* ============================================================
 *   AKCJE Z MENU
 * ============================================================ */
function cmd_reorderProposalFBE(){
  const ui = SpreadsheetApp.getUi();
  const resp = ui.prompt(
    'Propozycja zamówienia (FBE)',
    'Podaj docelowe pokrycie w miesiącach (może być ułamek, np. 1.5).',
    ui.ButtonSet.OK_CANCEL
  );
  if (resp.getSelectedButton() !== ui.Button.OK) return;
  const targetM = Number(String(resp.getResponseText()||'').replace(',','.'));
  if (!isFinite(targetM) || targetM <= 0) { ui.alert('Podaj dodatnią liczbę (np. 2 lub 1.5).'); return; }

  const lock = LockService.getDocumentLock();
  if (!lock.tryLock(8000)) { ui.alert('Program jest zajęty. Spróbuj za chwilę.'); return; }
  try {
    if (typeof ensureAuthWithPartner_ === 'function') ensureAuthWithPartner_();
    if (typeof requireActivePartner_ === 'function')  requireActivePartner_();
    applyReorderProposalFBE(targetM);
    SpreadsheetApp.getActive().toast('Propozycja zamówienia gotowa ✅', 'Programy', 5);
  } catch (e) {
    ui.alert('Błąd propozycji zamówienia:\n' + (e && (e.stack || e.message) || e));
  } finally { lock.releaseLock(); }
}

// ========= Custom kolumny z ADS_FBE → Trendy_FBE (po PNK) =========
const CUSTOM_COL_PREFIX = 'ADS: ';
const CUSTOM_COL_MAX    = 10;

function cmd_addCustomColumns() {
  try {
    ensureAuthWithPartner_();
    requireActivePartner_();
  } catch (e) {
    SpreadsheetApp.getUi().alert('Błąd autoryzacji/partnera:\n' + (e && (e.message || e)));
    return;
  }
  const t = HtmlService.createTemplateFromFile('AddCustomColumns');
  const html = t.evaluate().setWidth(520).setHeight(420);
  SpreadsheetApp.getUi().showModalDialog(html, 'Dodaj własne kolumny z ADS_FBE');
}

function getAdsFbeHeadersForPicker_() {
  const ss = SpreadsheetApp.getActive();
  const sh = ss.getSheetByName('ADS_FBE');
  if (!sh) throw new Error('Brak arkusza "ADS_FBE".');
  const lc = sh.getLastColumn();
  if (lc < 1) return [];
  const hdr = sh.getRange(1, 1, 1, lc).getDisplayValues()[0];
  return hdr.map(h => String(h || '').trim()).filter(Boolean);
}

function applySelectedAdsFbeColumns(selectedHeaders) {
  if (!Array.isArray(selectedHeaders)) selectedHeaders = [];
  const uniq = Array.from(new Set(selectedHeaders.map(s => String(s || '').trim()).filter(Boolean))).slice(0, CUSTOM_COL_MAX);

  const ss    = SpreadsheetApp.getActive();
  const outSh = ss.getSheetByName(OUT_SHEET_TRENDS);
  const adsSh = ss.getSheetByName('ADS_FBE');
  if (!outSh) throw new Error('Brak arkusza "' + OUT_SHEET_TRENDS + '".');
  if (!adsSh) throw new Error('Brak arkusza "ADS_FBE".');

  const lrO = outSh.getLastRow(), lcO = outSh.getLastColumn();
  if (lrO < 2) throw new Error('Arkusz "' + OUT_SHEET_TRENDS + '" jest pusty.');

  (function removeOldCustom_(){
    const hdr = outSh.getRange(1, 1, 1, lcO).getDisplayValues()[0];
    const toDelete = [];
    for (let i = 0; i < hdr.length; i++) {
      const h = String(hdr[i] || '');
      if (h.startsWith(CUSTOM_COL_PREFIX)) toDelete.push(i + 1);
    }
    for (let k = toDelete.length - 1; k >= 0; k--) {
      try { outSh.deleteColumn(toDelete[k]); } catch(_) {}
    }
  })();

  if (uniq.length === 0) {
    SpreadsheetApp.getActive().toast('Usunięto poprzednie kolumny ADS. Nic nie dodano.', 'Programy', 4);
    return;
  }

  const lrA = adsSh.getLastRow(), lcA = adsSh.getLastColumn();
  if (lrA < 2 || lcA < 1) throw new Error('"ADS_FBE" nie zawiera danych.');
  const hdrA = adsSh.getRange(1, 1, 1, lcA).getDisplayValues()[0].map(x => String(x || '').trim());
  const idxByHeader = new Map(); hdrA.forEach((name, i) => { if (name) idxByHeader.set(name, i); });

  const lowerA = hdrA.map(h => h.toLowerCase());
  const idxKeyAds0 = (function() {
    const cands = ['part number key (pnk)','pnk','part_number_key','part number key'];
    for (const c of cands) { const i = lowerA.indexOf(c); if (i !== -1) return i; }
    return -1;
  })();
  if (idxKeyAds0 === -1) throw new Error('W "ADS_FBE" nie znaleziono kolumny PNK.');

  const hdrOut = outSh.getRange(1, 1, 1, outSh.getLastColumn()).getDisplayValues()[0];
  const idxKeyOut1 = colIndexByHeaderSoft_FBE(hdrOut, 'part_number_key');
  if (!idxKeyOut1) throw new Error('W "' + OUT_SHEET_TRENDS + '" nie znaleziono kolumny "part_number_key".');

  const adsVals = adsSh.getRange(2, 1, lrA - 1, lcA).getDisplayValues();
  const rowByPNK = new Map();
  for (let r = 0; r < adsVals.length; r++) {
    const k = String(adsVals[r][idxKeyAds0] || '').trim();
    if (k && !rowByPNK.has(k)) rowByPNK.set(k, adsVals[r]);
  }

  const keysOut = outSh.getRange(2, idxKeyOut1, lrO - 1, 1).getDisplayValues().map(r => String(r[0] || '').trim());

  let insertAt = outSh.getLastColumn() + 1;
  const insertedCols = [];

  uniq.forEach(selHdr => {
    const colIdx0 = idxByHeader.get(selHdr);
    if (typeof colIdx0 !== 'number') return;

    if (outSh.getLastColumn() < insertAt) {
      outSh.insertColumnAfter(outSh.getLastColumn());
    }
    const col1 = insertAt; insertAt++;

    outSh.getRange(1, col1).setValue(CUSTOM_COL_PREFIX + selHdr).setFontWeight('bold');

    const dataRng = outSh.getRange(2, col1, keysOut.length, 1);
    neutralizeRangeFormatting_(dataRng);

    const colData = keysOut.map(key => {
      const row = rowByPNK.get(key);
      return [row ? row[colIdx0] : ''];
    });
    dataRng.setValues(colData);

    insertedCols.push(col1);
  });

  try {
    if (insertedCols.length) {
      const rules = outSh.getConditionalFormatRules();
      const keep = rules.filter(rule => !rule.getRanges().some(r => {
        const c1 = r.getColumn();
        const c2 = c1 + r.getNumColumns() - 1;
        return insertedCols.some(col => c1 <= col && col <= c2);
      }));
      outSh.setConditionalFormatRules(keep);
    }
  } catch (_) {}

  try { ensureBoldHeadersAllSheets_ && ensureBoldHeadersAllSheets_(); } catch(_) {}
  SpreadsheetApp.getActive().toast('Dodano kolumny: ' + uniq.join(', '), 'Programy', 5);
}

function getAdsFbeHeadersForPicker() {
  return getAdsFbeHeadersForPicker_();
}

function debugColIndex() {
  Logger.log('=== TEST 1: Czy funkcja istnieje? ===');
  Logger.log('Typ: ' + typeof colIndexByHeaderSoft_);
  
  Logger.log('\n=== TEST 2: Prosty test ===');
  const result = colIndexByHeaderSoft_FBE(['A', 'B', 'C'], 'B');
  Logger.log('Wynik: ' + result + ' (oczekiwane: 2)');
  Logger.log('Typ wyniku: ' + typeof result);
  
  Logger.log('\n=== TEST 3: Z prawdziwymi danymi ===');
  const ss = SpreadsheetApp.getActive();
  const sh = ss.getSheetByName('Trendy_FBE');
  if (!sh) {
    Logger.log('❌ Brak arkusza Trendy_FBE');
    return;
  }
  
  const headers = sh.getRange(1, 1, 1, sh.getLastColumn()).getDisplayValues()[0];
  Logger.log('Nagłówki z arkusza: ' + JSON.stringify(headers));
  
  Logger.log('\n=== TEST 4: Szukanie kolumn ===');
  const cKey = colIndexByHeaderSoft_FBE(headers, 'part_number_key');
  Logger.log('part_number_key → ' + cKey + ' (typ: ' + typeof cKey + ')');
  
  const cStock = colIndexByHeaderSoft_FBE(headers, 'Stock');
  Logger.log('Stock → ' + cStock + ' (typ: ' + typeof cStock + ')');
  
  const cF1 = colIndexByHeaderSoft_FBE(headers, 'Prognoza_nast_M');
  Logger.log('Prognoza_nast_M → ' + cF1 + ' (typ: ' + typeof cF1 + ')');
  
  if (cKey && cStock && cF1) {
    Logger.log('\n✅ WSZYSTKIE KOLUMNY ZNALEZIONE!');
  } else {
    Logger.log('\n❌ PROBLEM: Funkcja zwraca undefined mimo że kolumny istnieją');
  }
}

/* ============================================================
 *   ROTACJA / POZOSTAŁE SZTUKI → NADPISUJEMY KOLUMNĘ AB
 * ============================================================ */
function applyRotationCheckFBE(targetM){
  const ss   = SpreadsheetApp.getActive();
  const out  = ss.getSheetByName(OUT_SHEET_TRENDS);
  const inSh = ss.getSheetByName(IN_SHEET_TRENDS);
  if (!out)  throw new Error('Brak arkusza "'+OUT_SHEET_TRENDS+'". Najpierw policz trendy.');
  if (!inSh) throw new Error('Brak arkusza "'+IN_SHEET_TRENDS+'".');

  ensureColumns_(out, COL_ROTATION_LEFTOVER);

  const lrO = out.getLastRow(), lcO0 = out.getLastColumn();
  if (lrO < 2) return;

  const hdrO = out.getRange(1,1,1,lcO0).getDisplayValues()[0];

  // Funkcja pomocnicza do szukania kolumny po prefiksie
  const findColByPrefix = (headers, prefix) => {
    for (let i = 0; i < headers.length; i++) {
      const h = String(headers[i] || '').trim();
      if (h === prefix || h.startsWith(prefix + ' [') || h.startsWith(prefix + '[')) {
        return i + 1;
      }
    }
    return 0;
  };

  const cKey    = colIndexByHeaderSoft_FBE(hdrO, 'part_number_key');
  const cF1     = findColByPrefix(hdrO, 'Prognoza_nast_M');
  const cF2     = findColByPrefix(hdrO, 'Prognoza_+2M');
  const cF3     = findColByPrefix(hdrO, 'Prognoza_+3M');
  const cSlope  = colIndexByHeaderSoft_FBE(hdrO, 'Slope_6M');
  const cStock  = colIndexByHeaderSoft_FBE(hdrO, 'Stock');
  
  if (!(cKey && cF1 && cF2 && cF3 && cSlope && cStock)) {
    Logger.log('DEBUG kolumny w Trendy_FBE (rotacja):');
    Logger.log('Nagłówki: ' + JSON.stringify(hdrO));
    Logger.log('cKey=' + cKey + ', cF1=' + cF1 + ', cF2=' + cF2 + ', cF3=' + cF3 + ', cSlope=' + cSlope + ', cStock=' + cStock);
    throw new Error('Brakuje kolumn: part_number_key / Prognoza_* / Slope_6M / Stock w "'+OUT_SHEET_TRENDS+'". Sprawdź logi.');
  }

  const label = `Sztuki_pozostałe_po_[${String(targetM).replace('.',',')}M]`;
  setColumnHeaderAndClear_(out, COL_ROTATION_LEFTOVER, label);

  const body = out.getRange(2, 1, lrO-1, Math.max(lcO0, cStock, cF3, cSlope)).getValues();
  const result = [];
  const M = Math.max(0, Number(targetM) || 0);
  const Mfull = Math.floor(M);
  const Mfrac = M - Mfull;

  for (let i=0; i<body.length; i++){
    const row    = body[i];
    const f1     = Math.max(0, Number(row[cF1-1]||0));
    const f2     = Math.max(0, Number(row[cF2-1]||0));
    const f3     = Math.max(0, Number(row[cF3-1]||0));
    const slope  = Number(row[cSlope-1]||0);
    let   stock  = Math.max(0, Number(row[cStock-1]||0));

    let need = 0;

    if (Mfull >= 1) need += f1;
    if (Mfull >= 2) need += f2;
    if (Mfull >= 3) need += f3;

    if (Mfull > 3) {
      let prev = f3;
      for (let k = 4; k <= Mfull; k++){
        prev = Math.max(0, Math.round(prev + slope));
        need += prev;
        if (stock - need <= 0) break;
      }
    }

    if (Mfrac > 0) {
      let nextMonthForecast;
      if (Mfull === 0)       nextMonthForecast = f1;
      else if (Mfull === 1)  nextMonthForecast = f2;
      else if (Mfull === 2)  nextMonthForecast = f3;
      else                   nextMonthForecast = Math.max(0, Math.round(f3 + slope * (Mfull - 2)));
      need += nextMonthForecast * Mfrac;
    }

    const leftover = Math.max(0, Math.ceil(stock - need));
    result.push([leftover]);
  }

  out.getRange(2, COL_ROTATION_LEFTOVER, result.length, 1).setValues(result);

  try {
    const rng = out.getRange(2, COL_ROTATION_LEFTOVER, result.length, 1);
    const bg = result.map(([v]) => [Number(v||0) === 0 ? '#e6f4ea' : '#fde7e9']);
    rng.setBackgrounds(bg);
  } catch(_) {}
}

function cmd_rotationCheckFBE(){
  const ui = SpreadsheetApp.getUi();
  const resp = ui.prompt(
    'Rotacja / pozostałe sztuki (FBE)',
    'Podaj horyzont w miesiącach (może być ułamek, np. 2.5). Po tym czasie pokażemy, ile sztuk zostanie.',
    ui.ButtonSet.OK_CANCEL
  );
  if (resp.getSelectedButton() !== ui.Button.OK) return;
  const targetM = Number(String(resp.getResponseText()||'').replace(',','.'));
  if (!isFinite(targetM) || targetM <= 0) { ui.alert('Podaj dodatnią liczbę (np. 2 lub 1.5).'); return; }

  const lock = LockService.getDocumentLock();
  if (!lock.tryLock(8000)) { ui.alert('Program jest zajęty. Spróbuj za chwilę.'); return; }
  try {
    if (typeof ensureAuthWithPartner_ === 'function') ensureAuthWithPartner_();
    if (typeof requireActivePartner_ === 'function')  requireActivePartner_();
    applyRotationCheckFBE(targetM);
    SpreadsheetApp.getActive().toast('Obliczono pozostałe sztuki ✅', 'Programy', 5);
  } catch (e) {
    ui.alert('Błąd kalkulacji rotacji:\n' + (e && (e.stack || e.message) || e));
  } finally { lock.releaseLock(); }
}
