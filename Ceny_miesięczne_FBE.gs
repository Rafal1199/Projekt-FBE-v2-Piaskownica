/************************************************************
 * Ceny_miesięczne_FBE — średnie ceny miesięczne (od 01-2025)
 * Źródło: aktywny partner → plik RO/FBE → arkusz "Zamówienia"
 *  - Status (kol. J) ∈ {1,2,3,4}
 *  - Pozycje zamówień (kol. M, JSON) → part_number_key, sale_price,
 *    quantity, created (YYYY-MM-DD hh:mm:ss)
 * Wynik: arkusz "Ceny_miesięczne_FBE"
 *  - A: part_number_key
 *  - B..: MM-YYYY (od 01-2025 do ostatniego miesiąca w danych)
 *  - komórki: średnia ważona cena netto (po quantity), 2 miejsca po przecinku
 *  - komentarz: lista unikalnych cen użytych w miesiącu (po przecinku)
 ************************************************************/

// ===== LOKALNE STAŁE (unikalne nazwy, żeby nie kolidowały z innymi plikami) =====
const PRICING_OUT_SHEET     = 'Ceny_miesięczne_FBE';
const PRICING_ORDERS_SHEET  = 'Zamówienia';
const PRICING_ALLOWED_STAT  = new Set([1,2,3,4]); // 0/5 pomijamy
const PRICING_FIRST_Y       = 2025;
const PRICING_FIRST_M       = 1;                  // 1 = styczeń

/**
 * GŁÓWNA: przelicz średnie ceny miesięczne (wykorzystuje aktywnego partnera).
 * - Nie pyta o link — bierze z „Partnerzy” tej samej logiki co inne moduły.
 * - Uruchamiaj PO przeliczeniu zamówień i PRZED trendami.
 */
function recomputeMonthlyPricesFBE() {
  // Używamy Twoich istniejących helperów z pliku „Przyciski.gs”
  ensureAuthWithPartner_();
  const partnerId = requireActivePartner_();

  // 1) Otwórz plik źródłowy partnera (RO/FBE)
  const srcSS = getSS_RoFbe_();
  if (!srcSS) throw new Error('Brak pliku źródłowego RO/FBE dla aktywnego partnera: ' + partnerId);

  const ordSh = srcSS.getSheetByName(PRICING_ORDERS_SHEET);
  if (!ordSh) throw new Error('W pliku partnera brak arkusza "' + PRICING_ORDERS_SHEET + '".');

  // 2) Wczytaj zamówienia
  const lr = ordSh.getLastRow(), lc = ordSh.getLastColumn();
  if (lr < 2) {
    writeEmptyPricingTable_();
    SpreadsheetApp.getActive().toast('Brak danych zamówień. Utworzono pusty nagłówek.', 'Ceny', 4);
    return;
  }

  const hdr = ordSh.getRange(1, 1, 1, lc).getDisplayValues()[0].map(x => String(x||'').trim().toLowerCase());
  const idxStatus0  = findCol_(hdr, ['status','order status','status zamówienia','j']);     // J
  const idxProducts0= findCol_(hdr, ['products','pozycje','m']);                            // M
  if (idxStatus0 === -1 || idxProducts0 === -1) {
    throw new Error('Nie znaleziono kolumn: Status (J) i/lub Products (M) w "'+PRICING_ORDERS_SHEET+'".');
  }

  const data = ordSh.getRange(2, 1, lr-1, lc).getDisplayValues();

  // 3) Agreguj: { pnk → { yyyymm → {sumPQ, sumQ, prices:Set}} }
  const agg = new Map();
  let maxMonthKey = null;

  for (let i=0; i<data.length; i++) {
    const row = data[i];
    const status = PRICING_normalizeStatus_(row[idxStatus0]);
    if (status === null || !PRICING_ALLOWED_STAT.has(status)) continue;

    const items = PRICING_parseProductsCell_(row[idxProducts0]);
    if (!items.length) continue;

    for (const it of items) {
      const pnk  = String(it.partKey || it.part_number_key || '').trim();
      if (!pnk) continue;

      const qty  = Number(it.quantity || it.initial_qty || it.qty || 0);
      const price= PRICING_toNum_(it.sale_price);
      if (!(qty > 0) || !(price > 0)) continue;

      // miesiąc z pola created; jeśli brak, pomiń
      const created = PRICING_parseDate_(it.created);
      if (!created) continue;
      const y = created.getFullYear(), m = created.getMonth()+1;
      if (y < PRICING_FIRST_Y || (y === PRICING_FIRST_Y && m < PRICING_FIRST_M)) continue;

      const keyYM = y + '-' + String(m).padStart(2, '0');
      if (!maxMonthKey || maxMonthKey < keyYM) maxMonthKey = keyYM;

      let pnkMap = agg.get(pnk);
      if (!pnkMap) { pnkMap = new Map(); agg.set(pnk, pnkMap); }

      let mm = pnkMap.get(keyYM);
      if (!mm) { mm = { sumPQ:0, sumQ:0, byPrice:new Map() }; pnkMap.set(keyYM, mm); }

      mm.sumPQ += price * qty;
      mm.sumQ  += qty;
      mm.byPrice.set(price, (mm.byPrice.get(price) || 0) + qty);
    }
  }

  // 4) Jeśli brak danych → pusty nagłówek
  if (!agg.size) {
    writeEmptyPricingTable_();
    SpreadsheetApp.getActive().toast('Brak danych po filtrach (statusy 1–4, od 01-2025).', 'Ceny', 4);
    return;
  }

  // 5) Oś czasu od 01-2025 do maxMonthKey (włącznie)
  const maxY = Number(maxMonthKey.slice(0,4));
  const maxM = Number(maxMonthKey.slice(5,7));
  const months = PRICING_buildMonthAxis_(new Date(PRICING_FIRST_Y, PRICING_FIRST_M-1, 1), {y:maxY, m:maxM});
  const monthKeys = months.map(({y,m}) => y + '-' + String(m).padStart(2,'0'));
  const monthLabels = months.map(({y,m}) => (String(m).padStart(2,'0') + '-' + y));

  // 6) Przygotuj arkusz wynikowy i zapisz
  const out = SpreadsheetApp.getActive().getSheetByName(PRICING_OUT_SHEET) || SpreadsheetApp.getActive().insertSheet(PRICING_OUT_SHEET);

  // wyczyść poprzednią zawartość (zostawimy same treści/komentarze; formaty nagłówka możesz dopieścić globalnie)
  const lrOut = Math.max(2, out.getLastRow());
  const lcOut = Math.max(1, out.getLastColumn());
  try { out.getRange(1,1,lrOut,lcOut).clearContent().clearNote(); } catch(_){}

  // nagłówki
  const header = ['part_number_key'].concat(monthLabels);
  out.getRange(1,1,1, header.length).setValues([header]);
  out.setFrozenRows(1);

  // ciało
  const pnks = Array.from(agg.keys()).sort();
  const rows = [];
  const notes = [];

  for (const pnk of pnks) {
    const pnkMap = agg.get(pnk);
    const row = [pnk];
    const noteRow = [''];

    for (const ym of monthKeys) {
      const rec = pnkMap.get(ym);
      if (!rec || !(rec.sumQ > 0)) {
        row.push('');
        noteRow.push('');
      } else {
        const avg = PRICING_round2_(rec.sumPQ / rec.sumQ);
        row.push(avg);

        // komentarz: unikalne ceny rosnąco, po przecinku
        const list = Array.from(rec.byPrice.entries())
          .sort((a,b) => a[0] - b[0])                         // sort po cenie
          .map(([p, q]) => `${PRICING_fmt2_(p)}×${q}`);       // np. 125.21×3
        noteRow.push(list.join(', '));
      }
    }
    rows.push(row);
    notes.push(noteRow);
  }

  // Zapis wartości
  setValuesChunked_(out, 2, 1, rows, 20000);
  // Zapis komentarzy (notatek)
  try {
    const rngNotes = out.getRange(2, 1, notes.length, notes[0].length);
    rngNotes.setNotes(notes);
  } catch(_) {}

  // Format liczbowy dla zakresu cen
  try {
    const c1 = 2, r1 = 2;
    const nRows = rows.length, nCols = monthLabels.length;
    if (nRows > 0 && nCols > 0) {
      out.getRange(r1, c1, nRows, nCols).setNumberFormat('0.00');
    }
  } catch(_) {}

  SpreadsheetApp.getActive().toast('Ceny miesięczne przeliczone ✅', 'Ceny', 5);
}

/* ========================= HELPERY (lokalne, bez kolizji) ========================= */

function PRICING_buildMonthAxis_(fromDate, toYM) {
  // używa Twojej monthRange_ jeśli istnieje; w razie czego fallback
  if (typeof monthRange_ === 'function') {
    const end = new Date(toYM.y, toYM.m-1, 1);
    return monthRange_(fromDate.getFullYear(), fromDate.getMonth()+1, end)
      .map(k => ({y: Number(k.key.slice(0,4)), m: Number(k.key.slice(5,7))}));
  }
  // fallback
  const out = [];
  let y = fromDate.getFullYear(), m = fromDate.getMonth()+1;
  while (y < toYM.y || (y === toYM.y && m <= toYM.m)) {
    out.push({y, m});
    m++; if (m>12){ m=1; y++; }
  }
  return out;
}

function PRICING_normalizeStatus_(v) {
  if (typeof normalizeStatus_ === 'function') {
    const n = normalizeStatus_(v);
    return (typeof n === 'number' && Number.isFinite(n)) ? Math.trunc(n) : null;
  }

  if (v === null || typeof v === 'undefined' || v === '') return null;

  if (v instanceof Date) {
    const msPerDay = 24 * 60 * 60 * 1000;
    const epoch = Date.UTC(1899, 11, 30);
    const utc = Date.UTC(v.getFullYear(), v.getMonth(), v.getDate());
    const serial = Math.round((utc - epoch) / msPerDay);
    return Number.isFinite(serial) ? serial : null;
  }

  if (typeof v === 'number' && Number.isFinite(v)) {
    return Math.trunc(v);
  }

  const s = String(v).trim();
  const m = s.match(/-?\d+/);
  if (!m) return null;
  const n = Number(m[0]);
  return Number.isFinite(n) ? Math.trunc(n) : null;
}

function PRICING_parseProductsCell_(cell) {
  // Kolumna M: JSON lub „prawie JSON”.
  // Trzymamy się Twojej konwencji parsowania z extractItemsFromProducts_, ale tutaj
  // chcemy mieć pewność na sale_price/quantity/created/part_number_key
  try {
    if (cell == null) return [];
    if (typeof cell === 'object' && !Array.isArray(cell)) return [cell];
    if (Array.isArray(cell)) return cell;

    const s = String(cell).trim();
    if (!s) return [];

    // najpierw próbuj jako JSON
    if ((s.startsWith('[') && s.endsWith(']')) || (s.startsWith('{') && s.endsWith('}'))) {
      try {
        const obj = JSON.parse(s);
        return Array.isArray(obj) ? obj : [obj];
      } catch(_){}
    }

    // fallback – lekka normalizacja cudzysłowów dla prostych przypadków
    const fixed = s
      .replace(/([{,]\s*)'([^']+?)'\s*:/g, '$1"$2":')   // 'key': → "key":
      .replace(/:\s*'([^']*?)'(\s*[},])/g, ':"$1"$2');   // : 'val' → :"val"
    const obj2 = JSON.parse(fixed);
    return Array.isArray(obj2) ? obj2 : [obj2];
  } catch(e) {
    // ostatecznie spróbuj użyć istniejącego extractItemsFromProducts_ (bez cen)
    try {
      if (typeof extractItemsFromProducts_ === 'function') {
        return extractItemsFromProducts_(cell);
      }
    } catch(_){}
    return [];
  }
}

function PRICING_toNum_(v){
  const n = Number(String(v ?? '').replace(',', '.'));
  return isFinite(n) ? n : 0;
}
function PRICING_fmt2_(n){ return (Math.round((n+Number.EPSILON)*100)/100).toFixed(2); }
function PRICING_round2_(n){ return Math.round((n+Number.EPSILON)*100)/100; }

function PRICING_parseDate_(v){
  // jeśli masz globalne normalizeDate_ — użyj jej
  try {
    if (typeof normalizeDate_ === 'function') return normalizeDate_(v);
  } catch(_){}
  const s = String(v || '').trim();
  if (!s) return null;
  const d = new Date(s);
  return (d instanceof Date && !isNaN(d)) ? new Date(d.getFullYear(), d.getMonth(), d.getDate()) : null;
}

function writeEmptyPricingTable_(outSh = (SpreadsheetApp.getActive().getSheetByName(PRICING_OUT_SHEET) || SpreadsheetApp.getActive().insertSheet(PRICING_OUT_SHEET))) {
  // domyślna oś czasu do końca bieżącego miesiąca
  const today = new Date();
  const months = PRICING_buildMonthAxis_(new Date(PRICING_FIRST_Y, PRICING_FIRST_M-1, 1), {y:today.getFullYear(), m:today.getMonth()+1})
                   .map(({y,m}) => `${String(m).padStart(2,'0')}-${y}`);
  outSh.getRange(1,1,1, 1 + months.length).setValues([['part_number_key'].concat(months)]);
  outSh.setFrozenRows(1);
}

function setValuesChunked_(sheet, r1, c1, data, batchRows) {
  // korzysta z Twojej wersji, ale dodaję fallback gdyby nie było w zasięgu
  if (typeof this.setValuesChunked_ === 'function' && this.setValuesChunked_ !== setValuesChunked_) {
    return this.setValuesChunked_.apply(this, arguments);
  }
  batchRows = batchRows || 20000;
  if (!data || !data.length) return;
  let i = 0;
  while (i < data.length) {
    const part = data.slice(i, i + batchRows);
    sheet.getRange(r1 + i, c1, part.length, part[0].length).setValues(part);
    i += part.length;
  }
}
