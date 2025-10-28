/**
 * Klocek #1 — Sprzedaż_okresowa_FBE
 * Czyta „Zamówienia” z pliku [Partner] RO FBE i buduje „Sprzedaż_okresowa_FBE” w AKTYWNYM skoroszycie.
 * Dodatek: filtruje zamówienia po statusie z kolumny J (lub po nagłówku „status” itp.).
 * Uwzględniane statusy: 1,2,3,4 (pomija 0 i 5).
 */

// ───────────────────────────────────────────────────────────────────────────────
// USTAWIENIA
// ───────────────────────────────────────────────────────────────────────────────

const SRC_SALES_SHEET   = 'Zamówienia';
const OUT_SHEET         = 'Sprzedaż_okresowa_FBE';

const FALLBACK_COL_DATE   = 5;   // jeśli nie znajdzie kolumny daty po nagłówkach → E (1-indexed)
const FALLBACK_COL_STATUS = 10;  // jeśli nie znajdzie kolumny statusu po nagłówkach → J (1-indexed)

const START_YEAR  = 2025;        // zakres miesięcy od 01-2025 do „dziś”
const START_MONTH = 1;

const SETVALUES_BATCH = 20000;   // Limit Google API (porcjowanie)

// dozwolone statusy zamówień
const ALLOWED_ORDER_STATUSES = new Set([1, 2, 3, 4]);  // Set = szybkie sprawdzanie

// ───────────────────────────────────────────────────────────────────────────────
// POMOCNIKI LOKALNE (bez zależności na resztę projektu)
// ───────────────────────────────────────────────────────────────────────────────

/**
 * Normalizuje pole „status” do liczby całkowitej.
 * Obsługa:
 *  - Number → int
 *  - String → pierwsza liczba w tekście (np. "1 - shipped" → 1)
 *  - Date   → numer seryjny arkusza (np. 1..4 gdy kolumna błędnie sformatowana jako data)
 * Zwraca null, jeśli brak możliwej liczby.
 */
function normalizeStatus_(v) {
  if (v === null || typeof v === 'undefined' || v === '') return null;

  // Data → numer seryjny arkusza (Sheets epoch 1899-12-30)
  if (v instanceof Date) {
    const msPerDay = 24 * 60 * 60 * 1000;
    const epoch = Date.UTC(1899, 11, 30);
    const utc = Date.UTC(v.getFullYear(), v.getMonth(), v.getDate());
    const serial = Math.round((utc - epoch) / msPerDay);
    return Number.isFinite(serial) ? serial : null;
  }

  // Liczba
  if (typeof v === 'number' && Number.isFinite(v)) {
    return Math.trunc(v);
  }

  // Tekst → pierwsza liczba całkowita (dopuszcza minus, na wszelki wypadek)
  const s = String(v).trim();
  const m = s.match(/-?\d+/);
  if (!m) return null;
  const n = Number(m[0]);
  return Number.isFinite(n) ? Math.trunc(n) : null;
}

// ───────────────────────────────────────────────────────────────────────────────
// GŁÓWNA FUNKCJA
// ───────────────────────────────────────────────────────────────────────────────

function buildSalesPeriods() {
  requireActivePartner_();
  const t0 = new Date();
  const home = getHome_();

  // 1) Źródło z rejestru Partnerzy (aktywny Partner RO FBE)
  const srcSS = getSS_RoFbe_();
  if (!srcSS) throw new Error('Brak URL dla [Partner] RO FBE w „Partnerzy”.');
  const srcSh = srcSS.getSheetByName(SRC_SALES_SHEET);
  if (!srcSh) throw new Error(`W źródle brak zakładki "${SRC_SALES_SHEET}".`);

  // 2) Dane
  const vals = withRetry_(() => srcSh.getDataRange().getValues(), 'getValues');
  if (vals.length < 2) throw new Error('Źródło nie zawiera danych (nagłówki + min. 1 wiersz).');

  const header = vals[0].map(h => String(h || '').trim().toLowerCase());

  // 3) Indeksy kolumn
  const idxProducts = findCol_(header, ['products']);
  if (idxProducts === -1) throw new Error('Nie znaleziono kolumny "products" w źródle.');

  const idxDateHdr = findCol_(header, ['created', 'order_date', 'created_at', 'data zamówienia', 'data_zamowienia']);
  const idxDate = (idxDateHdr !== -1) ? idxDateHdr : (FALLBACK_COL_DATE - 1);
  if (idxDate < 0) throw new Error('Nie znaleziono kolumny z datą (ani po nagłówkach, ani fallback do E).');

  // kolumna statusu zamówienia (po nagłówkach lub fallback do J)
  const idxStatusHdr = findCol_(header, ['status', 'order_status', 'status zamówienia', 'status_zamowienia']);
  const idxStatus = (idxStatusHdr !== -1) ? idxStatusHdr : (FALLBACK_COL_STATUS - 1);

  // 4) Zakres miesięcy
  const months = monthRange_(START_YEAR, START_MONTH, new Date());
  const monthIndex = new Map(months.map((m, i) => [m.key, i]));

  // 5) Agregacja
  const agg = new Map(); // key -> { allQty:number, monthly:number[] }
  let kept = 0, dropped = 0;

  for (let r = 1; r < vals.length; r++) {
    const row = vals[r];

    // (a) filtr po statusie (jeśli mamy kolumnę)
    let pass = true; // jeżeli z jakiegoś powodu statusu nie mamy, nie blokujemy przepływu
    if (idxStatus >= 0 && idxStatus < row.length) {
      const st = normalizeStatus_(row[idxStatus]);
      pass = (st !== null) && ALLOWED_ORDER_STATUSES.has(st);
    }
    if (!pass) { dropped++; continue; }

    // (b) data
    const date = normalizeDate_(row[idxDate]);
    if (!date) { dropped++; continue; }
    const idxM = monthIndex.get(yyyymmKey_(date));

    // (c) pozycje z kolumny "products"
    const items = extractItemsFromProducts_(row[idxProducts]);
    if (!items.length) { dropped++; continue; }

    // (d) agregacja
    for (const it of items) {
      const partKey = (it.partKey || '').trim();
      if (!partKey) continue;
      const qty = Number(it.qty || 0) || 1;

      if (!agg.has(partKey)) {
        agg.set(partKey, { allQty: 0, monthly: Array(months.length).fill(0) });
      }
      const rec = agg.get(partKey);
      rec.allQty += qty;
      if (idxM !== undefined) rec.monthly[idxM] += qty;
      kept++;
    }
  }

  // 6) Wynik (nagłówek + wiersze)
  const headerOut = ['part_number_key', 'Zamówienia_od_początku', ...months.map(m => m.label)];
  const out = [headerOut];
  for (const [key, rec] of agg) {
    out.push([key, rec.allQty, ...rec.monthly]);
  }

  // 7) Zapis do OUT_SHEET
  let outSh = home.getSheetByName(OUT_SHEET) || home.insertSheet(OUT_SHEET);
  outSh.clearContents();
  setValuesChunked_(outSh, 1, 1, out, SETVALUES_BATCH);
  outSh.setFrozenRows(1);

  // 8) Podsumowanie
  const secs = ((new Date() - t0) / 1000).toFixed(2);
  toast_(`Sprzedaż_okresowa_FBE: wierszy OK=${kept}, odrzuconych=${dropped}. Czas: ${secs}s`, OUT_SHEET, 6);
}
