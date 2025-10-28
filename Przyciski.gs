/* ================================================================================
 * PRZYCISKI.JS - MENU + MULTI-PARTNER + WSPÓLNE HELPERY
 * ================================================================================ */

// ================================================================================
// 1. STAŁE GLOBALNE
// ================================================================================

var PARTNERS_SHEET = 'Partnerzy';
var BACKLOG_SHEET = 'Backlog';
var BACKLOG_ACTIVE_CELL = 'A2';
var USER_PROP_ACTIVE = 'ACTIVE_PARTNER';

var COLS = {
  id: 1,
  roFbe: 2,
  prodRoFbe: 3,
  ro: 4,
  prodRo: 5
};

var CFG = {
  REG: {
    IN_SHEET_TRENDS: 'Sprzedaż_okresowa_KZ',
    OUT_SHEET_TRENDS: 'Trendy_KZ'
  }
};

// ================================================================================
// 2. LIFECYCLE
// ================================================================================

function onOpen() {
  try {
    installProgramMenu_();
  } catch (err) {
    SpreadsheetApp.getUi().alert('Nie udało się zainstalować menu "Programy".\n' + (err && (err.message || err)));
  }
  
  try {
    ensureBoldHeadersAllSheets_();
  } catch (_) {}
}

function onInstall(e) {
  onOpen(e);
}

// ================================================================================
// 3. MENU (USUNIĘTO: ⏳ Rotacja / pozostałe sztuki (FBE)…)
// ================================================================================

function installProgramMenu_() {
  var ui = SpreadsheetApp.getUi();
  var menu = ui.createMenu('Programy 💻')
    .addItem('🧹 Wyczyść arkusze partnera', 'cmd_clearSheetsForPartner')
    .addSeparator()
    .addItem('👤 Wybierz partnera…', 'showPartnerPicker')
    .addSeparator()
    .addItem('🔐 Autoryzacja', 'cmd_authorize')
    .addItem('⟳ Przeliczenie zamówień z FBE', 'cmd_recomputeSales')
    .addItem('💲 Przelicz średnie ceny miesięczne (FBE)', 'cmd_przeliczSrednieCenyMiesieczne')
    .addItem('➚ Przeliczenie Trendu z FBE', 'cmd_computeTrends')
    .addItem('📦 Propozycja zamówienia (FBE)…', 'cmd_reorderProposalFBE')
    .addItem('⏳ Rotacja / pozostałe sztuki (FBE)…', 'cmd_rotationCheckFBE')
    .addItem('➕ Dodaj własne kolumny (ADS_FBE)…', 'cmd_addCustomColumns')
    .addItem('⟳ Przeliczenie zamówień (konto zwykłe)', 'cmd_recomputeSalesRegular')
    .addItem('💲 Przelicz średnie ceny miesięczne (konto zwykłe)', 'cmd_przeliczSrednieCenyMiesieczneKZ')
    .addItem('➚ Przeliczenie Trendu z konta zwykłego', 'cmd_computeTrendsRegularPrompt')
    .addItem('➕ Dodaj własne kolumny (ADS_KZ)…', 'cmd_addCustomColumnsRegular')
    .addItem('⚖️ Porównanie sprzedaży KZ vs FBE', 'cmd_buildSalesComparisonKzFbe')
    .addSeparator()
    .addItem('📧 Generuj raport partnera...', 'cmd_generatePartnerReport')
    .addItem('🎨 Odśwież styl nagłówków (globalnie)', 'cmd_fixHeadersAllSheets')
    .addItem('📧 Zgłoś błąd / sugestię', 'cmd_reportBug');

  try {
    menu.addToUi();
  } catch (err) {
    Logger.log('Menu installation failed: ' + (err && (err.stack || err.message) || err));
    throw err;
  }
}

// ================================================================================
// POZOSTAŁA CZĘŚĆ PLIKU BEZ ZMIAN (autoryzacja, partnery, czyszczenie itd.)
// ================================================================================

function cmd_authorize() {
  try {
    ensureAuthBasic_();
    SpreadsheetApp.getUi().alert('✅ Autoryzacja potwierdzona.');
  } catch (e) {
    SpreadsheetApp.getUi().alert('🔐 Wymagana autoryzacja.\n\n' + (e && (e.message || e)));
    throw e;
  }
}

function ensureAuthBasic_() {
  SpreadsheetApp.getActive().getSheets()[0].getName();
  ScriptApp.getOAuthToken();
}

function ensureAuthWithPartner_() {
  ensureAuthBasic_();
  try { var a = getSS_RoFbe_(); if (a) a.getSheets()[0].getName(); } catch (_) {}
  try { var b = getSS_Ro_(); if (b) b.getSheets()[0].getName(); } catch (_) {}
}

function ensureAuth_() {
  return ensureAuthWithPartner_();
}

function showPartnerPicker() {
  ensureAuthBasic_();
  var ids = listPartnerIds_();
  var t = HtmlService.createTemplateFromFile('PartnerPicker');
  t.ids = ids;
  var html = t.evaluate().setWidth(420).setHeight(260);
  SpreadsheetApp.getUi().showModalDialog(html, 'Wybierz partnera');
}

function setActivePartner(id) {
  var reg = getPartnersRegistry_();
  if (!reg[id]) throw new Error('Nie znaleziono partnera: ' + id);
  PropertiesService.getUserProperties().setProperty(USER_PROP_ACTIVE, id);
  updateBacklogActivePartner_(id);
  toast_('Aktywny partner: ' + id, 'Programy', 3);
}

function getActivePartnerId_() {
  return PropertiesService.getUserProperties().getProperty(USER_PROP_ACTIVE) || '';
}

function requireActivePartner_() {
  var id = getActivePartnerId_();
  if (!id) throw new Error('Nie wybrano partnera. Użyj: Programy → „Wybierz partnera…".');
  return id;
}

function getPartnersRegistry_() {
  var sh = SpreadsheetApp.getActive().getSheetByName(PARTNERS_SHEET);
  if (!sh) throw new Error('Brak karty: ' + PARTNERS_SHEET);
  
  var values = sh.getDataRange().getValues();
  if (values.length < 2) return {};
  
  var map = {};
  for (var i = 1; i < values.length; i++) {
    var row = values[i];
    var id = String(row[COLS.id - 1] || '').trim();
    if (!id) continue;
    
    map[id] = {
      id: id,
      roFbeUrl: String(row[COLS.roFbe - 1] || '').trim(),
      prodRoFbeUrl: String(row[COLS.prodRoFbe - 1] || '').trim(),
      roUrl: String(row[COLS.ro - 1] || '').trim(),
      prodRoUrl: String(row[COLS.prodRo - 1] || '').trim()
    };
  }
  return map;
}

function listPartnerIds_() {
  return Object.keys(getPartnersRegistry_()).sort();
}

function updateBacklogActivePartner_(value) {
  try {
    var ss = SpreadsheetApp.getActive();
    var sh = ss.getSheetByName(BACKLOG_SHEET);
    if (!sh) return;
    sh.getRange(BACKLOG_ACTIVE_CELL).setValue(value || '');
  } catch (_) {}
}

function getSS_RoFbe_() {
  var u = getPartnersRegistry_()[requireActivePartner_()].roFbeUrl;
  return u ? SpreadsheetApp.openByUrl(u) : null;
}

function getSS_Ro_() {
  var u = getPartnersRegistry_()[requireActivePartner_()].roUrl;
  return u ? SpreadsheetApp.openByUrl(u) : null;
}

function getSS_ProdRoFbe_() {
  var u = getPartnersRegistry_()[requireActivePartner_()].prodRoFbeUrl;
  return u ? SpreadsheetApp.openByUrl(u) : null;
}

function getSS_ProdRo_() {
  var u = getPartnersRegistry_()[requireActivePartner_()].prodRoUrl;
  return u ? SpreadsheetApp.openByUrl(u) : null;
}

function cmd_recomputeSales() {
  var ui = SpreadsheetApp.getUi();
  var lock = LockService.getDocumentLock();
  if (!lock.tryLock(8000)) {
    ui.alert('Program jest zajęty. Spróbuj za chwilę.');
    return;
  }
  
  try {
    ensureAuthWithPartner_();
    requireActivePartner_();
    SpreadsheetApp.getActive().toast('Liczenie sprzedaży (FBE)…', 'Programy', 5);
    buildSalesPeriods();
    SpreadsheetApp.getActive().toast('Gotowe ✅', 'Programy', 5);
  } catch (e) {
    ui.alert('Błąd przeliczenia (FBE):\n' + (e && (e.stack || e.message) || e));
  } finally {
    lock.releaseLock();
  }
}

function cmd_computeTrends() {
  try {
    ensureAuthWithPartner_();
    requireActivePartner_();
  } catch (e) {
    SpreadsheetApp.getUi().alert('Błąd autoryzacji/partnera:\n' + (e && (e.message || e)));
    return;
  }
  
  // Pokaż popup z mnożnikami
  var t = HtmlService.createTemplateFromFile('TrendMultipliersPrompt');
  var html = t.evaluate().setWidth(480).setHeight(520);
  SpreadsheetApp.getUi().showModalDialog(html, 'Mnożniki dla prognoz FBE');
}

function computeTrendsWithMultipliers(mult1, mult2, mult3) {
  var lock = LockService.getDocumentLock();
  if (!lock.tryLock(8000)) {
    throw new Error('Program jest zajęty. Spróbuj za chwilę.');
  }
  
  try {
    ensureAuthWithPartner_();
    requireActivePartner_();
    SpreadsheetApp.getActive().toast('Liczenie trendów (FBE) z mnożnikami…', 'Programy', 5);
    computeTrendsFromSalesPeriods(mult1, mult2, mult3);
    SpreadsheetApp.getActive().toast('Trendy gotowe ✅', 'Programy', 5);
  } catch (e) {
    throw new Error('Błąd trendów (FBE):\n' + (e && (e.stack || e.message) || e));
  } finally {
    lock.releaseLock();
  }
}

// ================================================================================
// KONTO ZWYKŁE - TRENDY Z MNOŻNIKAMI
// ================================================================================

function cmd_computeTrendsRegularPrompt() {
  try {
    ensureAuthWithPartner_();
    requireActivePartner_();
  } catch (e) {
    SpreadsheetApp.getUi().alert('Błąd autoryzacji/partnera:\n' + (e && (e.message || e)));
    return;
  }
  
  var t = HtmlService.createTemplateFromFile('TrendMultipliersPromptKZ');
  var html = t.evaluate().setWidth(480).setHeight(520);
  SpreadsheetApp.getUi().showModalDialog(html, 'Mnożniki dla prognoz (Konto Zwykłe)');
}

function computeTrendsRegularWithMultipliers(mult1, mult2, mult3) {
  var lock = LockService.getDocumentLock();
  if (!lock.tryLock(8000)) {
    throw new Error('Program jest zajęty. Spróbuj za chwilę.');
  }
  
  try {
    ensureAuthWithPartner_();
    requireActivePartner_();
    SpreadsheetApp.getActive().toast('Liczenie trendów (konto zwykłe) z mnożnikami…', 'Programy', 5);
    computeTrendsFromSalesPeriodsRegular(mult1, mult2, mult3);
    SpreadsheetApp.getActive().toast('Trendy (konto zwykłe) gotowe ✅', 'Programy', 5);
  } catch (e) {
    throw new Error('Błąd trendów (konto zwykłe):\n' + (e && (e.stack || e.message) || e));
  } finally {
    lock.releaseLock();
  }
}

function cmd_przeliczSrednieCenyMiesieczne() {
  var ui = SpreadsheetApp.getUi();
  var lock = LockService.getDocumentLock();
  if (!lock.tryLock(8000)) {
    ui.alert('Program jest zajęty. Spróbuj za chwilę.');
    return;
  }
  
  try {
    ensureAuthWithPartner_();
    requireActivePartner_();
    SpreadsheetApp.getActive().toast('Liczenie średnich cen miesięcznych…', 'Ceny', 5);
    recomputeMonthlyPricesFBE();
    SpreadsheetApp.getActive().toast('Gotowe ✅', 'Ceny', 5);
  } catch (e) {
    ui.alert('Błąd przeliczenia cen:\n' + (e && (e.stack || e.message) || e));
  } finally {
    lock.releaseLock();
  }
}

function cmd_przeliczSrednieCenyMiesieczneKZ() {
  var ui = SpreadsheetApp.getUi();
  var lock = LockService.getDocumentLock();
  if (!lock.tryLock(8000)) {
    ui.alert('Program jest zajęty. Spróbuj za chwilę.');
    return;
  }
  
  try {
    ensureAuthWithPartner_();
    requireActivePartner_();
    SpreadsheetApp.getActive().toast('Liczenie średnich cen (KZ)…', 'Programy', 5);
    cmd_przeliczSrednieCenyMiesieczneKZ();
    SpreadsheetApp.getActive().toast('Ceny miesięczne KZ gotowe ✅', 'Programy', 5);
  } catch (e) {
    ui.alert('Błąd liczenia cen (KZ):\n' + (e && (e.stack || e.message) || e));
  } finally {
    lock.releaseLock();
  }
}

function cmd_fixHeadersAllSheets() {
  try {
    ensureGlobalHeaderStyling_();
    openMeasureDialogForAll_();
    SpreadsheetApp.getActive().toast('Nagłówki odświeżone ✅', 'Programy', 4);
  } catch (e) {
    SpreadsheetApp.getUi().alert('Błąd stylowania nagłówków:\n' + (e && (e.message || e)));
  }
}

function cmd_clearSheetsForPartner() {
  var ui = SpreadsheetApp.getUi();
  var lock = LockService.getDocumentLock();
  if (!lock.tryLock(8000)) {
    ui.alert('Program jest zajęty. Spróbuj za chwilę.');
    return;
  }
  
  try {
    var ss = SpreadsheetApp.getActive();
    var sheets = ss.getSheets();
    
    for (var i = 0; i < sheets.length; i++) {
      var sh = sheets[i];
      var name = sh.getName();
      
      if (name === 'Partnerzy' || name === 'LOGS') continue;
      
      var maxRows = Math.max(1, sh.getMaxRows());
      var maxCols = Math.max(1, sh.getMaxColumns());
      var bodyRows = Math.max(0, maxRows - 1);
      
      if (bodyRows > 0) {
        var bodyRange = sh.getRange(2, 1, bodyRows, maxCols);
        try { bodyRange.clearContent(); } catch (_) {}
        try { bodyRange.clearNote(); } catch (_) {}
        try { bodyRange.clearDataValidations(); } catch (_) {}
        try { bodyRange.clearFormat(); } catch (_) {}
        try { bodyRange.breakApart(); } catch (_) {}
      }
      
      try { if (sh.getFilter) { var filter = sh.getFilter(); if (filter) filter.remove(); } } catch (_) {}
      try { if (sh.setConditionalFormatRules) sh.setConditionalFormatRules([]); } catch (_) {}
      try { if (sh.getCharts) { var charts = sh.getCharts(); for (var c = 0; c < charts.length; c++) sh.removeChart(charts[c]); } } catch (_) {}
      try { if (sh.getSlicers) { var slicers = sh.getSlicers(); for (var s = 0; s < slicers.length; s++) slicers[s].remove(); } } catch (_) {}
      try { if (sh.getPivotTables) { var pivots = sh.getPivotTables(); for (var p = 0; p < pivots.length; p++) pivots[p].remove(); } } catch (_) {}
      try { if (sh.getDrawings) { var drawings = sh.getDrawings(); for (var d = 0; d < drawings.length; d++) drawings[d].remove(); } } catch (_) {}
      
      try { sh.setFrozenRows(0); } catch (_) {}
      try { sh.setFrozenColumns(0); } catch (_) {}
      
      var headerReset = false;
      try { headerReset = resetPartnerSheetHeader_(sh); } catch (_) {}
      try { sh.setFrozenRows(headerReset ? 1 : 0); } catch (_) {}
    }
    
    try {
    if (typeof ensureGlobalHeaderStyling_ === 'function') {
      ensureGlobalHeaderStyling_();
    } else if (typeof ensureBoldHeadersAllSheets_ === 'function') {
      ensureBoldHeadersAllSheets_();
    }
    
    sheets.forEach(function(sh) {
      var name = sh.getName();
      if (name === 'Partnerzy' || name === 'LOGS') return;
      
      try {
        applyDefaultHeaderWidths_(sh);
      } catch (_) {}
    });
    
  } catch (_) {}
  
  ss.toast('Arkusze partnera wyczyszczone ✅', 'Programy', 5);
  } finally {
    try { updateBacklogActivePartner_(''); } catch (_) {}
    lock.releaseLock();
  }
}

function resetPartnerSheetHeader_(sh) {
  if (!sh) return false;
  
  var name = sh.getName();
  if (!name) return false;
  
  if (name === 'Ceny_miesięczne_FBE' && typeof writeEmptyPricingTable_ === 'function') {
    var headerFbe = buildPartnerPriceHeaderFbe_();
    ensureMinColumns_(sh, headerFbe.length);
    var maxColsFbe = Math.max(headerFbe.length, sh.getMaxColumns());
    clearHeaderRow_(sh, maxColsFbe);
    writeEmptyPricingTable_(sh);
    if (maxColsFbe > headerFbe.length) {
      try { sh.getRange(1, headerFbe.length + 1, 1, maxColsFbe - headerFbe.length).clearContent(); } catch (_) {}
    }
    try { applyDefaultHeaderWidths_(sh, headerFbe); } catch (_) {}
    return true;
  }
  
  var header = null;
  
  if (name === 'Sprzedaż_okresowa_FBE' || name === 'Sprzedaż_okresowa_KZ') {
    header = buildPartnerSalesHeader_(typeof START_YEAR === 'number' ? START_YEAR : 2025, typeof START_MONTH === 'number' ? START_MONTH : 1);
  } else if (name === 'Trendy_FBE') {
    // ZMIANA: usunięto Deficyt_w_miesiącu_1 i Miesiące_pokrycia (duplikat)
    header = ['part_number_key', 'ID', 'SKU', 'EAN', '3 miesiące wstecz', 'Sprzedaż dwa miesiące wstecz', 'Sprzedaż ubiegły miesiąc', 'Sprzedaż w obecnym miesiącu', 'Trend_M0/M-1_%', 'Trend_M-1/M-2_%', 'Trend_M-2/M-3_%', 'Suma_3M', 'Średnia_3M', 'Suma_6M', 'Średnia_6M', 'Slope_6M', 'Klasyfikacja', 'Performance', 'Prognoza_nast_M', 'Prognoza_+2M', 'Prognoza_+3M', 'Stock', 'Cena_netto', 'Cena_brutto', 'Pokrycie_magazynu'];
  } else if (name === 'Trendy_KZ') {
    header = ['part_number_key', 'ID', 'SKU', 'EAN', '3 miesiące wstecz', 'Sprzedaż dwa miesiące wstecz', 'Sprzedaż ubiegły miesiąc', 'Sprzedaż w obecnym miesiącu', 'Trend_M0/M-1_%', 'Trend_M-1/M-2_%', 'Trend_M-2/M-3_%', 'Suma_3M', 'Średnia_3M', 'Suma_6M', 'Średnia_6M', 'Slope_6M', 'Klasyfikacja', 'Performance', 'Prognoza_nast_M', 'Prognoza_+2M', 'Prognoza_+3M', 'Stock', 'Cena_netto', 'Cena_brutto', 'Pokrycie_magazynu', 'Obrót_3M_netto', 'Udział_w_obrócie_%', 'Rotacja_magazynu'];
  } else if (name === 'Ceny_miesięczne_KZ') {
    header = buildPartnerPriceHeaderKz_();
  }
  
  if (!header || !header.length) return false;
  
  ensureMinColumns_(sh, header.length);
  var maxCols = Math.max(header.length, sh.getMaxColumns());
  clearHeaderRow_(sh, maxCols);
  sh.getRange(1, 1, 1, header.length).setValues([header]);
  
  if (maxCols > header.length) {
    try { sh.getRange(1, header.length + 1, 1, maxCols - header.length).clearContent(); } catch (_) {}
  }
  
  if (name === 'Trendy_FBE' && typeof ensurePersistentHelperColumns_ === 'function') {
    try { ensurePersistentHelperColumns_(sh); } catch (_) {}
  }
  
  try { applyDefaultHeaderWidths_(sh, header); } catch (_) {}
  return true;
}

function buildPartnerSalesHeader_(startYear, startMonth) {
  var labels = buildPartnerMonthLabels_(startYear, startMonth);
  return ['part_number_key', 'Zamówienia_od_początku'].concat(labels);
}

function buildPartnerPriceHeaderKz_() {
  var cfg = (typeof KZ_REG_CFG === 'object' && KZ_REG_CFG) || {};
  var startYear = typeof cfg.START_YEAR === 'number' ? cfg.START_YEAR : 2025;
  var startMonth = typeof cfg.START_MONTH === 'number' ? cfg.START_MONTH : 1;
  var labels = buildPartnerMonthLabels_(startYear, startMonth);
  return ['part_number_key'].concat(labels);
}

function buildPartnerPriceHeaderFbe_() {
  var startYear = (typeof PRICING_FIRST_Y === 'number') ? PRICING_FIRST_Y : 2025;
  var startMonth = (typeof PRICING_FIRST_M === 'number') ? PRICING_FIRST_M : 1;
  var labels = buildPartnerMonthLabels_(startYear, startMonth);
  return ['part_number_key'].concat(labels);
}

function buildPartnerMonthLabels_(startYear, startMonth) {
  if (typeof monthRange_ !== 'function') return [];
  var months = monthRange_(startYear, startMonth, new Date()) || [];
  var labels = [];
  for (var i = 0; i < months.length; i++) {
    labels.push(months[i].label);
  }
  return labels;
}

function clearHeaderRow_(sh, maxCols) {
  if (!sh) return;
  maxCols = Math.max(1, maxCols || sh.getMaxColumns());
  var headerRange = sh.getRange(1, 1, 1, maxCols);
  try { headerRange.clearContent(); } catch (_) {}
  try { headerRange.clearFormat(); } catch (_) {}
  try { headerRange.clearNote(); } catch (_) {}
  try { headerRange.clearDataValidations(); } catch (_) {}
  try { headerRange.breakApart(); } catch (_) {}
}

function ensureMinColumns_(sh, minCols) {
  if (!sh || !(minCols > 0)) return;
  try {
    if (typeof ensureColumns_ === 'function') {
      ensureColumns_(sh, minCols);
    } else {
      var lc = sh.getLastColumn();
      if (lc < minCols) {
        sh.insertColumnsAfter(lc, minCols - lc);
      }
    }
  } catch (_) {}
}

function applyDefaultHeaderWidths_(sh, headerValues) {
  if (!sh) return;
  
  var lc = sh.getLastColumn();
  if (!(lc > 0)) return;
  
  var headers;
  if (headerValues && headerValues.length) {
    headers = [];
    for (var i = 0; i < lc; i++) {
      headers.push(i < headerValues.length ? headerValues[i] : '');
    }
  } else {
    try {
      headers = sh.getRange(1, 1, 1, lc).getDisplayValues()[0] || [];
    } catch (_) {
      headers = [];
    }
    if (headers.length < lc) {
      for (var j = headers.length; j < lc; j++) headers[j] = '';
    }
  }
  
  for (var c = 1; c <= lc; c++) {
    var header = String(headers[c - 1] || '').trim();
    var width = calculateOptimalWidth_(header);
    try { 
      sh.setColumnWidth(c, width); 
    } catch (_) {}
  }
}

function getHome_() {
  try {
    var ss = SpreadsheetApp.getActive();
    if (ss) return ss;
  } catch (e) {}
  throw new Error('Brak aktywnego skoroszytu.');
}

function toast_(msg, title, secs) {
  try {
    SpreadsheetApp.getActive().toast(msg, title || 'Info', secs || 5);
  } catch (e) {
    Logger.log((title || 'Info') + ': ' + msg);
  }
}

function withRetry_(fn, label, attempts, baseMs) {
  var last;
  attempts = attempts || 3;
  baseMs = baseMs || 400;
  
  for (var i = 0; i < attempts; i++) {
    try {
      return fn();
    } catch (e) {
      last = e;
      var m = String(e && (e.message || e));
      if (!/(internal|unavailable|temporar|timeout|backend|try again)/i.test(m) || i === attempts - 1) throw e;
      Utilities.sleep(baseMs * Math.pow(2, i));
    }
  }
  throw last;
}

function findCol_(headerLowercaseArray, candidates) {
  for (var i = 0; i < candidates.length; i++) {
    var idx = headerLowercaseArray.indexOf(String(candidates[i]).toLowerCase());
    if (idx !== -1) return idx;
  }
  return -1;
}

function monthRange_(startYear, startMonth1, endDate) {
  var out = [];
  var y = startYear;
  var m = startMonth1;
  var endY = endDate.getFullYear();
  var endM = endDate.getMonth() + 1;
  
  while (y < endY || (y === endY && m <= endM)) {
    out.push({
      key: y + '-' + String(m).padStart(2, '0'),
      label: String(m).padStart(2, '0') + '-' + y
    });
    m++;
    if (m > 12) {
      m = 1;
      y++;
    }
  }
  return out;
}

function yyyymmKey_(d) {
  return d.getFullYear() + '-' + String(d.getMonth() + 1).padStart(2, '0');
}

function normalizeDate_(v) {
  if (v instanceof Date && !isNaN(v)) {
    return new Date(v.getFullYear(), v.getMonth(), v.getDate());
  }
  var s = String(v || '').trim();
  if (!s) return null;
  var d = new Date(s);
  return (d instanceof Date && !isNaN(d)) ? new Date(d.getFullYear(), d.getMonth(), d.getDate()) : null;
}

var ALLOW_EVAL = false;

function extractItemsFromProducts_(cell) {
  var out = [];
  if (cell == null) return out;
  if (Array.isArray(cell)) {
    for (var i = 0; i < cell.length; i++) pushItemFromObj_(cell[i], out);
    return out;
  }
  if (typeof cell === 'object') {
    pushItemFromObj_(cell, out);
    return out;
  }
  
  var s = String(cell).trim();
  if (!s) return out;
  
  if ((s.startsWith('[') && s.endsWith(']')) || (s.startsWith('{') && s.endsWith('}'))) {
    try {
      var obj = JSON.parse(s);
      if (Array.isArray(obj)) {
        for (var j = 0; j < obj.length; j++) pushItemFromObj_(obj[j], out);
      } else {
        pushItemFromObj_(obj, out);
      }
      if (out.length) return out;
    } catch (_) {}
  }
  
  if (ALLOW_EVAL) {
    try {
      var obj2 = eval('(' + s + ')');
      if (Array.isArray(obj2)) {
        for (var k = 0; k < obj2.length; k++) pushItemFromObj_(obj2[k], out);
      } else {
        pushItemFromObj_(obj2, out);
      }
      if (out.length) return out;
    } catch (_) {}
  }
  
  var reKey = /part[_ ]?number[_ ]?key['"]?\s*[:=]\s*['"]([^'"]+)['"]/gi;
  var m;
  while ((m = reKey.exec(s)) !== null) {
    var partKey = m[1];
    var slice = s.slice(Math.max(0, m.index - 160), Math.min(s.length, m.index + 160));
    var mQty = /(?:initial_qty|quantity|qty)\s*[:=]\s*(\d+)/i.exec(slice);
    var qty = mQty ? Number(mQty[1]) : 1;
    out.push({ partKey: partKey, qty: qty });
  }
  return out;
}

function pushItemFromObj_(o, out) {
  if (!o) return;
  var partKey = (o.part_number_key != null) ? String(o.part_number_key) : (o.partNumberKey != null) ? String(o.partNumberKey) : '';
  if (!partKey) return;
  
  var qty = 0;
  if (o.initial_qty != null) qty = Number(o.initial_qty);
  else if (o.quantity != null) qty = Number(o.quantity);
  else if (o.qty != null) qty = Number(o.qty);
  if (!qty || !isFinite(qty)) qty = 1;
  
  out.push({ partKey: partKey, qty: qty });
}

function setValuesChunked_(sheet, r1, c1, data, batchRows) {
  batchRows = batchRows || 20000;
  if (!data || !data.length) return;
  
  var i = 0;
  while (i < data.length) {
    var part = data.slice(i, i + batchRows);
    withRetry_(function () {
      sheet.getRange(r1 + i, c1, part.length, part[0].length).setValues(part);
    }, 'setValues');
    i += part.length;
  }
}

function ensureBoldHeadersAllSheets_() {
  var ss = SpreadsheetApp.getActive();
  ss.getSheets().forEach(function (sh) {
    var lr = sh.getLastRow();
    var lc = sh.getLastColumn();
    if (lr < 1 || lc < 1) return;
    sh.getRange(1, 1, 1, lc).setFontWeight('bold');
  });
}

function ensureGlobalHeaderStyling_() {
  var ss = SpreadsheetApp.getActive();
  var LIGHT_GREEN = '#b7e1cd';
  
  ss.getSheets().forEach(function (sh) {
    var lc = sh.getLastColumn();
    if (lc < 1) return;
    
    var hdrRng = sh.getRange(1, 1, 1, lc);
    var hdrVals = hdrRng.getDisplayValues()[0];
    
    hdrRng
      .setFontFamily('Arial')
      .setFontSize(10)
      .setFontWeight('bold')
      .setHorizontalAlignment('center')
      .setVerticalAlignment('middle');
    
    var row = [];
    for (var i = 0; i < lc; i++) {
      row.push(String(hdrVals[i] || '').trim() ? LIGHT_GREEN : '#ffffff');
    }
    hdrRng.setBackgrounds([row]);
  });
}

function openMeasureDialogForAll_() {
  var ss = SpreadsheetApp.getActive();
  var sheets = ss.getSheets();
  
  var payload = sheets.map(function (sh) {
    var lc = sh.getLastColumn();
    var headers = lc > 0 ? sh.getRange(1, 1, 1, lc).getDisplayValues()[0] : [];
    return { name: sh.getName(), headers: headers };
  });
  
  var t = HtmlService.createTemplateFromFile('MeasureHeaders');
  t.payload = JSON.stringify(payload);
  var html = t.evaluate().setWidth(10).setHeight(10);
  SpreadsheetApp.getUi().showModalDialog(html, 'Dopasowuję szerokości…');
}

function cmd_reportBug() {
  try {
    var t = HtmlService.createTemplateFromFile('ReportBug');
    
    try {
      t.activePartner = getActivePartnerId_() || 'Brak';
    } catch (_) {
      t.activePartner = 'Brak';
    }
    
    var html = t.evaluate()
      .setWidth(520)
      .setHeight(450);
    
    SpreadsheetApp.getUi().showModalDialog(html, '📧 Zgłoś błąd lub sugestię');
  } catch (e) {
    SpreadsheetApp.getUi().alert('Błąd otwierania formularza:\n' + (e && (e.message || e)));
  }
}

function sendBugReport(userMessage) {
  var YOUR_EMAIL = 'rafalporeba@olimpmarketplace.com';
  
  if (!userMessage || !userMessage.trim()) {
    throw new Error('Wiadomość nie może być pusta');
  }
  
  var context = {
    timestamp: new Date().toLocaleString('pl-PL'),
    user: Session.getActiveUser().getEmail(),
    spreadsheetName: SpreadsheetApp.getActive().getName(),
    spreadsheetUrl: SpreadsheetApp.getActive().getUrl(),
    activePartner: 'Brak'
  };
  
  try {
    context.activePartner = getActivePartnerId_() || 'Brak';
  } catch (_) {}
  
  var subject = '🐛 Zgłoszenie z aplikacji: ' + context.spreadsheetName;
  
  var body = 
    '=== ZGŁOSZENIE OD UŻYTKOWNIKA ===\n\n' +
    userMessage + '\n\n' +
    '=== DANE KONTEKSTOWE ===\n' +
    'Data/czas: ' + context.timestamp + '\n' +
    'Użytkownik: ' + context.user + '\n' +
    'Skoroszyt: ' + context.spreadsheetName + '\n' +
    'URL: ' + context.spreadsheetUrl + '\n' +
    'Aktywny partner: ' + context.activePartner + '\n';
  
  MailApp.sendEmail({
    to: YOUR_EMAIL,
    subject: subject,
    body: body
  });
  
  return true;
}

function applyMeasuredHeaderWidths_(resultMap) {
  var ss = SpreadsheetApp.getActive();
  Object.keys(resultMap || {}).forEach(function (name) {
    var sh = ss.getSheetByName(name);
    if (!sh) return;
    
    var widths = resultMap[name] || [];
    var lc = Math.min(sh.getLastColumn(), widths.length);
    
    for (var c = 1; c <= lc; c++) {
      var w = widths[c - 1];
      if (isFinite(w) && w > 0) {
        try {
          sh.setColumnWidth(c, Math.round(w));
        } catch (_) {}
      }
    }
  });
}

function calculateOptimalWidth_(headerText) {
  var s = String(headerText || '').trim();
  if (!s) return 70;
  
  if (/^\d{2}-\d{4}$/.test(s)) return 80;
  
  var len = s.length;
  var baseWidth = len * 7.5;
  
  var wideChars = (s.match(/[WM@#%&ĄĆĘŁŃÓŚŹŻ]/g) || []).length;
  baseWidth += wideChars * 2;
  
  var narrowChars = (s.match(/[il.:,;|!''´`"'()[\]{}]/g) || []).length;
  baseWidth -= narrowChars * 2;
  
  var padding = 38;
  var totalWidth = Math.round(baseWidth + padding);
  
  var MIN_WIDTH = 70;
  var MAX_WIDTH = 250;
  
  return Math.max(MIN_WIDTH, Math.min(MAX_WIDTH, totalWidth));
}
