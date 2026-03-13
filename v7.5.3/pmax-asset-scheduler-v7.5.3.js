/*******************************************************
 * PMax Asset Scheduler v7.5.3 (2026-03-12)
 * Időzített TEXT + IMAGE + VIDEO asset hozzáadás és törlés
 * Performance Max asset group-okban.
 *
 * © 2025 Klára Bognár – All rights reserved.
 * Author: Klára Bognár (Impresszió Online Marketing)
 * https://impresszio.hu
 * Created with assistance from Google Ads Script Sensei © Nils Rooijmans and Claude Code.
 *
 * Telepítés:
 * 1) Illeszd be ezt a kódot a módosítandó Google Ads fiókban új Scriptként.
 * 2) Készíts Google Sheet-et ennek a sablonnak a másolásával: https://docs.google.com/spreadsheets/d/14cWq7Ly8XHIkcUkIhYHeIwhwEb3w5Qjl8npcf6SwaiI/copy
 * 3) Állítsd be a SPREADSHEET_URL és NOTIFICATION_EMAIL konfigot.
 * 4) Ütemezd óránként a scriptet.
 * 5) Opcionális: Engedélyezd a YouTube API-t a videó validációhoz (Script → Advanced APIs → YouTube).
 *******************************************************/

/*** ===================== KONFIG ===================== ***/

// Kötelező: Google Sheet URL
const SPREADSHEET_URL = 'YOUR_SPREADSHEET_URL_HERE';

// Kötelező: Értesítési e-mail(ek), vesszővel elválasztva
const NOTIFICATION_EMAIL = 'YOUR_EMAIL_HERE';


// Lapnevek
const TEXT_SHEET_NAME = 'TextAssets';
const IMAGE_SHEET_NAME = 'ImageAssets';
const VIDEO_SHEET_NAME = 'VideoAssets';
const PREVIEW_RESULTS_SHEET_NAME = 'Preview Results';
const RESULTS_SHEET_NAME = 'Results';

// GAQL API verzió
const API_VERSION = 'v21';

// Időablakok (percekben a nap elejétől)
// EXCLUSIVE upper bound: [from, to) ahol 'to' NEM tartozik bele
const ADD_WINDOW_FROM_MIN = 0;      // 00:00 (inclusive)
const ADD_WINDOW_TO_MIN = 60;       // 01:00 (exclusive) → 00:00-00:59
const REMOVE_WINDOW_FROM_MIN = 1380; // 23:00 (inclusive)
const REMOVE_WINDOW_TO_MIN = 1440;   // 24:00 (exclusive) → 23:00-23:59

// Validációs időablak (napokban a jövőbe nézve)
// Mai nap + következő X nap sorai kerülnek validálásra
const VALIDATION_FUTURE_DAYS = 30;

// Preview beállítások
// Ha true: MINDIG küld preview email-t (sheet hash változás detektálás kikapcsolva)
// Ha false: Csak akkor küld preview-t, ha a sheet tartalma változott az előző futás óta
const ALWAYS_SEND_PREVIEW = false;

// Text asset limitek (Google PMax követelmények)
const TEXT_LIMITS = {
  HEADLINE: { min: 3, max: 15, maxLen: 30, warnThreshold: 2 },
  LONG_HEADLINE: { min: 1, max: 5, maxLen: 90, warnThreshold: 1 },
  DESCRIPTION: { min: 2, max: 5, maxLen: 90, warnThreshold: 1 }
};

// Image asset limitek
const IMAGE_LIMITS = {
  TOTAL: { max: 20 },
  MARKETING_IMAGE: { min: 1 }, // HORIZONTAL (1.91:1, pl. 1200×628)
  SQUARE_MARKETING_IMAGE: { min: 1 }, // SQUARE (1:1)
  PORTRAIT_MARKETING_IMAGE: { min: 0 }, // VERTICAL (4:5, pl. 960×1200)
  TALL_PORTRAIT_MARKETING_IMAGE: { min: 0 } // VERTICAL (9:16, pl. 900×1600) - NEW in API v19
};

// Video asset limitek
const VIDEO_LIMITS = {
  YOUTUBE_VIDEO: {
    min: 0,              // Opcionális (nem kötelező videó)
    max: 5,              // API v21 limit
    warnThreshold: 0     // Csak OK vagy ERROR (WARNING-ok kikapcsolva)
  }
};

// Retry beállítások (race condition + átmeneti hibákra)
const RETRY_MAX_ATTEMPTS = 5; // 1 + 5 retry = összesen 6 próbálkozás
const RETRY_BASE_DELAY_MS = 1000; // 1s, 2s, 4s, 8s, 16s exponential backoff + jitter (±25%)

// Log előnézet
const LOG_PREVIEW_COUNT = 10;
const EMAIL_ROW_LIMIT = 50;

// Színkódok (PreviewResults és Results sheetekhez)
const COLOR_OK = '#d4edda';
const COLOR_WARNING = '#fff3cd';
const COLOR_ERROR = '#f8d7da';
const COLOR_SUCCESS = '#d4edda';
const COLOR_SUCCESS_WITH_WARNING = '#ffc107';
const COLOR_FAILED = '#f8d7da';

// Account timezone cache (performance optimization - Claude.ai javaslat #3)
// Egyszer lekérdezzük main()-ben, aztán session végéig cache-elt érték
let ACCOUNT_TIMEZONE = null;

/*** ===================== MAIN ===================== ***/

function main() {
  console.log('=== PMax Asset Scheduler ===');

  // Fail-fast: Konfiguráció ellenőrzés
  if (!SPREADSHEET_URL || !SPREADSHEET_URL.trim()) {
    console.log('❌ HIBA: SPREADSHEET_URL nincs beállítva! Állítsd be a konfig szekcióban.');
    return;
  }

  // Cache timezone once per session (eliminates ~6000+ redundant API calls)
  ACCOUNT_TIMEZONE = AdsApp.currentAccount().getTimeZone();
  const tz = ACCOUNT_TIMEZONE;
  const now = new Date();
  const todayStr = Utilities.formatDate(now, tz, 'yyyy-MM-dd');
  const hour = parseInt(Utilities.formatDate(now, tz, 'H'), 10);
  const min = parseInt(Utilities.formatDate(now, tz, 'm'), 10);
  const nowMinutes = hour * 60 + min;
  const timeStr = Utilities.formatDate(now, tz, 'HH:mm');

  console.log(`⏰ Aktuális idő (${tz}): ${timeStr} (${nowMinutes} perc a napból)`);
  console.log(`📅 Mai dátum: ${todayStr}`);

  if (!NOTIFICATION_EMAIL || !NOTIFICATION_EMAIL.trim()) {
    console.log('⚠️ FIGYELMEZTETÉS: NOTIFICATION_EMAIL nincs beállítva - nem megy email értesítés.');
  }

  const ss = SpreadsheetApp.openByUrl(SPREADSHEET_URL);

  // ===== FÁZIS 0: Sheet Struktúra Validáció =====
  console.log('\n🔍 Sheet struktúra ellenőrzése...');
  try {
    validateSheetStructure(ss);
    setupVideoAssetsSheet(ss);
    setupImageAssetsThumbnail(ss);
    console.log('  ✅ Sheet struktúra rendben');
  } catch (e) {
    console.log(`❌ ${e.message}`);
    sendCriticalErrorEmail(e.message);
    return;
  }

  // ===== FÁZIS 1: Sheet Beolvasás =====
  console.log('\n📖 Sheet beolvasás...');
  const textInput = readSheet(ss, TEXT_SHEET_NAME);
  const imageInput = readSheet(ss, IMAGE_SHEET_NAME);
  const videoInput = readSheet(ss, VIDEO_SHEET_NAME);

  console.log(`  Text sorok: ${textInput.rows.length}, Image sorok: ${imageInput.rows.length}, Video sorok: ${videoInput.rows.length}`);

  // ===== FÁZIS 2: Üres Text/Asset ID/Video ID Szűrés =====
  console.log('\n🔍 Üres sorok szűrése...');
  const textFiltered = filterNonEmptyRows(textInput, 'Text', TEXT_SHEET_NAME);
  const imageFiltered = filterNonEmptyRows(imageInput, 'Asset ID', IMAGE_SHEET_NAME);
  const videoFiltered = filterNonEmptyRows(videoInput, 'Video URL or ID / Asset ID', VIDEO_SHEET_NAME);

  console.log(`  Text valid: ${textFiltered.length}, Image valid: ${imageFiltered.length}, Video valid: ${videoFiltered.length}`);

  // Thumbnail frissítés (ha van Thumbnail oszlop)
  updateImageThumbnails(ss, imageFiltered);
  updateVideoThumbnails(ss, videoFiltered);

  if (textFiltered.length === 0 && imageFiltered.length === 0 && videoFiltered.length === 0) {
    console.log('✅ Nincs feldolgozható sor (minden Text/Asset ID/Video ID üres).');
    clearPreviewResultsSheet(ss);
    return;
  }

  // ===== FÁZIS 3: Dátum Szűrés (mai + jövő 30 nap) =====
  console.log(`\n📅 Dátum szűrés (mai + ${VALIDATION_FUTURE_DAYS} nap)...`);
  const textInRange = filterByDateRange(textFiltered, todayStr);
  const imageInRange = filterByDateRange(imageFiltered, todayStr);
  const videoInRange = filterByDateRange(videoFiltered, todayStr);

  console.log(`  Text range-ben: ${textInRange.length}, Image range-ben: ${imageInRange.length}, Video range-ben: ${videoInRange.length}`);

  if (textInRange.length === 0 && imageInRange.length === 0 && videoInRange.length === 0) {
    console.log('✅ Nincs releváns dátumú sor (mai + 30 nap).');
    clearPreviewResultsSheet(ss);
    return;
  }

  // ===== FÁZIS 3.5: Sheet Change Detection =====
  console.log('\n🔍 Sheet változás ellenőrzése...');
  const currentHash = calculateSheetHash(textInRange, imageInRange, videoInRange);
  const previousHash = getPreviousSheetHash();
  const sheetChanged = (currentHash !== previousHash);

  console.log(`  Hash: ${currentHash}`);
  console.log(`  Változott: ${sheetChanged ? 'IGEN' : 'NEM'}`);

  // ===== FÁZIS 4: Időablak Előszűrés (GAQL optimalizálás) =====
  console.log('\n⏰ Időablak szűrés...');
  const textInWindow = filterByTimeWindow(textInRange, nowMinutes);
  const imageInWindow = filterByTimeWindow(imageInRange, nowMinutes);
  const videoInWindow = filterByTimeWindow(videoInRange, nowMinutes);

  console.log(`  Text időablakban: ${textInWindow.valid.length}, skipped: ${textInWindow.skipped.length}`);
  console.log(`  Image időablakban: ${imageInWindow.valid.length}, skipped: ${imageInWindow.skipped.length}`);
  console.log(`  Video időablakban: ${videoInWindow.valid.length}, skipped: ${videoInWindow.skipped.length}`);

  const hasExecutableItems = (textInWindow.valid.length > 0 || imageInWindow.valid.length > 0 || videoInWindow.valid.length > 0);

  // ===== FÁZIS 4.5: Skip Decision - Sheet nem változott ÉS nincs időablakban =====
  // FIX v7.4.6: ALWAYS_SEND_PREVIEW kapcsoló - ha true, soha nem skip-elünk
  const shouldSendPreview = ALWAYS_SEND_PREVIEW || sheetChanged;
  if (!hasExecutableItems && !shouldSendPreview) {
    console.log('⏭️ SKIP: Sheet nem változott és nincs végrehajtandó művelet (időablakon kívül).');
    console.log('   → Validáció kihagyva, nincs email.');
    console.log('   💡 Tipp: Állítsd ALWAYS_SEND_PREVIEW = true, ha mindig szeretnél preview-t kapni.');
    return;
  }

  // Ha ide jutottunk: VAGY változott a sheet, VAGY ALWAYS_SEND_PREVIEW, VAGY időablakban vagyunk → validálás fut!
  if (!hasExecutableItems && shouldSendPreview) {
    if (ALWAYS_SEND_PREVIEW && !sheetChanged) {
      console.log('📧 ALWAYS_SEND_PREVIEW aktív! Validáció fut, preview email küldése...');
    } else {
      console.log('📧 Sheet változott! Validáció fut, preview email küldése...');
    }
  } else if (hasExecutableItems) {
    console.log('⏰ Időablakban vagyunk! Validáció + végrehajtás fut...');
  }

  // ===== FÁZIS 5: Kampányok Gyűjtése =====
  console.log('\n🔎 Kampányok gyűjtése...');
  // MINDIG az ÖSSZES range-beli sort validáljuk (preview + jövőbeli sorok)
  let textForValidation = textInRange;
  let imageForValidation = imageInRange;
  let videoForValidation = videoInRange;

  const campaigns = gatherCampaigns(textForValidation, imageForValidation, videoForValidation);
  console.log(`  Kampányok (${campaigns.length}): ${campaigns.slice(0, LOG_PREVIEW_COUNT).join(', ')}`);

  if (campaigns.length === 0) {
    console.log('⚠️ Nincs kampány a feldolgozandó sorokban.');
    clearPreviewResultsSheet(ss);
    saveSheetHash(currentHash); // Hash mentése, hogy ne validáljunk újra
    return;
  }

  // ===== FÁZIS 6: Batch Kampány ID Lekérés =====
  console.log('\n🔍 Kampány ID-k lekérése (batch)...');
  const campaignIdMap = getCampaignIdsByNamesBatch(campaigns);
  console.log(`  Talált kampányok: ${Object.keys(campaignIdMap).length}`);

  // ===== FÁZIS 7: Batch Asset Group States Lekérés =====
  console.log('\n🔍 Asset group állapotok lekérése (batch)...');
  const validCampaignIds = Object.values(campaignIdMap);
  const groupStates = fetchAssetGroupStatesBatch(validCampaignIds);
  console.log(`  Asset group-ok: ${Object.keys(groupStates.text).length} text, ${Object.keys(groupStates.images).length} image, ${Object.keys(groupStates.videos).length} video`);

  // ===== FÁZIS 7.5: WindowRowNums Set (időablakos sorok azonosítása) =====
  // KRITIKUS: TEXT, IMAGE és VIDEO külön sheet-ek → külön Set-ek kellenek!
  const textWindowRowNums = new Set();
  const imageWindowRowNums = new Set();
  const videoWindowRowNums = new Set();
  if (hasExecutableItems) {
    textInWindow.valid.forEach(item => textWindowRowNums.add(item.sheetRowNum));
    imageInWindow.valid.forEach(item => imageWindowRowNums.add(item.sheetRowNum));
    videoInWindow.valid.forEach(item => videoWindowRowNums.add(item.sheetRowNum));
    console.log(`\n🔍 Időablakos sorok: TEXT ${textWindowRowNums.size}, IMAGE ${imageWindowRowNums.size}, VIDEO ${videoWindowRowNums.size}`);
  }

  // ===== FÁZIS 7.6: Időablakos itemek action beállítása a validáció előtt =====
  // FIX v7.5.1: Az időablakban lévő itemeknek nincs action-jük (csak filterByDateRange fut rajtuk),
  // ezért a validateAll preview-mode-ban hívta getClosestAction()-t, ami null-t adott vissza
  // "mai napos REMOVE" esetén → LOGIKAI HIBA ERROR. A fix: az action-t a filterByTimeWindow
  // eredményéből (textInWindow.valid) vesszük át, hogy execution-mode-ban validáljon.
  if (hasExecutableItems) {
    const buildWindowActionMap = (windowItems) => {
      const map = new Map();
      windowItems.forEach(wItem => {
        map.set(wItem.sheetRowNum, { action: wItem.action, customHour: wItem.customHour });
      });
      return map;
    };
    const enrichWithWindowAction = (rows, windowActionMap) => {
      if (windowActionMap.size === 0) return rows;
      return rows.map(item => {
        const wMeta = windowActionMap.get(item.sheetRowNum);
        return wMeta ? { ...item, action: wMeta.action, customHour: wMeta.customHour } : item;
      });
    };
    textForValidation = enrichWithWindowAction(textForValidation, buildWindowActionMap(textInWindow.valid));
    imageForValidation = enrichWithWindowAction(imageForValidation, buildWindowActionMap(imageInWindow.valid));
    videoForValidation = enrichWithWindowAction(videoForValidation, buildWindowActionMap(videoInWindow.valid));
  }

  // ===== FÁZIS 8: Validáció =====
  console.log('\n✅ Validáció...');
  const validationResults = validateAll(
    textForValidation,
    imageForValidation,
    videoForValidation,
    campaignIdMap,
    groupStates,
    textWindowRowNums,
    imageWindowRowNums,
    videoWindowRowNums
  );

  console.log(`  OK: ${validationResults.ok}, WARNING: ${validationResults.warnings}, ERROR: ${validationResults.errors}`);
  console.log(`  Jövőbeli sorok: ${validationResults.futureRows.length}, Időablakos sorok: ${validationResults.windowRows.length}`);

  // ===== FÁZIS 8.5: Cross-Row ADD+REMOVE Konfliktus Detektálás =====
  console.log('\n🔍 Cross-row ADD+REMOVE konfliktus ellenőrzése...');
  const conflictResult = detectCrossRowAddRemoveConflicts(
    validationResults.rows,
    validationResults.futureRows,
    validationResults.windowRows
  );

  if (conflictResult.conflictCount > 0) {
    console.log(`⚠️ ${conflictResult.conflictCount} ADD+REMOVE konfliktus (sorok között, ugyanaz az óra) → ERROR státusz`);

    // Frissítjük a validationResults-ot az új ERROR sorokkal
    validationResults.rows = conflictResult.allRows;
    validationResults.futureRows = conflictResult.futureRows;
    validationResults.windowRows = conflictResult.windowRows;

    // Counter újraszámolás (ERROR-ok nőttek)
    const newCounts = countStatuses(validationResults.rows);
    validationResults.ok = newCounts.ok;
    validationResults.warnings = newCounts.warnings;
    validationResults.errors = newCounts.errors;

    console.log(`  Frissített számok: OK: ${validationResults.ok}, WARNING: ${validationResults.warnings}, ERROR: ${validationResults.errors}`);
  }

  // ===== FÁZIS 9: PreviewResults Sheet Írás (csak jövőbeli sorok!) =====
  // KRITIKUS v7.3.26: Csak akkor töröljük/írjuk a lapot, ha vannak új futureRows!
  // Ha csak execution fut (nincs futureRows), maradjon az előző preview eredmény.
  // FIX v7.4.6: Ha lesz execution, SKIP - majd execution után friss állapottal írjuk
  if (validationResults.futureRows.length > 0 && !hasExecutableItems) {
    writePreviewResultsSheet(ss, validationResults.futureRows, campaignIdMap, groupStates);
  }

  // ===== FÁZIS 9.5: Hash Mentés =====
  saveSheetHash(currentHash);

  // ===== FÁZIS 9.6: Preview Email (ha vannak jövőbeli sorok) =====
  // FIX v7.4.6: Ha lesz execution, SKIP - majd execution után friss állapottal küldjük
  if (validationResults.futureRows.length > 0 && !hasExecutableItems) {
    console.log(`\n📧 Preview email küldése (${validationResults.futureRows.length} jövőbeli művelet)...`);
    sendEmail('VALIDATION_PREVIEW', validationResults.futureRows, []);
  }

  // ===== FÁZIS 9.7: Early Return Ha Nincs Időablakban =====
  if (!hasExecutableItems) {
    // Nincs időablakban → preview email már elment (ha volt), most return
    console.log('\n=== PMax Asset Scheduler DONE (Preview Only) ===');
    return;
  }

  // ===== FÁZIS 10: Executable Items Szűrés =====
  // Csak OK és WARNING sorok, ÉS ha időablakban vagyunk → csak az időablakban lévők
  let executableTextItems = validationResults.validTextItems || [];
  let executableImageItems = validationResults.validImageItems || [];
  let executableVideoItems = validationResults.validVideoItems || [];

  if (hasExecutableItems) {
    // KRITIKUS FIX: Map<sheetRowNum, ARRAY of {action, customHour}>
    // Egy sor TÖBBSZÖR is szerepelhet (pl. ADD 10:00 és REMOVE 23:00)
    const textWindowMap = new Map();
    textInWindow.valid.forEach(item => {
      if (!textWindowMap.has(item.sheetRowNum)) {
        textWindowMap.set(item.sheetRowNum, []);
      }
      textWindowMap.get(item.sheetRowNum).push({ action: item.action, customHour: item.customHour });
    });

    const imageWindowMap = new Map();
    imageInWindow.valid.forEach(item => {
      if (!imageWindowMap.has(item.sheetRowNum)) {
        imageWindowMap.set(item.sheetRowNum, []);
      }
      imageWindowMap.get(item.sheetRowNum).push({ action: item.action, customHour: item.customHour });
    });

    const videoWindowMap = new Map();
    videoInWindow.valid.forEach(item => {
      if (!videoWindowMap.has(item.sheetRowNum)) {
        videoWindowMap.set(item.sheetRowNum, []);
      }
      videoWindowMap.get(item.sheetRowNum).push({ action: item.action, customHour: item.customHour });
    });

    // Szűrés ÉS DUPLIKÁLÁS (minden action-re külön executable item!)
    executableTextItems = [];
    (validationResults.validTextItems || []).forEach(entry => {
      if (!textWindowMap.has(entry.item.sheetRowNum)) return;

      const windowMetas = textWindowMap.get(entry.item.sheetRowNum);

      // Minden action-re külön executable item!
      windowMetas.forEach(meta => {
        executableTextItems.push({
          item: { ...entry.item, action: meta.action, customHour: meta.customHour },
          validGroups: entry.validGroups
        });
      });
    });

    executableImageItems = [];
    (validationResults.validImageItems || []).forEach(entry => {
      if (!imageWindowMap.has(entry.item.sheetRowNum)) return;

      const windowMetas = imageWindowMap.get(entry.item.sheetRowNum);

      // Minden action-re külön executable item!
      windowMetas.forEach(meta => {
        executableImageItems.push({
          item: { ...entry.item, action: meta.action, customHour: meta.customHour },
          validGroups: entry.validGroups
        });
      });
    });

    executableVideoItems = [];
    (validationResults.validVideoItems || []).forEach(entry => {
      if (!videoWindowMap.has(entry.item.sheetRowNum)) return;

      const windowMetas = videoWindowMap.get(entry.item.sheetRowNum);

      // Minden action-re külön executable item!
      windowMetas.forEach(meta => {
        executableVideoItems.push({
          item: { ...entry.item, action: meta.action, customHour: meta.customHour },
          validGroups: entry.validGroups,
          parsedIdentifier: entry.parsedIdentifier
        });
      });
    });

    console.log(`\n⏰ Időablak szűrés: ${executableTextItems.length} TEXT item, ${executableImageItems.length} IMAGE item, ${executableVideoItems.length} VIDEO item az időablakban`);
  }

  // Végrehajtható SOROK száma (group-onként külön számolva!)
  const executableRowCount = validationResults.rows.filter(row => row[7] === 'OK' || row[7] === 'WARNING').length;

  console.log(`\n📋 Végrehajtható műveletek: ${executableRowCount} (OK + WARNING sorok)`);

  if (validationResults.errors > 0) {
    const errorRowCount = validationResults.rows.filter(row => row[7] === 'ERROR').length;
    console.log(`⚠️ ${errorRowCount} ERROR sor kihagyva (nem kerül végrehajtásra)`);
  }

  if (executableTextItems.length === 0 && executableImageItems.length === 0 && executableVideoItems.length === 0) {
    console.log('ℹ️ Nincs végrehajtható művelet (csak ERROR sorok).');

    // Időablakban vagyunk, de minden ERROR → Results sheet + Execution email
    appendResultsSheet(ss, validationResults.windowRows, campaignIdMap, groupStates);

    if (validationResults.windowRows.length > 0) {
      console.log(`\n📧 Execution email küldése (${validationResults.windowRows.length} időablakos művelet - csak ERROR-ok)...`);
      sendEmail('EXECUTION_COMPLETE', validationResults.windowRows, []);
    }

    // FIX v7.4.6: Post-execution preview az "All Errors" ágban is!
    // FIX v7.5.3: Ugyanolyan feltétel mint a normál ágban (futureRows VAGY ALWAYS_SEND_PREVIEW)
    // FIX v7.5.3: textInRange/imageInRange/videoInRange (nem enrichelt!) - v7.5.2 fix alkalmazása
    const strictFutureText = filterStrictFutureRows(textInRange);
    const strictFutureImage = filterStrictFutureRows(imageInRange);
    const strictFutureVideo = filterStrictFutureRows(videoInRange);

    console.log(`  📅 Szigorúan jövőbeli sorok: Text ${strictFutureText.length}, Image ${strictFutureImage.length}, Video ${strictFutureVideo.length}`);

    const shouldRunAllErrPreview = strictFutureText.length > 0 || strictFutureImage.length > 0 || strictFutureVideo.length > 0 || ALWAYS_SEND_PREVIEW;
    if (shouldRunAllErrPreview) {
      console.log('\n🔄 Post-execution preview validáció (friss állapottal)...');

      const freshGroupStates = fetchAssetGroupStatesBatch(Object.values(campaignIdMap));
      const freshValidation = validateAll(
        strictFutureText,
        strictFutureImage,
        strictFutureVideo,
        campaignIdMap,
        freshGroupStates,
        new Set(),  // Nincs időablakos sor → minden futureRows lesz
        new Set(),
        new Set()
      );

      console.log(`  OK: ${freshValidation.ok}, WARNING: ${freshValidation.warnings}, ERROR: ${freshValidation.errors}`);
      console.log(`  Jövőbeli sorok: ${freshValidation.futureRows.length}`);

      const strictFutureResults = filterStrictFutureValidationRows(freshValidation.futureRows);
      console.log(`  📅 Szigorúan jövőbeli eredmények: ${strictFutureResults.length} (szűrés előtt: ${freshValidation.futureRows.length})`);

      if (strictFutureResults.length > 0) {
        writePreviewResultsSheet(ss, strictFutureResults, campaignIdMap, freshGroupStates);
        console.log(`\n📧 Preview email küldése (${strictFutureResults.length} jövőbeli művelet, friss állapot)...`);
        sendEmail('VALIDATION_PREVIEW', strictFutureResults, []);
      } else {
        console.log('  ℹ️ Nincs szigorúan jövőbeli művelet a preview-hoz (szűrés után).');
      }
    } else {
      console.log('  ℹ️ Nincs jövőbeli művelet, preview kihagyva.');
    }

    console.log('\n=== PMax Asset Scheduler DONE (Execution - All Errors) ===');
    return;
  }

  // ===== FÁZIS 10.5: Cross-Row Duplikáció Detektálás & Deduplikáció =====
  console.log('\n🔍 Cross-row duplikáció ellenőrzése...');
  const dedupResult = deduplicateExecutableItems(executableTextItems, executableImageItems, executableVideoItems, campaignIdMap, groupStates);
  executableTextItems = dedupResult.textItems;
  executableImageItems = dedupResult.imageItems;
  executableVideoItems = dedupResult.videoItems;

  if (dedupResult.duplicateCount > 0) {
    console.log(`⚠️ ${dedupResult.duplicateCount} duplikált művelet kiszűrve (sorok között azonos asset + action + hour + group)`);
  }

  // ===== FÁZIS 11: Végrehajtás =====
  console.log('\n▶️ Végrehajtás...');
  const executionResults = executeAll(
    executableTextItems,
    executableImageItems,
    executableVideoItems,
    campaignIdMap,
    groupStates
  );

  console.log(`  Végrehajtott műveletek: ${executionResults.length}`);

  // ===== FÁZIS 12: Post-Verification =====
  console.log('\n🔍 Post-verification...');
  const verifiedResults = postVerifyAll(executionResults, groupStates);

  // ===== FÁZIS 12.5: Merge Window + Verified Results (EMAIL ÉS SHEET SZÁMÁRA!) =====
  // KRITIKUS FIX: windowRows merge-elése verifiedResults-szal
  // Email-ben ÉS Results sheet-ben is a merge-elt eredmények kellenek!
  const mergedWindowRows = mergeWindowAndExecutionRows(validationResults.windowRows, verifiedResults);

  // ===== FÁZIS 13: Results Sheet Írás (APPEND merge-elt sorok!) =====
  // KRITIKUS FIX: Merge-elt eredményeket írjuk (SUCCESS [Verified ✓] státusszal!)
  appendResultsSheet(ss, mergedWindowRows, campaignIdMap, groupStates);

  // ===== FÁZIS 14: Execution Email (ha vannak időablakos sorok) =====
  if (mergedWindowRows.length > 0) {
    console.log(`\n📧 Execution email küldése (${mergedWindowRows.length} időablakos művelet)...`);
    // Email-ben már nem kell merge (már merge-elt sorokat kapott!)
    // Második paraméter = previewRows (amit a táblázatban mutat)
    // Harmadik paraméter = executionRows (üres, mert már merge-elve van)
    sendEmail('EXECUTION_COMPLETE', mergedWindowRows, []);
  }

  // ===== FÁZIS 15: Post-Execution Preview (ha vannak jövőbeli sorok VAGY ALWAYS_SEND_PREVIEW) =====
  // FIX v7.4.6: Execution után újra validálni a jövőbeli sorokat a FRISS groupStates alapján
  // Így a preview az execution UTÁNI állapotot tükrözi
  // FIX v7.4.6: ALWAYS_SEND_PREVIEW esetén mindig futtatjuk
  const shouldRunPostExecPreview = validationResults.futureRows.length > 0 || ALWAYS_SEND_PREVIEW;
  if (shouldRunPostExecPreview) {
    console.log('\n🔄 Post-execution preview validáció (friss állapottal)...');

    // FIX v7.4.6: Kiszűrjük a jelenlegi órában lévő műveleteket (azok már végrehajtódtak)
    // Csak a SZIGORÚAN jövőbeli műveletek maradnak
    // FIX v7.5.2: textInRange/imageInRange/videoInRange (eredeti, enrichment nélkül) szükséges!
    // textForValidation enrichelt (window item action beállítva), ami felülírja a strictFutureAction-t
    // → ADD-nak enrichelt sor REMOVE-ot is tartalmazva helytelenül ADD-ként validálna preview-ban.
    const strictFutureText = filterStrictFutureRows(textInRange);
    const strictFutureImage = filterStrictFutureRows(imageInRange);
    const strictFutureVideo = filterStrictFutureRows(videoInRange);

    console.log(`  📅 Szigorúan jövőbeli sorok: Text ${strictFutureText.length}, Image ${strictFutureImage.length}, Video ${strictFutureVideo.length}`);

    if (strictFutureText.length === 0 && strictFutureImage.length === 0 && strictFutureVideo.length === 0) {
      console.log('  ℹ️ Nincs szigorúan jövőbeli művelet a preview-hoz.');
    } else {
      // Újra lekérni a groupStates-t az execution utáni állapottal
      console.log('  🔍 Friss asset group állapotok lekérése...');
      const freshGroupStates = fetchAssetGroupStatesBatch(Object.values(campaignIdMap));

      // Újra validálni a jövőbeli sorokat
      console.log('  ✅ Jövőbeli sorok újravalidálása...');
      const freshValidation = validateAll(
        strictFutureText,
        strictFutureImage,
        strictFutureVideo,
        campaignIdMap,
        freshGroupStates,
        new Set(),  // Nincs időablakos sor (csak preview)
        new Set(),
        new Set()
      );

      console.log(`  OK: ${freshValidation.ok}, WARNING: ${freshValidation.warnings}, ERROR: ${freshValidation.errors}`);
      console.log(`  Friss jövőbeli sorok: ${freshValidation.futureRows.length}`);

      // FIX v7.4.6: Kiszűrjük a validation eredményekből is a jelenlegi órás action-öket
      const strictFutureResults = filterStrictFutureValidationRows(freshValidation.futureRows);
      console.log(`  📅 Szigorúan jövőbeli eredmények: ${strictFutureResults.length} (szűrés előtt: ${freshValidation.futureRows.length})`);

      // PreviewResults sheet írás friss adatokkal
      if (strictFutureResults.length > 0) {
        writePreviewResultsSheet(ss, strictFutureResults, campaignIdMap, freshGroupStates);

        // Preview email küldése friss adatokkal
        console.log(`\n📧 Preview email küldése (${strictFutureResults.length} jövőbeli művelet, friss állapot)...`);
        sendEmail('VALIDATION_PREVIEW', strictFutureResults, []);
      } else {
        console.log('  ℹ️ Nincs szigorúan jövőbeli művelet a preview-hoz (szűrés után).');
      }
    }
  }

  console.log('\n=== PMax Asset Scheduler DONE ===');
}

/*** ===================== SHEET BEOLVASÁS ===================== ***/

function validateSheetStructure(ss) {
  const requiredSheets = {
    [TEXT_SHEET_NAME]: ['Campaign Name', 'Asset Group Name', 'Text Type', 'Text', 'Add Date', 'Add Hour', 'Remove Date', 'Remove Hour'],
    [IMAGE_SHEET_NAME]: ['Campaign Name', 'Asset Group Name', 'Image Type', 'Asset ID', 'Add Date', 'Add Hour', 'Remove Date', 'Remove Hour'],
    [VIDEO_SHEET_NAME]: ['Campaign Name', 'Asset Group Name', 'Video Aspect (optional)', 'Video URL or ID / Asset ID', 'Add Date', 'Add Hour', 'Remove Date', 'Remove Hour']
  };

  for (const [sheetName, columns] of Object.entries(requiredSheets)) {
    const sh = ss.getSheetByName(sheetName);
    if (!sh) throw new Error(`Hiányzó sheet: ${sheetName}`);

    const headers = sh.getRange(1, 1, 1, sh.getLastColumn()).getValues()[0];
    const missing = columns.filter(col => !headers.includes(col));
    if (missing.length > 0) {
      throw new Error(`Hiányzó oszlopok (${sheetName}): ${missing.join(', ')}`);
    }
  }
}

function setupImageAssetsThumbnail(ss) {
  const sh = ss.getSheetByName(IMAGE_SHEET_NAME);
  if (!sh) return; // Ha nincs ImageAssets sheet, nem csinálunk semmit

  const headers = sh.getRange(1, 1, 1, sh.getLastColumn()).getValues()[0];
  if (!headers.includes('Thumbnail')) {
    // Hozzáadjuk a végéhez
    const lastCol = sh.getLastColumn() + 1;
    sh.getRange(1, lastCol).setValue('Thumbnail').setFontWeight('bold').setBackground('#f6f8fa');
    sh.setColumnWidth(lastCol, 120);
    console.log('  ✅ ImageAssets: Thumbnail oszlop hozzáadva');
  }
}

function setupVideoAssetsSheet(ss) {
  let sh = ss.getSheetByName(VIDEO_SHEET_NAME);
  if (!sh) {
    sh = ss.insertSheet(VIDEO_SHEET_NAME);

    // Header sor
    sh.appendRow([
      'Campaign Name',
      'Asset Group Name',
      'Video Aspect (optional)',
      'Video URL or ID / Asset ID',
      'Add Date',
      'Add Hour',
      'Remove Date',
      'Remove Hour',
      'Thumbnail'
    ]);

    sh.getRange(1, 1, 1, 9).setFontWeight('bold').setBackground('#f6f8fa');

    // Video Aspect legördülő (C oszlop, C2:C1000 range)
    const aspectRule = SpreadsheetApp.newDataValidation()
      .requireValueInList([
        '',
        'HORIZONTAL (16:9)',
        'VERTICAL (9:16)',
        'SQUARE (1:1)'
      ], true)
      .setAllowInvalid(false)
      .build();

    sh.getRange('C2:C1000').setDataValidation(aspectRule);

    // Thumbnail oszlop szélesség
    sh.setColumnWidth(9, 120);

    console.log('  ✅ VideoAssets sheet létrehozva + legördülő validáció beállítva');
  } else {
    // Meglévő sheet - ellenőrizzük, hogy van-e Thumbnail oszlop
    const headers = sh.getRange(1, 1, 1, sh.getLastColumn()).getValues()[0];
    if (!headers.includes('Thumbnail')) {
      // Hozzáadjuk a végéhez
      const lastCol = sh.getLastColumn() + 1;
      sh.getRange(1, lastCol).setValue('Thumbnail').setFontWeight('bold').setBackground('#f6f8fa');
      sh.setColumnWidth(lastCol, 120);
      console.log('  ✅ VideoAssets: Thumbnail oszlop hozzáadva');
    }
  }
}

function readSheet(ss, sheetName) {
  const sh = ss.getSheetByName(sheetName);
  if (!sh) {
    console.log(`  ⚠️ "${sheetName}" lap nem található.`);
    return { headers: [], rows: [] };
  }

  const data = sh.getDataRange().getValues();
  if (!data || data.length < 2) {
    console.log(`  ℹ️ "${sheetName}" lap üres vagy nincs adat.`);
    return { headers: [], rows: [] };
  }

  const headers = data[0].map(String);
  const rows = data.slice(1);

  return { headers, rows };
}

function filterNonEmptyRows(input, keyColumn, sheetName) {
  if (!input.headers || !input.rows) return [];

  const idx = input.headers.indexOf(keyColumn);
  if (idx === -1) {
    console.log(`  ⚠️ "${keyColumn}" oszlop nem található (${sheetName}).`);
    return [];
  }

  const out = [];

  input.rows.forEach((row, i) => {
    const sheetRowNum = i + 2; // 1. sor a header
    const value = String(row[idx] || '').trim();
    if (!value) {
      console.log(`    ℹ️ ${sheetName} Row ${sheetRowNum} skipped (empty ${keyColumn})`);
      return;
    }
    out.push({ rawRow: row, headers: input.headers, sheetRowNum });
  });

  return out;
}

/*** ===================== DÁTUM ÉS IDŐABLAK SZŰRÉS ===================== ***/

function filterByDateRange(rows, todayStr) {
  const tz = ACCOUNT_TIMEZONE;
  const now = new Date();
  const nowMinutes = parseInt(Utilities.formatDate(now, tz, 'H'), 10) * 60 + parseInt(Utilities.formatDate(now, tz, 'm'), 10);

  // Mai dátum (string formátum)
  const todayDate = new Date(todayStr + 'T00:00:00');

  // Jövőbeli max dátum (mai + VALIDATION_FUTURE_DAYS nap)
  const maxFutureDate = new Date(todayDate);
  maxFutureDate.setDate(maxFutureDate.getDate() + VALIDATION_FUTURE_DAYS);
  const maxFutureStr = Utilities.formatDate(maxFutureDate, tz, 'yyyy-MM-dd');

  console.log(`  📅 Dátum szűrés: ${todayStr} → ${maxFutureStr} (${VALIDATION_FUTURE_DAYS} nap), now: ${minutesToHHMM(nowMinutes)}`);

  const filtered = [];

  rows.forEach(item => {
    const { rawRow, headers, sheetRowNum } = item;

    const idxAddDate = headers.indexOf('Add Date');
    const idxAddHour = headers.indexOf('Add Hour');
    const idxRemoveDate = headers.indexOf('Remove Date');
    const idxRemoveHour = headers.indexOf('Remove Hour');

    let hasAddInRange = false;
    let hasRemoveInRange = false;

    // Add Date check: ma VAGY jövő VALIDATION_FUTURE_DAYS napon belül
    if (idxAddDate !== -1 && rawRow[idxAddDate]) {
      try {
        const addDate = new Date(rawRow[idxAddDate]);
        const addDateStr = Utilities.formatDate(addDate, tz, 'yyyy-MM-dd');

        // String alapú összehasonlítás (időzóna-biztos)
        if (addDateStr >= todayStr && addDateStr <= maxFutureStr) {
          // Ha MA van a dátum, ellenőrizzük az órát is
          if (addDateStr === todayStr) {
            const addHourRaw = idxAddHour !== -1 ? String(rawRow[idxAddHour] ?? '').trim() : '';
            const addHourParsed = parseCustomHourToInt(addHourRaw);

            if (addHourParsed.ok) {
              // Van megadott óra
              if (addHourParsed.value !== null) {
                // Custom hour: ellenőrizzük hogy még nem telt el
                const addHourEnd = (addHourParsed.value + 1) * 60; // Exclusive upper bound
                if (nowMinutes < addHourEnd) {
                  hasAddInRange = true;
                  console.log(`    ✅ Row ${sheetRowNum} Add Date+Hour in range: ${addDateStr} ${pad2(addHourParsed.value)}:00 (current: ${minutesToHHMM(nowMinutes)})`);
                } else {
                  console.log(`    ⏭️ Row ${sheetRowNum} Add Date OK but Hour elapsed: ${addDateStr} ${pad2(addHourParsed.value)}:00 (current: ${minutesToHHMM(nowMinutes)})`);
                }
              } else {
                // Nincs custom hour → default 00:00-01:00
                if (nowMinutes < ADD_WINDOW_TO_MIN) {
                  hasAddInRange = true;
                  console.log(`    ✅ Row ${sheetRowNum} Add Date in range (default hour): ${addDateStr} 00:00-00:59 (current: ${minutesToHHMM(nowMinutes)})`);
                } else {
                  console.log(`    ⏭️ Row ${sheetRowNum} Add Date OK but default hour elapsed: ${addDateStr} (current: ${minutesToHHMM(nowMinutes)})`);
                }
              }
            } else {
              // Érvénytelen óra formátum → skip
              console.log(`    ⚠️ Row ${sheetRowNum} Invalid Add Hour: ${addHourRaw}`);
            }
          } else {
            // Jövőbeli dátum → mindig in range
            hasAddInRange = true;
            console.log(`    ✅ Row ${sheetRowNum} Add Date in range (future): ${addDateStr}`);
          }
        } else {
          console.log(`    ⏭️ Row ${sheetRowNum} Add Date out of range: ${addDateStr} (${todayStr}-${maxFutureStr})`);
        }
      } catch (e) {
        console.log(`    ⚠️ Row ${sheetRowNum} Invalid Add Date: ${rawRow[idxAddDate]}`);
      }
    }

    // Remove Date check: ma VAGY jövő VALIDATION_FUTURE_DAYS napon belül
    if (idxRemoveDate !== -1 && rawRow[idxRemoveDate]) {
      try {
        const remDate = new Date(rawRow[idxRemoveDate]);
        const remDateStr = Utilities.formatDate(remDate, tz, 'yyyy-MM-dd');

        // String alapú összehasonlítás (időzóna-biztos)
        if (remDateStr >= todayStr && remDateStr <= maxFutureStr) {
          // Ha MA van a dátum, ellenőrizzük az órát is
          if (remDateStr === todayStr) {
            const remHourRaw = idxRemoveHour !== -1 ? String(rawRow[idxRemoveHour] ?? '').trim() : '';
            const remHourParsed = parseCustomHourToInt(remHourRaw);

            if (remHourParsed.ok) {
              // Van megadott óra
              if (remHourParsed.value !== null) {
                // Custom hour: ellenőrizzük hogy még nem telt el
                const remHourEnd = (remHourParsed.value + 1) * 60; // Exclusive upper bound
                if (nowMinutes < remHourEnd) {
                  hasRemoveInRange = true;
                  console.log(`    ✅ Row ${sheetRowNum} Remove Date+Hour in range: ${remDateStr} ${pad2(remHourParsed.value)}:00 (current: ${minutesToHHMM(nowMinutes)})`);
                } else {
                  console.log(`    ⏭️ Row ${sheetRowNum} Remove Date OK but Hour elapsed: ${remDateStr} ${pad2(remHourParsed.value)}:00 (current: ${minutesToHHMM(nowMinutes)})`);
                }
              } else {
                // Nincs custom hour → default 23:00-24:00
                if (nowMinutes < REMOVE_WINDOW_TO_MIN) {
                  hasRemoveInRange = true;
                  console.log(`    ✅ Row ${sheetRowNum} Remove Date in range (default hour): ${remDateStr} 23:00-23:59 (current: ${minutesToHHMM(nowMinutes)})`);
                } else {
                  console.log(`    ⏭️ Row ${sheetRowNum} Remove Date OK but default hour elapsed: ${remDateStr} (current: ${minutesToHHMM(nowMinutes)})`);
                }
              }
            } else {
              // Érvénytelen óra formátum → skip
              console.log(`    ⚠️ Row ${sheetRowNum} Invalid Remove Hour: ${remHourRaw}`);
            }
          } else {
            // Jövőbeli dátum → mindig in range
            hasRemoveInRange = true;
            console.log(`    ✅ Row ${sheetRowNum} Remove Date in range (future): ${remDateStr}`);
          }
        } else {
          console.log(`    ⏭️ Row ${sheetRowNum} Remove Date out of range: ${remDateStr} (${todayStr}-${maxFutureStr})`);
        }
      } catch (e) {
        console.log(`    ⚠️ Row ${sheetRowNum} Invalid Remove Date: ${rawRow[idxRemoveDate]}`);
      }
    }

    // Csak akkor adjuk hozzá ha legalább egy művelet in range
    if (hasAddInRange || hasRemoveInRange) {
      filtered.push({ ...item, hasAddInRange, hasRemoveInRange });
    }
  });

  return filtered;
}

function filterByTimeWindow(rows, nowMinutes) {
  const valid = [];
  const skipped = [];
  const tz = ACCOUNT_TIMEZONE;
  const now = new Date();
  const todayStr = Utilities.formatDate(now, tz, 'yyyy-MM-dd');

  rows.forEach(item => {
    const { rawRow, headers, sheetRowNum } = item;

    const idxAddDate = headers.indexOf('Add Date');
    const idxAddHour = headers.indexOf('Add Hour');
    const idxRemoveDate = headers.indexOf('Remove Date');
    const idxRemoveHour = headers.indexOf('Remove Hour');

    // Ellenőrizzük hogy Add Date ma van-e
    let hasAddToday = false;
    if (idxAddDate !== -1 && rawRow[idxAddDate]) {
      const addDate = new Date(rawRow[idxAddDate]);
      const addDateStr = Utilities.formatDate(addDate, tz, 'yyyy-MM-dd');
      if (addDateStr === todayStr) hasAddToday = true;
    }

    // Ellenőrizzük hogy Remove Date ma van-e
    let hasRemoveToday = false;
    if (idxRemoveDate !== -1 && rawRow[idxRemoveDate]) {
      const remDate = new Date(rawRow[idxRemoveDate]);
      const remDateStr = Utilities.formatDate(remDate, tz, 'yyyy-MM-dd');
      if (remDateStr === todayStr) hasRemoveToday = true;
    }

    // Ha egyik sem ma (safety check, filterByDate-nél már kiszűrtük)
    if (!hasAddToday && !hasRemoveToday) {
      skipped.push({ item, reason: 'Nincs Add/Remove Date ma' });
      return;
    }

    // Process ADD if today
    if (hasAddToday) {
      const hourRaw = idxAddHour !== -1 ? String(rawRow[idxAddHour] ?? '').trim() : '';
      const hourParse = parseCustomHourToInt(hourRaw);

      if (!hourParse.ok) {
        skipped.push({ item, reason: `INVALID ADD Hour: ${hourParse.error}` });
      } else {
        const within = isWithinTimeWindow(nowMinutes, hourParse.value, ADD_WINDOW_FROM_MIN, ADD_WINDOW_TO_MIN);
        if (within) {
          valid.push({ ...item, action: 'ADD', customHour: hourParse.value });
        } else {
          const windowStr = formatTimeWindow(hourParse.value, ADD_WINDOW_FROM_MIN, ADD_WINDOW_TO_MIN);
          skipped.push({ item, reason: `ADD ablakon kívül (${windowStr})` });
        }
      }
    }

    // Process REMOVE if today
    if (hasRemoveToday) {
      const hourRaw = idxRemoveHour !== -1 ? String(rawRow[idxRemoveHour] ?? '').trim() : '';
      const hourParse = parseCustomHourToInt(hourRaw);

      if (!hourParse.ok) {
        skipped.push({ item, reason: `INVALID REMOVE Hour: ${hourParse.error}` });
      } else {
        const within = isWithinTimeWindow(nowMinutes, hourParse.value, REMOVE_WINDOW_FROM_MIN, REMOVE_WINDOW_TO_MIN);
        if (within) {
          valid.push({ ...item, action: 'REMOVE', customHour: hourParse.value });
        } else {
          const windowStr = formatTimeWindow(hourParse.value, REMOVE_WINDOW_FROM_MIN, REMOVE_WINDOW_TO_MIN);
          skipped.push({ item, reason: `REMOVE ablakon kívül (${windowStr})` });
        }
      }
    }
  });

  return { valid, skipped };
}

function parseCustomHourToInt(raw) {
  const s = String(raw || '').trim();
  if (!s) return { ok: true, value: null };

  // KÖZEPES PRIORITÁSÚ FIX: Megengedjük a "10" és "10:00" formátumot is
  const m = /^(\d{1,2})(?::00)?$/.exec(s);
  if (!m) return { ok: false, error: `Rossz formátum: "${s}". Várt: 0-23 vagy HH:00 (pl. 10 vagy 10:00).` };

  const hour = parseInt(m[1], 10);
  if (isNaN(hour) || hour < 0 || hour > 23) {
    return { ok: false, error: `Érvénytelen óraszám: "${s}". Várt: 0-23.` };
  }

  return { ok: true, value: hour };
}

function isWithinTimeWindow(nowMinutes, customHourIntOrNull, defaultFrom, defaultTo) {
  let from, to;

  if (customHourIntOrNull !== null && customHourIntOrNull !== undefined) {
    // Custom hour: H:00 - (H+1):00 (exclusive upper bound)
    // Pl. hour=14 → 840-900 → [14:00, 15:00) → 14:00-14:59
    from = customHourIntOrNull * 60;
    to = customHourIntOrNull * 60 + 60;
  } else {
    // Default window
    from = defaultFrom;
    to = defaultTo;
  }

  // EXCLUSIVE upper bound: >= from AND < to
  return nowMinutes >= from && nowMinutes < to;
}

function formatTimeWindow(customHourIntOrNull, defaultFrom, defaultTo) {
  if (customHourIntOrNull !== null && customHourIntOrNull !== undefined) {
    const h = customHourIntOrNull;
    return `${pad2(h)}:00-${pad2(h)}:59`;
  }

  // EXCLUSIVE upper bound: defaultTo-1 az utolsó tényleges perc
  const fromH = Math.floor(defaultFrom / 60);
  const fromM = defaultFrom % 60;
  const actualTo = defaultTo - 1;  // Exclusive upper → utolsó tényleges perc
  const toH = Math.floor(actualTo / 60);
  const toM = actualTo % 60;

  return `${pad2(fromH)}:${pad2(fromM)}-${pad2(toH)}:${pad2(toM)}`;
}

function pad2(n) {
  return (n < 10 ? '0' : '') + n;
}

function minutesToHHMM(mins) {
  const h = Math.floor(mins / 60);
  const m = mins % 60;
  return `${pad2(h)}:${pad2(m)}`;
}

/**
 * Kiszűri azokat a sorokat, ahol az action (ADD/REMOVE) a jelenlegi órában van.
 * Csak a SZIGORÚAN jövőbeli műveleteket hagyja meg.
 * FIX v7.4.6: Post-execution preview-hoz, hogy ne jelenjenek meg az épp végrehajtott/próbált műveletek.
 * FIX v7.4.6: Beállítja az item.action mezőt a jövőbeli action-re (ADD vagy REMOVE)!
 * @param {Array} rows - A filterByDateRange által visszaadott sorok
 * @returns {Array} Csak a jövőbeli sorok, módosított action mezővel
 */
function filterStrictFutureRows(rows) {
  const tz = ACCOUNT_TIMEZONE;
  const now = new Date();
  const todayStr = Utilities.formatDate(now, tz, 'yyyy-MM-dd');
  const currentHour = parseInt(Utilities.formatDate(now, tz, 'H'), 10);

  const result = [];

  rows.forEach(item => {
    const { rawRow, headers } = item;

    const idxAddDate = headers.indexOf('Add Date');
    const idxAddHour = headers.indexOf('Add Hour');
    const idxRemoveDate = headers.indexOf('Remove Date');
    const idxRemoveHour = headers.indexOf('Remove Hour');

    let hasStrictFutureAdd = false;
    let hasStrictFutureRemove = false;

    // ADD: csak ha dátum > ma, VAGY (dátum === ma ÉS óra > jelenlegi óra)
    if (idxAddDate !== -1 && rawRow[idxAddDate]) {
      try {
        const addDate = new Date(rawRow[idxAddDate]);
        const addDateStr = Utilities.formatDate(addDate, tz, 'yyyy-MM-dd');

        if (addDateStr > todayStr) {
          hasStrictFutureAdd = true;  // Jövőbeli nap
        } else if (addDateStr === todayStr) {
          const addHourRaw = idxAddHour !== -1 ? String(rawRow[idxAddHour] ?? '').trim() : '';
          const addHourParsed = parseCustomHourToInt(addHourRaw);
          if (addHourParsed.ok) {
            // FIX v7.5.3: null = üres cella → default ADD ablak kezdőórája
            const effectiveAddHour = addHourParsed.value !== null ? addHourParsed.value : Math.floor(ADD_WINDOW_FROM_MIN / 60);
            if (effectiveAddHour > currentHour) {
              hasStrictFutureAdd = true;  // Ma, de későbbi óra
            }
          }
        }
      } catch (e) { /* skip */ }
    }

    // REMOVE: csak ha dátum > ma, VAGY (dátum === ma ÉS óra > jelenlegi óra)
    if (idxRemoveDate !== -1 && rawRow[idxRemoveDate]) {
      try {
        const remDate = new Date(rawRow[idxRemoveDate]);
        const remDateStr = Utilities.formatDate(remDate, tz, 'yyyy-MM-dd');

        if (remDateStr > todayStr) {
          hasStrictFutureRemove = true;  // Jövőbeli nap
        } else if (remDateStr === todayStr) {
          const remHourRaw = idxRemoveHour !== -1 ? String(rawRow[idxRemoveHour] ?? '').trim() : '';
          const remHourParsed = parseCustomHourToInt(remHourRaw);
          if (remHourParsed.ok) {
            // FIX v7.5.3: null = üres cella → default REMOVE ablak kezdőórája
            const effectiveRemHour = remHourParsed.value !== null ? remHourParsed.value : Math.floor(REMOVE_WINDOW_FROM_MIN / 60);
            if (effectiveRemHour > currentHour) {
              hasStrictFutureRemove = true;  // Ma, de későbbi óra
            }
          }
        }
      } catch (e) { /* skip */ }
    }

    // Ha van legalább egy szigorúan jövőbeli action, megtartjuk a sort
    if (hasStrictFutureAdd || hasStrictFutureRemove) {
      // FIX v7.4.6: Beállítjuk az action mezőt a JÖVŐBELI action-re!
      // Ha mindkettő jövőbeli, akkor a korábbi (ADD ha van)
      // Ha csak az egyik jövőbeli, akkor azt
      let futureAction = null;
      if (hasStrictFutureAdd && hasStrictFutureRemove) {
        futureAction = 'ADD';  // Mindkettő jövő → ADD a korábbi
      } else if (hasStrictFutureAdd) {
        futureAction = 'ADD';
      } else {
        futureAction = 'REMOVE';
      }

      // Létrehozunk egy módosított item-et az explicit action-nel
      // FIX v7.4.6: strictFutureAction mezőt használunk, nem action-t!
      // Így a scheduledCol továbbra is allScheduledActions lesz (mindkét dátum)
      result.push({
        ...item,
        strictFutureAction: futureAction,
        // A customHour-t is be kell állítani az adott action-hez
        customHour: futureAction === 'ADD'
          ? (idxAddHour !== -1 ? parseCustomHourToInt(String(rawRow[idxAddHour] ?? '').trim()).value : null)
          : (idxRemoveHour !== -1 ? parseCustomHourToInt(String(rawRow[idxRemoveHour] ?? '').trim()).value : null)
      });
    }
  });

  return result;
}

/**
 * KRITIKUS FIX: Egyszerűsített logika - csak az effectiveAction időpontját nézzük
 *
 * PROBLÉMA volt: Ha egy sorban ADD már végrehajtódott (pl. 10:00) de van jövőbeli REMOVE is (pl. 23:00),
 * a függvény megtartotta a sort, de az effectiveAction továbbra is 'ADD' maradt.
 * Így a preview-ban megjelent "ADD 10:00" ami már végrehajtódott - zavaró!
 *
 * MEGOLDÁS: Csak azt nézzük, hogy az effectiveAction időpontja SZIGORÚAN jövőbeli-e.
 * Ha nem → skip (nem jelenik meg a preview-ban)
 * Ha igen → megtartjuk
 */
function filterStrictFutureValidationRows(validationRows) {
  const tz = ACCOUNT_TIMEZONE;
  const now = new Date();
  const todayStr = Utilities.formatDate(now, tz, 'yyyy-MM-dd');
  const currentHour = parseInt(Utilities.formatDate(now, tz, 'H'), 10);

  console.log(`    🔍 filterStrictFutureValidationRows: todayStr=${todayStr}, currentHour=${currentHour}`);

  return validationRows.filter(row => {
    const scheduledCol = row[5] || '';
    const effectiveAction = row[6] || '';

    // Parse scheduledCol
    const scheduleMatches = scheduledCol.match(/(\d{4}-\d{2}-\d{2}) (\d{2}):00/g);

    if (!scheduleMatches || scheduleMatches.length === 0) {
      return true; // Ha nem tudjuk parse-olni, megtartjuk
    }

    // KRITIKUS FIX: Csak az effectiveAction időpontját nézzük,
    // nem foglalkozunk a többi action-nel a scheduledCol-ban

    // ADD vagy ADD+REMOVE esetén
    if (effectiveAction === 'ADD' || effectiveAction === 'ADD+REMOVE') {
      const addMatch = scheduledCol.match(/ADD:\s*(\d{4}-\d{2}-\d{2}) (\d{2}):00/);
      const firstMatch = addMatch || scheduledCol.match(/^(\d{4}-\d{2}-\d{2}) (\d{2}):00/);

      if (firstMatch) {
        const dateStr = firstMatch[1];
        const hour = parseInt(firstMatch[2], 10);
        // Skip ha NEM jövőbeli (ma ÉS hour <= current)
        if (dateStr === todayStr && hour <= currentHour) {
          return false;
        }
      }
    }

    // REMOVE esetén
    if (effectiveAction === 'REMOVE') {
      const removeMatch = scheduledCol.match(/REMOVE:\s*(\d{4}-\d{2}-\d{2}) (\d{2}):00/);
      const firstMatch = removeMatch || scheduledCol.match(/^(\d{4}-\d{2}-\d{2}) (\d{2}):00/);

      if (firstMatch) {
        const dateStr = firstMatch[1];
        const hour = parseInt(firstMatch[2], 10);
        if (dateStr === todayStr && hour <= currentHour) {
          return false;
        }
      }
    }

    return true; // Megtartjuk (jövőbeli vagy nem tudjuk eldönteni)
  });
}

/**
 * Meghatározza a legközelebbi jövőbeli műveletet egy sorban
 * @param {object} item - A sor (rawRow, headers)
 * @param {Date} now - Jelenlegi időpont
 * @returns {string} 'ADD', 'REMOVE', 'ADD+REMOVE', vagy null
 */
function getClosestAction(item, now) {
  const { rawRow, headers } = item;
  const tz = ACCOUNT_TIMEZONE;

  const idxAddDate = headers.indexOf('Add Date');
  const idxAddHour = headers.indexOf('Add Hour');
  const idxRemoveDate = headers.indexOf('Remove Date');
  const idxRemoveHour = headers.indexOf('Remove Hour');

  // Aktuális időpont timezone-aware formátumban (string összehasonlításhoz)
  const nowStr = Utilities.formatDate(now, tz, 'yyyy-MM-dd HH:mm:ss');

  let addEndStr = null;
  let removeEndStr = null;

  // ADD művelet óra vége (KRITIKUS FIX: tisztán string alapú, elkerülve Date timezone problémákat!)
  if (rawRow[idxAddDate]) {
    const addDate = new Date(rawRow[idxAddDate]);
    const addDateStr = Utilities.formatDate(addDate, tz, 'yyyy-MM-dd');
    const addHourRaw = idxAddHour !== -1 ? String(rawRow[idxAddHour] ?? '').trim() : '';
    const addHourParsed = parseCustomHourToInt(addHourRaw);
    const addHour = (addHourParsed.ok && addHourParsed.value !== null) ? addHourParsed.value : 0;

    // Óra vég: addHour + 1 (pl. hour 23 → 24:00 = next day 00:00)
    const endHour = addHour + 1;

    if (endHour === 24) {
      // Következő nap 00:00 (string művelettel!)
      const nextDay = new Date(addDate);
      nextDay.setDate(nextDay.getDate() + 1);
      const nextDayStr = Utilities.formatDate(nextDay, tz, 'yyyy-MM-dd');
      addEndStr = `${nextDayStr} 00:00:00`;
    } else {
      // Ugyanazon nap
      addEndStr = `${addDateStr} ${pad2(endHour)}:00:00`;
    }
  }

  // REMOVE művelet óra vége (KRITIKUS FIX: ugyanaz a logika)
  if (rawRow[idxRemoveDate]) {
    const remDate = new Date(rawRow[idxRemoveDate]);
    const remDateStr = Utilities.formatDate(remDate, tz, 'yyyy-MM-dd');
    const remHourRaw = idxRemoveHour !== -1 ? String(rawRow[idxRemoveHour] ?? '').trim() : '';
    const remHourParsed = parseCustomHourToInt(remHourRaw);
    const remHour = (remHourParsed.ok && remHourParsed.value !== null) ? remHourParsed.value : 23;

    const endHour = remHour + 1;

    if (endHour === 24) {
      const nextDay = new Date(remDate);
      nextDay.setDate(nextDay.getDate() + 1);
      const nextDayStr = Utilities.formatDate(nextDay, tz, 'yyyy-MM-dd');
      removeEndStr = `${nextDayStr} 00:00:00`;
    } else {
      removeEndStr = `${remDateStr} ${pad2(endHour)}:00:00`;
    }
  }

  // Szűrjük ki a TELJESEN elmúlt műveleteket (óra vége után)
  // String összehasonlítás (ISO formátum lexikografikusan helyes!)
  if (addEndStr && addEndStr <= nowStr) addEndStr = null;
  if (removeEndStr && removeEndStr <= nowStr) removeEndStr = null;

  // Ha nincs jövőbeli művelet
  if (!addEndStr && !removeEndStr) return null;

  // Ha csak egyik van
  if (addEndStr && !removeEndStr) return 'ADD';
  if (!addEndStr && removeEndStr) return 'REMOVE';

  // Ha mindkettő van, melyik közelebb?
  if (addEndStr === removeEndStr) {
    return 'ADD+REMOVE'; // Ugyanabban az órában (ezt később ERROR-ként kezeljük)
  }

  return addEndStr < removeEndStr ? 'ADD' : 'REMOVE';
}

/**
 * Videó azonosító parse-olása (YouTube URL, Video ID, Asset ID)
 * @param {string} input - User input (URL, Video ID vagy Asset ID)
 * @returns {object} { type: 'ASSET_ID'|'VIDEO_ID'|'INVALID', value: string, error?: string }
 */
function parseVideoIdentifier(input) {
  const s = String(input || '').trim();

  if (!s) {
    return { type: 'INVALID', error: 'Hiányzik Video URL or ID / Asset ID' };
  }

  // 1. Asset ID (csak számok - mint image-eknél)
  if (/^\d+$/.test(s)) {
    return { type: 'ASSET_ID', value: s };
  }

  // 2. YouTube Video ID (11 karakter: betű, szám, - vagy _)
  if (/^[a-zA-Z0-9_-]{11}$/.test(s)) {
    return { type: 'VIDEO_ID', value: s };
  }

  // 3. YouTube URL parse
  const urlPatterns = [
    /(?:youtube\.com\/watch\?v=)([a-zA-Z0-9_-]{11})/,        // https://youtube.com/watch?v=...
    /(?:youtu\.be\/)([a-zA-Z0-9_-]{11})/,                    // https://youtu.be/...
    /(?:youtube\.com\/embed\/)([a-zA-Z0-9_-]{11})/           // https://youtube.com/embed/...
  ];

  for (const pattern of urlPatterns) {
    const match = s.match(pattern);
    if (match && match[1]) {
      return { type: 'VIDEO_ID', value: match[1] };
    }
  }

  return {
    type: 'INVALID',
    error: `Érvénytelen formátum: "${s.substring(0, 50)}...". Várt: YouTube URL, Video ID (11 karakter) vagy Asset ID (csak számok).`
  };
}

/**
 * ISO 8601 duration parse (PT10M30S → másodpercek)
 * @param {string} duration - ISO 8601 formátumú időtartam (pl. PT1H2M30S, PT10M, PT30S)
 * @returns {number} - Időtartam másodpercben
 */
function parseYouTubeDuration(duration) {
  if (!duration) return 0;

  const match = duration.match(/PT(?:(\d+)H)?(?:(\d+)M)?(?:(\d+)S)?/);
  if (!match) return 0;

  const hours = parseInt(match[1] || '0', 10);
  const minutes = parseInt(match[2] || '0', 10);
  const seconds = parseInt(match[3] || '0', 10);

  return hours * 3600 + minutes * 60 + seconds;
}

/**
 * YouTube videó validáció YouTube Data API v3 használatával
 * Ellenőrzi: létezés, publikálási állapot, beágyazhatóság, videó hossz
 *
 * @param {string} videoId - YouTube Video ID (11 karakter)
 * @returns {object} { valid: boolean, error?: string, warning?: string, details?: object }
 *
 * FONTOS: A Google Ads Script beállításokban be kell kapcsolni a YouTube API-t!
 * (Script → Advanced APIs → YouTube)
 */
function validateYouTubeVideo(videoId) {
  try {
    // Ellenőrizzük, hogy a YouTube service elérhető-e
    if (typeof YouTube === 'undefined') {
      return {
        valid: true,
        warning: 'YouTube API nincs engedélyezve - validáció kihagyva. Kapcsold be: Script → Advanced APIs → YouTube',
        skipped: true
      };
    }

    // YouTube.Videos.list() - beépített service használata
    const response = YouTube.Videos.list('status,contentDetails', {
      id: videoId
    });

    // Ha a videó nem található (üres items tömb)
    if (!response.items || response.items.length === 0) {
      return {
        valid: false,
        error: `YouTube videó nem található (ID: ${videoId}) - ellenőrizd a videó adatait`,
        notFound: true
      };
    }

    const video = response.items[0];
    const status = video.status || {};
    const contentDetails = video.contentDetails || {};

    const privacyStatus = status.privacyStatus; // public, unlisted, private
    const embeddable = status.embeddable;
    const duration = contentDetails.duration; // ISO 8601: PT10M30S
    const durationSeconds = parseYouTubeDuration(duration);

    const details = {
      privacyStatus,
      embeddable,
      durationSeconds,
      durationFormatted: duration
    };

    // Ellenőrzések
    const errors = [];

    // 1. Privát videó (unlisted = nem nyilvános, de OK hirdetésekhez)
    if (privacyStatus === 'private') {
      errors.push('Videó privát - nem használható hirdetésekben');
    }

    // 2. Videó hossz (PMax minimum 10 másodperc)
    if (durationSeconds > 0 && durationSeconds < 10) {
      errors.push(`Videó túl rövid: ${durationSeconds}s (minimum 10 másodperc szükséges PMax-hoz)`);
    }

    // Megjegyzés: embeddable ellenőrzés eltávolítva - nem befolyásolja a Google Ads hirdetéseket

    if (errors.length > 0) {
      return { valid: false, error: errors.join('; '), details };
    }

    return { valid: true, details };

  } catch (e) {
    // Ha bármilyen hiba történik, ne blokkoljuk a folyamatot
    const errorMsg = e.message || String(e);

    // Ha "YouTube is not defined" vagy hasonló
    if (errorMsg.includes('YouTube') || errorMsg.includes('not defined')) {
      return {
        valid: true,
        warning: 'YouTube API nincs engedélyezve - validáció kihagyva. Kapcsold be: Script → Advanced APIs → YouTube',
        skipped: true
      };
    }

    console.log(`⚠️ YouTube API validáció hiba: ${errorMsg}`);
    return {
      valid: true,
      warning: `YouTube API hiba: ${errorMsg} - validáció kihagyva`,
      skipped: true
    };
  }
}

/*** ===================== KAMPÁNYOK GYŰJTÉSE ===================== ***/

function gatherCampaigns(textRows, imageRows, videoRows) {
  const campaigns = new Set();

  textRows.forEach(item => {
    const { rawRow, headers } = item;
    const idxCampaign = headers.indexOf('Campaign Name');
    if (idxCampaign !== -1) {
      const campaign = String(rawRow[idxCampaign] || '').trim();
      if (campaign) campaigns.add(campaign);
    }
  });

  imageRows.forEach(item => {
    const { rawRow, headers } = item;
    const idxCampaign = headers.indexOf('Campaign Name');
    if (idxCampaign !== -1) {
      const campaign = String(rawRow[idxCampaign] || '').trim();
      if (campaign) campaigns.add(campaign);
    }
  });

  videoRows.forEach(item => {
    const { rawRow, headers } = item;
    const idxCampaign = headers.indexOf('Campaign Name');
    if (idxCampaign !== -1) {
      const campaign = String(rawRow[idxCampaign] || '').trim();
      if (campaign) campaigns.add(campaign);
    }
  });

  return Array.from(campaigns);
}

/*** ===================== BATCH GAQL LEKÉRDEZÉSEK ===================== ***/

// Helper: nagy IN-listák batchelése (200-as chunk limit)
function runBatchedQuery(items, buildQueryFn, context, chunkSize = 200) {
  if (!items || items.length === 0) return [];

  const allResults = [];
  const totalChunks = Math.ceil(items.length / chunkSize);

  for (let i = 0; i < items.length; i += chunkSize) {
    const chunk = items.slice(i, i + chunkSize);
    const chunkIndex = Math.floor(i / chunkSize) + 1;

    if (totalChunks > 1) {
      console.log(`  📦 ${context} - batch ${chunkIndex}/${totalChunks} (${chunk.length} items)`);
    }

    const query = buildQueryFn(chunk);
    const rows = runReportSafe(query, `${context}:batch${chunkIndex}`);
    allResults.push(...rows);
  }

  return allResults;
}

function runReportSafe(gaql, context) {
  let attempt = 0;
  const maxRetries = RETRY_MAX_ATTEMPTS;

  while (attempt <= maxRetries) {
    try {
      const report = AdsApp.report(gaql, { apiVersion: API_VERSION });
      const it = report.rows();
      const rows = [];
      while (it.hasNext()) rows.push(it.next());
      return rows;
    } catch (e) {
      const msg = String(e.message || e);
      const isTransient = /RESOURCE_EXHAUSTED|INTERNAL|BACKEND_ERROR|DEADLINE_EXCEEDED|temporarily|rate limit/i.test(msg);

      if (isTransient && attempt < maxRetries) {
        const baseWait = Math.pow(2, attempt) * RETRY_BASE_DELAY_MS;
        const jitter = baseWait * (0.75 + Math.random() * 0.5);
        const waitMs = Math.floor(jitter);
        console.log(`  ⚠️ Átmeneti GAQL hiba (${context}): ${msg} | Retry #${attempt + 1} ~${waitMs}ms múlva`);
        Utilities.sleep(waitMs);
        attempt++;
        continue;
      }

      console.log(`  ❌ GAQL hiba (${context}): ${msg}`);
      throw e;
    }
  }
}

function getCampaignIdsByNamesBatch(campaigns) {
  const out = {};
  if (!campaigns || campaigns.length === 0) return out;

  const inList = campaigns.map(name => `'${gaqlEscapeSingleQuote(name)}'`).join(',');
  const q = `
    SELECT campaign.id, campaign.name
    FROM campaign
    WHERE campaign.name IN (${inList})
  `;

  const rows = runReportSafe(q, 'getCampaignIdsByNamesBatch');
  rows.forEach(r => {
    const name = String(r['campaign.name']);
    const id = String(r['campaign.id']);
    out[name] = id;
  });

  console.log(`  📊 Kampány ID-k: ${Object.keys(out).length}/${campaigns.length}`);
  return out;
}

function fetchAssetGroupStatesBatch(campaignIds) {
  const textStates = {};
  const imageStates = {};
  const videoStates = {};

  if (!campaignIds || campaignIds.length === 0) {
    return { text: textStates, images: imageStates, videos: videoStates };
  }

  const customerId = AdsApp.currentAccount().getCustomerId().replace(/-/g, '');

  // 1. Enabled asset group-ok lekérése
  const campaignResourceNames = campaignIds.map(id => `customers/${customerId}/campaigns/${id}`);
  const inListCampaigns = campaignResourceNames.map(rn => `'${rn}'`).join(',');

  const groupQuery = `
    SELECT asset_group.resource_name, asset_group.name,
           campaign.name, campaign.resource_name, campaign.id
    FROM asset_group
    WHERE asset_group.campaign IN (${inListCampaigns})
      AND asset_group.status = ENABLED
  `;

  const groups = runReportSafe(groupQuery, 'fetchAssetGroupStatesBatch:groups');
  const groupResourceNames = groups.map(r => r['asset_group.resource_name']);

  if (groupResourceNames.length === 0) {
    console.log(`  ℹ️ Nincs ENABLED asset group ezekben a kampányokban.`);
    return { text: textStates, images: imageStates, videos: videoStates };
  }

  // 2. HEADLINE presence check (feed-only szűrés) - BATCHELT 200-as csomag limit
  const headlinePresence = {};
  groupResourceNames.forEach(rn => headlinePresence[rn] = false);

  const headlineRows = runBatchedQuery(
    groupResourceNames,
    (chunk) => {
      const inList = chunk.map(rn => `'${rn}'`).join(',');
      return `
        SELECT asset_group_asset.asset_group
        FROM asset_group_asset
        WHERE asset_group_asset.asset_group IN (${inList})
          AND asset_group_asset.field_type = HEADLINE
          AND asset_group_asset.status = ENABLED
          AND asset_group_asset.source = ADVERTISER
      `;
    },
    'HEADLINE presence check'
  );

  headlineRows.forEach(r => {
    headlinePresence[r['asset_group_asset.asset_group']] = true;
  });

  const nonFeedOnlyGroups = groupResourceNames.filter(rn => headlinePresence[rn]);

  console.log(`  📊 Batch fetch: ${groups.length} ENABLED group, ${nonFeedOnlyGroups.length} nem feed-only (szűrés validációnál)`);

  if (nonFeedOnlyGroups.length === 0) {
    console.log(`  ℹ️ Minden asset group feed-only (nincs HEADLINE).`);
    return { text: textStates, images: imageStates, videos: videoStates };
  }

  // 3. TEXT asset linkek lekérése - BATCHELT 200-as csomag limit
  const textRows = runBatchedQuery(
    nonFeedOnlyGroups,
    (chunk) => {
      const inList = chunk.map(rn => `'${rn}'`).join(',');
      return `
        SELECT
          asset_group_asset.asset_group,
          asset_group_asset.resource_name,
          asset_group_asset.asset,
          asset_group_asset.field_type,
          asset_group_asset.status,
          asset.text_asset.text
        FROM asset_group_asset
        WHERE asset_group_asset.asset_group IN (${inList})
          AND asset_group_asset.field_type IN (HEADLINE, LONG_HEADLINE, DESCRIPTION)
          AND asset_group_asset.status = ENABLED
          AND asset_group_asset.source = ADVERTISER
      `;
    },
    'TEXT asset linkek'
  );

  textRows.forEach(r => {
    const groupRN = r['asset_group_asset.asset_group'];
    if (!textStates[groupRN]) textStates[groupRN] = [];

    textStates[groupRN].push({
      agaResource: r['asset_group_asset.resource_name'],
      assetResource: r['asset_group_asset.asset'],
      fieldType: r['asset_group_asset.field_type'],
      text: String(r['asset.text_asset.text'] || '')
    });
  });

  // 4. IMAGE asset linkek lekérése - BATCHELT 200-as csomag limit
  const imageRows = runBatchedQuery(
    nonFeedOnlyGroups,
    (chunk) => {
      const inList = chunk.map(rn => `'${rn}'`).join(',');
      return `
        SELECT
          asset_group_asset.asset_group,
          asset_group_asset.resource_name,
          asset_group_asset.asset,
          asset_group_asset.field_type,
          asset_group_asset.status
        FROM asset_group_asset
        WHERE asset_group_asset.asset_group IN (${inList})
          AND asset_group_asset.field_type IN (MARKETING_IMAGE, SQUARE_MARKETING_IMAGE, PORTRAIT_MARKETING_IMAGE, TALL_PORTRAIT_MARKETING_IMAGE)
          AND asset_group_asset.status = ENABLED
          AND asset_group_asset.source = ADVERTISER
      `;
    },
    'IMAGE asset linkek'
  );

  imageRows.forEach(r => {
    const groupRN = r['asset_group_asset.asset_group'];
    if (!imageStates[groupRN]) imageStates[groupRN] = [];

    imageStates[groupRN].push({
      agaResource: r['asset_group_asset.resource_name'],
      assetResource: r['asset_group_asset.asset'],
      fieldType: r['asset_group_asset.field_type']
    });
  });

  // 5. VIDEO asset linkek lekérése - BATCHELT 200-as csomag limit
  const videoRows = runBatchedQuery(
    nonFeedOnlyGroups,
    (chunk) => {
      const inList = chunk.map(rn => `'${rn}'`).join(',');
      return `
        SELECT
          asset_group_asset.asset_group,
          asset_group_asset.resource_name,
          asset_group_asset.asset,
          asset_group_asset.field_type,
          asset_group_asset.status
        FROM asset_group_asset
        WHERE asset_group_asset.asset_group IN (${inList})
          AND asset_group_asset.field_type = YOUTUBE_VIDEO
          AND asset_group_asset.status = ENABLED
          AND asset_group_asset.source = ADVERTISER
      `;
    },
    'VIDEO asset linkek'
  );

  videoRows.forEach(r => {
    const groupRN = r['asset_group_asset.asset_group'];
    if (!videoStates[groupRN]) videoStates[groupRN] = [];

    videoStates[groupRN].push({
      agaResource: r['asset_group_asset.resource_name'],
      assetResource: r['asset_group_asset.asset'],
      fieldType: r['asset_group_asset.field_type']
    });
  });

  // 6. Asset group mapping (resource name -> név, campaign)
  const groupMap = {};
  groups.forEach(r => {
    const rn = r['asset_group.resource_name'];
    groupMap[rn] = {
      name: r['asset_group.name'],
      campaignName: r['campaign.name'],
      campaignResourceName: r['campaign.resource_name'],
      campaignId: String(r['campaign.id']),
      hasHeadline: headlinePresence[rn]
    };
  });

  console.log(`  📊 Batch fetch kész: ${Object.keys(textStates).length} group TEXT, ${Object.keys(imageStates).length} group IMAGE, ${Object.keys(videoStates).length} group VIDEO assets`);

  return { text: textStates, images: imageStates, videos: videoStates, groupMap };
}

function gaqlEscapeSingleQuote(s) {
  return String(s)
    .replace(/\\/g, '\\\\')   // Backslash escape (ELŐSZÖR!)
    .replace(/'/g, "\\'")      // Single quote escape
    .replace(/\n/g, '\\n')     // Newline escape
    .replace(/\r/g, '\\r')     // Carriage return escape
    .replace(/\t/g, '\\t');    // Tab escape
}

function normalizeText(s) {
  return String(s || '').trim().toLowerCase();
}

/*** ===================== VALIDÁCIÓ ===================== ***/

function validateAll(textRows, imageRows, videoRows, campaignIdMap, groupStates, textWindowRowNums, imageWindowRowNums, videoWindowRowNums) {
  const tz = ACCOUNT_TIMEZONE;
  const now = new Date();
  const timestamp = Utilities.formatDate(now, tz, 'yyyy-MM-dd HH:mm:ss');

  const validationRows = [];  // Minden sor (backward compatibility)
  const futureRows = [];      // Csak jövőbeli sorok (időablakon kívül)
  const windowRows = [];      // Csak időablakos sorok (bármi státusz: OK/WARNING/ERROR)
  const validTextItems = [];   // OK és WARNING text items
  const validImageItems = [];  // OK és WARNING image items
  const validVideoItems = [];  // OK és WARNING video items
  let okCount = 0, warningCount = 0, errorCount = 0;

  // Text assets validáció
  textRows.forEach(item => {
    const result = validateTextAsset(item, campaignIdMap, groupStates, timestamp);

    // FONTOS: result.rows egy array! (group-onként külön sorok)
    const isInWindow = textWindowRowNums && textWindowRowNums.has(item.sheetRowNum);

    result.rows.forEach(row => {
      validationRows.push(row);
      if (isInWindow) {
        windowRows.push(row);
      } else {
        futureRows.push(row);
      }
    });

    if (result.status === 'OK') {
      okCount++;
    } else if (result.status === 'WARNING') {
      warningCount++;
    } else if (result.status === 'ERROR') {
      errorCount++;
    }

    // KRITIKUS: Csak azokat a group-okat hajtjuk végre amik OK vagy WARNING!
    if (result.validGroups && result.validGroups.length > 0) {
      validTextItems.push({ item, validGroups: result.validGroups });
    }
  });

  // Image assets validáció
  imageRows.forEach(item => {
    const result = validateImageAsset(item, campaignIdMap, groupStates, timestamp);

    // FONTOS: result.rows egy array! (group-onként külön sorok)
    const isInWindow = imageWindowRowNums && imageWindowRowNums.has(item.sheetRowNum);

    result.rows.forEach(row => {
      validationRows.push(row);
      if (isInWindow) {
        windowRows.push(row);
      } else {
        futureRows.push(row);
      }
    });

    if (result.status === 'OK') {
      okCount++;
    } else if (result.status === 'WARNING') {
      warningCount++;
    } else if (result.status === 'ERROR') {
      errorCount++;
    }

    // KRITIKUS: Csak azokat a group-okat hajtjuk végre amik OK vagy WARNING!
    if (result.validGroups && result.validGroups.length > 0) {
      validImageItems.push({ item, validGroups: result.validGroups });
    }
  });

  // Video assets validáció
  videoRows.forEach(item => {
    const result = validateVideoAsset(item, campaignIdMap, groupStates, timestamp);

    // FONTOS: result.rows egy array! (group-onként külön sorok)
    const isInWindow = videoWindowRowNums && videoWindowRowNums.has(item.sheetRowNum);

    result.rows.forEach(row => {
      validationRows.push(row);
      if (isInWindow) {
        windowRows.push(row);
      } else {
        futureRows.push(row);
      }
    });

    if (result.status === 'OK') {
      okCount++;
    } else if (result.status === 'WARNING') {
      warningCount++;
    } else if (result.status === 'ERROR') {
      errorCount++;
    }

    // KRITIKUS: Csak azokat a group-okat hajtjuk végre amik OK vagy WARNING!
    if (result.validGroups && result.validGroups.length > 0) {
      validVideoItems.push({ item, validGroups: result.validGroups, parsedIdentifier: result.parsedIdentifier });
    }
  });

  return {
    ok: okCount,
    warnings: warningCount,
    errors: errorCount,
    rows: validationRows,  // Backward compatibility
    futureRows,            // Jövőbeli sorok (validation)
    windowRows,            // Időablakos sorok (execution preview)
    validTextItems,    // Csak OK és WARNING items
    validImageItems,   // Csak OK és WARNING items
    validVideoItems    // Csak OK és WARNING items
  };
}

/**
 * Cross-row ADD+REMOVE konfliktus detektálás
 *
 * Ha ugyanaz az asset (campaign, group, type, text/assetId) ugyanarra az órára:
 * - ADD és REMOVE is be van ütemezve (különböző sorokban)
 * → Mindkét sor ERROR státuszra válik
 *
 * Row struktúra: [timestamp, campaign, groupName, assetType, textOrId, action, scheduled, status, message]
 */
function detectCrossRowAddRemoveConflicts(allRows, futureRows, windowRows) {
  // Kulcs: campaign|groupName|assetType|textOrId|hour
  const operationsByKey = new Map();

  // Gyűjtjük össze az összes műveletet (csak OK és WARNING sorokból)
  allRows.forEach((row, idx) => {
    const [timestamp, campaign, groupName, assetType, textOrId, allScheduled, action, status, message] = row;

    // Csak OK és WARNING soroknál ellenőrzünk
    if (status !== 'OK' && status !== 'WARNING') return;

    // Óra kinyerése a allScheduled mezőből
    // Formátumok:
    // - "2025-11-16 10:00-10:59" (execution mode - egy művelet)
    // - "ADD: 2025-11-16 00:00-00:59 | REMOVE: 2025-11-16 22:00-22:59" (preview mode - több művelet)

    const hours = extractHoursFromScheduled(allScheduled, action);

    hours.forEach(hour => {
      const key = `${campaign}|${groupName}|${assetType}|${textOrId}|${hour}`;

      if (!operationsByKey.has(key)) {
        operationsByKey.set(key, { adds: [], removes: [] });
      }

      const ops = operationsByKey.get(key);

      if (action === 'ADD' || action === 'ADD+REMOVE') {
        ops.adds.push({ rowIndex: idx, row });
      }
      if (action === 'REMOVE' || action === 'ADD+REMOVE') {
        ops.removes.push({ rowIndex: idx, row });
      }
    });
  });

  // Konfliktusok detektálása
  const conflictedRowIndices = new Set();
  let conflictCount = 0;

  operationsByKey.forEach((ops, key) => {
    const [campaign, groupName, assetType, textOrId, hour] = key.split('|');

    // Ha van ADD és REMOVE is ugyanarra az órára (különböző sorokban!)
    if (ops.adds.length > 0 && ops.removes.length > 0) {
      // Ellenőrizzük hogy NEM ugyanaz a sor (ADD+REMOVE egy sorban már ERROR a validációnál)
      const addRows = new Set(ops.adds.map(x => x.rowIndex));
      const removeRows = new Set(ops.removes.map(x => x.rowIndex));

      // Van-e KÜLÖNBÖZŐ sorok között konfliktus?
      const hasConflict = [...addRows].some(addIdx => !removeRows.has(addIdx)) ||
                          [...removeRows].some(remIdx => !addRows.has(remIdx));

      if (hasConflict) {
        conflictCount++;
        const allAffectedRows = [...ops.adds, ...ops.removes];
        const rowNumbers = allAffectedRows.map(x => x.row[0]); // timestamp helyett inkább index?

        console.log(`  ⚠️ ADD+REMOVE konfliktus: ${assetType}="${textOrId}", group=${groupName}, hour=${hour} → ${allAffectedRows.length} sor ERROR-rá alakul`);

        allAffectedRows.forEach(({ rowIndex }) => {
          conflictedRowIndices.add(rowIndex);
        });
      }
    }
  });

  // Ha nincs konfliktus, return eredeti sorokkal
  if (conflictedRowIndices.size === 0) {
    return {
      allRows,
      futureRows,
      windowRows,
      conflictCount: 0
    };
  }

  // Módosítjuk a konfliktusban lévő sorok státuszát ERROR-ra
  const updatedAllRows = allRows.map((row, idx) => {
    if (conflictedRowIndices.has(idx)) {
      const [timestamp, campaign, groupName, assetType, textOrId, allScheduled, action, status, message] = row;
      return [timestamp, campaign, groupName, assetType, textOrId, allScheduled, action, 'ERROR', 'ADD és REMOVE ugyanarra az órára van ütemezve (sorok között)'];
    }
    return row;
  });

  const updatedFutureRows = futureRows.map(row => {
    const idx = allRows.indexOf(row);
    if (idx !== -1 && conflictedRowIndices.has(idx)) {
      const [timestamp, campaign, groupName, assetType, textOrId, allScheduled, action, status, message] = row;
      return [timestamp, campaign, groupName, assetType, textOrId, allScheduled, action, 'ERROR', 'ADD és REMOVE ugyanarra az órára van ütemezve (sorok között)'];
    }
    return row;
  });

  const updatedWindowRows = windowRows.map(row => {
    const idx = allRows.indexOf(row);
    if (idx !== -1 && conflictedRowIndices.has(idx)) {
      const [timestamp, campaign, groupName, assetType, textOrId, allScheduled, action, status, message] = row;
      return [timestamp, campaign, groupName, assetType, textOrId, allScheduled, action, 'ERROR', 'ADD és REMOVE ugyanarra az órára van ütemezve (sorok között)'];
    }
    return row;
  });

  return {
    allRows: updatedAllRows,
    futureRows: updatedFutureRows,
    windowRows: updatedWindowRows,
    conflictCount
  };
}

/**
 * Óra(k) kinyerése a scheduled mezőből
 * Formátumok:
 * - "2025-11-16 10:00-10:59" → [10]
 * - "ADD: 2025-11-16 00:00-00:59 | REMOVE: 2025-11-16 22:00-22:59" → [0, 22] (ha action=ADD+REMOVE)
 */
function extractHoursFromScheduled(scheduled, action) {
  const hours = [];

  // Regex: YYYY-MM-DD HH:MM-HH:MM
  const regex = /(\d{4}-\d{2}-\d{2})\s+(\d{1,2}):(\d{2})-(\d{1,2}):(\d{2})/g;
  let match;

  while ((match = regex.exec(scheduled)) !== null) {
    const hour = parseInt(match[2], 10);
    hours.push(hour);
  }

  // Ha action specifikus (ADD vagy REMOVE), csak az adott action-höz tartozó órát adjuk vissza
  if (action === 'ADD' && scheduled.includes('ADD:')) {
    // Csak ADD: után következő órát
    const addMatch = /ADD:\s*(\d{4}-\d{2}-\d{2})\s+(\d{1,2}):(\d{2})-(\d{1,2}):(\d{2})/.exec(scheduled);
    if (addMatch) {
      return [parseInt(addMatch[2], 10)];
    }
  }

  if (action === 'REMOVE' && scheduled.includes('REMOVE:')) {
    // Csak REMOVE: után következő órát
    const removeMatch = /REMOVE:\s*(\d{4}-\d{2}-\d{2})\s+(\d{1,2}):(\d{2})-(\d{1,2}):(\d{2})/.exec(scheduled);
    if (removeMatch) {
      return [parseInt(removeMatch[2], 10)];
    }
  }

  return hours;
}

/**
 * Státuszok számolása (OK, WARNING, ERROR)
 */
function countStatuses(rows) {
  let ok = 0, warnings = 0, errors = 0;

  rows.forEach(row => {
    const status = row[7]; // Status a 8. pozíción (0-indexed: 7)
    if (status === 'OK') ok++;
    else if (status === 'WARNING') warnings++;
    else if (status === 'ERROR') errors++;
  });

  return { ok, warnings, errors };
}

function validateTextAsset(item, campaignIdMap, groupStates, timestamp) {
  // FIX v7.4.6: strictFutureAction támogatás (filterStrictFutureRows által beállítva)
  const { rawRow, headers, sheetRowNum, action, strictFutureAction, customHour, hasAddInRange, hasRemoveInRange } = item;

  const idxCampaign = headers.indexOf('Campaign Name');
  const idxGroup = headers.indexOf('Asset Group Name');
  const idxTextType = headers.indexOf('Text Type');
  const idxText = headers.indexOf('Text');
  const idxAddDate = headers.indexOf('Add Date');
  const idxAddHour = headers.indexOf('Add Hour');
  const idxRemoveDate = headers.indexOf('Remove Date');
  const idxRemoveHour = headers.indexOf('Remove Hour');

  const campaign = String(rawRow[idxCampaign] || '').trim();
  const assetGroup = String(rawRow[idxGroup] || '').trim();
  const textType = String(rawRow[idxTextType] || '').trim().toUpperCase();
  const text = String(rawRow[idxText] || '').trim();

  const errors = [];
  const warnings = [];

  // Action és Scheduled meghatározás
  // FIX v7.4.6: strictFutureAction-t is figyelembe vesszük
  let effectiveAction = action || strictFutureAction;
  let scheduled = '';
  let allScheduledActions = ''; // Minden jövőbeli művelet (preview táblázathoz)

  if (!effectiveAction) {
    // Preview mode: csak a legközelebbi jövőbeli műveletet validáljuk
    const now = new Date();
    const closestAction = getClosestAction(item, now);

    // Biztonsági ellenőrzés: ha nincs jövőbeli művelet → ERROR
    // (Ez nem kellene előforduljon, mert filterByDateRange() már kiszűrte)
    if (!closestAction) {
      console.log(`    ⚠️ LOGIKAI HIBA: Nincs meghatározható jövőbeli művelet Row ${sheetRowNum} (${campaign} / ${assetGroup})`);
      const addDate = rawRow[idxAddDate] ? String(rawRow[idxAddDate]).trim() : '';
      const removeDate = rawRow[idxRemoveDate] ? String(rawRow[idxRemoveDate]).trim() : '';
      const parts = [];
      if (addDate) parts.push(`ADD: ${addDate}`);
      if (removeDate) parts.push(`REMOVE: ${removeDate}`);
      const scheduledCol = parts.join(' | ') || 'N/A';
      return {
        status: 'ERROR',
        rows: [[timestamp, campaign, assetGroup || '(all groups)', textType, text, scheduledCol, 'N/A', 'ERROR', 'Nem meghatározható a következő művelet - ellenőrizd a dátumokat']]
      };
    }

    effectiveAction = closestAction;
  }

  // KRITIKUS v7.3.29: Preview mode-ban validáljuk az órák érvényességét MIELŐTT használnánk őket
  if (!action) {
    if (rawRow[idxAddDate]) {
      const addHourRaw = idxAddHour !== -1 ? String(rawRow[idxAddHour] ?? '').trim() : '';
      if (addHourRaw) {
        const addHourParsed = parseCustomHourToInt(addHourRaw);
        if (!addHourParsed.ok) {
          errors.push(`Érvénytelen Add Hour: ${addHourParsed.error}`);
        }
      }
    }
    if (rawRow[idxRemoveDate]) {
      const remHourRaw = idxRemoveHour !== -1 ? String(rawRow[idxRemoveHour] ?? '').trim() : '';
      if (remHourRaw) {
        const remHourParsed = parseCustomHourToInt(remHourRaw);
        if (!remHourParsed.ok) {
          errors.push(`Érvénytelen Remove Hour: ${remHourParsed.error}`);
        }
      }
    }
  }

  // All Scheduled Actions formázás (preview táblázathoz - minden jövőbeli művelet)
  // Formátum: időintervallum "10:00-10:59" (nem csak "10:00")
  if (!action && rawRow[idxAddDate]) {
    const dateStr = Utilities.formatDate(new Date(rawRow[idxAddDate]), ACCOUNT_TIMEZONE, 'yyyy-MM-dd');
    const addHourRaw = idxAddHour !== -1 ? String(rawRow[idxAddHour] ?? '').trim() : '';
    const addHourParsed = parseCustomHourToInt(addHourRaw);
    const hour = (addHourParsed.ok && addHourParsed.value !== null) ? addHourParsed.value : 0;
    const hourRangeStr = `${pad2(hour)}:00-${pad2(hour)}:59`;
    allScheduledActions = `ADD: ${dateStr} ${hourRangeStr}`;
  }

  if (!action && rawRow[idxRemoveDate]) {
    const dateStr = Utilities.formatDate(new Date(rawRow[idxRemoveDate]), ACCOUNT_TIMEZONE, 'yyyy-MM-dd');
    const remHourRaw = idxRemoveHour !== -1 ? String(rawRow[idxRemoveHour] ?? '').trim() : '';
    const remHourParsed = parseCustomHourToInt(remHourRaw);
    const hour = (remHourParsed.ok && remHourParsed.value !== null) ? remHourParsed.value : 23;
    const hourRangeStr = `${pad2(hour)}:00-${pad2(hour)}:59`;
    const separator = allScheduledActions ? ' | ' : '';
    allScheduledActions += `${separator}REMOVE: ${dateStr} ${hourRangeStr}`;
  }

  // Scheduled időpont formázás (csak az aktuális művelet - execution vagy legközelebbi preview-nál)
  // Formátum: "2025-11-15 10:00-10:59" (időtartomány)
  if (effectiveAction === 'ADD' || effectiveAction === 'ADD+REMOVE') {
    if (rawRow[idxAddDate]) {
      const dateStr = Utilities.formatDate(new Date(rawRow[idxAddDate]), ACCOUNT_TIMEZONE, 'yyyy-MM-dd');
      const addHourRaw = idxAddHour !== -1 ? String(rawRow[idxAddHour] ?? '').trim() : '';
      const addHourParsed = parseCustomHourToInt(addHourRaw);
      const hour = (addHourParsed.ok && addHourParsed.value !== null) ? addHourParsed.value : 0;
      const hourRangeStr = `${pad2(hour)}:00-${pad2(hour)}:59`;
      scheduled = `${dateStr} ${hourRangeStr}`;
    }
  }

  if (effectiveAction === 'REMOVE' || effectiveAction === 'ADD+REMOVE') {
    if (rawRow[idxRemoveDate]) {
      const dateStr = Utilities.formatDate(new Date(rawRow[idxRemoveDate]), ACCOUNT_TIMEZONE, 'yyyy-MM-dd');
      const remHourRaw = idxRemoveHour !== -1 ? String(rawRow[idxRemoveHour] ?? '').trim() : '';
      const remHourParsed = parseCustomHourToInt(remHourRaw);
      const hour = (remHourParsed.ok && remHourParsed.value !== null) ? remHourParsed.value : 23;
      const hourRangeStr = `${pad2(hour)}:00-${pad2(hour)}:59`;
      const separator = scheduled ? ', ' : '';
      scheduled += `${separator}${dateStr} ${hourRangeStr}`;
    }
  }

  // 0. ADD+REMOVE ugyanabban az órában ellenőrzés
  if (rawRow[idxAddDate] && rawRow[idxRemoveDate]) {
    const addDateStr = Utilities.formatDate(new Date(rawRow[idxAddDate]), ACCOUNT_TIMEZONE, 'yyyy-MM-dd');
    const removeDateStr = Utilities.formatDate(new Date(rawRow[idxRemoveDate]), ACCOUNT_TIMEZONE, 'yyyy-MM-dd');

    if (addDateStr === removeDateStr) {
      const addHourRaw = idxAddHour !== -1 ? String(rawRow[idxAddHour] ?? '').trim() : '';
      const removeHourRaw = idxRemoveHour !== -1 ? String(rawRow[idxRemoveHour] ?? '').trim() : '';
      const addHourParsed = parseCustomHourToInt(addHourRaw);
      const removeHourParsed = parseCustomHourToInt(removeHourRaw);

      const addHour = (addHourParsed.ok && addHourParsed.value !== null) ? addHourParsed.value : 0;
      const removeHour = (removeHourParsed.ok && removeHourParsed.value !== null) ? removeHourParsed.value : 23;

      if (addHour === removeHour) {
        errors.push('Hozzáadás és Törlés ugyanabban az órában - nem végrehajtható');
      }
    }
  }

  // 1. Bemeneti validáció
  if (!campaign) errors.push('Hiányzik Campaign Name');
  if (!textType || !['HEADLINE', 'LONG_HEADLINE', 'DESCRIPTION'].includes(textType)) {
    errors.push(`Érvénytelen Text Type: "${textType}"`);
  }
  if (!text) errors.push('Hiányzik Text');

  // 2. Karakterlimit és felkiáltójel (csak ADD műveletnél!)
  const limits = TEXT_LIMITS[textType];
  if (limits && text && (effectiveAction === 'ADD' || effectiveAction === 'ADD+REMOVE')) {
    if (text.length > limits.maxLen) {
      errors.push(`Túl hosszú (${text.length}/${limits.maxLen} char)`);
    }
    // Felkiáltójel ellenőrzés csak HEADLINE és LONG_HEADLINE-ra
    if (text.indexOf('!') >= 0) {
      if (textType === 'HEADLINE') {
        warnings.push('Felkiáltójel a címsorban (Nem ajánlott)');
      } else if (textType === 'LONG_HEADLINE') {
        warnings.push('Felkiáltójel a hosszú címsorban (Nem ajánlott)');
      }
      // DESCRIPTION-nél NEM ellenőrizzük a felkiáltójelet
    }
  }

  // 3. Kampány ellenőrzés (NEM early return! Aggregáljuk a hibákat)
  const campaignId = campaignIdMap[campaign];
  if (!campaignId) {
    errors.push('Kampány nem található');
  }

  // Ha kritikus hiba van (nincs alapadatok) → RETURN összes hibával
  if (errors.length > 0) {
    const scheduledCol = action ? scheduled : allScheduledActions;
    return {
      status: 'ERROR',
      rows: [[timestamp, campaign, assetGroup || '(all groups)', textType, text, scheduledCol, effectiveAction, 'ERROR', errors.join('; ')]]
    };
  }

  // 4. Target groups
  const targetGroups = getTargetGroups(campaign, assetGroup, campaignId, groupStates);

  // Debug: megmutatjuk hogy konkrétan hány group-ra szűrt
  const groupNameList = targetGroups.length <= 3
    ? targetGroups.map(g => g.name).join(', ')
    : `${targetGroups.slice(0, 3).map(g => g.name).join(', ')} (+${targetGroups.length - 3} további)`;
  console.log(`    🎯 Target groups: ${targetGroups.length} (${assetGroup || 'all'}): ${groupNameList}`);

  if (targetGroups.length === 0) {
    const scheduledCol = action ? scheduled : allScheduledActions;
    return {
      status: 'ERROR',
      rows: [[timestamp, campaign, assetGroup || '(all groups)', textType, text, scheduledCol, effectiveAction, 'ERROR', 'Nincs aktív (nem feed-only) asset group']]
    };
  }

  // 5. Limitek ellenőrzése (grouponként) - KÜLÖN SOR MINDEN GROUP-RA!
  const rows = [];
  const validGroups = [];  // OK és WARNING group-ok (végrehajtáshoz)
  let hasError = false;
  let hasWarning = false;

  targetGroups.forEach(g => {
    const groupErrors = [];
    const groupWarnings = [...warnings]; // Globális warnings másolása (pl. felkiáltójel)

    const currentAssets = groupStates.text[g.resourceName] || [];
    const ofType = currentAssets.filter(x => x.fieldType === textType);
    const currentCount = ofType.length;

    // Használjuk az effectiveAction-t a validációhoz
    if (effectiveAction === 'ADD' || effectiveAction === 'ADD+REMOVE') {
      const afterAdd = currentCount + 1;

      if (afterAdd > limits.max) {
        groupErrors.push(`MAX limit túllépés (${currentCount}+1 > ${limits.max})`);
      } else if (currentCount >= (limits.max - limits.warnThreshold)) {
        groupWarnings.push(`Limit közel (current=${currentCount}, after=${afterAdd}, max=${limits.max})`);
      }

      // Duplikáció ellenőrzés (case-sensitive, konzisztens az API-val)
      const exists = ofType.some(x => x.text === text);
      if (exists) {
        groupErrors.push('Text asset már létezik a csoportban');  // ERROR, nem WARNING - ne próbálja újra hozzáadni!
      }
    }

    if (effectiveAction === 'REMOVE' || effectiveAction === 'ADD+REMOVE') {
      const afterRemove = currentCount - 1;

      // Létezik? (case-sensitive, konzisztens az API-val)
      const exists = ofType.some(x => x.text === text);
      if (!exists && effectiveAction === 'REMOVE') {
        // Csak REMOVE esetén hiba ha nem létezik (ADD+REMOVE esetén még nem létezhet)
        groupErrors.push('Text asset nem található a csoportban');
      }

      // MIN limit figyelmeztetés
      if (exists && afterRemove < limits.min) {
        groupWarnings.push(`⚠️ MIN limit alatt! ${afterRemove} hirdető által hozzáadott ${textType} marad törlés után (min=${limits.min}, + esetleges Google által generált)`);
      }
    }

    // Per-group status és message
    const groupStatus = groupErrors.length > 0 ? 'ERROR' : groupWarnings.length > 0 ? 'WARNING' : 'OK';
    const groupMessage = groupErrors.length > 0 ? groupErrors.join('; ') : groupWarnings.length > 0 ? groupWarnings.join('; ') : 'OK';

    if (groupStatus === 'ERROR') hasError = true;
    if (groupStatus === 'WARNING') hasWarning = true;

    // Csak OK és WARNING group-ok végrehajthatók!
    if (groupStatus === 'OK' || groupStatus === 'WARNING') {
      validGroups.push(g);
    }

    // KÜLÖN SOR minden group-ra!
    // Preview mode: allScheduledActions, Execution mode: scheduled
    const scheduledCol = action ? scheduled : allScheduledActions;
    rows.push([timestamp, campaign, g.name, textType, text, scheduledCol, effectiveAction, groupStatus, groupMessage]);
  });

  // Overall status (a legrosszabb group alapján)
  const finalStatus = hasError ? 'ERROR' : hasWarning ? 'WARNING' : 'OK';

  return {
    status: finalStatus,
    rows: rows,  // Array of rows (group-onként külön)
    validGroups: validGroups  // Csak OK és WARNING group-ok
  };
}

function validateImageAsset(item, campaignIdMap, groupStates, timestamp) {
  // FIX v7.4.6: strictFutureAction támogatás (filterStrictFutureRows által beállítva)
  const { rawRow, headers, sheetRowNum, action, strictFutureAction, customHour, hasAddInRange, hasRemoveInRange } = item;

  const idxCampaign = headers.indexOf('Campaign Name');
  const idxGroup = headers.indexOf('Asset Group Name');
  const idxImageType = headers.indexOf('Image Type');
  const idxAssetId = headers.indexOf('Asset ID');
  const idxAddDate = headers.indexOf('Add Date');
  const idxAddHour = headers.indexOf('Add Hour');
  const idxRemoveDate = headers.indexOf('Remove Date');
  const idxRemoveHour = headers.indexOf('Remove Hour');

  const campaign = String(rawRow[idxCampaign] || '').trim();
  const assetGroup = String(rawRow[idxGroup] || '').trim();
  const imageType = String(rawRow[idxImageType] || '').trim();
  const assetId = String(rawRow[idxAssetId] || '').trim();

  // Asset Type megjelenítéshez (email, preview results)
  const displayAssetType = imageType ? `Image ${imageType}` : 'Image';

  const errors = [];
  const warnings = [];

  // Action és Scheduled meghatározás
  // FIX v7.4.6: strictFutureAction-t is figyelembe vesszük
  let effectiveAction = action || strictFutureAction;
  let scheduled = '';
  let allScheduledActions = ''; // Minden jövőbeli művelet (preview táblázathoz)

  if (!effectiveAction) {
    // Preview mode: csak a legközelebbi jövőbeli műveletet validáljuk
    const now = new Date();
    const closestAction = getClosestAction(item, now);

    // Biztonsági ellenőrzés: ha nincs jövőbeli művelet → ERROR
    // (Ez nem kellene előforduljon, mert filterByDateRange() már kiszűrte)
    if (!closestAction) {
      console.log(`    ⚠️ LOGIKAI HIBA: Nincs meghatározható jövőbeli művelet Row ${sheetRowNum} (${campaign} / ${assetGroup})`);
      const addDate = rawRow[idxAddDate] ? String(rawRow[idxAddDate]).trim() : '';
      const removeDate = rawRow[idxRemoveDate] ? String(rawRow[idxRemoveDate]).trim() : '';
      const parts = [];
      if (addDate) parts.push(`ADD: ${addDate}`);
      if (removeDate) parts.push(`REMOVE: ${removeDate}`);
      const scheduledCol = parts.join(' | ') || 'N/A';
      return {
        status: 'ERROR',
        rows: [[timestamp, campaign, assetGroup || '(all groups)', displayAssetType, assetId, scheduledCol, 'N/A', 'ERROR', 'Nem meghatározható a következő művelet - ellenőrizd a dátumokat']]
      };
    }

    effectiveAction = closestAction;
  }

  // KRITIKUS v7.3.29: Preview mode-ban validáljuk az órák érvényességét MIELŐTT használnánk őket
  if (!action) {
    if (rawRow[idxAddDate]) {
      const addHourRaw = idxAddHour !== -1 ? String(rawRow[idxAddHour] ?? '').trim() : '';
      if (addHourRaw) {
        const addHourParsed = parseCustomHourToInt(addHourRaw);
        if (!addHourParsed.ok) {
          errors.push(`Érvénytelen Add Hour: ${addHourParsed.error}`);
        }
      }
    }
    if (rawRow[idxRemoveDate]) {
      const remHourRaw = idxRemoveHour !== -1 ? String(rawRow[idxRemoveHour] ?? '').trim() : '';
      if (remHourRaw) {
        const remHourParsed = parseCustomHourToInt(remHourRaw);
        if (!remHourParsed.ok) {
          errors.push(`Érvénytelen Remove Hour: ${remHourParsed.error}`);
        }
      }
    }
  }

  // All Scheduled Actions formázás (preview táblázathoz - minden jövőbeli művelet)
  // Formátum: időintervallum "10:00-10:59" (nem csak "10:00")
  if (!action && rawRow[idxAddDate]) {
    const dateStr = Utilities.formatDate(new Date(rawRow[idxAddDate]), ACCOUNT_TIMEZONE, 'yyyy-MM-dd');
    const addHourRaw = idxAddHour !== -1 ? String(rawRow[idxAddHour] ?? '').trim() : '';
    const addHourParsed = parseCustomHourToInt(addHourRaw);
    const hour = (addHourParsed.ok && addHourParsed.value !== null) ? addHourParsed.value : 0;
    const hourRangeStr = `${pad2(hour)}:00-${pad2(hour)}:59`;
    allScheduledActions = `ADD: ${dateStr} ${hourRangeStr}`;
  }

  if (!action && rawRow[idxRemoveDate]) {
    const dateStr = Utilities.formatDate(new Date(rawRow[idxRemoveDate]), ACCOUNT_TIMEZONE, 'yyyy-MM-dd');
    const remHourRaw = idxRemoveHour !== -1 ? String(rawRow[idxRemoveHour] ?? '').trim() : '';
    const remHourParsed = parseCustomHourToInt(remHourRaw);
    const hour = (remHourParsed.ok && remHourParsed.value !== null) ? remHourParsed.value : 23;
    const hourRangeStr = `${pad2(hour)}:00-${pad2(hour)}:59`;
    const separator = allScheduledActions ? ' | ' : '';
    allScheduledActions += `${separator}REMOVE: ${dateStr} ${hourRangeStr}`;
  }

  // Scheduled időpont formázás (csak az aktuális művelet - execution vagy legközelebbi preview-nál)
  // Formátum: "2025-11-15 10:00-10:59" (időtartomány)
  if (effectiveAction === 'ADD' || effectiveAction === 'ADD+REMOVE') {
    if (rawRow[idxAddDate]) {
      const dateStr = Utilities.formatDate(new Date(rawRow[idxAddDate]), ACCOUNT_TIMEZONE, 'yyyy-MM-dd');
      const addHourRaw = idxAddHour !== -1 ? String(rawRow[idxAddHour] ?? '').trim() : '';
      const addHourParsed = parseCustomHourToInt(addHourRaw);
      const hour = (addHourParsed.ok && addHourParsed.value !== null) ? addHourParsed.value : 0;
      const hourRangeStr = `${pad2(hour)}:00-${pad2(hour)}:59`;
      scheduled = `${dateStr} ${hourRangeStr}`;
    }
  }

  if (effectiveAction === 'REMOVE' || effectiveAction === 'ADD+REMOVE') {
    if (rawRow[idxRemoveDate]) {
      const dateStr = Utilities.formatDate(new Date(rawRow[idxRemoveDate]), ACCOUNT_TIMEZONE, 'yyyy-MM-dd');
      const remHourRaw = idxRemoveHour !== -1 ? String(rawRow[idxRemoveHour] ?? '').trim() : '';
      const remHourParsed = parseCustomHourToInt(remHourRaw);
      const hour = (remHourParsed.ok && remHourParsed.value !== null) ? remHourParsed.value : 23;
      const hourRangeStr = `${pad2(hour)}:00-${pad2(hour)}:59`;
      const separator = scheduled ? ', ' : '';
      scheduled += `${separator}${dateStr} ${hourRangeStr}`;
    }
  }

  // 0. ADD+REMOVE ugyanabban az órában ellenőrzés
  if (rawRow[idxAddDate] && rawRow[idxRemoveDate]) {
    const addDateStr = Utilities.formatDate(new Date(rawRow[idxAddDate]), ACCOUNT_TIMEZONE, 'yyyy-MM-dd');
    const removeDateStr = Utilities.formatDate(new Date(rawRow[idxRemoveDate]), ACCOUNT_TIMEZONE, 'yyyy-MM-dd');

    if (addDateStr === removeDateStr) {
      const addHourRaw = idxAddHour !== -1 ? String(rawRow[idxAddHour] ?? '').trim() : '';
      const removeHourRaw = idxRemoveHour !== -1 ? String(rawRow[idxRemoveHour] ?? '').trim() : '';
      const addHourParsed = parseCustomHourToInt(addHourRaw);
      const removeHourParsed = parseCustomHourToInt(removeHourRaw);

      const addHour = (addHourParsed.ok && addHourParsed.value !== null) ? addHourParsed.value : 0;
      const removeHour = (removeHourParsed.ok && removeHourParsed.value !== null) ? removeHourParsed.value : 23;

      if (addHour === removeHour) {
        errors.push('Hozzáadás és Törlés ugyanabban az órában - nem végrehajtható');
      }
    }
  }

  // 1. Bemeneti validáció
  if (!campaign) errors.push('Hiányzik Campaign Name');
  if (!assetId) errors.push('Hiányzik Asset ID');

  // Asset ID formátum ellenőrzés (csak számok)
  if (assetId && !/^\d+$/.test(assetId)) {
    errors.push('Asset ID érvénytelen formátum (csak számok)');
  }

  const normalizedImageType = normalizeImageType(imageType);
  if (normalizedImageType === 'UNKNOWN') {
    errors.push(`Érvénytelen Image Type: "${imageType}"`);
  }

  // 2. Asset ID létezik? + Aspect ratio ellenőrzés (csak ha formátum valid!)
  // KRITIKUS v7.3.21: correctFieldType a tényleges kép méretek alapján (MINDEN típusnál!)
  let correctFieldType = normalizedImageType; // fallback
  let typeMismatch = false;
  let mismatchDetails = '';

  if (assetId && /^\d+$/.test(assetId)) {
    const assetDetails = fetchAssetDetails(assetId);
    if (!assetDetails) {
      errors.push('Asset ID nem található');
    } else if (assetDetails.type !== 'IMAGE') {
      errors.push(`Asset típusa ${assetDetails.type}, nem IMAGE`);
    } else {
      // Pontos fieldType meghatározása a TÉNYLEGES kép méretek alapján!
      const typeResult = determineImageFieldType(imageType, assetDetails.width, assetDetails.height);

      // KRITIKUS FIX #1 & #4: Ha aspect ratio NEM TÁMOGATOTT → ERROR!
      if (typeResult.error) {
        errors.push(typeResult.error);
      }

      correctFieldType = typeResult.fieldType;
      typeMismatch = typeResult.mismatch;

      // KRITIKUS FIX #4: Ha fieldType UNKNOWN → ERROR (ne próbálja végrehajtani!)
      if (correctFieldType === 'UNKNOWN' && !typeResult.error) {
        errors.push('Nem sikerült meghatározni Image Type-ot (ismeretlen aspect ratio vagy érvénytelen méretek)');
      }

      // KRITIKUS FIX #2: WARNING részletesebben a limitekkel kapcsolatban
      if (typeMismatch) {
        mismatchDetails = `⚠️ Image Type eltérés! Sheetben "${imageType}" (${typeResult.userExpected}) de kép ténylegesen ${typeResult.actualDetected}. Művelet és limitek a TÉNYLEGES típusra (${correctFieldType}) vonatkoznak! Ha ${typeResult.userExpected} limiteket szeretnél ellenőrizni, használj megfelelő aspect ratio-jú képet.`;
        warnings.push(mismatchDetails);
      }

      // Aspect ratio ellenőrzés (csak ADD műveleteknél és ha correctFieldType valid)
      if ((effectiveAction === 'ADD' || effectiveAction === 'ADD+REMOVE') && correctFieldType !== 'UNKNOWN') {
        const aspectCheck = validateAspectRatio(assetDetails.width, assetDetails.height, correctFieldType);
        if (!aspectCheck.valid) {
          errors.push(aspectCheck.message);
        }
      }
    }
  }

  // 3. Kampány (aggregáljuk a hibákat)
  const campaignId = campaignIdMap[campaign];
  if (!campaignId) {
    errors.push('Kampány nem található');
  }

  // Ha kritikus hiba van → RETURN összes hibával
  if (errors.length > 0) {
    const scheduledCol = action ? scheduled : allScheduledActions;
    return {
      status: 'ERROR',
      rows: [[timestamp, campaign, assetGroup || '(all groups)', displayAssetType, assetId, scheduledCol, effectiveAction, 'ERROR', errors.join('; ')]]
    };
  }

  // 4. Target groups
  const targetGroups = getTargetGroups(campaign, assetGroup, campaignId, groupStates);

  // Debug: megmutatjuk hogy konkrétan hány group-ra szűrt
  const groupNameList = targetGroups.length <= 3
    ? targetGroups.map(g => g.name).join(', ')
    : `${targetGroups.slice(0, 3).map(g => g.name).join(', ')} (+${targetGroups.length - 3} további)`;
  console.log(`    🎯 Target groups: ${targetGroups.length} (${assetGroup || 'all'}): ${groupNameList}`);

  if (targetGroups.length === 0) {
    const scheduledCol = action ? scheduled : allScheduledActions;
    return {
      status: 'ERROR',
      rows: [[timestamp, campaign, assetGroup || '(all groups)', displayAssetType, assetId, scheduledCol, effectiveAction, 'ERROR', 'Nincs aktív (nem feed-only) asset group']]
    };
  }

  // 5. Limitek - KÜLÖN SOR MINDEN GROUP-RA!
  const customerId = AdsApp.currentAccount().getCustomerId().replace(/-/g, '');
  const assetResourceName = `customers/${customerId}/assets/${assetId}`;

  const rows = [];
  const validGroups = [];  // OK és WARNING group-ok (végrehajtáshoz)
  let hasError = false;
  let hasWarning = false;

  targetGroups.forEach(g => {
    const groupErrors = [];
    const groupWarnings = [];

    // KRITIKUS v7.3.26: Globális warnings hozzáadása (pl. image type mismatch)
    if (warnings.length > 0) {
      groupWarnings.push(...warnings);
    }

    const currentImages = groupStates.images[g.resourceName] || [];
    const totalImages = currentImages.length;

    // Használjuk az effectiveAction-t a validációhoz
    if (effectiveAction === 'ADD' || effectiveAction === 'ADD+REMOVE') {
      const afterAdd = totalImages + 1;

      if (afterAdd > IMAGE_LIMITS.TOTAL.max) {
        groupErrors.push(`Képlimit túllépés (${totalImages}+1 > ${IMAGE_LIMITS.TOTAL.max})`);
      }

      const exists = currentImages.some(img => img.assetResource === assetResourceName);
      if (exists) {
        groupErrors.push('Image asset már létezik a csoportban');  // ERROR, nem WARNING - ne próbálja újra hozzáadni!
      }
      // Megjegyzés: Az aspect ratio ellenőrzés most már a validációban történik (validateAspectRatio)
    }

    if (effectiveAction === 'REMOVE' || effectiveAction === 'ADD+REMOVE') {
      const exists = currentImages.some(img => img.assetResource === assetResourceName);
      if (!exists && effectiveAction === 'REMOVE') {
        // Csak REMOVE esetén hiba ha nem létezik (ADD+REMOVE esetén még nem létezhet)
        groupErrors.push('Image asset nem található a csoportban');
      }

      // MIN limit figyelmeztetés SQUARE/HORIZONTAL esetén
      if (exists && correctFieldType === 'SQUARE_MARKETING_IMAGE') {
        const squares = currentImages.filter(img => img.fieldType === 'SQUARE_MARKETING_IMAGE');
        if (squares.length === 1 && squares.some(img => img.assetResource === assetResourceName)) {
          groupWarnings.push('⚠️ MIN limit alatt! 0 hirdető által hozzáadott SQUARE kép marad törlés után (min=1, + esetleges Google által generált)');
        }
      }
      if (exists && correctFieldType === 'MARKETING_IMAGE') {
        const horizontals = currentImages.filter(img => img.fieldType === 'MARKETING_IMAGE');
        if (horizontals.length === 1 && horizontals.some(img => img.assetResource === assetResourceName)) {
          groupWarnings.push('⚠️ MIN limit alatt! 0 hirdető által hozzáadott HORIZONTAL kép marad törlés után (min=1, + esetleges Google által generált)');
        }
      }
    }

    // Per-group status és message
    const groupStatus = groupErrors.length > 0 ? 'ERROR' : groupWarnings.length > 0 ? 'WARNING' : 'OK';
    const groupMessage = groupErrors.length > 0 ? groupErrors.join('; ') : groupWarnings.length > 0 ? groupWarnings.join('; ') : 'OK';

    if (groupStatus === 'ERROR') hasError = true;
    if (groupStatus === 'WARNING') hasWarning = true;

    // Csak OK és WARNING group-ok végrehajthatók!
    if (groupStatus === 'OK' || groupStatus === 'WARNING') {
      validGroups.push(g);
    }

    // KÜLÖN SOR minden group-ra!
    // Preview mode: allScheduledActions, Execution mode: scheduled
    const scheduledCol = action ? scheduled : allScheduledActions;
    rows.push([timestamp, campaign, g.name, displayAssetType, assetId, scheduledCol, effectiveAction, groupStatus, groupMessage]);
  });

  // Overall status (a legrosszabb group alapján)
  const finalStatus = hasError ? 'ERROR' : hasWarning ? 'WARNING' : 'OK';

  return {
    status: finalStatus,
    rows: rows,  // Array of rows (group-onként külön)
    validGroups: validGroups  // Csak OK és WARNING group-ok
  };
}

function getTargetGroups(campaign, assetGroupName, campaignId, groupStates) {
  const groups = [];

  // groupMap tartalmazza az összes groupot
  Object.keys(groupStates.groupMap || {}).forEach(groupRN => {
    const g = groupStates.groupMap[groupRN];

    // Kampányhoz tartozik? (KRITIKUS: ID alapján, nem resource name parse!)
    if (g.campaignId !== String(campaignId)) return;

    // Feed-only?
    if (!g.hasHeadline) return;

    // Név egyezés?
    if (assetGroupName) {
      if (normalizeText(g.name) === normalizeText(assetGroupName)) {
        groups.push({ resourceName: groupRN, name: g.name });
      }
    } else {
      // Minden aktív, nem feed-only group
      groups.push({ resourceName: groupRN, name: g.name });
    }
  });

  return groups;
}

function normalizeImageType(rawType) {
  const s = String(rawType || '').toUpperCase();

  if (s.includes('HORIZONTAL') || s.includes('19')) return 'MARKETING_IMAGE';
  if (s.includes('SQUARE') || s.includes('1:1')) return 'SQUARE_MARKETING_IMAGE';
  if (s.includes('9:16') || s.includes('TALL')) return 'TALL_PORTRAIT_MARKETING_IMAGE';
  if (s.includes('VERTICAL') || s.includes('4:5')) return 'PORTRAIT_MARKETING_IMAGE';

  return 'UNKNOWN';
}

/**
 * Meghatározza a pontos Google Ads API fieldType-ot a tényleges kép méretek alapján.
 * KRITIKUS FIX v7.3.21: MINDEN képtípusnál a tényleges méretek alapján dönt!
 *
 * MIÉRT? Limit check működéséhez kell:
 * - User írhat SQUARE-t de a kép lehet HORIZONTAL → rossz limit warning!
 * - REMOVE előtt tudnunk kell pontosan hány SQUARE/HORIZONTAL marad
 *
 * @returns {object} { fieldType: string, mismatch: boolean, userExpected: string, actualDetected: string }
 */
function determineImageFieldType(rawImageType, width, height) {
  const s = String(rawImageType || '').toUpperCase();
  const tolerance = 0.02; // ±2% tűrés

  // Aspect ratio követelmények
  const ratioHorizontal = 1.91;  // MARKETING_IMAGE
  const ratioSquare = 1.0;       // SQUARE_MARKETING_IMAGE
  const ratio45 = 0.8;           // PORTRAIT_MARKETING_IMAGE (4:5)
  const ratio916 = 0.5625;       // TALL_PORTRAIT_MARKETING_IMAGE (9:16)

  // User által VÁRT típus meghatározása a sheet alapján
  let userExpectedType = 'UNKNOWN';
  if (s.includes('HORIZONTAL') || s.includes('1.91') || s.includes('19')) {
    userExpectedType = 'MARKETING_IMAGE';
  } else if (s.includes('SQUARE') || s.includes('1:1')) {
    userExpectedType = 'SQUARE_MARKETING_IMAGE';
  } else if (s.includes('9:16')) {
    userExpectedType = 'TALL_PORTRAIT_MARKETING_IMAGE';
  } else if (s.includes('4:5') || s.includes('VERTICAL')) {
    userExpectedType = 'PORTRAIT_MARKETING_IMAGE';
  }

  // Ha nincs méret adat → visszaadjuk a user által megadottat (fallback)
  if (!width || !height || width === 0 || height === 0) {
    return {
      fieldType: userExpectedType !== 'UNKNOWN' ? userExpectedType : 'PORTRAIT_MARKETING_IMAGE',
      mismatch: false,
      userExpected: userExpectedType,
      actualDetected: null
    };
  }

  // TÉNYLEGES aspect ratio számítás
  const actualRatio = width / height;
  let actualFieldType = 'UNKNOWN';
  let actualLabel = '';

  // Melyik arányhoz van legközelebb?
  if (Math.abs(actualRatio - ratioHorizontal) <= tolerance) {
    actualFieldType = 'MARKETING_IMAGE';
    actualLabel = 'HORIZONTAL (1.91:1)';
  } else if (Math.abs(actualRatio - ratioSquare) <= tolerance) {
    actualFieldType = 'SQUARE_MARKETING_IMAGE';
    actualLabel = 'SQUARE (1:1)';
  } else if (Math.abs(actualRatio - ratio916) <= tolerance) {
    actualFieldType = 'TALL_PORTRAIT_MARKETING_IMAGE';
    actualLabel = 'VERTICAL (9:16)';
  } else if (Math.abs(actualRatio - ratio45) <= tolerance) {
    actualFieldType = 'PORTRAIT_MARKETING_IMAGE';
    actualLabel = 'VERTICAL (4:5)';
  } else {
    // Nem felismerhető arány → ERROR lesz a validateAspectRatio()-ban
    actualFieldType = 'UNKNOWN';
    actualLabel = `Ismeretlen (${actualRatio.toFixed(2)}:1)`;
  }

  // KRITIKUS FIX #1: Ha actualFieldType UNKNOWN és vannak méretek → NEM TÁMOGATOTT aspect ratio!
  if (actualFieldType === 'UNKNOWN' && width && height && width !== 0 && height !== 0) {
    return {
      fieldType: 'UNKNOWN',
      mismatch: false,
      userExpected: userExpectedType,
      actualDetected: actualLabel,
      error: `Nem támogatott aspect ratio: ${width}×${height} (${actualRatio.toFixed(2)}:1). Google PMax követelmények: 1.91:1 (HORIZONTAL), 1:1 (SQUARE), 4:5 vagy 9:16 (VERTICAL).`
    };
  }

  // Mismatch detektálás
  const mismatch = (userExpectedType !== 'UNKNOWN' && actualFieldType !== 'UNKNOWN' && userExpectedType !== actualFieldType);

  return {
    fieldType: actualFieldType !== 'UNKNOWN' ? actualFieldType : userExpectedType,
    mismatch: mismatch,
    userExpected: userExpectedType,
    actualDetected: actualLabel
  };
}

function validateAspectRatio(width, height, expectedImageType) {
  if (!width || !height || width === 0 || height === 0) {
    return { valid: false, message: 'Érvénytelen kép méretek' };
  }

  const actualRatio = width / height;
  const tolerance = 0.02; // ±2% tűrés

  // Aspect ratio követelmények (Google PMax)
  const requirements = {
    MARKETING_IMAGE: { ratio: 1.91, label: '1.91:1 (HORIZONTAL, pl. 1200×628)' },
    SQUARE_MARKETING_IMAGE: { ratio: 1.0, label: '1:1 (SQUARE, pl. 1200×1200)' },
    // PORTRAIT_MARKETING_IMAGE: 4:5 (0.8) VAGY 9:16 (0.5625)
  };

  // MARKETING_IMAGE (HORIZONTAL): 1.91:1
  if (expectedImageType === 'MARKETING_IMAGE') {
    const expected = requirements.MARKETING_IMAGE.ratio;
    if (Math.abs(actualRatio - expected) <= tolerance) {
      return { valid: true, message: `OK (${width}×${height} = ${actualRatio.toFixed(2)}:1)` };
    }
    return {
      valid: false,
      message: `Aspect ratio hiba: ${width}×${height} = ${actualRatio.toFixed(2)}:1, várt: ${requirements.MARKETING_IMAGE.label}`
    };
  }

  // SQUARE_MARKETING_IMAGE: 1:1
  if (expectedImageType === 'SQUARE_MARKETING_IMAGE') {
    const expected = requirements.SQUARE_MARKETING_IMAGE.ratio;
    if (Math.abs(actualRatio - expected) <= tolerance) {
      return { valid: true, message: `OK (${width}×${height} = ${actualRatio.toFixed(2)}:1)` };
    }
    return {
      valid: false,
      message: `Aspect ratio hiba: ${width}×${height} = ${actualRatio.toFixed(2)}:1, várt: ${requirements.SQUARE_MARKETING_IMAGE.label}`
    };
  }

  // PORTRAIT_MARKETING_IMAGE: 4:5 (0.8)
  if (expectedImageType === 'PORTRAIT_MARKETING_IMAGE') {
    const ratio45 = 0.8;    // 4:5

    if (Math.abs(actualRatio - ratio45) <= tolerance) {
      return { valid: true, message: `OK (${width}×${height} = 4:5)` };
    }
    return {
      valid: false,
      message: `Aspect ratio hiba: ${width}×${height} = ${actualRatio.toFixed(2)}:1, várt: 4:5 (0.8:1) VERTICAL`
    };
  }

  // TALL_PORTRAIT_MARKETING_IMAGE: 9:16 (0.5625)
  if (expectedImageType === 'TALL_PORTRAIT_MARKETING_IMAGE') {
    const ratio916 = 0.5625; // 9:16

    if (Math.abs(actualRatio - ratio916) <= tolerance) {
      return { valid: true, message: `OK (${width}×${height} = 9:16)` };
    }
    return {
      valid: false,
      message: `Aspect ratio hiba: ${width}×${height} = ${actualRatio.toFixed(2)}:1, várt: 9:16 (0.56:1) TALL VERTICAL`
    };
  }

  return { valid: true, message: 'Nem ellenőrzött típus' };
}

function fetchAssetType(assetId) {
  const details = fetchAssetDetails(assetId);
  return details ? details.type : null;
}

function fetchAssetDetails(assetId) {
  const q = `
    SELECT
      asset.id,
      asset.type,
      asset.image_asset.full_size.width_pixels,
      asset.image_asset.full_size.height_pixels,
      asset.image_asset.full_size.url,
      asset.youtube_video_asset.youtube_video_id
    FROM asset
    WHERE asset.id = ${assetId}
    LIMIT 1
  `;

  try {
    const rows = runReportSafe(q, 'fetchAssetDetails');
    if (rows.length > 0) {
      const row = rows[0];
      return {
        type: String(row['asset.type'] || ''),
        width: parseInt(row['asset.image_asset.full_size.width_pixels'] || 0, 10),
        height: parseInt(row['asset.image_asset.full_size.height_pixels'] || 0, 10),
        imageUrl: String(row['asset.image_asset.full_size.url'] || ''),
        youtubeVideoId: String(row['asset.youtube_video_asset.youtube_video_id'] || '')
      };
    }
  } catch (e) {
    console.log(`  ⚠️ Asset ID ${assetId} lekérdezési hiba: ${e.message}`);
  }

  return null;
}

function validateVideoAsset(item, campaignIdMap, groupStates, timestamp) {
  // FIX v7.4.6: strictFutureAction támogatás (filterStrictFutureRows által beállítva)
  const { rawRow, headers, sheetRowNum, action, strictFutureAction, customHour, hasAddInRange, hasRemoveInRange } = item;

  const idxCampaign = headers.indexOf('Campaign Name');
  const idxGroup = headers.indexOf('Asset Group Name');
  const idxVideoAspect = headers.indexOf('Video Aspect (optional)');
  const idxVideoId = headers.indexOf('Video URL or ID / Asset ID');
  const idxAddDate = headers.indexOf('Add Date');
  const idxAddHour = headers.indexOf('Add Hour');
  const idxRemoveDate = headers.indexOf('Remove Date');
  const idxRemoveHour = headers.indexOf('Remove Hour');

  const campaign = String(rawRow[idxCampaign] || '').trim();
  const assetGroup = String(rawRow[idxGroup] || '').trim();
  const videoAspect = String(rawRow[idxVideoAspect] || '').trim();
  const videoInput = String(rawRow[idxVideoId] || '').trim();

  // Asset Type megjelenítéshez (email, preview results)
  const displayAssetType = videoAspect ? `Video ${videoAspect}` : 'Video';

  const errors = [];
  const warnings = [];

  // Action és Scheduled meghatározás
  // FIX v7.4.6: strictFutureAction-t is figyelembe vesszük
  let effectiveAction = action || strictFutureAction;
  let scheduled = '';
  let allScheduledActions = '';

  if (!effectiveAction) {
    const now = new Date();
    const closestAction = getClosestAction(item, now);
    if (!closestAction) {
      console.log(`    ⚠️ LOGIKAI HIBA: Nincs meghatározható jövőbeli művelet Row ${sheetRowNum} (${campaign} / ${assetGroup})`);
      const addDate = rawRow[idxAddDate] ? String(rawRow[idxAddDate]).trim() : '';
      const removeDate = rawRow[idxRemoveDate] ? String(rawRow[idxRemoveDate]).trim() : '';
      const parts = [];
      if (addDate) parts.push(`ADD: ${addDate}`);
      if (removeDate) parts.push(`REMOVE: ${removeDate}`);
      const scheduledCol = parts.join(' | ') || 'N/A';
      return {
        status: 'ERROR',
        rows: [[timestamp, campaign, assetGroup || '(all groups)', displayAssetType, videoInput, scheduledCol, 'N/A', 'ERROR', 'Nem meghatározható a következő művelet - ellenőrizd a dátumokat']]
      };
    }
    effectiveAction = closestAction;
  }

  // Preview mode órák validálása
  if (!action) {
    if (rawRow[idxAddDate]) {
      const addHourRaw = idxAddHour !== -1 ? String(rawRow[idxAddHour] ?? '').trim() : '';
      if (addHourRaw) {
        const addHourParsed = parseCustomHourToInt(addHourRaw);
        if (!addHourParsed.ok) {
          errors.push(`Érvénytelen Add Hour: ${addHourParsed.error}`);
        }
      }
    }
    if (rawRow[idxRemoveDate]) {
      const remHourRaw = idxRemoveHour !== -1 ? String(rawRow[idxRemoveHour] ?? '').trim() : '';
      if (remHourRaw) {
        const remHourParsed = parseCustomHourToInt(remHourRaw);
        if (!remHourParsed.ok) {
          errors.push(`Érvénytelen Remove Hour: ${remHourParsed.error}`);
        }
      }
    }
  }

  // All Scheduled Actions formázás
  if (!action && rawRow[idxAddDate]) {
    const dateStr = Utilities.formatDate(new Date(rawRow[idxAddDate]), ACCOUNT_TIMEZONE, 'yyyy-MM-dd');
    const addHourRaw = idxAddHour !== -1 ? String(rawRow[idxAddHour] ?? '').trim() : '';
    const addHourParsed = parseCustomHourToInt(addHourRaw);
    const hour = (addHourParsed.ok && addHourParsed.value !== null) ? addHourParsed.value : 0;
    const hourRangeStr = `${pad2(hour)}:00-${pad2(hour)}:59`;
    allScheduledActions = `ADD: ${dateStr} ${hourRangeStr}`;
  }

  if (!action && rawRow[idxRemoveDate]) {
    const dateStr = Utilities.formatDate(new Date(rawRow[idxRemoveDate]), ACCOUNT_TIMEZONE, 'yyyy-MM-dd');
    const remHourRaw = idxRemoveHour !== -1 ? String(rawRow[idxRemoveHour] ?? '').trim() : '';
    const remHourParsed = parseCustomHourToInt(remHourRaw);
    const hour = (remHourParsed.ok && remHourParsed.value !== null) ? remHourParsed.value : 23;
    const hourRangeStr = `${pad2(hour)}:00-${pad2(hour)}:59`;
    const separator = allScheduledActions ? ' | ' : '';
    allScheduledActions += `${separator}REMOVE: ${dateStr} ${hourRangeStr}`;
  }

  // Scheduled időpont formázás
  if (effectiveAction === 'ADD' || effectiveAction === 'ADD+REMOVE') {
    if (rawRow[idxAddDate]) {
      const dateStr = Utilities.formatDate(new Date(rawRow[idxAddDate]), ACCOUNT_TIMEZONE, 'yyyy-MM-dd');
      const addHourRaw = idxAddHour !== -1 ? String(rawRow[idxAddHour] ?? '').trim() : '';
      const addHourParsed = parseCustomHourToInt(addHourRaw);
      const hour = (addHourParsed.ok && addHourParsed.value !== null) ? addHourParsed.value : 0;
      const hourRangeStr = `${pad2(hour)}:00-${pad2(hour)}:59`;
      scheduled = `${dateStr} ${hourRangeStr}`;
    }
  }

  if (effectiveAction === 'REMOVE' || effectiveAction === 'ADD+REMOVE') {
    if (rawRow[idxRemoveDate]) {
      const dateStr = Utilities.formatDate(new Date(rawRow[idxRemoveDate]), ACCOUNT_TIMEZONE, 'yyyy-MM-dd');
      const remHourRaw = idxRemoveHour !== -1 ? String(rawRow[idxRemoveHour] ?? '').trim() : '';
      const remHourParsed = parseCustomHourToInt(remHourRaw);
      const hour = (remHourParsed.ok && remHourParsed.value !== null) ? remHourParsed.value : 23;
      const hourRangeStr = `${pad2(hour)}:00-${pad2(hour)}:59`;
      const separator = scheduled ? ', ' : '';
      scheduled += `${separator}${dateStr} ${hourRangeStr}`;
    }
  }

  // ADD+REMOVE ugyanabban az órában ellenőrzés
  if (rawRow[idxAddDate] && rawRow[idxRemoveDate]) {
    const addDateStr = Utilities.formatDate(new Date(rawRow[idxAddDate]), ACCOUNT_TIMEZONE, 'yyyy-MM-dd');
    const removeDateStr = Utilities.formatDate(new Date(rawRow[idxRemoveDate]), ACCOUNT_TIMEZONE, 'yyyy-MM-dd');

    if (addDateStr === removeDateStr) {
      const addHourRaw = idxAddHour !== -1 ? String(rawRow[idxAddHour] ?? '').trim() : '';
      const removeHourRaw = idxRemoveHour !== -1 ? String(rawRow[idxRemoveHour] ?? '').trim() : '';
      const addHourParsed = parseCustomHourToInt(addHourRaw);
      const removeHourParsed = parseCustomHourToInt(removeHourRaw);

      const addHour = (addHourParsed.ok && addHourParsed.value !== null) ? addHourParsed.value : 0;
      const removeHour = (removeHourParsed.ok && removeHourParsed.value !== null) ? removeHourParsed.value : 23;

      if (addHour === removeHour) {
        errors.push('Hozzáadás és Törlés ugyanabban az órában - nem végrehajtható');
      }
    }
  }

  // 1. Bemeneti validáció
  if (!campaign) errors.push('Hiányzik Campaign Name');

  // 2. Video azonosító parse + validáció
  const parsed = parseVideoIdentifier(videoInput);

  if (parsed.type === 'INVALID') {
    errors.push(parsed.error);
  }

  let videoIdentifier = videoInput;
  let assetResourceName = null;

  if (parsed.type === 'ASSET_ID') {
    const assetDetails = fetchAssetDetails(parsed.value);
    if (!assetDetails) {
      errors.push('Asset ID nem található');
    } else if (assetDetails.type !== 'YOUTUBE_VIDEO') {
      errors.push(`Asset típusa ${assetDetails.type}, nem YOUTUBE_VIDEO`);
    } else {
      const customerId = AdsApp.currentAccount().getCustomerId().replace(/-/g, '');
      assetResourceName = `customers/${customerId}/assets/${parsed.value}`;

      // YouTube API validáció Asset ID esetén is, de csak ADD-nál (REMOVE-nál nem releváns)
      const isAddAction = action === 'ADD' || effectiveAction === 'ADD' || effectiveAction === 'ADD+REMOVE';
      if (assetDetails.youtubeVideoId && isAddAction) {
        videoIdentifier = assetDetails.youtubeVideoId;
        const ytValidation = validateYouTubeVideo(assetDetails.youtubeVideoId);

        if (!ytValidation.valid) {
          // YouTube API nem találja → ERROR (lehet private lett közben)
          errors.push(ytValidation.error);
        } else if (ytValidation.warning) {
          // API nem elérhető → WARNING
          warnings.push(ytValidation.warning);
        } else if (ytValidation.details) {
          const d = ytValidation.details;
          console.log(`    🎬 YouTube validáció (Asset ID → ${assetDetails.youtubeVideoId}): ${d.privacyStatus}, ${d.durationSeconds}s`);
        }
      }
    }
  }

  if (parsed.type === 'VIDEO_ID') {
    videoIdentifier = parsed.value;

    // Ellenőrizzük, hogy ez a Video ID már létezik-e asset-ként a fiókban
    const existingAssetRn = findExistingVideoAsset(parsed.value);
    if (existingAssetRn) {
      assetResourceName = existingAssetRn;
      console.log(`    ♻️ Video ID már létezik asset-ként: ${existingAssetRn}`);
    }

    const isAddAction = action === 'ADD' || effectiveAction === 'ADD' || effectiveAction === 'ADD+REMOVE';
    const isRemoveOnlyAction = (action === 'REMOVE' || effectiveAction === 'REMOVE') && !isAddAction;

    // REMOVE validáció: VIDEO_ID-nak léteznie kell asset-ként a fiókban
    if (isRemoveOnlyAction && !existingAssetRn) {
      errors.push('Video ID nem található asset-ként a fiókban (REMOVE-hoz léteznie kell)');
    }

    // YouTube API validáció csak ADD esetén (REMOVE-nál nem releváns)
    if (isAddAction) {
      const ytValidation = validateYouTubeVideo(parsed.value);

      if (!ytValidation.valid) {
        // YouTube API nem találja vagy hibás → ERROR
        // Private videót nem lehet hirdetni, nem létező videót sem
        errors.push(ytValidation.error);
      } else if (ytValidation.warning) {
        // API nem elérhető (nincs engedélyezve, quota) → WARNING
        warnings.push(ytValidation.warning);
      }

      // Logolás a validáció eredményéről
      if (ytValidation.details) {
        const d = ytValidation.details;
        console.log(`    🎬 YouTube validáció: ${parsed.value} - ${d.privacyStatus}, ${d.durationSeconds}s`);
      } else if (ytValidation.skipped) {
        console.log(`    ⚠️ YouTube validáció kihagyva: ${ytValidation.warning}`);
      }
    }
  }

  // 3. Kampány ellenőrzés
  const campaignId = campaignIdMap[campaign];
  if (!campaignId) {
    errors.push('Kampány nem található');
  }

  // Ha kritikus hiba van → RETURN
  if (errors.length > 0) {
    const scheduledCol = action ? scheduled : allScheduledActions;
    return {
      status: 'ERROR',
      // FIX: videoInput használata a konzisztens merge kulcs miatt
      rows: [[timestamp, campaign, assetGroup || '(all groups)', displayAssetType, videoInput, scheduledCol, effectiveAction, 'ERROR', errors.join('; ')]]
    };
  }

  // 4. Target groups
  const targetGroups = getTargetGroups(campaign, assetGroup, campaignId, groupStates);

  const groupNameList = targetGroups.length <= 3
    ? targetGroups.map(g => g.name).join(', ')
    : `${targetGroups.slice(0, 3).map(g => g.name).join(', ')} (+${targetGroups.length - 3} további)`;
  console.log(`    🎯 Target groups: ${targetGroups.length} (${assetGroup || 'all'}): ${groupNameList}`);

  if (targetGroups.length === 0) {
    const scheduledCol = action ? scheduled : allScheduledActions;
    return {
      status: 'ERROR',
      // FIX: videoInput használata a konzisztens merge kulcs miatt
      rows: [[timestamp, campaign, assetGroup || '(all groups)', displayAssetType, videoInput, scheduledCol, effectiveAction, 'ERROR', 'Nincs aktív (nem feed-only) asset group']]
    };
  }

  // 5. Limitek ellenőrzése - KÜLÖN SOR MINDEN GROUP-RA!
  const limits = VIDEO_LIMITS.YOUTUBE_VIDEO;
  const rows = [];
  const validGroups = [];
  let hasError = false;
  let hasWarning = false;

  targetGroups.forEach(g => {
    const groupErrors = [];
    const groupWarnings = [...warnings];

    const currentVideos = groupStates.videos[g.resourceName] || [];
    const currentCount = currentVideos.length;

    // ADD művelet
    if (effectiveAction === 'ADD' || effectiveAction === 'ADD+REMOVE') {
      const afterAdd = currentCount + 1;

      if (afterAdd > limits.max) {
        groupErrors.push(`Videólimit túllépés (${currentCount}+1 > ${limits.max})`);
      } else if (limits.warnThreshold > 0 && currentCount >= (limits.max - limits.warnThreshold)) {
        groupWarnings.push(`Limit közel (current=${currentCount}, after=${afterAdd}, max=${limits.max})`);
      }

      // Duplikáció ellenőrzés - FIX: VIDEO_ID esetén is működik, ha van assetResourceName
      let exists = false;
      if (assetResourceName) {
        exists = currentVideos.some(v => v.assetResource === assetResourceName);
      }
      if (exists) {
        groupErrors.push('Video asset már létezik a csoportban');
      }
    }

    // REMOVE művelet
    if (effectiveAction === 'REMOVE' || effectiveAction === 'ADD+REMOVE') {
      let exists = false;
      // FIX: VIDEO_ID esetén is működik, ha van assetResourceName
      if (assetResourceName) {
        exists = currentVideos.some(v => v.assetResource === assetResourceName);
      }

      // FIX: VIDEO_ID esetén is hibajelzés, ha nincs linkelve
      if (!exists && effectiveAction === 'REMOVE' && assetResourceName) {
        groupErrors.push('Video asset nem található a csoportban');
      }
    }

    // Per-group status
    const groupStatus = groupErrors.length > 0 ? 'ERROR' : groupWarnings.length > 0 ? 'WARNING' : 'OK';
    const groupMessage = groupErrors.length > 0 ? groupErrors.join('; ') : groupWarnings.length > 0 ? groupWarnings.join('; ') : 'OK';

    if (groupStatus === 'ERROR') hasError = true;
    if (groupStatus === 'WARNING') hasWarning = true;

    if (groupStatus === 'OK' || groupStatus === 'WARNING') {
      validGroups.push(g);
    }

    const scheduledCol = action ? scheduled : allScheduledActions;
    // FIX: videoInput használata (eredeti sheet érték) a konzisztens merge kulcs miatt
    // (executeVideoAsset is videoInput-ot használ, nem a konvertált videoIdentifier-t)
    rows.push([timestamp, campaign, g.name, displayAssetType, videoInput, scheduledCol, effectiveAction, groupStatus, groupMessage]);
  });

  const finalStatus = hasError ? 'ERROR' : hasWarning ? 'WARNING' : 'OK';

  return {
    status: finalStatus,
    rows: rows,
    validGroups: validGroups,
    parsedIdentifier: parsed
  };
}

/*** ===================== DEDUPLIKÁCIÓ ===================== ***/

/**
 * Cross-row duplikáció detektálás és deduplikáció (EXECUTION szinten)
 *
 * Problémák amiket kezel:
 * 1. Ugyanaz az asset többször szerepel különböző sorokban (ugyanaz a campaign, group, type, text/assetId, action, hour)
 * 2. Asset Group overlap: konkrét group + "all groups" (üres Asset Group Name)
 *
 * Kulcs: campaign|groupResourceName|assetType|text/assetId|action|hour
 *
 * Ha duplikációt talál: csak az ELSŐ előfordulást tartja meg, a többit eltávolítja.
 *
 * Megjegyzés: ADD+REMOVE konfliktust (sorok között, ugyanaz az óra) már Phase 8.5 kezeli ERROR-ként!
 */
function deduplicateExecutableItems(textItems, imageItems, videoItems, campaignIdMap, groupStates) {
  const seen = new Map(); // kulcs -> { itemIndex, groupIndex, sheetRowNum, assetType }
  let duplicateCount = 0;

  // Helper: kulcs generálás
  function makeKey(campaign, groupResourceName, assetType, textOrId, action, hour) {
    return `${campaign}|${groupResourceName}|${assetType}|${textOrId}|${action}|${hour || 'null'}`;
  }

  // Helper: asset azonosító kinyerése (text vagy assetId)
  function getAssetIdentifier(item, isText) {
    const { rawRow, headers } = item;
    if (isText) {
      const idxText = headers.indexOf('Text');
      return String(rawRow[idxText] || '').trim();
    } else {
      const idxAssetId = headers.indexOf('Asset ID');
      return String(rawRow[idxAssetId] || '').trim();
    }
  }

  // Helper: campaign név kinyerése
  function getCampaign(item) {
    const { rawRow, headers } = item;
    const idxCampaign = headers.indexOf('Campaign Name');
    return String(rawRow[idxCampaign] || '').trim();
  }

  // Helper: asset type kinyerése
  function getAssetType(item, isText) {
    const { rawRow, headers } = item;
    if (isText) {
      const idxTextType = headers.indexOf('Text Type');
      return String(rawRow[idxTextType] || '').trim().toUpperCase();
    } else {
      const idxImageType = headers.indexOf('Image Type');
      return String(rawRow[idxImageType] || '').trim();
    }
  }

  // TEXT items deduplikáció
  const dedupedTextItems = [];
  textItems.forEach((entry, itemIdx) => {
    const { item, validGroups } = entry;
    const campaign = getCampaign(item);
    const assetType = getAssetType(item, true);
    const text = getAssetIdentifier(item, true);
    const { action, customHour } = item;

    const remainingGroups = [];

    validGroups.forEach((group, groupIdx) => {
      const key = makeKey(campaign, group.resourceName, assetType, text, action, customHour);

      if (seen.has(key)) {
        const original = seen.get(key);
        console.log(`  ⚠️ DUPLIKÁCIÓ: Row ${item.sheetRowNum} (${assetType}="${text}", action=${action}, hour=${customHour || 'null'}, group=${group.name}) már szerepelt Row ${original.sheetRowNum}-ben → kihagyva`);
        duplicateCount++;
      } else {
        seen.set(key, { itemIndex: itemIdx, groupIndex: groupIdx, sheetRowNum: item.sheetRowNum, assetType });
        remainingGroups.push(group);
      }
    });

    // Ha maradtak group-ok, tartjuk meg az item-et
    if (remainingGroups.length > 0) {
      dedupedTextItems.push({ item, validGroups: remainingGroups });
    }
  });

  // IMAGE items deduplikáció
  const dedupedImageItems = [];
  imageItems.forEach((entry, itemIdx) => {
    const { item, validGroups } = entry;
    const campaign = getCampaign(item);
    const assetType = getAssetType(item, false);
    const assetId = getAssetIdentifier(item, false);
    const { action, customHour } = item;

    const remainingGroups = [];

    validGroups.forEach((group, groupIdx) => {
      const key = makeKey(campaign, group.resourceName, assetType, assetId, action, customHour);

      if (seen.has(key)) {
        const original = seen.get(key);
        console.log(`  ⚠️ DUPLIKÁCIÓ: Row ${item.sheetRowNum} (${assetType} assetId=${assetId}, action=${action}, hour=${customHour || 'null'}, group=${group.name}) már szerepelt Row ${original.sheetRowNum}-ben → kihagyva`);
        duplicateCount++;
      } else {
        seen.set(key, { itemIndex: itemIdx, groupIndex: groupIdx, sheetRowNum: item.sheetRowNum, assetType });
        remainingGroups.push(group);
      }
    });

    // Ha maradtak group-ok, tartjuk meg az item-et
    if (remainingGroups.length > 0) {
      dedupedImageItems.push({ item, validGroups: remainingGroups });
    }
  });

  // VIDEO items deduplikáció
  const dedupedVideoItems = [];
  videoItems.forEach((entry, itemIdx) => {
    const { item, validGroups, parsedIdentifier } = entry;
    const campaign = getCampaign(item);
    const assetType = 'YOUTUBE_VIDEO';  // Mindig ez
    const { action, customHour } = item;

    // Azonosító a parsedIdentifier alapján
    let identifier = '';
    if (parsedIdentifier.type === 'ASSET_ID') {
      identifier = `ASSET_ID:${parsedIdentifier.value}`;
    } else if (parsedIdentifier.type === 'VIDEO_ID') {
      identifier = `VIDEO_ID:${parsedIdentifier.value}`;
    }

    const remainingGroups = [];

    validGroups.forEach((group, groupIdx) => {
      const key = makeKey(campaign, group.resourceName, assetType, identifier, action, customHour);

      if (seen.has(key)) {
        const original = seen.get(key);
        console.log(`  ⚠️ DUPLIKÁCIÓ: Row ${item.sheetRowNum} (VIDEO ${identifier}, action=${action}, hour=${customHour || 'null'}, group=${group.name}) már szerepelt Row ${original.sheetRowNum}-ben → kihagyva`);
        duplicateCount++;
      } else {
        seen.set(key, { itemIndex: itemIdx, groupIndex: groupIdx, sheetRowNum: item.sheetRowNum, assetType });
        remainingGroups.push(group);
      }
    });

    if (remainingGroups.length > 0) {
      dedupedVideoItems.push({ item, validGroups: remainingGroups, parsedIdentifier });
    }
  });

  return {
    textItems: dedupedTextItems,
    imageItems: dedupedImageItems,
    videoItems: dedupedVideoItems,
    duplicateCount
  };
}

/*** ===================== VÉGREHAJTÁS ===================== ***/

function executeAll(textItemsWithGroups, imageItemsWithGroups, videoItemsWithGroups, campaignIdMap, groupStates) {
  const results = [];

  // Text assets
  textItemsWithGroups.forEach(({ item, validGroups }) => {
    const result = executeTextAsset(item, validGroups, campaignIdMap, groupStates);
    results.push(...result);
  });

  // Image assets
  imageItemsWithGroups.forEach(({ item, validGroups }) => {
    const result = executeImageAsset(item, validGroups, campaignIdMap, groupStates);
    results.push(...result);
  });

  // Video assets
  videoItemsWithGroups.forEach(({ item, validGroups, parsedIdentifier }) => {
    const result = executeVideoAsset(item, validGroups, parsedIdentifier, campaignIdMap, groupStates);
    results.push(...result);
  });

  return results;
}

function executeTextAsset(item, validGroups, campaignIdMap, groupStates) {
  const { rawRow, headers, action, customHour } = item;

  const idxCampaign = headers.indexOf('Campaign Name');
  const idxGroup = headers.indexOf('Asset Group Name');
  const idxTextType = headers.indexOf('Text Type');
  const idxText = headers.indexOf('Text');
  const idxAddDate = headers.indexOf('Add Date');
  const idxRemoveDate = headers.indexOf('Remove Date');

  const campaign = String(rawRow[idxCampaign] || '').trim();
  const assetGroup = String(rawRow[idxGroup] || '').trim();
  const textType = String(rawRow[idxTextType] || '').trim().toUpperCase();
  const text = String(rawRow[idxText] || '').trim();

  // Scheduled időpont kiszámítása (KRITIKUS FIX: RANGE formátum, mint a validációnál!)
  let scheduled = '';
  if (action === 'ADD' && rawRow[idxAddDate]) {
    const dateStr = Utilities.formatDate(new Date(rawRow[idxAddDate]), ACCOUNT_TIMEZONE, 'yyyy-MM-dd');
    const hour = customHour !== null ? customHour : 0;
    const hourRangeStr = `${pad2(hour)}:00-${pad2(hour)}:59`;
    scheduled = `${dateStr} ${hourRangeStr}`;
  } else if (action === 'REMOVE' && rawRow[idxRemoveDate]) {
    const dateStr = Utilities.formatDate(new Date(rawRow[idxRemoveDate]), ACCOUNT_TIMEZONE, 'yyyy-MM-dd');
    const hour = customHour !== null ? customHour : 23;
    const hourRangeStr = `${pad2(hour)}:00-${pad2(hour)}:59`;
    scheduled = `${dateStr} ${hourRangeStr}`;
  }

  // KRITIKUS: Csak a validGroups-okon hajtjuk végre (nem getTargetGroups()!)
  const targetGroups = validGroups;

  const results = [];

  if (action === 'ADD') {
    // Text asset létrehozása vagy keresése
    let assetResourceName = findExistingTextAsset(text);
    if (!assetResourceName) {
      assetResourceName = createTextAsset(text);
    }

    // Linkelés groupokhoz
    targetGroups.forEach(g => {
      const result = executeSingleMutation(() => {
        return linkTextAssetToGroup(g.resourceName, assetResourceName, textType);
      }, campaign, g.name, textType, text, 'ADD', scheduled);

      results.push(result);

      // Inter-operation delay: 750ms + random jitter (0-250ms) = 750-1000ms
      // Racing condition elkerülésére (v3.9.7)
      Utilities.sleep(750 + Math.floor(Math.random() * 250));
    });
  } else if (action === 'REMOVE') {
    targetGroups.forEach(g => {
      const currentAssets = groupStates.text[g.resourceName] || [];
      const ofType = currentAssets.filter(x => x.fieldType === textType);
      // Case-sensitive keresés (konzisztens az API-val)
      const match = ofType.find(x => x.text === text);

      if (match) {
        const result = executeSingleMutation(() => {
          return unlinkAssetFromGroup(match.agaResource);
        }, campaign, g.name, textType, text, 'REMOVE', scheduled);

        results.push(result);
        // Inter-operation delay: 750ms + random jitter (0-250ms) = 750-1000ms
        Utilities.sleep(750 + Math.floor(Math.random() * 250));
      }
    });
  }

  return results;
}

function executeImageAsset(item, validGroups, campaignIdMap, groupStates) {
  const { rawRow, headers, action, customHour } = item;

  const idxCampaign = headers.indexOf('Campaign Name');
  const idxGroup = headers.indexOf('Asset Group Name');
  const idxImageType = headers.indexOf('Image Type');
  const idxAssetId = headers.indexOf('Asset ID');
  const idxAddDate = headers.indexOf('Add Date');
  const idxRemoveDate = headers.indexOf('Remove Date');

  const campaign = String(rawRow[idxCampaign] || '').trim();
  const assetGroup = String(rawRow[idxGroup] || '').trim();
  const imageType = String(rawRow[idxImageType] || '').trim();
  const assetId = String(rawRow[idxAssetId] || '').trim();

  // Asset Type megjelenítéshez (execution results)
  const displayAssetType = imageType ? `Image ${imageType}` : 'Image';

  // Scheduled időpont kiszámítása (KRITIKUS FIX: RANGE formátum, mint a validációnál!)
  let scheduled = '';
  if (action === 'ADD' && rawRow[idxAddDate]) {
    const dateStr = Utilities.formatDate(new Date(rawRow[idxAddDate]), ACCOUNT_TIMEZONE, 'yyyy-MM-dd');
    const hour = customHour !== null ? customHour : 0;
    const hourRangeStr = `${pad2(hour)}:00-${pad2(hour)}:59`;
    scheduled = `${dateStr} ${hourRangeStr}`;
  } else if (action === 'REMOVE' && rawRow[idxRemoveDate]) {
    const dateStr = Utilities.formatDate(new Date(rawRow[idxRemoveDate]), ACCOUNT_TIMEZONE, 'yyyy-MM-dd');
    const hour = customHour !== null ? customHour : 23;
    const hourRangeStr = `${pad2(hour)}:00-${pad2(hour)}:59`;
    scheduled = `${dateStr} ${hourRangeStr}`;
  }

  // KRITIKUS: Csak a validGroups-okon hajtjuk végre (nem getTargetGroups()!)
  const targetGroups = validGroups;

  const customerId = AdsApp.currentAccount().getCustomerId().replace(/-/g, '');
  const assetResourceName = `customers/${customerId}/assets/${assetId}`;

  // KRITIKUS v7.3.21: Pontos fieldType meghatározása a TÉNYLEGES kép méretek alapján!
  // MINDEN képtípusnál: HORIZONTAL, SQUARE, PORTRAIT (4:5), TALL_PORTRAIT (9:16)
  let correctFieldType = normalizeImageType(imageType); // fallback
  const assetDetails = fetchAssetDetails(assetId);
  if (assetDetails && assetDetails.type === 'IMAGE') {
    const typeResult = determineImageFieldType(imageType, assetDetails.width, assetDetails.height);
    correctFieldType = typeResult.fieldType;

    // KÖZEPES FIX #3: Log ha mismatch van execution során
    if (typeResult.mismatch) {
      console.log(`  ⚠️ EXECUTION Image Type mismatch: Sheet="${imageType}" → Actual="${typeResult.actualDetected}", using fieldType=${correctFieldType}`);
    }
  }

  const results = [];

  if (action === 'ADD') {
    targetGroups.forEach(g => {
      const result = executeSingleMutation(() => {
        return linkImageAssetToGroup(g.resourceName, assetResourceName, correctFieldType);
      }, campaign, g.name, displayAssetType, assetId, 'ADD', scheduled);

      results.push(result);
      // Inter-operation delay: 750ms + random jitter (0-250ms) = 750-1000ms
      Utilities.sleep(750 + Math.floor(Math.random() * 250));
    });
  } else if (action === 'REMOVE') {
    targetGroups.forEach(g => {
      const currentImages = groupStates.images[g.resourceName] || [];
      const match = currentImages.find(img => img.assetResource === assetResourceName);

      if (match) {
        const result = executeSingleMutation(() => {
          return unlinkAssetFromGroup(match.agaResource);
        }, campaign, g.name, displayAssetType, assetId, 'REMOVE', scheduled);

        results.push(result);
        // Inter-operation delay: 750ms + random jitter (0-250ms) = 750-1000ms
        Utilities.sleep(750 + Math.floor(Math.random() * 250));
      }
    });
  }

  return results;
}

function executeVideoAsset(item, validGroups, parsedIdentifier, campaignIdMap, groupStates) {
  const { rawRow, headers, action, customHour } = item;

  const idxCampaign = headers.indexOf('Campaign Name');
  const idxGroup = headers.indexOf('Asset Group Name');
  const idxVideoAspect = headers.indexOf('Video Aspect (optional)');
  const idxVideoId = headers.indexOf('Video URL or ID / Asset ID');
  const idxAddDate = headers.indexOf('Add Date');
  const idxRemoveDate = headers.indexOf('Remove Date');

  const campaign = String(rawRow[idxCampaign] || '').trim();
  const assetGroup = String(rawRow[idxGroup] || '').trim();
  const videoAspect = String(rawRow[idxVideoAspect] || '').trim();
  const videoInput = String(rawRow[idxVideoId] || '').trim();

  // Asset Type megjelenítéshez (execution results)
  const displayAssetType = videoAspect ? `Video ${videoAspect}` : 'Video';

  // Scheduled időpont
  let scheduled = '';
  if (action === 'ADD' && rawRow[idxAddDate]) {
    const dateStr = Utilities.formatDate(new Date(rawRow[idxAddDate]), ACCOUNT_TIMEZONE, 'yyyy-MM-dd');
    const hour = customHour !== null ? customHour : 0;
    const hourRangeStr = `${pad2(hour)}:00-${pad2(hour)}:59`;
    scheduled = `${dateStr} ${hourRangeStr}`;
  } else if (action === 'REMOVE' && rawRow[idxRemoveDate]) {
    const dateStr = Utilities.formatDate(new Date(rawRow[idxRemoveDate]), ACCOUNT_TIMEZONE, 'yyyy-MM-dd');
    const hour = customHour !== null ? customHour : 23;
    const hourRangeStr = `${pad2(hour)}:00-${pad2(hour)}:59`;
    scheduled = `${dateStr} ${hourRangeStr}`;
  }

  const targetGroups = validGroups;
  const customerId = AdsApp.currentAccount().getCustomerId().replace(/-/g, '');

  let assetResourceName = null;

  // Asset resource name meghatározása
  if (parsedIdentifier.type === 'ASSET_ID') {
    assetResourceName = `customers/${customerId}/assets/${parsedIdentifier.value}`;
  } else if (parsedIdentifier.type === 'VIDEO_ID') {
    // YouTube Video ID → keressük, és csak ADD esetén hozzuk létre
    assetResourceName = findExistingVideoAsset(parsedIdentifier.value);
    if (!assetResourceName && action === 'ADD') {
      try {
        assetResourceName = createVideoAsset(parsedIdentifier.value);
      } catch (e) {
        console.log(`  ❌ Video asset létrehozás sikertelen: ${e.message}`);
        return [[
          Utilities.formatDate(new Date(), ACCOUNT_TIMEZONE, 'yyyy-MM-dd HH:mm:ss'),
          campaign, assetGroup || '(all groups)', displayAssetType, videoInput,
          scheduled, action, 'ERROR', `Video asset létrehozás sikertelen: ${e.message}`
        ]];
      }
    }
  }

  if (!assetResourceName) {
    const errorMsg = action === 'REMOVE'
      ? 'Video asset nem található a fiókban (REMOVE-hoz léteznie kell)'
      : 'Asset resource name nem határozható meg';
    return [[
      Utilities.formatDate(new Date(), ACCOUNT_TIMEZONE, 'yyyy-MM-dd HH:mm:ss'),
      campaign, assetGroup || '(all groups)', displayAssetType, videoInput,
      scheduled, action, 'ERROR', errorMsg
    ]];
  }

  const results = [];

  if (action === 'ADD') {
    targetGroups.forEach(g => {
      const result = executeSingleMutation(() => {
        return linkVideoAssetToGroup(g.resourceName, assetResourceName);
      }, campaign, g.name, displayAssetType, videoInput, 'ADD', scheduled);

      results.push(result);
      Utilities.sleep(750 + Math.floor(Math.random() * 250));
    });
  } else if (action === 'REMOVE') {
    targetGroups.forEach(g => {
      const currentVideos = groupStates.videos[g.resourceName] || [];
      const match = currentVideos.find(v => v.assetResource === assetResourceName);

      if (match) {
        const result = executeSingleMutation(() => {
          return unlinkAssetFromGroup(match.agaResource);
        }, campaign, g.name, displayAssetType, videoInput, 'REMOVE', scheduled);

        results.push(result);
        Utilities.sleep(750 + Math.floor(Math.random() * 250));
      }
    });
  }

  return results;
}

function executeSingleMutation(mutationFn, campaign, groupName, assetType, textOrId, action, scheduled) {
  const tz = ACCOUNT_TIMEZONE;
  const timestamp = Utilities.formatDate(new Date(), tz, 'yyyy-MM-dd HH:mm:ss');

  let attempt = 0;
  const maxRetries = RETRY_MAX_ATTEMPTS;

  while (attempt <= maxRetries) {
    try {
      const operation = mutationFn();
      const res = AdsApp.mutateAll([operation], { partialFailure: true, apiVersion: API_VERSION });

      const r = res[0];
      if (r.isSuccessful && r.isSuccessful()) {
        return [timestamp, campaign, groupName, assetType, textOrId, scheduled, action, 'SUCCESS', `Végrehajtva (attempt ${attempt + 1})`];
      } else {
        // KRITIKUS: Google Ads Scripts MutateResult API - getErrorMessages() tömböt ad!
        const errs = r.getErrorMessages ? r.getErrorMessages() : ['ismeretlen hiba'];
        const errorText = errs.join('; ');
        const isRaceCondition = /Another task is also trying to change|CONCURRENT_MODIFICATION|Multiple requests were attempting to modify the same resource/i.test(errorText);

        if (isRaceCondition && attempt < maxRetries) {
          const baseWait = Math.pow(2, attempt) * RETRY_BASE_DELAY_MS;
          // Jitter: ±25% randomizálás, hogy ne ütközzenek a párhuzamos mutation-ök
          const jitter = baseWait * (0.75 + Math.random() * 0.5); // 0.75-1.25x
          const waitMs = Math.floor(jitter);
          console.log(`  ⚠️ Race condition, retry #${attempt + 1} (~${waitMs}ms)`);
          Utilities.sleep(waitMs);
          attempt++;
          continue;
        }

        return [timestamp, campaign, groupName, assetType, textOrId, scheduled, action, 'ERROR', `Hiba: ${errorText}`];
      }
    } catch (e) {
      const isTransient = /RESOURCE_EXHAUSTED|INTERNAL|DEADLINE_EXCEEDED/i.test(e.message);

      if (isTransient && attempt < maxRetries) {
        const baseWait = Math.pow(2, attempt) * RETRY_BASE_DELAY_MS;
        // Jitter: ±25% randomizálás
        const jitter = baseWait * (0.75 + Math.random() * 0.5);
        const waitMs = Math.floor(jitter);
        console.log(`  ⚠️ Átmeneti hiba, retry #${attempt + 1} (~${waitMs}ms)`);
        Utilities.sleep(waitMs);
        attempt++;
        continue;
      }

      return [timestamp, campaign, groupName, assetType, textOrId, scheduled, action, 'ERROR', `Exception: ${e.message}`];
    }
  }

  return [timestamp, campaign, groupName, assetType, textOrId, scheduled, action, 'ERROR', `Max retry elérve (${maxRetries})`];
}

function findExistingTextAsset(text) {
  const safe = gaqlEscapeSingleQuote(text);
  const q = `
    SELECT asset.resource_name
    FROM asset
    WHERE asset.type = TEXT
      AND asset.text_asset.text = '${safe}'
    LIMIT 1
  `;

  try {
    const rows = runReportSafe(q, 'findExistingTextAsset');
    if (rows.length > 0) {
      return rows[0]['asset.resource_name'];
    }
  } catch (e) {}

  return null;
}

function createTextAsset(text) {
  const ops = [{
    assetOperation: {
      create: { textAsset: { text: text } }
    }
  }];

  const res = AdsApp.mutateAll(ops, { partialFailure: true, apiVersion: API_VERSION });

  // KRITIKUS: Google Ads Scripts MutateResult API
  const r = res[0];
  if (r.isSuccessful && r.isSuccessful()) {
    // Sikeres létrehozás - getResourceName() metódus
    const rn = r.getResourceName ? r.getResourceName() : null;
    if (rn) {
      console.log(`  ✅ TEXT asset létrehozva: ${rn}`);
      return rn;
    }

    // Fallback: lookup a sikeres létrehozás után (ha getResourceName nem működött)
    console.log(`  ⚠️ Mutation sikeres, de getResourceName() null, fallback lookup...`);
    Utilities.sleep(1000);
    const foundRn = findExistingTextAsset(text);
    if (foundRn) return foundRn;
  }

  // Mutation sikertelen vagy nem találjuk az assetet
  const errs = r.getErrorMessages ? r.getErrorMessages() : ['ismeretlen hiba'];
  throw new Error(`TEXT asset létrehozás sikertelen: ${errs.join('; ')}`);
}

function linkTextAssetToGroup(groupResourceName, assetResourceName, fieldType) {
  return {
    assetGroupAssetOperation: {
      create: {
        assetGroup: groupResourceName,
        asset: assetResourceName,
        fieldType: fieldType
      }
    }
  };
}

function linkImageAssetToGroup(groupResourceName, assetResourceName, fieldType) {
  return {
    assetGroupAssetOperation: {
      create: {
        assetGroup: groupResourceName,
        asset: assetResourceName,
        fieldType: fieldType
      }
    }
  };
}

function unlinkAssetFromGroup(assetGroupAssetResourceName) {
  return {
    assetGroupAssetOperation: {
      remove: assetGroupAssetResourceName
    }
  };
}

/**
 * Meglévő YouTube Video asset keresése (YouTube Video ID alapján)
 */
function findExistingVideoAsset(youtubeVideoId) {
  const safe = gaqlEscapeSingleQuote(youtubeVideoId);
  const q = `
    SELECT asset.resource_name
    FROM asset
    WHERE asset.type = YOUTUBE_VIDEO
      AND asset.youtube_video_asset.youtube_video_id = '${safe}'
    LIMIT 1
  `;

  try {
    const rows = runReportSafe(q, 'findExistingVideoAsset');
    if (rows.length > 0) {
      return rows[0]['asset.resource_name'];
    }
  } catch (e) {}

  return null;
}

/**
 * YouTube Video asset létrehozása
 */
function createVideoAsset(youtubeVideoId) {
  const ops = [{
    assetOperation: {
      create: {
        type: 'YOUTUBE_VIDEO',
        youtubeVideoAsset: {
          youtubeVideoId: youtubeVideoId
        }
      }
    }
  }];

  const res = AdsApp.mutateAll(ops, { partialFailure: true, apiVersion: API_VERSION });

  const r = res[0];
  if (r.isSuccessful && r.isSuccessful()) {
    const rn = r.getResourceName ? r.getResourceName() : null;
    if (rn) {
      console.log(`  ✅ VIDEO asset létrehozva: ${rn}`);
      return rn;
    }

    // Fallback lookup
    console.log(`  ⚠️ Mutation sikeres, de getResourceName() null, fallback lookup...`);
    Utilities.sleep(1000);
    const foundRn = findExistingVideoAsset(youtubeVideoId);
    if (foundRn) return foundRn;
  }

  const errs = r.getErrorMessages ? r.getErrorMessages() : ['ismeretlen hiba'];
  throw new Error(`VIDEO asset létrehozás sikertelen: ${errs.join('; ')}`);
}

/**
 * Video asset linkelése asset group-hoz
 */
function linkVideoAssetToGroup(groupResourceName, assetResourceName) {
  return {
    assetGroupAssetOperation: {
      create: {
        assetGroup: groupResourceName,
        asset: assetResourceName,
        fieldType: 'YOUTUBE_VIDEO'
      }
    }
  };
}

/*** ===================== POST-VERIFICATION ===================== ***/

function verifyVideoAssetLinked(groupResourceName, assetResourceName, action) {
  const q = `
    SELECT asset_group_asset.asset
    FROM asset_group_asset
    WHERE asset_group_asset.asset_group = '${groupResourceName}'
      AND asset_group_asset.status = ENABLED
      AND asset_group_asset.field_type = YOUTUBE_VIDEO
      AND asset_group_asset.source = ADVERTISER
  `;

  try {
    const rows = runReportSafe(q, 'verifyVideoAssetLinked');
    const exists = rows.some(r => r['asset_group_asset.asset'] === assetResourceName);

    if (action === 'ADD') return exists;
    if (action === 'REMOVE') return !exists;
  } catch (e) {}

  return false;
}

function verifyVideoAssetLinkedWithRetry(groupResourceName, assetResourceName, action) {
  const maxAttempts = 5;
  const delayMs = 1000;

  for (let attempt = 1; attempt <= maxAttempts; attempt++) {
    if (attempt > 1) {
      Utilities.sleep(delayMs);
    } else {
      Utilities.sleep(500);
    }

    const verified = verifyVideoAssetLinked(groupResourceName, assetResourceName, action);
    if (verified) {
      if (attempt > 1) {
        console.log(`  ✅ Post-verify OK (attempt ${attempt}/${maxAttempts})`);
      }
      return true;
    }

    if (attempt < maxAttempts) {
      console.log(`  ⏳ Post-verify várakozás (attempt ${attempt}/${maxAttempts}), ${delayMs}ms delay...`);
    }
  }

  console.log(`  ❌ Post-verify FAILED (${maxAttempts} attempts)`);
  return false;
}

function postVerifyAll(executionResults, groupStates) {
  console.log('  🔍 Post-verification...');

  const verifiedResults = executionResults.map(result => {
    const [timestamp, campaign, groupName, assetType, textOrId, scheduled, action, status, message] = result;

    if (status !== 'SUCCESS') return result;

    // Group resource name keresés - KRITIKUS: campaign ÉS groupName alapján!
    let groupResourceName = null;
    Object.keys(groupStates.groupMap || {}).forEach(rn => {
      const g = groupStates.groupMap[rn];
      if (normalizeText(g.campaignName) === normalizeText(campaign) &&
          normalizeText(g.name) === normalizeText(groupName)) {
        groupResourceName = rn;
      }
    });

    if (!groupResourceName) {
      return [timestamp, campaign, groupName, assetType, textOrId, scheduled, action, 'ERROR', `Post-verify: group nem található (${campaign} / ${groupName})`];
    }

    // TEXT asset verify (retry loop propagációs késleltetésre)
    if (['HEADLINE', 'LONG_HEADLINE', 'DESCRIPTION'].includes(assetType)) {
      const verified = verifyTextAssetLinkedWithRetry(groupResourceName, assetType, textOrId, action);
      if (!verified) {
        return [timestamp, campaign, groupName, assetType, textOrId, scheduled, action, 'ERROR', 'Post-verify FAILED'];
      }
      return [timestamp, campaign, groupName, assetType, textOrId, scheduled, action, 'SUCCESS', message + ' [Verified ✓]'];
    }

    // IMAGE asset verify (retry loop)
    // FIX: case-insensitive check - displayAssetType = 'Image SQUARE (1:1)' nem 'IMAGE'
    if (assetType.toLowerCase().includes('image') || ['MARKETING_IMAGE', 'SQUARE_MARKETING_IMAGE', 'PORTRAIT_MARKETING_IMAGE', 'TALL_PORTRAIT_MARKETING_IMAGE'].includes(assetType)) {
      const verified = verifyImageAssetLinkedWithRetry(groupResourceName, textOrId, action);
      if (!verified) {
        return [timestamp, campaign, groupName, assetType, textOrId, scheduled, action, 'ERROR', 'Post-verify FAILED'];
      }
      return [timestamp, campaign, groupName, assetType, textOrId, scheduled, action, 'SUCCESS', message + ' [Verified ✓]'];
    }

    // VIDEO asset verify (retry loop)
    // FIX: case-insensitive check - displayAssetType = 'Video HORIZONTAL...' nem 'VIDEO'
    if (assetType.toLowerCase().includes('video')) {
      // Asset resource name reconstruct
      const parsed = parseVideoIdentifier(textOrId);
      let assetResourceName = null;
      const customerId = AdsApp.currentAccount().getCustomerId().replace(/-/g, '');

      if (parsed.type === 'ASSET_ID') {
        assetResourceName = `customers/${customerId}/assets/${parsed.value}`;
      } else if (parsed.type === 'VIDEO_ID') {
        // FIX: Retry logika propagációs késleltetés kezelésére
        // Ha az executeVideoAsset most hozta létre az assetet, a GAQL még nem biztos, hogy látja
        for (let attempt = 1; attempt <= 3; attempt++) {
          assetResourceName = findExistingVideoAsset(parsed.value);
          if (assetResourceName) break;
          if (attempt < 3) {
            console.log(`  ⏳ Video asset propagáció várása (${attempt}/3)...`);
            Utilities.sleep(1500);
          }
        }
      }

      if (!assetResourceName) {
        return [timestamp, campaign, groupName, assetType, textOrId, scheduled, action, 'ERROR', 'Post-verify: asset resource name nem található'];
      }

      const verified = verifyVideoAssetLinkedWithRetry(groupResourceName, assetResourceName, action);
      if (!verified) {
        return [timestamp, campaign, groupName, assetType, textOrId, scheduled, action, 'ERROR', 'Post-verify FAILED'];
      }
      return [timestamp, campaign, groupName, assetType, textOrId, scheduled, action, 'SUCCESS', message + ' [Verified ✓]'];
    }

    return result;
  });

  return verifiedResults;
}

function verifyTextAssetLinked(groupResourceName, fieldType, text, action) {
  const q = `
    SELECT asset_group_asset.asset_group, asset.text_asset.text
    FROM asset_group_asset
    WHERE asset_group_asset.asset_group = '${groupResourceName}'
      AND asset_group_asset.field_type = ${fieldType}
      AND asset_group_asset.status = ENABLED
      AND asset_group_asset.source = ADVERTISER
  `;

  try {
    const rows = runReportSafe(q, 'verifyTextAssetLinked');
    // Case-sensitive ellenőrzés (konzisztens az API-val)
    const exists = rows.some(r => r['asset.text_asset.text'] === text);

    if (action === 'ADD') return exists;
    if (action === 'REMOVE') return !exists;
  } catch (e) {}

  return false;
}

function verifyImageAssetLinked(groupResourceName, assetId, action) {
  const customerId = AdsApp.currentAccount().getCustomerId().replace(/-/g, '');
  const assetResourceName = `customers/${customerId}/assets/${assetId}`;

  const q = `
    SELECT asset_group_asset.asset
    FROM asset_group_asset
    WHERE asset_group_asset.asset_group = '${groupResourceName}'
      AND asset_group_asset.status = ENABLED
      AND asset_group_asset.field_type IN (MARKETING_IMAGE, SQUARE_MARKETING_IMAGE, PORTRAIT_MARKETING_IMAGE, TALL_PORTRAIT_MARKETING_IMAGE)
      AND asset_group_asset.source = ADVERTISER
  `;

  try {
    const rows = runReportSafe(q, 'verifyImageAssetLinked');
    const exists = rows.some(r => r['asset_group_asset.asset'] === assetResourceName);

    if (action === 'ADD') return exists;
    if (action === 'REMOVE') return !exists;
  } catch (e) {}

  return false;
}

// Retry wrapper TEXT asset verify-hez (propagációs késleltetésre)
function verifyTextAssetLinkedWithRetry(groupResourceName, fieldType, text, action) {
  const maxAttempts = 5;
  const delayMs = 1000;  // 1s backoff

  for (let attempt = 1; attempt <= maxAttempts; attempt++) {
    // Várakozás propagációra (kivéve első kísérlet)
    if (attempt > 1) {
      Utilities.sleep(delayMs);
    } else {
      // Első kísérlet előtt minimális várakozás
      Utilities.sleep(500);
    }

    const verified = verifyTextAssetLinked(groupResourceName, fieldType, text, action);
    if (verified) {
      if (attempt > 1) {
        console.log(`  ✅ Post-verify OK (attempt ${attempt}/${maxAttempts})`);
      }
      return true;
    }

    if (attempt < maxAttempts) {
      console.log(`  ⏳ Post-verify várakozás (attempt ${attempt}/${maxAttempts}), ${delayMs}ms delay...`);
    }
  }

  console.log(`  ❌ Post-verify FAILED (${maxAttempts} attempts)`);
  return false;
}

// Retry wrapper IMAGE asset verify-hez (propagációs késleltetésre)
function verifyImageAssetLinkedWithRetry(groupResourceName, assetId, action) {
  const maxAttempts = 5;
  const delayMs = 1000;  // 1s backoff

  for (let attempt = 1; attempt <= maxAttempts; attempt++) {
    // Várakozás propagációra (kivéve első kísérlet)
    if (attempt > 1) {
      Utilities.sleep(delayMs);
    } else {
      // Első kísérlet előtt minimális várakozás
      Utilities.sleep(500);
    }

    const verified = verifyImageAssetLinked(groupResourceName, assetId, action);
    if (verified) {
      if (attempt > 1) {
        console.log(`  ✅ Post-verify OK (attempt ${attempt}/${maxAttempts})`);
      }
      return true;
    }

    if (attempt < maxAttempts) {
      console.log(`  ⏳ Post-verify várakozás (attempt ${attempt}/${maxAttempts}), ${delayMs}ms delay...`);
    }
  }

  console.log(`  ❌ Post-verify FAILED (${maxAttempts} attempts)`);
  return false;
}

/*** ===================== SHEET ÍRÁS ===================== ***/

function clearPreviewResultsSheet(ss) {
  let sh = ss.getSheetByName(PREVIEW_RESULTS_SHEET_NAME);
  if (!sh) {
    sh = ss.insertSheet(PREVIEW_RESULTS_SHEET_NAME);
  }
  sh.clear();
  sh.appendRow(['Timestamp', 'Campaign', 'Asset Group', 'Asset Type', 'Text/Asset ID', 'All Scheduled Actions', 'Next Action', 'Status', 'Validation Message']);
  sh.getRange(1, 1, 1, 9).setFontWeight('bold').setBackground('#f6f8fa');
}

function writePreviewResultsSheet(ss, rows, campaignIdMap, groupStates) {
  clearPreviewResultsSheet(ss);

  if (!rows || rows.length === 0) return;

  const sh = ss.getSheetByName(PREVIEW_RESULTS_SHEET_NAME);
  sh.getRange(2, 1, rows.length, 9).setValues(rows);

  // FIX v7.4.6: Hyperlink-ek Campaign, Asset Group, Asset ID cellákra
  // Debug: Logoljuk az első sort hogy lássuk milyen adatokat kapunk
  if (rows.length > 0) {
    console.log(`  🔍 Preview hyperlink debug (első sor):`);
    console.log(`    Campaign: "${rows[0][1]}"`);
    console.log(`    Asset Group: "${rows[0][2]}"`);
    console.log(`    Asset Type: "${rows[0][3]}"`);
    console.log(`    Asset ID/Text: "${rows[0][4]}"`);
  }

  rows.forEach((r, idx) => {
    const rowNum = idx + 2;
    const assetType = String(r[3] || '').toLowerCase();
    const assetIdOrText = String(r[4] || '');

    // IMAGE: Asset ID → Google Ads asset link
    try {
      if (idx === 0) {
        console.log(`  🔍 IMAGE check (első sor): assetType="${assetType}", includes('image')=${assetType.includes('image')}`);
      }
      if (assetType.includes('image')) {
        const assetIdMatch = assetIdOrText.match(/(\d{10,})/);
        if (idx === 0) {
          console.log(`    assetIdOrText="${assetIdOrText}", match result=${assetIdMatch ? assetIdMatch[1] : 'null'}`);
        }
        if (assetIdMatch) {
          const assetId = assetIdMatch[1];
          // fetchAssetDetails → asset.image_asset.full_size.url = tpc.googlesyndication.com URL
          const details = fetchAssetDetails(assetId);
          if (details && details.imageUrl) {
            sh.getRange(rowNum, 5).setFormula(`=HYPERLINK("${details.imageUrl}"; "🖼️ ${assetId}")`);
          } else {
            sh.getRange(rowNum, 5).setValue(`🖼️ ${assetId}`);
          }
          if (idx === 0) {
            console.log(`    ✅ IMAGE link beállítva (row ${rowNum}): ${details && details.imageUrl ? details.imageUrl : 'nincs URL'}`);
          }
        }
      }
    } catch (e) {
      console.log(`  ⚠️ IMAGE hyperlink hiba (row ${rowNum}): ${e.message}`);
    }

    // VIDEO: YouTube ID → YouTube link
    try {
      if (idx === 0) {
        console.log(`  🔍 VIDEO check (első sor): assetType="${assetType}", includes('video')=${assetType.includes('video')}`);
      }
      if (assetType.includes('video')) {
        const parsed = parseVideoIdentifier(assetIdOrText);
        if (idx === 0) {
          console.log(`    assetIdOrText="${assetIdOrText}", parsed.type="${parsed.type}", parsed.value="${parsed.value}"`);
        }
        if (parsed.type === 'VIDEO_ID') {
          const videoId = parsed.value;
          const youtubeUrl = `https://www.youtube.com/watch?v=${videoId}`;
          sh.getRange(rowNum, 5).setFormula(`=HYPERLINK("${youtubeUrl}"; "▶️ ${videoId}")`);
          if (idx === 0) {
            console.log(`    ✅ VIDEO hyperlink beállítva (row ${rowNum})`);
          }
        }
      }
    } catch (e) {
      console.log(`  ⚠️ VIDEO hyperlink hiba (row ${rowNum}): ${e.message}`);
    }
  });

  // Színezés
  const colors = rows.map(r => {
    const status = String(r[7] || '');
    let color;
    if (status === 'OK') color = COLOR_OK;
    else if (status === 'WARNING') color = COLOR_WARNING;
    else if (status === 'ERROR') color = COLOR_ERROR;
    else color = 'white';
    return Array(9).fill(color);
  });

  sh.getRange(2, 1, rows.length, 9).setBackgrounds(colors);

  try { sh.autoResizeColumns(1, 9); } catch (_) {}
}

function appendResultsSheet(ss, rows, campaignIdMap, groupStates) {
  if (!rows || rows.length === 0) return;

  let sh = ss.getSheetByName(RESULTS_SHEET_NAME);
  if (!sh) {
    sh = ss.insertSheet(RESULTS_SHEET_NAME);
    // v7.5.2: 'All Scheduled Actions' oszlop eltávolítva (execution módban félrevezető volt)
    // Ha a sheet korábban 9 oszloppal volt létrehozva, töröld és futtasd újra!
    sh.appendRow(['Timestamp', 'Campaign', 'Asset Group', 'Asset Type', 'Text/Asset ID', 'Executed Action', 'Status', 'Execution Message']);
    sh.getRange(1, 1, 1, 8).setFontWeight('bold').setBackground('#f6f8fa');
  }

  // Belső sor formátum: [ts, campaign, group, type, textOrId, scheduledCol(skip), action, status, msg]
  // Az 5. index (scheduledCol) nem kerül kiírásra – execution módban csak egyetlen időpont lenne.
  const mappedRows = rows.map(r => [r[0], r[1], r[2], r[3], r[4], r[6], r[7], r[8]]);
  const startRow = sh.getLastRow() + 1;
  sh.getRange(startRow, 1, mappedRows.length, 8).setValues(mappedRows);

  // FIX v7.4.6: Hyperlink-ek Campaign, Asset Group, Asset ID cellákra
  rows.forEach((r, idx) => {
    const rowNum = startRow + idx;
    const assetType = String(r[3] || '').toLowerCase();
    const assetIdOrText = String(r[4] || '');

    // IMAGE: Asset ID → Google Ads asset link
    try {
      if (assetType.includes('image')) {
        const assetIdMatch = assetIdOrText.match(/(\d{10,})/);
        if (assetIdMatch) {
          const assetId = assetIdMatch[1];
          // fetchAssetDetails → asset.image_asset.full_size.url = tpc.googlesyndication.com URL
          const details = fetchAssetDetails(assetId);
          if (details && details.imageUrl) {
            sh.getRange(rowNum, 5).setFormula(`=HYPERLINK("${details.imageUrl}"; "🖼️ ${assetId}")`);
          } else {
            sh.getRange(rowNum, 5).setValue(`🖼️ ${assetId}`);
          }
        }
      }
    } catch (e) {
      console.log(`  ⚠️ IMAGE hyperlink hiba (row ${rowNum}): ${e.message}`);
    }

    // VIDEO: YouTube ID → YouTube link
    try {
      if (assetType.includes('video')) {
        const parsed = parseVideoIdentifier(assetIdOrText);
        if (parsed.type === 'VIDEO_ID') {
          const videoId = parsed.value;
          const youtubeUrl = `https://www.youtube.com/watch?v=${videoId}`;
          sh.getRange(rowNum, 5).setFormula(`=HYPERLINK("${youtubeUrl}"; "▶️ ${videoId}")`);
        }
      }
    } catch (e) {
      console.log(`  ⚠️ VIDEO hyperlink hiba (row ${rowNum}): ${e.message}`);
    }
  });

  // Háttérszín kódolás (mint writePreviewResultsSheet-ben)
  const colors = rows.map(r => {
    const status = String(r[7] || '');
    let color;
    if (status === 'SUCCESS') color = COLOR_SUCCESS;
    else if (status === 'SUCCESS_WITH_WARNING') color = COLOR_SUCCESS_WITH_WARNING;
    else if (status === 'ERROR') color = COLOR_FAILED;
    else if (status === 'OK') color = COLOR_OK;
    else if (status === 'WARNING') color = COLOR_WARNING;
    else color = 'white';
    return Array(8).fill(color);
  });
  sh.getRange(startRow, 1, rows.length, 8).setBackgrounds(colors);
}

/*** ===================== SHEET HASH & CHANGE DETECTION ===================== ***/

function calculateSheetHash(textRows, imageRows, videoRows) {
  // Egyszerű hash számítás a sheet tartalmából
  // Sorokat JSON-né alakítjuk és összefűzzük
  const textContent = textRows.map(item => JSON.stringify(item.rawRow)).join('|');
  const imageContent = imageRows.map(item => JSON.stringify(item.rawRow)).join('|');
  const videoContent = videoRows.map(item => JSON.stringify(item.rawRow)).join('|');
  const combined = textContent + '||' + imageContent + '||' + videoContent;

  // Egyszerű hash algoritmus (nem kriptográfiai, de elég change detection-höz)
  let hash = 0;
  for (let i = 0; i < combined.length; i++) {
    const char = combined.charCodeAt(i);
    hash = ((hash << 5) - hash) + char;
    hash = hash & hash; // Convert to 32bit integer
  }

  return hash.toString();
}

function getPreviousSheetHash() {
  try {
    const props = PropertiesService.getScriptProperties();
    return props.getProperty('SHEET_HASH') || '';
  } catch (e) {
    console.log(`⚠️ PropertiesService read error: ${e.message}`);
    return '';
  }
}

function saveSheetHash(hash) {
  try {
    const props = PropertiesService.getScriptProperties();
    props.setProperty('SHEET_HASH', hash);
  } catch (e) {
    console.log(`⚠️ PropertiesService write error: ${e.message}`);
  }
}

/*** ===================== THUMBNAIL UPDATE ===================== ***/

/**
 * Thumbnail oszlop frissítése a VideoAssets sheet-ben
 * Validáció után hívódik meg, beírja az =IMAGE() formulákat
 *
 * @param {Spreadsheet} ss - Google Spreadsheet objektum
 * @param {Array} videoRows - Validált video sorok
 */
function updateVideoThumbnails(ss, videoRows) {
  if (!videoRows || videoRows.length === 0) return;

  const sh = ss.getSheetByName(VIDEO_SHEET_NAME);
  if (!sh) return;

  const headers = sh.getRange(1, 1, 1, sh.getLastColumn()).getValues()[0];
  const thumbnailColIdx = headers.indexOf('Thumbnail');
  if (thumbnailColIdx === -1) return; // Nincs Thumbnail oszlop

  const videoIdColIdx = headers.indexOf('Video URL or ID / Asset ID');
  if (videoIdColIdx === -1) return;

  let updatedCount = 0;

  videoRows.forEach(item => {
    const { rawRow, sheetRowNum } = item;
    const videoInput = String(rawRow[videoIdColIdx] || '').trim();
    if (!videoInput) return;

    // Parse video identifier
    const parsed = parseVideoIdentifier(videoInput);
    let videoId = null;

    if (parsed.type === 'VIDEO_ID') {
      videoId = parsed.value;
    } else if (parsed.type === 'ASSET_ID') {
      // Asset ID esetén lekérdezzük a YouTube Video ID-t
      const details = fetchAssetDetails(parsed.value);
      if (details && details.youtubeVideoId) {
        videoId = details.youtubeVideoId;
      }
    }

    if (videoId) {
      const thumbnailUrl = `https://img.youtube.com/vi/${videoId}/mqdefault.jpg`;
      const youtubeUrl = `https://www.youtube.com/watch?v=${videoId}`;
      // FIX v7.4.6: HYPERLINK formula kattintható thumbnail-hez
      const formula = `=HYPERLINK("${youtubeUrl}"; IMAGE("${thumbnailUrl}"))`;
      const cell = sh.getRange(sheetRowNum, thumbnailColIdx + 1);

      // Csak akkor írjuk felül, ha üres vagy más érték van
      const currentValue = cell.getValue();
      if (!currentValue || !String(currentValue).includes(videoId)) {
        cell.setFormula(formula);
        updatedCount++;
      }
    }
  });

  if (updatedCount > 0) {
    console.log(`  🖼️ VideoAssets: ${updatedCount} thumbnail frissítve`);
  }
}

/**
 * Thumbnail oszlop frissítése az ImageAssets sheet-ben
 * Validáció után hívódik meg, beírja az =IMAGE() formulákat
 *
 * @param {Spreadsheet} ss - Google Spreadsheet objektum
 * @param {Array} imageRows - Validált image sorok
 */
function updateImageThumbnails(ss, imageRows) {
  if (!imageRows || imageRows.length === 0) return;

  const sh = ss.getSheetByName(IMAGE_SHEET_NAME);
  if (!sh) return;

  const headers = sh.getRange(1, 1, 1, sh.getLastColumn()).getValues()[0];
  const thumbnailColIdx = headers.indexOf('Thumbnail');
  if (thumbnailColIdx === -1) return; // Nincs Thumbnail oszlop

  const assetIdColIdx = headers.indexOf('Asset ID');
  if (assetIdColIdx === -1) return;

  let updatedCount = 0;

  imageRows.forEach(item => {
    const { rawRow, sheetRowNum } = item;
    const assetId = String(rawRow[assetIdColIdx] || '').trim();
    if (!assetId || !/^\d+$/.test(assetId)) return; // Csak számok

    // Lekérdezzük a kép URL-t
    const details = fetchAssetDetails(assetId);
    if (details && details.imageUrl) {
      // details.imageUrl = asset.image_asset.full_size.url (Google Ads API) → valódi képlink
      const formula = `=HYPERLINK("${details.imageUrl}"; IMAGE("${details.imageUrl}"))`;
      const cell = sh.getRange(sheetRowNum, thumbnailColIdx + 1);

      // Csak akkor írjuk felül, ha üres vagy más érték van
      const currentValue = cell.getValue();
      if (!currentValue || !String(currentValue).includes(assetId)) {
        cell.setFormula(formula);
        updatedCount++;
      }
    }
  });

  if (updatedCount > 0) {
    console.log(`  🖼️ ImageAssets: ${updatedCount} thumbnail frissítve`);
  }
}

/*** ===================== EMAIL ===================== ***/

/**
 * Kritikus hiba e-mail küldése
 * Akkor használjuk, ha a script leáll egy kritikus hiba miatt (pl. hiányzó sheet)
 */
function sendCriticalErrorEmail(errorMessage) {
  if (!NOTIFICATION_EMAIL || !NOTIFICATION_EMAIL.trim()) {
    console.log('⚠️ Nincs NOTIFICATION_EMAIL beállítva - kritikus hiba e-mail nem küldhető.');
    return;
  }

  const to = NOTIFICATION_EMAIL.split(',').map(s => s.trim()).filter(Boolean).join(',');
  if (!to) return;

  const accountName = AdsApp.currentAccount().getName();
  const tz = ACCOUNT_TIMEZONE;
  const nowStr = Utilities.formatDate(new Date(), tz, 'yyyy-MM-dd HH:mm');

  const subject = `[${accountName}] PMax Asset Scheduler - KRITIKUS HIBA`;

  const htmlBody = `
    <div style="font:14px Arial,sans-serif;color:#111">
      <h2 style="color:#d32f2f;">⛔ PMax Asset Scheduler - Kritikus Hiba</h2>
      <p><strong>Fiók:</strong> ${accountName}<br>
      <strong>Időpont:</strong> ${nowStr} (${tz})</p>

      <div style="background:#ffebee;border-left:4px solid #d32f2f;padding:12px;margin:16px 0;">
        <strong>Hiba:</strong> ${errorMessage}
      </div>

      <p>A script leállt és nem hajtotta végre a műveleteket. Kérjük, ellenőrizd a beállításokat!</p>

      <h3>Gyakori okok:</h3>
      <ul>
        <li>Hiányzó sheet (TextAssets, ImageAssets, VideoAssets)</li>
        <li>Hiányzó oszlopok a sheet-ekben</li>
        <li>Érvénytelen SPREADSHEET_URL</li>
      </ul>

      <p style="margin-top:16px;">
        <a href="${SPREADSHEET_URL}" target="_blank" style="background:#d32f2f;color:#fff;padding:10px 16px;border-radius:6px;text-decoration:none;font-weight:bold;">
          📊 Google Sheet megnyitása
        </a>
      </p>
    </div>
  `;

  try {
    MailApp.sendEmail({ to, subject, htmlBody });
    console.log(`📧 Kritikus hiba e-mail elküldve: ${to}`);
  } catch (e) {
    console.log(`❌ Kritikus hiba e-mail küldési hiba: ${e.message}`);
  }
}

function sendEmail(phase, previewRows, executionRows) {
  if (!NOTIFICATION_EMAIL || !NOTIFICATION_EMAIL.trim()) {
    console.log('⚠️ Nincs NOTIFICATION_EMAIL beállítva.');
    return;
  }

  const to = NOTIFICATION_EMAIL.split(',').map(s => s.trim()).filter(Boolean).join(',');
  if (!to) return;

  const accountName = AdsApp.currentAccount().getName();
  const tz = ACCOUNT_TIMEZONE;
  const nowStr = Utilities.formatDate(new Date(), tz, 'yyyy-MM-dd HH:mm');

  // Beszédesebb subject és fejléc
  let subject, emailTitle, htmlBody;

  if (phase === 'VALIDATION_PREVIEW') {
    subject = `[${accountName}] PMax Asset Scheduler - Preview Report`;
    emailTitle = 'Preview Report';

    // Video disclaimer detektálása
    const hasVideoRows = previewRows.some(row => {
      const assetType = String(row[3] || '').toUpperCase();
      return assetType.includes('VIDEO') || assetType.includes('YOUTUBE');
    });

    // YouTube API elérhetőség ellenőrzése a dinamikus disclaimerhez
    const youtubeApiAvailable = typeof YouTube !== 'undefined';

    const videoDisclaimer = hasVideoRows ? (youtubeApiAvailable ? `
      <div style="background:#d4edda;border-left:4px solid #28a745;padding:12px;margin:16px 0;">
        <strong>✅ Videó validáció aktív:</strong><br>
        • A YouTube API engedélyezve van - a videó <strong>létezését, elérhetőségét és hosszúságát</strong> ellenőrizzük<br>
      </div>
    ` : `
      <div style="background:#fff3cd;border-left:4px solid #ff9800;padding:12px;margin:16px 0;">
        <strong>⚠️ Videó validáció inaktív:</strong><br>
        • A YouTube API <strong>nincs engedélyezve</strong> - videó validáció kihagyva<br>
        • A Performance Max kampányok <strong>minimum 10 másodperc</strong> hosszúságú videókat igényelnek<br>
        • Rövid videók hozzáadása sikeres lehet, de a kampány <strong>nem fog futni</strong><br>
        • <strong>Ellenőrizd a Google Ads felületen</strong>, hogy a videók megfelelnek-e a követelményeknek!<br>
        • Engedélyezd a YouTube API-t: Script → Advanced APIs → YouTube
      </div>
    `) : '';

    htmlBody = `
      <div style="font:14px Arial,sans-serif;color:#111">
        <h2>PMax Asset Scheduler - ${emailTitle}</h2>
        <p><strong>Fiók:</strong> ${accountName}<br>
        <strong>Időpont:</strong> ${nowStr} (${tz})</p>

         <div style="background:#fff3cd;border-left:4px solid #ff9800;padding:12px;margin:16px 0;">
          <strong>⚠️ Béta:</strong> Ez az előnézet segít áttekinteni a tervezett műveleteket, csak a jövőbeli sorokat tartalmazza, soronként a legközelebbi műveletre ellenőriz.<br>
          • Ugyanarra az elemcsoportra irányuló több egymás utáni művelet együttes hatását nem tudja figyelembe venni.<br>
          • A végrehajtás nem garantált - a Google Ads fiókban történő változások (pl. kampány szüneteltetés, új elemek hozzáadása) befolyásolhatják a művelet sikerességét<br><br>
          Ellenőrizd gondosan a megadott adatokat és az ütemezést!
        </div>
        ${videoDisclaimer}

        <h3>Jövőbeli, első műveletek előnézete (${previewRows.length} sor)</h3>
        ${generateTable(previewRows.slice(0, EMAIL_ROW_LIMIT), true)}
        ${previewRows.length > EMAIL_ROW_LIMIT ? `<p>... +${previewRows.length - EMAIL_ROW_LIMIT} további sor</p>` : ''}

        <p style="margin-top:16px;">
          <a href="${SPREADSHEET_URL}" target="_blank" style="background:#1a73e8;color:#fff;padding:10px 16px;border-radius:6px;text-decoration:none;font-weight:bold;">
            📊 Google Sheet megnyitása
          </a>
        </p>
      </div>
    `;
  } else if (phase === 'EXECUTION_COMPLETE') {
    subject = `[${accountName}] PMax Asset Scheduler - Execution Report`;
    emailTitle = 'Execution Report';

    // KRITIKUS FIX: Ha executionRows üres, akkor már merge-elt sorok érkeztek!
    // (main()-ben mergeWindowAndExecutionRows() már meghívódott)
    const mergedRows = executionRows.length > 0
      ? mergeWindowAndExecutionRows(previewRows, executionRows)
      : previewRows;

    // Video disclaimer detektálása
    const hasVideoRows = mergedRows.some(row => {
      const assetType = String(row[3] || '').toUpperCase();
      return assetType.includes('VIDEO') || assetType.includes('YOUTUBE');
    });

    // YouTube API elérhetőség ellenőrzése a dinamikus disclaimerhez
    const youtubeApiAvailable = typeof YouTube !== 'undefined';

    const videoDisclaimer = hasVideoRows ? (youtubeApiAvailable ? `
      <div style="background:#d4edda;border-left:4px solid #28a745;padding:12px;margin:16px 0;">
        <strong>✅ Videó validáció aktív:</strong><br>
        • A YouTube API engedélyezve van - a videók létezését, elérhetőségét és hosszúságát ellenőriztük<br>
        • A validációs hibák az eredménytáblázatban láthatók
      </div>
    ` : `
      <div style="background:#fff3cd;border-left:4px solid #ff9800;padding:12px;margin:16px 0;">
        <strong>⚠️ Videó validáció inaktív:</strong><br>
        • A YouTube API <strong>nincs engedélyezve</strong> - videó validáció kihagyva<br>
        • A Performance Max kampányok <strong>minimum 10 másodperc</strong> hosszúságú videókat igényelnek<br>
        • <strong>Ellenőrizd a Google Ads felületen</strong>, hogy a videók megfelelnek-e a követelményeknek!<br>
        • Engedélyezd a YouTube API-t: Script → Advanced APIs → YouTube
      </div>
    `) : '';

    htmlBody = `
      <div style="font:14px Arial,sans-serif;color:#111">
        <h2>PMax Asset Scheduler - ${emailTitle}</h2>
        <p><strong>Fiók:</strong> ${accountName}<br>
        <strong>Időpont:</strong> ${nowStr} (${tz})</p>

        <h3>Időablakban lévő műveletek (${mergedRows.length} sor)</h3>
        <p style="color:#666;font-size:13px;margin-top:-8px;">
          ⚠️ ERROR státuszú sorok nem kerültek végrehajtásra. SUCCESS státuszú sorok végrehajtva és ellenőrizve.<br>
          ℹ️ <strong>Fontos:</strong> A Google Ads irányelveinek való megfelelést a hozzáadott elemeknél a rendszer a végrehajtás után ellenőrzi. A Google Ads által jóváhagyott módosításokat a hirdetési fiókban tudod ellenőrizni.
        </p>
        ${videoDisclaimer}
        ${generateTable(mergedRows.slice(0, EMAIL_ROW_LIMIT), false)}
        ${mergedRows.length > EMAIL_ROW_LIMIT ? `<p>... +${mergedRows.length - EMAIL_ROW_LIMIT} további sor</p>` : ''}

        <p style="margin-top:16px;">
          <a href="${SPREADSHEET_URL}" target="_blank" style="background:#1a73e8;color:#fff;padding:10px 16px;border-radius:6px;text-decoration:none;font-weight:bold;">
            📊 Google Sheet megnyitása
          </a>
        </p>
      </div>
    `;
  } else {
    // Fallback (régi viselkedés)
    subject = `[${accountName}] PMax Asset Scheduler - ${phase}`;
    emailTitle = phase;

    htmlBody = `
      <div style="font:14px Arial,sans-serif;color:#111">
        <h2>PMax Asset Scheduler - ${phase}</h2>
        <p><strong>Fiók:</strong> ${accountName}<br>
        <strong>Időpont:</strong> ${nowStr} (${tz})</p>

        <h3>Művelet eredmény (${previewRows.length} sor)</h3>
        ${generateTable(previewRows.slice(0, EMAIL_ROW_LIMIT), true)}
        ${previewRows.length > EMAIL_ROW_LIMIT ? `<p>... +${previewRows.length - EMAIL_ROW_LIMIT} további sor</p>` : ''}

        ${executionRows.length > 0 ? `
          <h3>Végrehajtás (${executionRows.length} művelet)</h3>
          ${generateTable(executionRows.slice(0, EMAIL_ROW_LIMIT), false)}
        ` : ''}

        <p style="margin-top:16px;">
          <a href="${SPREADSHEET_URL}" target="_blank" style="background:#1a73e8;color:#fff;padding:10px 16px;border-radius:6px;text-decoration:none;font-weight:bold;">
            📊 Google Sheet megnyitása
          </a>
        </p>
      </div>
    `;
  }

  try {
    MailApp.sendEmail({ to, subject, htmlBody });
    console.log(`📧 Email elküldve: ${to}`);
  } catch (e) {
    console.log(`❌ Email hiba: ${e.message}`);
  }
}

function mergeWindowAndExecutionRows(windowRows, verifiedResults) {
  // windowRows: időablakos sorok validációs státusszal (OK/WARNING/ERROR)
  // verifiedResults: végrehajtott sorok post-verification státusszal (SUCCESS/FAILED)

  // KRITIKUS FIX v7.3.25: action + hour kell a kulcsba (allScheduled formátum különbözik!)
  // windowRows (validation): allScheduled = "ADD: 2025-11-16 00:00-00:59 | REMOVE: 2025-11-16 22:00-22:59"
  // verifiedResults (execution): scheduled = "2025-11-16 22:00-22:59"
  const verifiedMap = new Map();
  verifiedResults.forEach(row => {
    const [timestamp, campaign, groupName, assetType, textOrId, scheduled, action] = row;
    const hours = extractHoursFromScheduled(scheduled, action);
    hours.forEach(hour => {
      const key = `${campaign}|${groupName}|${assetType}|${textOrId}|${action}|${hour}`;
      verifiedMap.set(key, row);
    });
  });

  // Merge: ERROR sorok maradnak, OK/WARNING sorok átveszik a végrehajtási státuszt
  return windowRows.map(row => {
    const status = String(row[7] || '');

    if (status === 'ERROR') {
      // ERROR sorok nem hajtódtak végre, maradnak validációs státusszal
      return row;
    }

    // OK vagy WARNING sorok: keressük meg a végrehajtási eredményt (action + hour alapján!)
    // KRITIKUS FIX v7.3.25: windowRows validation rows → allScheduled!
    const [timestamp, campaign, groupName, assetType, textOrId, allScheduled, action] = row;
    const hours = extractHoursFromScheduled(allScheduled, action);

    // Több hour is lehet (pl. ADD: 10:00-10:59 | REMOVE: 22:00-22:59), próbáljuk mindet
    for (const hour of hours) {
      const key = `${campaign}|${groupName}|${assetType}|${textOrId}|${action}|${hour}`;
      const verified = verifiedMap.get(key);

      if (verified) {
        // Végrehajtási státusz és üzenet felülírása
        return [...row.slice(0, 7), verified[7], verified[8]];
      }
    }

    // Fallback: nem kellene előfordulni, de ha mégis, marad az eredeti
    return row;
  });
}

function generateTable(rows, isPreview) {
  if (!rows || rows.length === 0) return '<p>Nincs adat.</p>';

  // Preview mode: "All Scheduled Actions" + "Next Action" oszlopok
  // Execution mode: "All Scheduled Actions" oszlop nincs (v7.5.2), csak "Executed Action"
  const headerRow = isPreview
    ? `<tr><th>Timestamp</th><th>Campaign</th><th>Asset Group</th><th>Asset Type</th><th>Text/Asset ID</th><th style="font-size:10px;">All Scheduled Actions</th><th>Next Action</th><th>Status</th><th>Message</th></tr>`
    : `<tr><th>Timestamp</th><th>Campaign</th><th>Asset Group</th><th>Asset Type</th><th>Text/Asset ID</th><th>Executed Action</th><th>Status</th><th>Message</th></tr>`;

  const dataRows = rows.map(r => {
    const status = String(r[7] || '');
    let bgColor;
    if (status === 'OK') bgColor = COLOR_OK;
    else if (status === 'WARNING') bgColor = COLOR_WARNING;
    else if (status === 'ERROR') bgColor = COLOR_ERROR;
    else if (status === 'SUCCESS') bgColor = COLOR_SUCCESS;
    else if (status === 'SUCCESS_WITH_WARNING') bgColor = COLOR_SUCCESS_WITH_WARNING;
    else if (status === 'FAILED') bgColor = COLOR_FAILED;
    else bgColor = 'white';

    // Preview mode: Next Action = action + időpont (kikeressük az allScheduledActions-ből a next action időpontját)
    let actionCell = r[6]; // default: action name
    if (isPreview && r[6] && r[5]) {
      const action = String(r[6]).toUpperCase(); // ADD / REMOVE / ADD+REMOVE
      const allScheduled = String(r[5]); // "ADD: 2025-11-17 10:00-10:59 | REMOVE: ..."

      // Keressük meg a megfelelő action időpontját az allScheduledActions-ből
      const actionUpper = action.replace(/\+/g, '\\+'); // Escape + for regex
      const match = allScheduled.match(new RegExp(`${actionUpper}:\\s*([\\d\\-:\\s]+)`, 'i'));
      if (match && match[1]) {
        actionCell = `${action} ${match[1].trim()}`;
      }
    }

    // FIX v7.4.6: Thumbnail + link generálás IMAGE és VIDEO asset-ekhez
    const assetType = String(r[3] || '').toLowerCase();
    const assetIdOrText = String(r[4] || '');
    let assetCell = escapeHtml(assetIdOrText);

    // IMAGE: szöveges link → tpc.googlesyndication.com (fetchAssetDetails)
    if (assetType.includes('image')) {
      const assetIdMatch = assetIdOrText.match(/(\d{10,})/);
      if (assetIdMatch) {
        const assetId = assetIdMatch[1];
        const imgDetails = fetchAssetDetails(assetId);
        if (imgDetails && imgDetails.imageUrl) {
          assetCell = `<a href="${imgDetails.imageUrl}" target="_blank">🖼️ ${escapeHtml(assetId)}</a>`;
        } else {
          assetCell = `🖼️ ${escapeHtml(assetId)}`;
        }
      }
    }

    // VIDEO: szöveges YouTube link
    if (assetType.includes('video')) {
      const parsed = parseVideoIdentifier(assetIdOrText);
      if (parsed.type === 'VIDEO_ID') {
        const videoId = parsed.value;
        const youtubeUrl = `https://www.youtube.com/watch?v=${videoId}`;
        assetCell = `<a href="${youtubeUrl}" target="_blank">▶️ ${escapeHtml(videoId)}</a>`;
      } else if (parsed.type === 'ASSET_ID') {
        assetCell = `Asset ID: ${escapeHtml(assetIdOrText)}`;
      }
    }

    // Execution módban (v7.5.2): "All Scheduled Actions" (r[5]) oszlop nélkül
    if (!isPreview) {
      return `<tr style="background:${bgColor}">
        <td>${escapeHtml(r[0])}</td>
        <td>${escapeHtml(r[1])}</td>
        <td>${escapeHtml(r[2])}</td>
        <td>${escapeHtml(r[3])}</td>
        <td>${assetCell}</td>
        <td>${escapeHtml(actionCell)}</td>
        <td><strong>${escapeHtml(r[7])}</strong></td>
        <td>${escapeHtml(r[8])}</td>
      </tr>`;
    }
    return `<tr style="background:${bgColor}">
      <td>${escapeHtml(r[0])}</td>
      <td>${escapeHtml(r[1])}</td>
      <td>${escapeHtml(r[2])}</td>
      <td>${escapeHtml(r[3])}</td>
      <td>${assetCell}</td>
      <td style="font-size:10px;">${escapeHtml(r[5])}</td>
      <td>${escapeHtml(actionCell)}</td>
      <td><strong>${escapeHtml(r[7])}</strong></td>
      <td>${escapeHtml(r[8])}</td>
    </tr>`;
  }).join('');

  return `<table cellpadding="6" cellspacing="0" border="1" style="border-collapse:collapse;font-size:13px;">
    <thead>${headerRow}</thead>
    <tbody>${dataRows}</tbody>
  </table>`;
}

function escapeHtml(s) {
  return String(s || '')
    .replace(/&/g, '&amp;')
    .replace(/</g, '&lt;')
    .replace(/>/g, '&gt;')
    .replace(/"/g, '&quot;');
}