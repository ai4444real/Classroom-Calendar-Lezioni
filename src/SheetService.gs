/**
 * Servizio per lettura/scrittura dati dallo Sheet (SSOT)
 */

/**
 * Ottiene il foglio per nome, con gestione errore
 * @param {string} sheetName - Nome del foglio
 * @returns {GoogleAppsScript.Spreadsheet.Sheet}
 */
function getSheet_(sheetName) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(sheetName);
  if (!sheet) {
    throw new Error(`Foglio "${sheetName}" non trovato. Esegui setupSheets() per crearlo.`);
  }
  return sheet;
}

/**
 * Legge tutti i dati di un foglio come array di oggetti
 * @param {string} sheetName - Nome del foglio
 * @returns {Object[]} Array di oggetti con chiavi = intestazioni
 */
function getSheetData_(sheetName) {
  const sheet = getSheet_(sheetName);
  const data = sheet.getDataRange().getValues();

  if (data.length < 2) {
    return []; // Solo intestazioni o vuoto
  }

  const headers = data[0];
  const rows = data.slice(1);

  return rows.map((row, index) => {
    const obj = { _rowIndex: index + 2 }; // Riga effettiva nello sheet (1-based + header)
    headers.forEach((header, i) => {
      obj[header] = row[i];
    });
    return obj;
  });
}

// ============================================================
// CHANNELS
// ============================================================

/**
 * Ottiene tutti i canali (target)
 * @returns {Object[]} Array di {target_key, classroom_course_id, calendar_id, label, _rowIndex}
 */
function getChannels() {
  return getSheetData_(CONFIG.SHEETS.CHANNELS);
}

/**
 * Ottiene un canale per target_key
 * @param {string} targetKey - Es. 'PRACT_26A'
 * @returns {Object|null} Channel object o null se non trovato
 */
function getChannel(targetKey) {
  const channels = getChannels();
  return channels.find(ch => ch.target_key === targetKey) || null;
}

/**
 * Risolve un array di target_key in array di channel objects
 * @param {string[]} targetKeys - Array di target_key
 * @returns {Object[]} Array di {targetKey, channel} - channel può essere null se non trovato
 */
function resolveTargets(targetKeys) {
  const channels = getChannels();
  const channelMap = new Map(channels.map(ch => [ch.target_key, ch]));

  return targetKeys.map(targetKey => ({
    targetKey: targetKey,
    channel: channelMap.get(targetKey) || null
  }));
}

// ============================================================
// LESSONS
// ============================================================

/**
 * Ottiene tutte le lezioni
 * @returns {Object[]} Array di lesson objects
 */
function getLessons() {
  return getSheetData_(CONFIG.SHEETS.LESSONS);
}

/**
 * Ottiene una lezione per lesson_id
 * @param {string} lessonId - Es. 'L10-2026-02-05'
 * @returns {Object|null} Lesson object o null se non trovata
 */
function getLesson(lessonId) {
  const lessons = getLessons();
  return lessons.find(l => l.lesson_id === lessonId) || null;
}

/**
 * Ottiene la lezione dalla riga attualmente selezionata
 * @returns {Object|null} Lesson object o null se selezione non valida
 */
function getLessonBySelectedRow() {
  const lessons = getLessonsBySelectedRows();
  return lessons.length > 0 ? lessons[0] : null;
}

/**
 * Ottiene le lezioni dalle righe selezionate (supporta range)
 * @returns {Object[]} Array di Lesson objects (può essere vuoto)
 */
function getLessonsBySelectedRows() {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();

  // Verifica che siamo nel foglio Lezioni
  if (sheet.getName() !== CONFIG.SHEETS.LESSONS) {
    return [];
  }

  const range = SpreadsheetApp.getActiveRange();
  const startRow = range.getRow();
  const numRows = range.getNumRows();

  // Ignora riga intestazione
  if (startRow < 2) {
    return [];
  }

  const lessons = getLessons();
  const selectedLessons = [];

  for (let row = startRow; row < startRow + numRows; row++) {
    const lesson = lessons.find(l => l._rowIndex === row);
    // Salta righe vuote (senza lesson_id)
    if (lesson && lesson.lesson_id) {
      selectedLessons.push(lesson);
    }
  }

  return selectedLessons;
}

/**
 * Parsa il campo targets (CSV) in array
 * @param {string} targetsCSV - Es. 'PRACT_26A,EXAM_03'
 * @returns {string[]} Array di target_key
 */
function parseTargets(targetsCSV) {
  if (!targetsCSV || typeof targetsCSV !== 'string') {
    return [];
  }
  return targetsCSV
    .split(',')
    .map(t => t.trim())
    .filter(t => t.length > 0);
}

// ============================================================
// VALIDAZIONE
// ============================================================

/**
 * Trova lesson_id duplicati in un array di lezioni
 * @param {Object[]} lessons - Array di lesson objects
 * @returns {string[]} Array di lesson_id duplicati (vuoto se nessun duplicato)
 */
function findDuplicateLessonIds(lessons) {
  const seen = new Set();
  const duplicates = new Set();
  for (const lesson of lessons) {
    if (seen.has(lesson.lesson_id)) {
      duplicates.add(lesson.lesson_id);
    }
    seen.add(lesson.lesson_id);
  }
  return Array.from(duplicates);
}

// ============================================================
// LESSON TARGETS (mapping per idempotenza)
// ============================================================

/**
 * Ottiene tutti i LessonTargets
 * @returns {Object[]}
 */
function getLessonTargets() {
  return getSheetData_(CONFIG.SHEETS.LESSON_TARGETS);
}

/**
 * Cerca un LessonTarget esistente
 * @param {string} lessonId
 * @param {string} targetKey
 * @returns {Object|null}
 */
function getLessonTarget(lessonId, targetKey) {
  const all = getLessonTargets();
  return all.find(lt =>
    lt.lesson_id === lessonId && lt.target_key === targetKey
  ) || null;
}

/**
 * Salva o aggiorna un LessonTarget
 * @param {Object} data - {lesson_id, target_key, classroom_material_id, calendar_event_id, topic_id, published_pre_at, published_post_at}
 */
function saveLessonTarget(data) {
  const sheet = getSheet_(CONFIG.SHEETS.LESSON_TARGETS);
  const headers = CONFIG.HEADERS.LESSON_TARGETS;

  // Cerca se esiste già
  const existing = getLessonTarget(data.lesson_id, data.target_key);

  if (existing) {
    // Update: scrivi sulla riga esistente
    const rowValues = headers.map(h => data[h] !== undefined ? data[h] : existing[h]);
    sheet.getRange(existing._rowIndex, 1, 1, headers.length).setValues([rowValues]);
    Logger.log(`LessonTarget aggiornato: ${data.lesson_id} / ${data.target_key}`);
  } else {
    // Insert: nuova riga
    const rowValues = headers.map(h => data[h] || '');
    sheet.appendRow(rowValues);
    Logger.log(`LessonTarget creato: ${data.lesson_id} / ${data.target_key}`);
  }
}
