/**
 * Classroom-Calendar Lezioni Distribution
 * Main entry point
 */

// ============================================================
// FORMULE CUSTOM
// ============================================================

/**
 * Decodifica l'ID Classroom da base64 (quello nell'URL) a numero
 * Uso: =DECODE_CLASSROOM_ID(A1)
 * @param {string} encoded - ID codificato (es. "ODE4Njk5ODA2NDMw")
 * @returns {string} ID numerico (es. "818699806430")
 * @customfunction
 */
function DECODE_CLASSROOM_ID(encoded) {
  if (!encoded) return '';
  try {
    return Utilities.newBlob(Utilities.base64Decode(encoded)).getDataAsString();
  } catch (e) {
    return 'ERRORE: ' + e.message;
  }
}

/**
 * Crea il menu quando si apre lo Sheet
 */
function onOpen() {
  const ui = SpreadsheetApp.getUi();
  ui.createMenu('📚')
    .addItem('Crea Cartella Drive', 'createDriveFolderSelected')
    .addItem('Crea Evento Calendario', 'createEventSelected')
    .addItem('Pubblica Materiale', 'publishMaterialSelected')
    .addSeparator()
    .addSubMenu(ui.createMenu('Setup & Test')
      .addItem('Crea fogli SSOT', 'setupSheets')
      .addItem('Test connessione API', 'testConnection')
      .addItem('Test permessi Classroom', 'testClassroomPermissions')
      .addItem('Test lettura dati', 'testReadData'))
    .addToUi();
}

/**
 * Crea i fogli necessari se non esistono
 */
function setupSheets() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();

  Object.entries(CONFIG.SHEETS).forEach(([key, sheetName]) => {
    let sheet = ss.getSheetByName(sheetName);

    if (!sheet) {
      sheet = ss.insertSheet(sheetName);
      Logger.log(`Creato foglio: ${sheetName}`);
    }

    // Imposta intestazioni se foglio vuoto
    const headers = CONFIG.HEADERS[key];
    if (headers && sheet.getLastRow() === 0) {
      sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
      sheet.getRange(1, 1, 1, headers.length).setFontWeight('bold');
      sheet.setFrozenRows(1);
      Logger.log(`Intestazioni create per: ${sheetName}`);
    }
  });

  SpreadsheetApp.getUi().alert('Setup completato! Fogli creati: ' + Object.values(CONFIG.SHEETS).join(', '));
}

/**
 * Test connessione API Classroom e Calendar
 */
function testConnection() {
  const results = [];

  try {
    // Test Classroom API
    const courses = Classroom.Courses.list({ teacherId: 'me', pageSize: 100 });
    const courseCount = courses.courses ? courses.courses.length : 0;
    results.push(`✓ Classroom API OK - ${courseCount} corsi trovati`);

    if (courses.courses) {
      courses.courses.forEach(c => {
        results.push(`  - ${c.name} (ID: ${c.id})`);
      });
    }
  } catch (e) {
    results.push(`✗ Classroom API ERRORE: ${e.message}`);
  }

  try {
    // Test Calendar API
    const calendars = CalendarApp.getAllCalendars();
    results.push(`✓ Calendar API OK - ${calendars.length} calendari trovati`);
  } catch (e) {
    results.push(`✗ Calendar API ERRORE: ${e.message}`);
  }

  const message = results.join('\n');
  Logger.log(message);
  SpreadsheetApp.getUi().alert(message);
}

/**
 * Diagnostica permessi Classroom per i corsi nel foglio Corsi
 * Testa: get corso, list materiali, create materiale minimo
 */
function testClassroomPermissions() {
  const results = [];
  const channels = getChannels();

  // Mostra scope effettivi del token
  try {
    const token = ScriptApp.getOAuthToken();
    const resp = UrlFetchApp.fetch('https://www.googleapis.com/oauth2/v1/tokeninfo?access_token=' + token);
    const info = JSON.parse(resp.getContentText());
    results.push('SCOPE ATTIVI:');
    (info.scope || '').split(' ').forEach(s => results.push('  ' + s));
    results.push('');
  } catch (e) {
    results.push('Impossibile leggere scope: ' + e.message);
  }

  for (const ch of channels) {
    const courseId = String(ch.classroom_course_id);
    if (!courseId) continue;
    results.push(`--- ${ch.target_key} (corso: ${courseId}) ---`);

    // 1. Get corso
    try {
      const course = Classroom.Courses.get(courseId);
      results.push(`  ✓ GET corso: ${course.name} [stato: ${course.courseState}]`);
    } catch (e) {
      results.push(`  ✗ GET corso: ${e.message}`);
    }

    // 2. List materiali
    try {
      const mats = Classroom.Courses.CourseWorkMaterials.list(courseId);
      const count = (mats.courseWorkMaterial || []).length;
      results.push(`  ✓ LIST materiali: ${count} trovati`);
    } catch (e) {
      results.push(`  ✗ LIST materiali: ${e.message}`);
    }

    // 3. Create materiale minimo (DRAFT, solo titolo)
    try {
      const created = Classroom.Courses.CourseWorkMaterials.create(
        { title: '__TEST_DIAG__', state: 'DRAFT' },
        courseId
      );
      results.push(`  ✓ CREATE materiale OK (id: ${created.id})`);
      // Pulisci subito
      try {
        Classroom.Courses.CourseWorkMaterials.remove(courseId, created.id);
        results.push(`  ✓ DELETE test OK`);
      } catch (e2) {
        results.push(`  ⚠ DELETE test fallito: ${e2.message}`);
      }
    } catch (e) {
      results.push(`  ✗ CREATE materiale: ${e.message}`);
    }

    results.push('');
  }

  const message = results.join('\n');
  Logger.log(message);
  SpreadsheetApp.getUi().alert(message);
}



/**
 * Test lettura dati da fogli
 */
function testReadData() {
  const results = [];

  try {
    // Test Channels
    const channels = getChannels();
    results.push(`✓ Channels: ${channels.length} righe`);
    channels.forEach(ch => {
      results.push(`  - ${ch.target_key} → corso: ${ch.classroom_course_id}`);
    });
  } catch (e) {
    results.push(`✗ Channels ERRORE: ${e.message}`);
  }

  try {
    // Test Lessons
    const lessons = getLessons();
    results.push(`✓ Lessons: ${lessons.length} righe`);
    lessons.forEach(l => {
      const targets = parseTargets(l.targets);
      results.push(`  - ${l.lesson_id}: [${l.topic}] ${l.date} → targets: [${targets.join(', ')}]`);
    });
  } catch (e) {
    results.push(`✗ Lessons ERRORE: ${e.message}`);
  }

  // Test resolve targets (se ci sono lezioni)
  try {
    const lessons = getLessons();
    if (lessons.length > 0) {
      const firstLesson = lessons[0];
      const targetKeys = parseTargets(firstLesson.targets);
      const resolved = resolveTargets(targetKeys);

      results.push(`\n✓ Resolve targets per "${firstLesson.lesson_id}":`);
      resolved.forEach(r => {
        if (r.channel) {
          results.push(`  - ${r.targetKey} → corso: ${r.channel.classroom_course_id}, cal: ${r.channel.calendar_id}`);
        } else {
          results.push(`  - ${r.targetKey} → ✗ NON TROVATO in Channels!`);
        }
      });
    }
  } catch (e) {
    results.push(`✗ Resolve ERRORE: ${e.message}`);
  }

  const message = results.join('\n');
  Logger.log(message);
  SpreadsheetApp.getUi().alert(message);
}

/**
 * Pubblica Materiale per righe selezionate
 * Crea/ricrea Materiale con tutti i file dalla cartella Drive
 */
function publishMaterialSelected() {
  const ui = SpreadsheetApp.getUi();
  const lessons = getLessonsBySelectedRows();

  if (lessons.length === 0) {
    ui.alert('Seleziona una o più righe nel foglio "Lezioni" prima di pubblicare.');
    return;
  }

  const duplicates = findDuplicateLessonIds(lessons);
  if (duplicates.length > 0) {
    ui.alert(`Lezioni con ID duplicato: ${duplicates.join(', ')}\nCorreggi prima di procedere.`);
    return;
  }

  // Conferma se più di 10 righe
  if (lessons.length > 10) {
    const confirm = ui.alert(
      'Conferma',
      `Stai per pubblicare materiale per ${lessons.length} lezioni. Vuoi continuare?`,
      ui.ButtonSet.YES_NO
    );
    if (confirm !== ui.Button.YES) return;
  }

  const allResults = [];
  for (const lesson of lessons) {
    const results = publishMaterial(lesson);
    allResults.push({ lessonId: lesson.lesson_id, rowIndex: lesson._rowIndex, results: results });
  }

  // Feedback visivo + report
  showResults_(allResults, {
    successColumns: ['drive_folder_url'],
    successCondition: r => r.success,
    operationLabel: 'Pubblicazione'
  });
}

/**
 * Esegue pubblicazione Materiale per una lezione
 * @param {Object} lesson
 * @returns {Object[]} Array di risultati per ogni target
 */
function publishMaterial(lesson) {
  const targetKeys = parseTargets(lesson.targets);
  const results = [];

  if (targetKeys.length === 0) {
    return [{ targetKey: '(nessuno)', success: false, error: 'Nessun target specificato' }];
  }

  for (const targetKey of targetKeys) {
    const result = publishMaterialToTarget_(lesson, targetKey);
    results.push(result);
  }

  return results;
}

/**
 * Pubblica Materiale su un singolo target
 * Cancella materiale esistente e lo ricrea con tutti gli allegati
 * @param {Object} lesson
 * @param {string} targetKey
 * @returns {Object} Risultato
 */
function publishMaterialToTarget_(lesson, targetKey) {
  const result = { targetKey: targetKey, success: false };
  let courseId = null;

  try {
    // 1. Risolvi channel
    const channel = getChannel(targetKey);
    if (!channel) {
      result.error = `Target "${targetKey}" non trovato in Corsi`;
      return result;
    }

    courseId = channel.classroom_course_id;
    if (!courseId) {
      result.error = `classroom_course_id mancante per "${targetKey}"`;
      return result;
    }

    // 2. Trova materiali esistenti (PUBLISHED + DRAFT) e il topic
    let topicId = null;
    const existingTarget = getLessonTarget(lesson.lesson_id, targetKey);
    const existingMaterials = findMaterialsByMarker(courseId, lesson.lesson_id);

    if (existingTarget && existingTarget.topic_id) {
      topicId = existingTarget.topic_id;
    }

    // Prendi topicId dal primo materiale esistente (se presente)
    if (!topicId && existingMaterials.length > 0 && existingMaterials[0].topicId) {
      topicId = existingMaterials[0].topicId;
    }

    // Se non abbiamo un topic, crealo/trovalo
    if (!topicId) {
      topicId = ensureTopic(courseId, lesson.topic);
      if (!topicId) {
        result.error = 'Colonna "topic" vuota — inserisci un argomento prima di pubblicare';
        return result;
      }
    }

    // 3. Cancella TUTTI i materiali esistenti (published + draft orfani)
    for (const mat of existingMaterials) {
      deleteMaterial(courseId, mat.id);
      Logger.log(`Materiale esistente cancellato (${mat.state}): ${mat.id}`);
    }

    // 4. Crea nuovo materiale con tutti gli allegati dalla cartella
    const newMaterialId = createMaterialWithAttachments(courseId, topicId, lesson);

    // 5. Aggiorna LessonTargets con nuovo ID
    saveLessonTarget({
      lesson_id: lesson.lesson_id,
      target_key: targetKey,
      classroom_material_id: newMaterialId,
      topic_id: topicId,
      published_at: new Date().toISOString()
    });

    // Conta file allegati per il messaggio
    const fileCount = lesson.drive_folder_url ? getFilesFromFolder(lesson.drive_folder_url).length : 0;
    result.success = true;
    result.action = existingMaterials.length > 0
      ? `ricreato con ${fileCount} file`
      : `creato con ${fileCount} file`;
    Logger.log(`Materiale pubblicato: ${lesson.lesson_id} / ${targetKey} → ${newMaterialId}`);

  } catch (e) {
    const cId = courseId || '?';
    result.error = `${e.message} [corso: ${cId}]`;
    Logger.log(`Errore materiale ${lesson.lesson_id} / ${targetKey} (corso ${cId}): ${e.message}`);
  }

  return result;
}

// ============================================================
// CARTELLE DRIVE
// ============================================================

/**
 * Crea cartelle Drive per righe selezionate
 * Struttura: My Drive / Lezioni {anno} / {corso} / yyyymmdd
 */
function createDriveFolderSelected() {
  const ui = SpreadsheetApp.getUi();
  const lessons = getLessonsBySelectedRows();

  if (lessons.length === 0) {
    ui.alert('Seleziona una o più righe nel foglio "Lezioni" prima di creare le cartelle.');
    return;
  }

  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(CONFIG.SHEETS.LESSONS);
  const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
  const folderUrlColIndex = headers.indexOf('drive_folder_url') + 1;

  const allResults = [];
  for (const lesson of lessons) {
    const result = createDriveFolder_(lesson);

    // Scrivi URL nel foglio se successo e nuova cartella
    if (result.success && result.folderUrl && folderUrlColIndex > 0) {
      sheet.getRange(lesson._rowIndex, folderUrlColIndex).setValue(result.folderUrl);
    }

    allResults.push({
      lessonId: lesson.lesson_id,
      rowIndex: lesson._rowIndex,
      results: [{ targetKey: result.folderName || '—', success: result.success, action: result.action, error: result.error }]
    });
  }

  showResults_(allResults, {
    successColumns: ['drive_folder_url'],
    successCondition: r => r.success,
    operationLabel: 'Creazione cartelle'
  });
}

/**
 * Crea cartella Drive per una lezione
 * @param {Object} lesson
 * @returns {Object} Risultato con folderUrl, folderName, success, action/error
 */
function createDriveFolder_(lesson) {
  const result = { success: false };

  try {
    // Skip se drive_folder_url già presente
    if (lesson.drive_folder_url) {
      result.success = true;
      result.action = 'cartella già presente';
      return result;
    }

    if (!lesson.date) {
      result.error = 'Data mancante';
      return result;
    }

    // Primo target → nome cartella corso
    const targetKeys = parseTargets(lesson.targets);
    if (targetKeys.length === 0) {
      result.error = 'Nessun target specificato';
      return result;
    }

    const channel = getChannel(targetKeys[0]);
    if (!channel || !channel.folder) {
      result.error = `Colonna "folder" mancante per ${targetKeys[0]} nel foglio Corsi`;
      return result;
    }

    const courseFolderName = channel.folder;
    result.folderName = courseFolderName;

    // Anno dalla data della lezione
    const date = new Date(lesson.date);
    const year = date.getFullYear();
    const dateFolderName = formatDateForFolder_(date);

    // Struttura: root > Lezioni {anno} > {corso} > yyyymmdd
    const root = DriveApp.getRootFolder();
    const yearFolder = findOrCreateFolder_(root, 'Lezioni ' + year);
    const courseFolder = findOrCreateFolder_(yearFolder, courseFolderName);
    const dateFolder = findOrCreateFolder_(courseFolder, dateFolderName);

    result.folderUrl = dateFolder.getUrl();
    result.success = true;
    result.action = 'creata in ' + courseFolderName + '/' + dateFolderName;
    Logger.log('Cartella creata: ' + result.folderUrl);

  } catch (e) {
    result.error = e.message;
    Logger.log('Errore cartella ' + lesson.lesson_id + ': ' + e.message);
  }

  return result;
}

/**
 * Cerca una sottocartella per nome, la crea se non esiste
 * @param {GoogleAppsScript.Drive.Folder} parent
 * @param {string} name
 * @returns {GoogleAppsScript.Drive.Folder}
 */
function findOrCreateFolder_(parent, name) {
  const folders = parent.getFoldersByName(name);
  if (folders.hasNext()) {
    return folders.next();
  }
  return parent.createFolder(name);
}

/**
 * Formatta data come yyyymmdd per nome cartella
 * @param {Date|string} date
 * @returns {string}
 */
function formatDateForFolder_(date) {
  const d = new Date(date);
  const yyyy = d.getFullYear();
  const mm = String(d.getMonth() + 1).padStart(2, '0');
  const dd = String(d.getDate()).padStart(2, '0');
  return yyyy + mm + dd;
}

// ============================================================
// FEEDBACK VISIVO (consolidato)
// ============================================================

/**
 * Mostra risultati: colora celle di successo, rich text sui target, alert
 * Unico punto per tutto il feedback visivo post-operazione
 * @param {Object[]} allResults - Array di {lessonId, rowIndex, results}
 * @param {Object} options - {successColumns, successCondition, operationLabel}
 */
function showResults_(allResults, options) {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(CONFIG.SHEETS.LESSONS);
  if (!sheet) return;

  const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
  const targetsColIndex = headers.indexOf('targets') + 1;

  for (const lr of allResults) {
    if (!lr.rowIndex) continue;

    // 1. Rich text verde/rosso sui singoli target
    if (targetsColIndex > 0) {
      colorTargetsRichText_(sheet, lr.rowIndex, targetsColIndex, lr.results);
    }

    // 2. Sfondo verde sulle colonne di successo
    const anySuccess = lr.results.some(options.successCondition);
    if (anySuccess) {
      for (const colName of options.successColumns) {
        const colIndex = headers.indexOf(colName) + 1;
        if (colIndex > 0) {
          sheet.getRange(lr.rowIndex, colIndex).setBackground(CONFIG.COLORS.SUCCESS);
        }
      }
    }
  }

  // 3. Alert riepilogativo
  const message = allResults.map(lr => {
    const details = lr.results.map(r => {
      if (r.success) {
        return `  ✓ ${r.targetKey}: ${r.action}`;
      } else {
        return `  ✗ ${r.targetKey}: ${r.error}`;
      }
    }).join('\n');
    return `${lr.lessonId}:\n${details}`;
  }).join('\n\n');

  SpreadsheetApp.getUi().alert(
    `${options.operationLabel} completata (${allResults.length} lezioni):\n\n${message}`
  );
}

/**
 * Applica rich text alla cella targets: verde per successo, rosso per errore
 * @param {GoogleAppsScript.Spreadsheet.Sheet} sheet
 * @param {number} rowIndex - Riga della lezione
 * @param {number} colIndex - Colonna targets
 * @param {Object[]} results - Risultati per target [{targetKey, success}, ...]
 */
function colorTargetsRichText_(sheet, rowIndex, colIndex, results) {
  const cell = sheet.getRange(rowIndex, colIndex);
  const text = cell.getValue().toString();
  if (!text) return;

  const builder = SpreadsheetApp.newRichTextValue().setText(text);
  const greenStyle = SpreadsheetApp.newTextStyle().setForegroundColor('#006100').setBold(true).build();
  const redStyle = SpreadsheetApp.newTextStyle().setForegroundColor('#cc0000').setBold(true).build();

  for (const r of results) {
    const start = text.indexOf(r.targetKey);
    if (start < 0) continue;
    const end = start + r.targetKey.length;
    builder.setTextStyle(start, end, r.success ? greenStyle : redStyle);
  }

  cell.setRichTextValue(builder.build());
}

// ============================================================
// EVENTI CALENDARIO
// ============================================================

/**
 * Crea evento calendario per righe selezionate
 */
function createEventSelected() {
  const ui = SpreadsheetApp.getUi();
  const lessons = getLessonsBySelectedRows();

  if (lessons.length === 0) {
    ui.alert('Seleziona una o più righe nel foglio "Lezioni" prima di creare gli eventi.');
    return;
  }

  const duplicates = findDuplicateLessonIds(lessons);
  if (duplicates.length > 0) {
    ui.alert(`Lezioni con ID duplicato: ${duplicates.join(', ')}\nCorreggi prima di procedere.`);
    return;
  }

  // Conferma se più di 10 righe
  if (lessons.length > 10) {
    const confirm = ui.alert(
      'Conferma',
      `Stai per creare eventi per ${lessons.length} lezioni. Vuoi continuare?`,
      ui.ButtonSet.YES_NO
    );
    if (confirm !== ui.Button.YES) return;
  }

  const allResults = [];
  for (const lesson of lessons) {
    const results = createEvents(lesson);
    allResults.push({ lessonId: lesson.lesson_id, rowIndex: lesson._rowIndex, results: results });
  }

  // Feedback visivo + report
  showResults_(allResults, {
    successColumns: ['date', 'start_time', 'end_time'],
    successCondition: r => r.success && (r.action === 'creato' || r.action === 'aggiornato'),
    operationLabel: 'Creazione eventi'
  });
}

/**
 * Crea eventi per una lezione su tutti i target
 * @param {Object} lesson
 * @returns {Object[]} Array di risultati per ogni target
 */
function createEvents(lesson) {
  const targetKeys = parseTargets(lesson.targets);
  const results = [];

  if (targetKeys.length === 0) {
    return [{ targetKey: '(nessuno)', success: false, error: 'Nessun target specificato' }];
  }

  for (const targetKey of targetKeys) {
    const result = createEventToTarget_(lesson, targetKey);
    results.push(result);
  }

  return results;
}

/**
 * Crea evento su un singolo target
 * @param {Object} lesson
 * @param {string} targetKey
 * @returns {Object} Risultato
 */
function createEventToTarget_(lesson, targetKey) {
  const result = { targetKey: targetKey, success: false };

  try {
    // 0. Verifica dati obbligatori (date, start_time, end_time)
    if (!canCreateEvent(lesson)) {
      result.success = true;
      result.action = 'saltato (manca date/start_time/end_time)';
      return result;
    }

    // 1. Risolvi channel
    const channel = getChannel(targetKey);
    if (!channel) {
      result.error = `Target "${targetKey}" non trovato in Channels`;
      return result;
    }

    const calendarId = channel.calendar_id;
    if (!calendarId) {
      // Calendar opzionale: salta senza errore
      result.success = true;
      result.action = 'saltato (nessun calendar_id)';
      return result;
    }

    // 2. Cerca evento esistente per marker (idempotenza)
    const existingEvent = findEventByMarker(calendarId, lesson.lesson_id, lesson.date);
    if (existingEvent) {
      updateCalendarEvent(existingEvent, lesson);
      // Salva/aggiorna mapping
      saveLessonTarget({
        lesson_id: lesson.lesson_id,
        target_key: targetKey,
        calendar_event_id: existingEvent.getId()
      });
      result.success = true;
      result.action = 'aggiornato';
      return result;
    }

    // 3. Crea nuovo evento
    const eventId = createCalendarEvent(calendarId, lesson);

    // 4. Salva mapping
    saveLessonTarget({
      lesson_id: lesson.lesson_id,
      target_key: targetKey,
      calendar_event_id: eventId
    });

    result.success = true;
    result.action = 'creato';
    Logger.log(`Evento creato: ${lesson.lesson_id} / ${targetKey} → ${eventId}`);

  } catch (e) {
    result.error = e.message;
    Logger.log(`Errore evento ${lesson.lesson_id} / ${targetKey}: ${e.message}`);
  }

  return result;
}
