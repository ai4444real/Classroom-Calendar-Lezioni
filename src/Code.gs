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
    .addItem('Crea Evento Calendario', 'createEventSelected')
    .addItem('Pubblica Materiale', 'publishMaterialSelected')
    .addSeparator()
    .addSubMenu(ui.createMenu('Setup & Test')
      .addItem('Crea fogli SSOT', 'setupSheets')
      .addItem('Test connessione API', 'testConnection')
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

  // Colora celle per le pubblicazioni riuscite
  colorMaterialCells_(allResults);

  // Mostra risultati
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

  ui.alert(`Pubblicazione completata (${lessons.length} lezioni):\n\n${message}`);
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

  try {
    // 1. Risolvi channel
    const channel = getChannel(targetKey);
    if (!channel) {
      result.error = `Target "${targetKey}" non trovato in Corsi`;
      return result;
    }

    const courseId = channel.classroom_course_id;
    if (!courseId) {
      result.error = `classroom_course_id mancante per "${targetKey}"`;
      return result;
    }

    // 2. Trova materiale esistente e il suo topic
    let existingMaterial = null;
    let topicId = null;
    const existingTarget = getLessonTarget(lesson.lesson_id, targetKey);

    if (existingTarget && existingTarget.classroom_material_id) {
      existingMaterial = findMaterialByMarker(courseId, lesson.lesson_id);
      topicId = existingTarget.topic_id;
    } else {
      // Fallback: cerca per marker
      existingMaterial = findMaterialByMarker(courseId, lesson.lesson_id);
    }

    // Se esiste, prendi il topicId dal materiale
    if (existingMaterial && existingMaterial.topicId) {
      topicId = existingMaterial.topicId;
    }

    // Se non abbiamo un topic, crealo/trovalo
    if (!topicId) {
      topicId = ensureTopic(courseId, lesson.topic);
    }

    // 3. Cancella materiale esistente (se esiste)
    if (existingMaterial) {
      deleteMaterial(courseId, existingMaterial.id);
      Logger.log(`Materiale esistente cancellato: ${existingMaterial.id}`);
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
    result.action = existingMaterial
      ? `ricreato con ${fileCount} file`
      : `creato con ${fileCount} file`;
    Logger.log(`Materiale pubblicato: ${lesson.lesson_id} / ${targetKey} → ${newMaterialId}`);

  } catch (e) {
    result.error = e.message;
    Logger.log(`Errore materiale ${lesson.lesson_id} / ${targetKey}: ${e.message}`);
  }

  return result;
}

/**
 * Colora le celle drive_folder_url per le pubblicazioni riuscite
 * @param {Object[]} allResults
 */
function colorMaterialCells_(allResults) {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(CONFIG.SHEETS.LESSONS);
  if (!sheet) return;

  const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
  const folderColIndex = headers.indexOf('drive_folder_url') + 1;

  if (folderColIndex === 0) return;

  for (const lr of allResults) {
    // Colora solo se almeno un target ha avuto successo
    const anySuccess = lr.results.some(r => r.success);
    if (anySuccess && lr.rowIndex) {
      sheet.getRange(lr.rowIndex, folderColIndex).setBackground(CONFIG.COLORS.SUCCESS);
    }
  }
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

  // Colora celle per gli eventi creati/aggiornati
  colorEventCells_(allResults);

  // Mostra risultati
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

  ui.alert(`Creazione eventi completata (${lessons.length} lezioni):\n\n${message}`);
}

/**
 * Colora le celle date, start_time, end_time per gli eventi creati
 * @param {Object[]} allResults
 */
function colorEventCells_(allResults) {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(CONFIG.SHEETS.LESSONS);
  if (!sheet) return;

  const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
  const dateColIndex = headers.indexOf('date') + 1;
  const startColIndex = headers.indexOf('start_time') + 1;
  const endColIndex = headers.indexOf('end_time') + 1;

  for (const lr of allResults) {
    // Colora solo se almeno un target ha avuto successo con "creato" o "aggiornato"
    const anyCreated = lr.results.some(r => r.success && (r.action === 'creato' || r.action === 'aggiornato'));
    if (anyCreated && lr.rowIndex) {
      if (dateColIndex > 0) {
        sheet.getRange(lr.rowIndex, dateColIndex).setBackground(CONFIG.COLORS.SUCCESS);
      }
      if (startColIndex > 0) {
        sheet.getRange(lr.rowIndex, startColIndex).setBackground(CONFIG.COLORS.SUCCESS);
      }
      if (endColIndex > 0) {
        sheet.getRange(lr.rowIndex, endColIndex).setBackground(CONFIG.COLORS.SUCCESS);
      }
    }
  }
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
