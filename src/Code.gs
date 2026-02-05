/**
 * Classroom-Calendar Lezioni Distribution
 * Main entry point
 */

/**
 * Crea il menu quando si apre lo Sheet
 */
function onOpen() {
  const ui = SpreadsheetApp.getUi();
  ui.createMenu('Lezioni')
    .addItem('Crea Evento Calendario', 'createEventSelected')
    .addSeparator()
    .addItem('Pubblica PRE (Keypoints)', 'publishPreSelected')
    .addItem('Pubblica POST (Materiale)', 'publishPostSelected')
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
 * Pubblica PRE per riga selezionata
 * Crea Topic (se non esiste) + Materiale con keypoints su Classroom
 */
function publishPreSelected() {
  const ui = SpreadsheetApp.getUi();
  const lesson = getLessonBySelectedRow();

  if (!lesson) {
    ui.alert('Seleziona una riga nel foglio "Lessons" prima di pubblicare.');
    return;
  }

  const results = publishPre(lesson);

  // Mostra risultati
  const message = results.map(r => {
    if (r.success) {
      return `✓ ${r.targetKey}: materiale ${r.action} (ID: ${r.materialId})`;
    } else {
      return `✗ ${r.targetKey}: ${r.error}`;
    }
  }).join('\n');

  ui.alert(`Pubblicazione PRE completata:\n\n${message}`);
}

/**
 * Esegue pubblicazione PRE per una lezione
 * @param {Object} lesson
 * @returns {Object[]} Array di risultati per ogni target
 */
function publishPre(lesson) {
  const targetKeys = parseTargets(lesson.targets);
  const results = [];

  if (targetKeys.length === 0) {
    Logger.log(`Nessun target per lezione ${lesson.lesson_id}`);
    return [{ targetKey: '(nessuno)', success: false, error: 'Nessun target specificato' }];
  }

  for (const targetKey of targetKeys) {
    const result = publishPreToTarget_(lesson, targetKey);
    results.push(result);
  }

  return results;
}

/**
 * Pubblica PRE su un singolo target
 * @param {Object} lesson
 * @param {string} targetKey
 * @returns {Object} Risultato
 */
function publishPreToTarget_(lesson, targetKey) {
  const result = { targetKey: targetKey, success: false };

  try {
    // 1. Risolvi channel
    const channel = getChannel(targetKey);
    if (!channel) {
      result.error = `Target "${targetKey}" non trovato in Channels`;
      Logger.log(result.error);
      return result;
    }

    const courseId = channel.classroom_course_id;
    if (!courseId) {
      result.error = `classroom_course_id mancante per "${targetKey}"`;
      Logger.log(result.error);
      return result;
    }

    // 2. Verifica se già pubblicato (idempotenza)
    const existingTarget = getLessonTarget(lesson.lesson_id, targetKey);
    if (existingTarget && existingTarget.classroom_material_id) {
      // Già esiste, verifica che sia ancora presente su Classroom
      const existingMaterial = findMaterialByMarker(courseId, lesson.lesson_id);
      if (existingMaterial) {
        result.success = true;
        result.action = 'già esistente';
        result.materialId = existingTarget.classroom_material_id;
        Logger.log(`PRE già pubblicato per ${lesson.lesson_id} / ${targetKey}`);
        return result;
      }
    }

    // 3. Cerca materiale per marker (fallback se LessonTargets non aggiornato)
    const foundMaterial = findMaterialByMarker(courseId, lesson.lesson_id);
    if (foundMaterial) {
      // Salva mapping e ritorna
      saveLessonTarget({
        lesson_id: lesson.lesson_id,
        target_key: targetKey,
        classroom_material_id: foundMaterial.id,
        topic_id: foundMaterial.topicId || '',
        published_pre_at: new Date().toISOString()
      });
      result.success = true;
      result.action = 'trovato esistente';
      result.materialId = foundMaterial.id;
      return result;
    }

    // 4. Crea topic se necessario
    const topicId = ensureTopic(courseId, lesson.topic);

    // 5. Crea materiale
    const materialId = createMaterial(courseId, topicId, lesson);

    // 6. Salva mapping
    saveLessonTarget({
      lesson_id: lesson.lesson_id,
      target_key: targetKey,
      classroom_material_id: materialId,
      topic_id: topicId,
      published_pre_at: new Date().toISOString()
    });

    result.success = true;
    result.action = 'creato';
    result.materialId = materialId;
    Logger.log(`PRE pubblicato: ${lesson.lesson_id} / ${targetKey} → ${materialId}`);

  } catch (e) {
    result.error = e.message;
    Logger.log(`Errore PRE ${lesson.lesson_id} / ${targetKey}: ${e.message}`);
  }

  return result;
}

/**
 * Pubblica POST per riga selezionata
 * Aggiorna Materiale esistente aggiungendo drive_folder_url
 */
function publishPostSelected() {
  const ui = SpreadsheetApp.getUi();
  const lesson = getLessonBySelectedRow();

  if (!lesson) {
    ui.alert('Seleziona una riga nel foglio "Lessons" prima di pubblicare.');
    return;
  }

  const results = publishPost(lesson);

  // Mostra risultati
  const message = results.map(r => {
    if (r.success) {
      return `✓ ${r.targetKey}: ${r.action}`;
    } else {
      return `✗ ${r.targetKey}: ${r.error}`;
    }
  }).join('\n');

  ui.alert(`Pubblicazione POST completata:\n\n${message}`);
}

/**
 * Esegue pubblicazione POST per una lezione
 * @param {Object} lesson
 * @returns {Object[]} Array di risultati per ogni target
 */
function publishPost(lesson) {
  const targetKeys = parseTargets(lesson.targets);
  const results = [];

  if (targetKeys.length === 0) {
    return [{ targetKey: '(nessuno)', success: false, error: 'Nessun target specificato' }];
  }

  for (const targetKey of targetKeys) {
    const result = publishPostToTarget_(lesson, targetKey);
    results.push(result);
  }

  return results;
}

/**
 * Pubblica POST su un singolo target
 * Cancella materiale esistente e lo ricrea con tutti gli allegati
 * @param {Object} lesson
 * @param {string} targetKey
 * @returns {Object} Risultato
 */
function publishPostToTarget_(lesson, targetKey) {
  const result = { targetKey: targetKey, success: false };

  try {
    // 1. Risolvi channel
    const channel = getChannel(targetKey);
    if (!channel) {
      result.error = `Target "${targetKey}" non trovato in Channels`;
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

    // 4. Crea nuovo materiale con tutti gli allegati
    const newMaterialId = createMaterialWithAttachments(courseId, topicId, lesson);

    // 5. Aggiorna LessonTargets con nuovo ID
    saveLessonTarget({
      lesson_id: lesson.lesson_id,
      target_key: targetKey,
      classroom_material_id: newMaterialId,
      topic_id: topicId,
      published_post_at: new Date().toISOString()
    });

    // Conta file allegati per il messaggio
    const fileCount = lesson.drive_folder_url ? getFilesFromFolder(lesson.drive_folder_url).length : 0;
    result.success = true;
    result.action = existingMaterial
      ? `ricreato con ${fileCount} file`
      : `creato con ${fileCount} file`;
    Logger.log(`POST pubblicato: ${lesson.lesson_id} / ${targetKey} → ${newMaterialId}`);

  } catch (e) {
    result.error = e.message;
    Logger.log(`Errore POST ${lesson.lesson_id} / ${targetKey}: ${e.message}`);
  }

  return result;
}

// ============================================================
// EVENTI CALENDARIO
// ============================================================

/**
 * Crea evento calendario per riga selezionata
 */
function createEventSelected() {
  const ui = SpreadsheetApp.getUi();
  const lesson = getLessonBySelectedRow();

  if (!lesson) {
    ui.alert('Seleziona una riga nel foglio "Lessons" prima di creare l\'evento.');
    return;
  }

  const results = createEvents(lesson);

  // Mostra risultati
  const message = results.map(r => {
    if (r.success) {
      return `✓ ${r.targetKey}: evento ${r.action}`;
    } else {
      return `✗ ${r.targetKey}: ${r.error}`;
    }
  }).join('\n');

  ui.alert(`Creazione eventi completata:\n\n${message}`);
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
