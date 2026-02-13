/**
 * Servizio per interazione con Google Classroom API
 */

/**
 * Mostra alert di debug se DEBUG è attivo
 * @param {string} message
 */
function debug_(message) {
  Logger.log('[DEBUG] ' + message);
  if (CONFIG.DEBUG) {
    SpreadsheetApp.getUi().alert('DEBUG:\n\n' + message);
  }
}

/**
 * Costruisce il marker per idempotenza
 * @param {string} lessonId
 * @returns {string} Es: '[LESSON_ID=L01-2026-02-05]'
 */
function buildMarker(lessonId) {
  return CONFIG.MARKER_PREFIX + lessonId + CONFIG.MARKER_SUFFIX;
}

/**
 * Estrae il lesson_id da un testo contenente il marker
 * @param {string} text
 * @returns {string|null} lesson_id o null se non trovato
 */
function extractMarker(text) {
  if (!text) return null;
  const regex = /\[LESSON_ID=([^\]]+)\]/;
  const match = text.match(regex);
  return match ? match[1] : null;
}

// ============================================================
// TOPICS
// ============================================================

/**
 * Ottiene tutti i topic di un corso
 * @param {string} courseId
 * @returns {Object[]} Array di topic objects
 */
function getTopics_(courseId) {
  try {
    const response = Classroom.Courses.Topics.list(courseId);
    return response.topic || [];
  } catch (e) {
    Logger.log(`Errore getTopics per corso ${courseId}: ${e.message}`);
    return [];
  }
}

/**
 * Assicura che un topic esista, creandolo se necessario
 * @param {string} courseId
 * @param {string} topicName
 * @returns {string} topicId
 */
function ensureTopic(courseId, topicName) {
  // Sanitizza: rimuovi * iniziale e spazi superflui
  const cleanName = (topicName || '').replace(/^\*+/, '').trim();
  if (!cleanName) return null;

  // Cerca topic esistente
  const topics = getTopics_(courseId);
  const existing = topics.find(t => t.name === cleanName);

  if (existing) {
    Logger.log(`Topic "${topicName}" già esiste: ${existing.topicId}`);
    return existing.topicId;
  }

  // Crea nuovo topic
  const newTopic = Classroom.Courses.Topics.create({ name: cleanName }, courseId);
  Logger.log(`Topic "${cleanName}" creato: ${newTopic.topicId}`);
  return newTopic.topicId;
}

// ============================================================
// MATERIALS (CourseWorkMaterials)
// ============================================================

/**
 * Cerca tutti i materiali esistenti tramite marker nel corso (PUBLISHED + DRAFT)
 * @param {string} courseId
 * @param {string} lessonId
 * @returns {Object[]} Array di material objects (può essere vuoto)
 */
function findMaterialsByMarker(courseId, lessonId) {
  const marker = buildMarker(lessonId);
  const found = [];

  try {
    const response = Classroom.Courses.CourseWorkMaterials.list(courseId, {
      courseWorkMaterialStates: ['PUBLISHED', 'DRAFT']
    });
    const materials = response.courseWorkMaterial || [];

    for (const mat of materials) {
      if (mat.description && mat.description.includes(marker)) {
        found.push(mat);
      }
    }
    if (found.length > 0) {
      Logger.log(`Trovati ${found.length} materiali con marker ${marker}`);
    }
  } catch (e) {
    Logger.log(`Errore ricerca materiale: ${e.message}`);
  }

  return found;
}


/**
 * Formatta la data per il titolo del materiale
 * @param {Date|string} date
 * @returns {string} Es: '2026-02-05'
 */
function formatDateForTitle_(date) {
  if (date instanceof Date) {
    return Utilities.formatDate(date, 'Europe/Rome', 'dd.MM.yyyy');
  }
  // Se è già stringa, prova a convertirla
  const d = new Date(date);
  if (!isNaN(d.getTime())) {
    return Utilities.formatDate(d, 'Europe/Rome', 'dd.MM.yyyy');
  }
  return String(date);
}


/**
 * Estrae l'ID della cartella da un URL Drive
 * @param {string} folderUrl - es. https://drive.google.com/drive/folders/ABC123
 * @returns {string|null} folder ID o null
 */
function extractFolderId_(folderUrl) {
  if (!folderUrl) return null;
  const match = folderUrl.match(/folders\/([a-zA-Z0-9_-]+)/);
  return match ? match[1] : null;
}

/**
 * Blocca download/copia/stampa per un file (usato per video)
 * downloadRestrictions: blocca download/copia/stampa per tutti (inclusi editor)
 * @param {string} fileId
 */
function blockDownload_(fileId) {
  try {
    Drive.Files.update(
      {
        downloadRestrictions: {
          itemDownloadRestriction: {
            restrictedForWriters: true,
            restrictedForReaders: true
          }
        }
      },
      fileId
    );
    Logger.log(`Download bloccato per file: ${fileId}`);
  } catch (e) {
    Logger.log(`Errore blocco download: ${e.message}`);
  }
}


/**
 * Verifica se un file è un video
 * @param {string} mimeType
 * @returns {boolean}
 */
function isVideo_(mimeType) {
  return mimeType && mimeType.startsWith('video/');
}

/**
 * Legge tutti i file da una cartella Drive
 * Blocca automaticamente il download per i video
 * @param {string} folderUrl
 * @returns {Object[]} Array di {name, url, mimeType, id}
 */
function getFilesFromFolder(folderUrl) {
  const folderId = extractFolderId_(folderUrl);
  if (!folderId) {
    Logger.log(`Impossibile estrarre folder ID da: ${folderUrl}`);
    return [];
  }

  try {
    const folder = DriveApp.getFolderById(folderId);
    const files = folder.getFiles();
    const result = [];
    const debugInfo = [];

    while (files.hasNext()) {
      const file = files.next();
      const mimeType = file.getMimeType();
      const fileId = file.getId();
      const fileName = file.getName();

      debugInfo.push(`${fileName}\n  mimeType: ${mimeType}\n  isVideo: ${isVideo_(mimeType)}`);

      result.push({
        name: fileName,
        url: file.getUrl(),
        mimeType: mimeType,
        id: fileId
      });
    }

    // Ordina per nome
    result.sort((a, b) => a.name.localeCompare(b.name));
    Logger.log(`Trovati ${result.length} file nella cartella`);

    // Debug alert
    if (debugInfo.length > 0) {
      debug_(`File nella cartella:\n\n${debugInfo.join('\n\n')}`);
    }

    return result;
  } catch (e) {
    Logger.log(`Errore lettura cartella: ${e.message}`);
    debug_(`Errore lettura cartella: ${e.message}`);
    return [];
  }
}

/**
 * Cancella un materiale esistente
 * @param {string} courseId
 * @param {string} materialId
 */
function deleteMaterial(courseId, materialId) {
  try {
    Classroom.Courses.CourseWorkMaterials.remove(courseId, materialId);
    Logger.log(`Materiale cancellato: ${materialId}`);
  } catch (e) {
    Logger.log(`Errore cancellazione materiale: ${e.message}`);
    throw e;
  }
}

/**
 * Crea materiale con allegati dalla cartella Drive
 * @param {string} courseId
 * @param {string} topicId
 * @param {Object} lesson
 * @returns {string} materialId
 */
function createMaterialWithAttachments(courseId, topicId, lesson) {
  const title = formatDateForTitle_(lesson.date);
  const description = buildMarker(lesson.lesson_id);

  const material = {
    title: title,
    description: description,
    state: 'DRAFT',
    materials: []
  };
  if (topicId) material.topicId = String(topicId);

  // Aggiungi file dalla cartella Drive
  const videoFileIds = [];
  if (lesson.drive_folder_url) {
    const files = getFilesFromFolder(lesson.drive_folder_url);
    for (const file of files) {
      material.materials.push({
        driveFile: {
          driveFile: { id: file.id, title: file.name },
          shareMode: 'VIEW'
        }
      });
      if (isVideo_(file.mimeType)) {
        videoFileIds.push(file.id);
      }
    }
  }

  // Crea materiale come DRAFT
  const created = Classroom.Courses.CourseWorkMaterials.create(material, courseId);
  Logger.log(`Materiale creato come DRAFT con ${material.materials.length} allegati: ${created.id}`);

  // Blocca download sui video DOPO la creazione
  for (const videoId of videoFileIds) {
    blockDownload_(videoId);
  }

  // Pubblica
  Classroom.Courses.CourseWorkMaterials.patch(
    { state: 'PUBLISHED' },
    courseId,
    created.id,
    { updateMask: 'state' }
  );
  Logger.log(`Materiale pubblicato: ${created.id}`);
  return created.id;
}

