/**
 * Configurazione e costanti
 */

const CONFIG = {
  // Debug mode: mostra alert dettagliati
  DEBUG: false,

  // ID della cartella "Lezioni {anno}" nel Mio Drive di X.
  // Lo script la usa direttamente come cartella anno, quindi crea {corso}/yyyymmdd al suo interno.
  // Aggiornare ogni anno con l'ID della nuova cartella "Lezioni {anno}".
  // Se vuoto: usa il Mio Drive dell'utente che esegue (comportamento originale).
  DRIVE_ROOT_FOLDER_ID: '1NB4UFzVEKIeVYa6tMUn8J9ELjyFNrM6Q',

  // Nomi dei fogli nello Sheet
  SHEETS: {
    CHANNELS: 'Corsi',
    LESSONS: 'Lezioni',
    LESSON_TARGETS: 'LessonTargets'
  },

  // Intestazioni colonne per ogni foglio
  HEADERS: {
    CHANNELS: ['target_key', 'classroom_course_id', 'calendar_id', 'folder'],
    LESSONS: ['lesson_id', 'docente', 'tutor', 'titolo_evento', 'argomento', 'data', 'ora_inizio', 'ora_fine', 'destinatari', 'url_cartella_drive', 'url_zoom'],
    LESSON_TARGETS: ['lesson_id', 'target_key', 'classroom_material_id', 'calendar_event_id', 'topic_id', 'published_at']
  },

  // Marker per idempotenza
  MARKER_PREFIX: '[LESSON_ID=',
  MARKER_SUFFIX: ']',

  // Colori per status (verde chiaro)
  COLORS: {
    SUCCESS: '#d9ead3'
  }
};
