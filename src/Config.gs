/**
 * Configurazione e costanti
 */

const CONFIG = {
  // Debug mode: mostra alert dettagliati
  DEBUG: false,
  // Nomi dei fogli nello Sheet
  SHEETS: {
    CHANNELS: 'Corsi',
    LESSONS: 'Lezioni',
    LESSON_TARGETS: 'LessonTargets'
  },

  // Intestazioni colonne per ogni foglio
  HEADERS: {
    CHANNELS: ['target_key', 'classroom_course_id', 'calendar_id', 'label'],
    LESSONS: ['lesson_id', 'topic', 'event_title', 'date', 'start_time', 'end_time', 'targets', 'drive_folder_url', 'zoom_url'],
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
