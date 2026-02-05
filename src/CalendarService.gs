/**
 * Servizio per interazione con Google Calendar API
 */

/**
 * Costruisce la descrizione dell'evento (topic + Zoom + marker)
 * @param {Object} lesson
 * @returns {string}
 */
function buildEventDescription_(lesson) {
  const parts = [];

  // Topic come prima riga della descrizione
  if (lesson.topic) {
    parts.push(lesson.topic);
    parts.push('');
  }

  if (lesson.zoom_url) {
    parts.push(`Zoom: ${lesson.zoom_url}`);
  }

  // Marker per idempotenza (in fondo, con spazio)
  parts.push('');
  parts.push('');
  parts.push(buildMarker(lesson.lesson_id));

  return parts.join('\n');
}

/**
 * Verifica se una lezione ha i dati necessari per creare un evento
 * @param {Object} lesson
 * @returns {boolean}
 */
function canCreateEvent(lesson) {
  return !!(lesson.date && lesson.start_time && lesson.end_time);
}

/**
 * Parsa l'orario (HH:MM) e lo combina con una data
 * @param {Date|string} date
 * @param {string} timeStr - formato "HH:MM"
 * @returns {Date}
 */
function parseDateTime_(date, timeStr) {
  const baseDate = date instanceof Date ? new Date(date) : new Date(date);

  const [hours, minutes] = timeStr.split(':').map(Number);
  baseDate.setHours(hours, minutes, 0, 0);

  return baseDate;
}

/**
 * Cerca un evento esistente tramite marker nel calendario
 * @param {string} calendarId
 * @param {string} lessonId
 * @param {Date|string} date - per limitare la ricerca
 * @returns {GoogleAppsScript.Calendar.CalendarEvent|null}
 */
function findEventByMarker(calendarId, lessonId, date) {
  const marker = buildMarker(lessonId);

  try {
    const calendar = CalendarApp.getCalendarById(calendarId);
    if (!calendar) {
      Logger.log(`Calendario non trovato: ${calendarId}`);
      return null;
    }

    // Cerca eventi in un range di ±1 giorno dalla data
    const searchDate = date instanceof Date ? date : new Date(date);
    const startSearch = new Date(searchDate);
    startSearch.setDate(startSearch.getDate() - 1);
    const endSearch = new Date(searchDate);
    endSearch.setDate(endSearch.getDate() + 2);

    const events = calendar.getEvents(startSearch, endSearch);

    for (const event of events) {
      const desc = event.getDescription() || '';
      if (desc.includes(marker)) {
        Logger.log(`Evento trovato con marker ${marker}: ${event.getId()}`);
        return event;
      }
    }
  } catch (e) {
    Logger.log(`Errore ricerca evento: ${e.message}`);
  }

  return null;
}

/**
 * Crea un nuovo evento nel calendario
 * @param {string} calendarId
 * @param {Object} lesson
 * @returns {string} eventId
 */
function createCalendarEvent(calendarId, lesson) {
  const calendar = CalendarApp.getCalendarById(calendarId);
  if (!calendar) {
    throw new Error(`Calendario non trovato: ${calendarId}`);
  }

  const title = lesson.event_title || lesson.topic;
  const description = buildEventDescription_(lesson);

  // Parsa data e orari
  const startTime = parseDateTime_(lesson.date, lesson.start_time);
  const endTime = parseDateTime_(lesson.date, lesson.end_time);

  // Crea evento con orario
  const event = calendar.createEvent(title, startTime, endTime, {
    description: description
  });

  const eventId = event.getId();
  Logger.log(`Evento creato: ${eventId} - "${title}" ${lesson.start_time}-${lesson.end_time}`);
  return eventId;
}

/**
 * Aggiorna un evento esistente
 * @param {GoogleAppsScript.Calendar.CalendarEvent} event
 * @param {Object} lesson
 */
function updateCalendarEvent(event, lesson) {
  const title = lesson.event_title || lesson.topic;
  const description = buildEventDescription_(lesson);

  event.setTitle(title);
  event.setDescription(description);

  // Aggiorna orari
  const startTime = parseDateTime_(lesson.date, lesson.start_time);
  const endTime = parseDateTime_(lesson.date, lesson.end_time);
  event.setTime(startTime, endTime);

  Logger.log(`Evento aggiornato: ${event.getId()}`);
}
