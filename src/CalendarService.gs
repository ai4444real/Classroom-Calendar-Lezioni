/**
 * Servizio per interazione con Google Calendar API
 */

/**
 * Costruisce il titolo dell'evento: "Titolo [docente tutor]" con suffisso opzionale
 * @param {Object} lesson
 * @param {string|null} titleSuffix
 * @returns {string}
 */
function buildEventTitle_(lesson, titleSuffix) {
  const base = lesson.titolo_evento || lesson.argomento;
  const docente = lesson.docente || '  ';
  const tutor = lesson.tutor || '  ';
  let title = `${base} [${docente} ${tutor}]`;
  if (titleSuffix) title += ` (${titleSuffix})`;
  return title;
}

/**
 * Verifica se un evento (identificato dall'iCalUID restituito da event.getId())
 * esiste ancora nel calendario e non è cancellato/nel cestino.
 * @param {string} calendarId
 * @param {string} iCalUID - da event.getId()
 * @returns {boolean}
 */
function isCalendarEventActive_(calendarId, iCalUID) {
  try {
    const token = ScriptApp.getOAuthToken();
    const url = 'https://www.googleapis.com/calendar/v3/calendars/'
      + encodeURIComponent(calendarId) + '/events'
      + '?iCalUID=' + encodeURIComponent(iCalUID) + '&maxResults=1';
    const res = UrlFetchApp.fetch(url, {
      headers: { 'Authorization': 'Bearer ' + token },
      muteHttpExceptions: true
    });
    const code = res.getResponseCode();
    if (code !== 200) {
      Logger.log('isCalendarEventActive_ HTTP ' + code + ' — assume attivo');
      return true; // su errore API, assume attivo (meglio aggiornare che duplicare)
    }
    const items = JSON.parse(res.getContentText()).items || [];
    // senza showDeleted=true, gli eventi cancellati non compaiono → items vuoto = cancellato
    return items.length > 0;
  } catch (e) {
    Logger.log('isCalendarEventActive_ errore: ' + e.message + ' — assume attivo');
    return true;
  }
}

/**
 * Costruisce la descrizione dell'evento (topic + Zoom + marker)
 * @param {Object} lesson
 * @returns {string}
 */
function buildEventDescription_(lesson) {
  const parts = [];

  // Topic come prima riga della descrizione
  if (lesson.argomento) {
    parts.push(lesson.argomento);
    parts.push('');
  }

  if (lesson.url_zoom) {
    parts.push(`Zoom: ${lesson.url_zoom}`);
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
  return !!(lesson.data && lesson.ora_inizio && lesson.ora_fine);
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
 * @param {string|null} titleSuffix - Testo da aggiungere tra parentesi (calendari secondari)
 * @returns {string} eventId
 */
function createCalendarEvent(calendarId, lesson, titleSuffix) {
  const calendar = CalendarApp.getCalendarById(calendarId);
  if (!calendar) {
    throw new Error(`Calendario non trovato: ${calendarId}`);
  }

  const title = buildEventTitle_(lesson, titleSuffix);
  const description = buildEventDescription_(lesson);

  // Parsa data e orari
  const startTime = parseDateTime_(lesson.data, lesson.ora_inizio);
  const endTime = parseDateTime_(lesson.data, lesson.ora_fine);

  // Crea evento con orario
  const event = calendar.createEvent(title, startTime, endTime, {
    description: description
  });

  const eventId = event.getId();
  Logger.log(`Evento creato: ${eventId} - "${title}" ${lesson.ora_inizio}-${lesson.ora_fine}`);
  return eventId;
}

/**
 * Aggiorna un evento esistente
 * @param {GoogleAppsScript.Calendar.CalendarEvent} event
 * @param {Object} lesson
 * @param {string|null} titleSuffix - Testo da aggiungere tra parentesi (calendari secondari)
 */
function updateCalendarEvent(event, lesson, titleSuffix) {
  const title = buildEventTitle_(lesson, titleSuffix);
  const description = buildEventDescription_(lesson);

  event.setTitle(title);
  event.setDescription(description);

  // Aggiorna orari
  const startTime = parseDateTime_(lesson.data, lesson.ora_inizio);
  const endTime = parseDateTime_(lesson.data, lesson.ora_fine);
  event.setTime(startTime, endTime);

  Logger.log(`Evento aggiornato: ${event.getId()}`);
}
