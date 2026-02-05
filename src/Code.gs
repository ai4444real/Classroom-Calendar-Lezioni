/**
 * Classroom-Calendar Lezioni Distribution
 * Main entry point
 *
 * @see specs/Classroom-Calendar_Distribuzione-Lezioni_SPEC.pdf
 */

function onOpen() {
  const ui = SpreadsheetApp.getUi();
  ui.createMenu('Lezioni')
    .addItem('Pubblica PRE (Keypoints)', 'publishPre')
    .addItem('Pubblica POST (Materiale)', 'publishPost')
    .addSeparator()
    .addItem('Test connessione', 'testConnection')
    .addToUi();
}

/**
 * Test connection to Classroom and Calendar APIs
 */
function testConnection() {
  try {
    // Test Classroom API
    const courses = Classroom.Courses.list({ pageSize: 1 });
    Logger.log('Classroom API OK - Corsi trovati: ' + (courses.courses ? courses.courses.length : 0));

    // Test Calendar API
    const calendars = CalendarApp.getAllCalendars();
    Logger.log('Calendar API OK - Calendari trovati: ' + calendars.length);

    Logger.log('=== CONNESSIONE OK ===');
    return true;
  } catch (e) {
    Logger.log('ERRORE: ' + e.message);
    throw e;
  }
}
