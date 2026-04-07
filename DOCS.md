# Classroom-Calendar Lezioni

**Scopo:** usare un Google Spreadsheet come SSOT per distribuire lezioni su Google Classroom e Google Calendar, mantenendo i materiali in una cartella Drive dedicata.

---

## Panoramica

Lo script legge le righe selezionate nel foglio `Lezioni` e, per ciascun target indicato:

- crea una cartella Drive per i materiali, se richiesta
- pubblica o ricrea un materiale su Google Classroom
- crea o aggiorna un evento su Google Calendar

L'idempotenza è gestita tramite:

- marker testuale nei materiali Classroom e nelle descrizioni eventi Calendar
- tabella tecnica `LessonTargets`, che salva gli ID esterni già creati

---

## Struttura Fogli

### Corsi

Mapping dei canali di distribuzione.

| Colonna | Descrizione |
|---------|-------------|
| `target_key` | Identificatore stabile interno, usato nel foglio `Lezioni` |
| `classroom_course_id` | ID numerico del corso Google Classroom |
| `calendar_id` | ID del calendario Google da usare per gli eventi |
| `folder` | Nome cartella corso da usare sotto la root Drive annuale |

Note:

- `target_key` non è il nome visibile del corso.
- `calendar_id` è opzionale: se manca, l’evento viene saltato senza errore.
- `folder` è usato sia per la struttura Drive sia come suffisso nel titolo evento sui target secondari.

### Lezioni

Elenco lezioni da pubblicare.

| Colonna | Descrizione |
|---------|-------------|
| `lesson_id` | Identificatore univoco logico della lezione |
| `docente` | Nome docente mostrato nel titolo evento |
| `tutor` | Nome tutor mostrato nel titolo evento |
| `titolo_evento` | Titolo base dell’evento calendario |
| `argomento` | Argomento della lezione; usato per Topic Classroom e descrizione evento |
| `data` | Data lezione |
| `ora_inizio` | Ora inizio |
| `ora_fine` | Ora fine |
| `destinatari` | Lista `target_key` separati da virgola |
| `url_cartella_drive` | URL cartella Drive contenente i materiali |
| `url_zoom` | URL Zoom opzionale |

Note:

- Se `titolo_evento` è vuoto, il titolo evento usa `argomento`.
- Se `url_cartella_drive` è vuoto, il materiale Classroom viene creato senza allegati.
- Se manca una tra `data`, `ora_inizio`, `ora_fine`, l’evento calendario viene saltato.

### LessonTargets

Tabella tecnica di supporto. Non va modificata manualmente salvo attività di recovery o debug.

| Colonna | Descrizione |
|---------|-------------|
| `lesson_id` | Riferimento a `Lezioni.lesson_id` |
| `target_key` | Riferimento a `Corsi.target_key` |
| `classroom_material_id` | ID del materiale Classroom creato |
| `calendar_event_id` | ID evento Calendar salvato dallo script |
| `topic_id` | ID del Topic Classroom associato |
| `published_at` | Timestamp ultimo aggiornamento materiale Classroom |

Uso pratico:

- evita duplicazioni
- permette update/ricreazione invece di creare nuovi oggetti
- conserva lo stato tecnico tra una run e la successiva

---

## Operazioni Disponibili

Menu personalizzato: `📚`

### Crea Cartella Drive

Crea la struttura cartelle per le righe selezionate del foglio `Lezioni`.

Struttura:

```text
Lezioni {anno} / {folder} / yyyymmdd
```

Comportamento attuale:

- se `CONFIG.DRIVE_ROOT_FOLDER_ID` è valorizzato, quello è trattato come root effettiva
- se `url_cartella_drive` è già presente, la riga viene saltata come già pronta
- il nome cartella corso arriva da `Corsi.folder`
- l’URL della cartella finale viene scritto in `Lezioni.url_cartella_drive`

### Crea Evento Calendario

Crea o aggiorna l’evento calendario per ogni target della lezione selezionata.

Comportamento attuale:

- titolo: `titolo_evento` oppure `argomento`, seguito da `[docente tutor]`
- per i target secondari aggiunge un suffisso con `folder`
- descrizione: `argomento`, eventuale riga `Zoom: ...`, poi marker `[LESSON_ID=...]`
- se esiste già un evento, lo aggiorna
- se non esiste, lo crea
- se `calendar_id` manca nel target, la pubblicazione evento viene saltata senza errore

Feedback visivo:

- colora di verde `data`, `ora_inizio`, `ora_fine` se almeno un target ha creato o aggiornato l’evento
- colora i singoli `destinatari` in verde o rosso con rich text

### Pubblica Materiale

Pubblica un materiale Google Classroom per ogni target della lezione selezionata.

Comportamento attuale:

- usa `argomento` come nome del Topic Classroom
- titolo materiale: data formattata `dd.MM.yyyy`
- descrizione materiale: solo marker `[LESSON_ID=...]`
- legge tutti i file presenti in `url_cartella_drive`
- crea il materiale come `DRAFT`, poi lo pubblica come `PUBLISHED`
- se trova materiali esistenti con lo stesso marker, li cancella tutti e ricrea da zero

Feedback visivo:

- colora di verde `url_cartella_drive` se almeno un target è andato a buon fine
- colora i singoli `destinatari` in verde o rosso con rich text

### Archivia lezioni passate

Operazione solo visiva sul foglio `Lezioni`.

Comportamento attuale:

- mette in grigio le righe con `data < oggi`
- lascia nere le righe da oggi in avanti
- non tocca i background, quindi i verdi di stato restano visibili

### Setup & Test

Voci disponibili:

- `Crea fogli SSOT`
- `Test connessione API`
- `Test permessi Classroom`
- `Test lettura dati`

Servono per bootstrap e diagnostica dei permessi Google.

---

## Logica di Idempotenza

### Marker

Ogni materiale Classroom e ogni evento Calendar contiene il marker:

```text
[LESSON_ID=...]
```

Questo consente di:

- riconoscere oggetti già pubblicati
- riallineare oggetti creati in versioni precedenti
- evitare duplicati in caso di rilancio della stessa lezione

### LessonTargets

Lo script salva gli ID esterni in `LessonTargets`:

- `classroom_material_id`
- `calendar_event_id`
- `topic_id`

Per Calendar, la lookup avviene prima tramite `calendar_event_id`, poi tramite marker come fallback.

Per Classroom, la lookup usa il marker e aggiorna `LessonTargets` col nuovo materiale creato.

---

## Google Classroom

### Topic

- Se esiste già un Topic con nome uguale ad `argomento`, viene riutilizzato.
- Se non esiste, viene creato.
- Se `argomento` è vuoto, la pubblicazione materiale viene bloccata.

### Ricreazione materiale

Gli allegati dei `CourseWorkMaterials` non vengono aggiornati in place. Per questo il flusso è:

1. cerca tutti i materiali esistenti con lo stesso marker, in stato `PUBLISHED` o `DRAFT`
2. li cancella
3. ricrea il materiale da zero con tutti gli allegati attuali
4. lo pubblica
5. aggiorna `LessonTargets`

Questo evita materiali orfani o bozze rimaste da run parziali.

### `topic_id` stale

È già gestito nel codice attuale un caso di recovery:

- se `LessonTargets.topic_id` è presente ma non più valido nel corso
- e la create fallisce con `invalid argument`
- lo script rigenera il Topic e ritenta la create

Questo riduce il problema documentato in passato su topic eliminati dal corso.

---

## Google Drive

### Materiali

I file restano in Drive. Classroom allega riferimenti ai file, non copie.

### Video con download bloccato

Per i file `video/*`, dopo la creazione del materiale lo script applica:

```javascript
downloadRestrictions.itemDownloadRestriction = {
  restrictedForWriters: true,
  restrictedForReaders: true
}
```

Questo blocca download, copia e stampa.

Ordine corretto delle operazioni:

1. crea il materiale Classroom con allegati
2. solo dopo blocca il download dei video
3. pubblica il materiale

Questo ordine è importante perché l’applicazione anticipata delle restrizioni può interferire con l’allegato da parte di Classroom.

---

## Google Calendar

### Titolo evento

Formato attuale:

```text
{titolo_evento oppure argomento} [docente tutor]
```

Per i target secondari:

```text
{titolo} [docente tutor] ({folder})
```

### Descrizione evento

Formato attuale:

```text
{argomento}

Zoom: {url_zoom}


[LESSON_ID=...]
```

Le sezioni opzionali vengono omesse se i dati mancano.

### Evento cancellato manualmente

Il progetto contiene già una mitigazione rispetto al problema storico:

- prima di usare `calendar_event_id`, lo script verifica se l’evento risulta ancora attivo con una chiamata HTTP alla Calendar API v3 filtrata per `iCalUID`
- se l’evento non risulta più attivo, non tenta l’update e passa alla ricreazione

Limitazione:

- il progetto non ha ancora configurato il servizio avanzato `Calendar` in `appsscript.json`
- la verifica attuale è fatta via `UrlFetchApp`, non tramite Advanced Service

Vedi anche `docs/known-issues.md` per il contesto del problema.

---

## Validazioni e Comportamenti

### Errori bloccanti

- `target_key` non trovato in `Corsi`
- `classroom_course_id` mancante per la pubblicazione Classroom
- `argomento` vuoto durante la creazione di un materiale
- `folder` mancante nel foglio `Corsi` durante la creazione cartella Drive

### Skip non bloccanti

- `calendar_id` mancante: evento saltato
- `data` o orari mancanti: evento saltato
- `url_cartella_drive` già presente: cartella considerata già pronta
- `url_cartella_drive` vuoto: materiale Classroom creato senza allegati

### Duplicati

Prima di operare sulle righe selezionate, lo script controlla eventuali `lesson_id` duplicati nella selezione corrente e blocca l’esecuzione.

Nota:

- questo controllo è limitato alle righe selezionate, non all’intero foglio

---

## File Progetto

```text
src/
├── appsscript.json
├── Code.gs
├── Config.gs
├── SheetService.gs
├── ClassroomService.gs
└── CalendarService.gs
```

Ruoli:

- `Code.gs`: menu, entry point, orchestrazione, feedback visivo
- `Config.gs`: costanti, nomi fogli, colori, root Drive
- `SheetService.gs`: lettura fogli, parsing target, persistenza `LessonTargets`
- `ClassroomService.gs`: Topic, CourseWorkMaterials, lettura cartelle Drive, restrizioni download video
- `CalendarService.gs`: creazione e update eventi, marker, verifica esistenza evento

---

## Configurazione

### Debug

In `Config.gs`:

```javascript
DEBUG: false
```

Se impostato a `true`, lo script mostra alert diagnostici aggiuntivi.

### Root Drive annuale

In `Config.gs`:

```javascript
DRIVE_ROOT_FOLDER_ID: '...'
```

Comportamento:

- se valorizzato, lo script usa direttamente quella cartella come root operativa
- il commento nel codice prevede che punti alla cartella annuale `Lezioni {anno}`
- se vuoto, lo script crea o usa `Lezioni {anno}` nel Drive dell’utente che esegue

Attenzione:

- se si cambia anno e si continua a usare una root esplicita, l’ID va aggiornato manualmente

---

## Autorizzazioni Richieste

Configurate in [src/appsscript.json](C:/Users/simone/Dropbox/pnlevolution/rebekko/AppScripts/src/appsscript.json):

- `https://www.googleapis.com/auth/classroom.courses.readonly`
- `https://www.googleapis.com/auth/classroom.courseworkmaterials`
- `https://www.googleapis.com/auth/classroom.topics`
- `https://www.googleapis.com/auth/calendar`
- `https://www.googleapis.com/auth/spreadsheets`
- `https://www.googleapis.com/auth/drive`
- `https://www.googleapis.com/auth/script.external_request`

Advanced Services attivi:

- `Classroom v1`
- `Drive v3`

---

## Comandi Clasp

```bash
clasp push --force
clasp pull
clasp open
```

---

## Problemi Noti

### `lesson_id` modificato dopo la pubblicazione

Se `lesson_id` cambia dopo che una lezione è già stata pubblicata:

- i marker non corrispondono più
- `LessonTargets` perde il legame logico
- una nuova esecuzione può produrre duplicati

Possibili mitigazioni future:

- protezione colonna
- validazione globale univocità
- generazione automatica ID

### Calendar Advanced Service non ancora abilitato

Il progetto usa già una verifica HTTP verso Calendar API, ma non ha ancora il servizio avanzato `Calendar` attivo nel manifest.

### Recovery manuale

In caso di stato tecnico incoerente, il foglio `LessonTargets` resta il punto di recovery più diretto:

- cancellare o svuotare la riga/cella tecnica interessata
- rieseguire la pubblicazione della lezione
