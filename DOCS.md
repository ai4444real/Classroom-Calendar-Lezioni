# Classroom-Calendar Lezioni

**Scopo:** Pubblicazione deterministica di lezioni su Google Classroom (Materiale) e Google Calendar, mantenendo materiale unico su Drive.

---

## Struttura Fogli (Google Sheet)

### Corsi
Mapping dei "canali di distribuzione" = corsi Classroom.

| Colonna | Descrizione |
|---------|-------------|
| `target_key` | Identificatore stabile (es. `PRACT_26A`, `EXAM_03`) |
| `classroom_course_id` | ID numerico del corso Classroom |
| `calendar_id` | ID del calendario (opzionale) |
| `label` | Nome leggibile (opzionale) |

**Nota:** `target_key` NON è il nome visibile del corso. È un codice interno stabile.

### Lezioni
Elenco lezioni da pubblicare.

| Colonna | Descrizione |
|---------|-------------|
| `lesson_id` | Identificatore univoco (es. `L01-2026-02-05`) |
| `topic` | Argomento/categoria (usato per Topic Classroom e descrizione evento) |
| `event_title` | Titolo evento calendario |
| `date` | Data lezione (YYYY-MM-DD) |
| `start_time` | Ora inizio (HH:MM) |
| `end_time` | Ora fine (HH:MM) |
| `targets` | Lista target separati da virgola (es. `PRACT_26A, EXAM_03`) |
| `drive_folder_url` | Link cartella Drive con materiali |
| `zoom_url` | Link Zoom (opzionale) |

### LessonTargets
Tabella di lavoro per idempotenza. **NON MODIFICARE MANUALMENTE.**

| Colonna | Descrizione |
|---------|-------------|
| `lesson_id` | Riferimento a Lezioni |
| `target_key` | Riferimento a Corsi |
| `classroom_material_id` | ID materiale creato su Classroom |
| `calendar_event_id` | ID evento creato su Calendar |
| `topic_id` | ID topic Classroom |
| `published_at` | Timestamp pubblicazione |

---

## Operazioni Disponibili (Menu "📚")

### Crea Evento Calendario
- Crea evento con titolo = `event_title`, descrizione = `topic` + link Zoom
- **Requisiti:** `date`, `start_time`, `end_time` tutti compilati
- Se manca `calendar_id` nel corso → salta silenziosamente
- Se manca `zoom_url` → evento senza link Zoom
- **Feedback visivo:** celle `date`, `start_time`, `end_time` diventano verdi

### Pubblica Materiale
- Crea Topic su Classroom (se non esiste)
- Crea Materiale con titolo = data (formato dd.MM.yyyy)
- Allega tutti i file dalla cartella Drive indicata in `drive_folder_url`
- **Video:** download automaticamente bloccato per gli studenti
- Se materiale già esiste (per marker): lo cancella e ricrea
- **Feedback visivo:** cella `drive_folder_url` diventa verde

---

## Decisioni Architetturali

### Idempotenza (Marker)
Ogni materiale/evento contiene un marker nascosto nella descrizione:
```
[LESSON_ID=xxx]
```
Lo script cerca questo marker per capire se esiste già → aggiorna invece di duplicare.

### Materiale unico su Drive
I file restano su Drive. Classroom contiene solo allegati (riferimenti), non copie.

### Video protetti
I file video (`mimeType: video/*`) hanno download/copia/stampa bloccati tramite `downloadRestrictions` (Drive API v3). Gli studenti possono solo visualizzare in streaming.

**Importante — `downloadRestrictions` vs `writersCanShare`:**
- `downloadRestrictions` (restrictedForWriters + restrictedForReaders): blocca download/copia/stampa per TUTTI. È il lucchetto effettivo. Viene applicato DOPO la creazione del materiale Classroom.
- `writersCanShare: false`: NON si usa. Blocca la condivisione da parte degli editor, ma impedisce anche a Classroom di allegare il file. Se applicato su file di altri proprietari, diventa irrevocabile (solo l'owner può toglierlo). Causa: errore "The caller does not have permission" sulla create del materiale.

**Ordine delle operazioni (critico):**
1. Crea materiale Classroom con allegati (Classroom condivide i file con gli studenti)
2. DOPO → applica `blockDownload_` sui video (blocca download)

Se invertito, `downloadRestrictions` potrebbe interferire con la capacità di Classroom di processare il file.

### Ricreazione materiale
L'API Classroom non permette di modificare gli allegati dopo la creazione. Quindi:
1. Cerca TUTTI i materiali esistenti per marker (PUBLISHED + DRAFT)
2. Li cancella tutti (previene bozze orfane)
3. Ricrea con tutti gli allegati aggiornati

L'ID materiale cambia, ma LessonTargets viene aggiornato.

**Bozze orfane:** Se una creazione precedente fallisce dopo il DRAFT ma prima del PUBLISH, rimane una bozza. La ricerca include `courseWorkMaterialStates: ['PUBLISHED', 'DRAFT']` per trovarle e cancellarle.

### Feedback visivo
Le celle vengono colorate di verde (#d9ead3) quando l'operazione ha successo:
- **Evento creato:** date, start_time, end_time
- **Materiale pubblicato:** drive_folder_url

### Campi opzionali
| Campo mancante | Comportamento |
|----------------|---------------|
| `calendar_id` | Salta creazione evento |
| `drive_folder_url` | Materiale senza allegati |
| `zoom_url` | Evento senza link Zoom |
| `start_time` o `end_time` | Salta creazione evento |

### Errori bloccanti
- `target_key` non trovato in Corsi → errore esplicito

---

## File Progetto

```
src/
├── appsscript.json      # Manifest (API, OAuth scopes)
├── Code.gs              # Menu, entry points, orchestrazione
├── Config.gs            # Costanti, nomi fogli, DEBUG flag
├── SheetService.gs      # Lettura/scrittura fogli
├── ClassroomService.gs  # API Classroom (Topic, Material)
├── CalendarService.gs   # API Calendar (Eventi)
```

---

## Comandi Clasp

```bash
clasp push --force   # Carica su Google
clasp pull           # Scarica da Google
clasp open           # Apre editor online
```

---

## Problemi Noti / TODO

### lesson_id modificabile
Se qualcuno modifica `lesson_id` dopo la pubblicazione:
- LessonTargets perde il collegamento
- Rieseguendo → duplicati

**Possibili soluzioni future:**
- Validazione duplicati pre-esecuzione
- Protezione colonna in Sheets
- Auto-generazione ID

### File di proprietà altrui
Se un file video nella cartella Drive è di proprietà di un altro utente:
- `blockDownload_` (solo `downloadRestrictions`) funziona se l'account ha accesso in modifica
- NON impostare mai `writersCanShare: false` su file non propri: diventa irrevocabile
- Se un file risulta già bloccato con `writersCanShare: false` da un giro precedente, solo il proprietario può sbloccarlo

---

## Debug

In `Config.gs`:
```javascript
DEBUG: true   // Mostra alert dettagliati
DEBUG: false  // Produzione
```

---

## Autorizzazioni Richieste (OAuth Scopes)

- `classroom.courses.readonly` - Leggere corsi
- `classroom.courseworkmaterials` - Creare/modificare materiali
- `classroom.topics` - Creare topic
- `calendar` - Creare/modificare eventi
- `spreadsheets` - Leggere/scrivere il foglio SSOT
- `drive` - Leggere cartelle, bloccare download video
