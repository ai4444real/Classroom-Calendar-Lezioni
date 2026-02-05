# Classroom-Calendar Lezioni

**Scopo:** Pubblicazione deterministica di lezioni su Google Classroom (Materiale) e Google Calendar, mantenendo materiale unico su Drive.

---

## Struttura Fogli (Google Sheet)

### Corsi (ex Channels)
Mapping dei "canali di distribuzione" = corsi Classroom.

| Colonna | Descrizione |
|---------|-------------|
| `target_key` | Identificatore stabile (es. `PRACT_26A`, `EXAM_03`) |
| `classroom_course_id` | ID numerico del corso Classroom |
| `calendar_id` | ID del calendario (opzionale) |
| `label` | Nome leggibile (opzionale) |

**Nota:** `target_key` NON è il nome visibile del corso. È un codice interno stabile.

### Lezioni (ex Lessons)
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
| `keypoints_doc_url` | Link al doc Keypoints (opzionale) |
| `drive_folder_url` | Link cartella Drive con materiali (opzionale) |
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
| `published_pre_at` | Timestamp pubblicazione PRE |
| `published_post_at` | Timestamp pubblicazione POST |

---

## Operazioni Disponibili (Menu "Lezioni")

### Crea Evento Calendario
- Crea evento con titolo = `event_title`, descrizione = `topic` + link Zoom
- **Requisiti:** `date`, `start_time`, `end_time` tutti compilati
- Se manca `calendar_id` nel corso → salta silenziosamente
- Se manca `zoom_url` → evento senza link Zoom

### Pubblica PRE (Keypoints)
- Crea Topic su Classroom (se non esiste)
- Crea Materiale con titolo = data, allegato = link Keypoints
- **Non richiede** `drive_folder_url`

### Pubblica POST (Materiale)
- **Cancella e ricrea** il materiale (l'API non permette di aggiungere allegati)
- Allega: Keypoints + tutti i file dalla cartella Drive
- **Video:** download automaticamente bloccato per gli studenti

---

## Decisioni Architetturali

### Idempotenza (Marker)
Ogni materiale/evento contiene un marker nascosto nella descrizione:
```
[LESSON_ID=xxx]
```
Lo script cerca questo marker per capire se esiste già → aggiorna invece di duplicare.

### Materiale unico su Drive
I file restano su Drive. Classroom contiene solo link/allegati, non copie.

### Video protetti
I file video (`mimeType: video/*`) hanno download/copia/stampa bloccati automaticamente. Gli studenti possono solo visualizzare. I docenti/proprietari possono comunque scaricare (limite Google).

### POST ricrea il materiale
L'API Classroom non permette di modificare gli allegati dopo la creazione. Quindi POST:
1. Trova materiale esistente (per marker)
2. Lo cancella
3. Lo ricrea con tutti gli allegati aggiornati

L'ID materiale cambia, ma LessonTargets viene aggiornato.

### Campi opzionali
| Campo mancante | Comportamento |
|----------------|---------------|
| `calendar_id` | Salta creazione evento |
| `keypoints_doc_url` | Materiale senza keypoints |
| `drive_folder_url` | POST crea materiale senza file extra |
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
