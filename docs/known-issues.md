# Known Issues

## Evento calendario cancellato manualmente e ripubblicato

### Problema

Se un evento viene cancellato dal Google Calendar (finisce nel cestino) e poi si ripubblica dalla tabella, lo script dice "aggiornato" ma l'evento non appare. L'evento viene "aggiornato" nel cestino senza essere ripristinato.

### Causa tecnica

`CalendarApp.getEventById()` restituisce un oggetto non-null anche per eventi soft-deleted (nel cestino per 30 giorni). Il metodo `updateCalendarEvent()` ci lavora sopra senza lanciare errori, ma l'evento rimane nel cestino e non è visibile.

`CalendarApp` non espone `event.status` (confirmed/cancelled), quindi non è possibile rilevare eventi cancellati con il servizio standard.

### Soluzione definitiva (non ancora implementata)

Aggiungere il servizio **Advanced Calendar API** (`Calendar.Events.get`) nei servizi dello script — come già fatto per Classroom. Questo espone `event.status === 'cancelled'` in modo affidabile.

Passo manuale richiesto: nell'editor Apps Script → **Servizi → Google Calendar API**.

### Workaround attuale

Aprire il foglio **LessonTargets**, trovare la riga corrispondente alla lezione (lesson_id + target_key), e **svuotare la cella `calendar_event_id`**. Al prossimo publish lo script non trova l'ID salvato e crea l'evento da zero.
