# Backlog

## Da fare

<!-- Aggiungi qui le richieste degli utenti e le migliorie pianificate -->

## Noti / in sospeso

- **Advanced Calendar API** — aggiungere il servizio per rilevare eventi cancellati in modo affidabile (vedi `known-issues.md`)
- **topic_id non valido** — si è verificato su PRACTITIONER e TRAINING_AUTOGENO: il topic_id salvato in LessonTargets diventa stale (es. topic eliminato dal corso). Workaround: cancellare la riga in LessonTargets e ripubblicare. Miglioramento: in caso di "Request contains an invalid argument", rilevare automaticamente il topic_id non valido, svuotarlo e riprovare con un topic nuovo.
