# Slide

Copertina 16:9 per presentare un progetto a un ente o a un cliente, a schermo o in PDF.

- **Struttura:** `.scs-slide` (contenitore di scala, riferimento 1920×1080) con `.scs-slide-logo` in alto a sinistra, `.scs-slide-cert` in alto a destra, `.scs-slide-art.scs-frames` (doppia cornice; dentro si può mettere un `<img>` del render chiave), `.scs-slide-words` (occhiello, titolo `display`, `lead`), `.scs-slide-meta` (committente, data, responsabile) e `.scs-band` in basso.
- **Il consumatore fornisce:** fase, titolo dell'intervento (max 2 righe), frase di sintesi (max 2 righe), tre dati meta, opzionale immagine chiave.
- **Temi:** `light` per PDF e stampa; `dark` per proiezione in sala. I loghi restano su piastra bianca.
- **Slide interne:** stesso margine (120px di riferimento), titolo `headline`, testo `body`/`lead`, banda bicolore in basso; numero pagina in `blue-ink` in basso a destra.
