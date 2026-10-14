# Cartiglio

Colonna dati formali sul lato destro della tavola: numero, titolo, scala, formato, revisioni, progettista.

- **Struttura:** `.scs-cart` > `.scs-cblk` impilati (etichetta `.scs-label` + valore `.scs-value`); il numero in `.num` (`tav-number`, `cqy-navy`); la scala con `.acc` (`blue-ink`); revisioni in `table.scs-rev`, la riga corrente con `tr.cur` (punto `cqy-cyan`).
- **Il consumatore fornisce:** numero tavola a due cifre, titolo sintetico, scala, revisioni (codice, mm/aa, sigla), nome e iscrizione all'Ordine del progettista.
- **Sì:** l'ultimo blocco (progettista) va in fondo; una sola riga `cur`.
- **No:** numeri tavola senza zero iniziale («7»); più di quattro righe di revisione — archiviare le vecchie.
