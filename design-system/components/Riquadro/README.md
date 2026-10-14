# Riquadro

Cella della griglia della tavola (A–F): testata con lettera, titolo e scala, poi l'elaborato.

- **Struttura:** `.scs-cell` > `.scs-cell-hdr` (`.scs-chip` lettera, `.scs-cell-title`, `.scs-cell-scale`) + `.scs-cell-body` (un `<img>` dell'elaborato, oppure `.scs-ph` segnaposto).
- **Varianti:** `.span-c` occupa due colonne, `.span-r` due righe. Griglia `.scs-grid` 2×3; `.g-1x3`, `.g-2x2`, `.g-1` per le altre disposizioni.
- **Il consumatore fornisce:** lettera, titolo (es. «Planimetria di progetto»), scala («1:200», «Schema», «Render», «NTS»), elaborato esportato da ArchiCAD/CAD in PDF o PNG ≥300 dpi.
- **Sì:** lettere in ordine di lettura (sinistra→destra, alto→basso); planimetrie con Simboli (nord + scala grafica).
- **No:** colorare il fondo del riquadro; mettere più di un elaborato per riquadro senza titolo.
