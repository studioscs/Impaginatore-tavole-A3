# Intestazione

Testata della tavola A3: identifica studio, oggetto e fase in una sola fascia alta 132px di riferimento.

- **Struttura:** `.scs-band` (bicolore) sopra, poi `.scs-hdr` con quattro zone: `.scs-hdr-logo` (logo SCS su bianco), `.scs-hdr-proj` (occhiello `.scs-eyebrow`, oggetto `.scs-hdr-title`, riga committente `.scs-hdr-sub`), `.scs-hdr-meta` (griglia 2×2 etichetta/valore), `.scs-hdr-cert` (marchio CQY).
- **Il consumatore fornisce:** testo dell'occhiello (tipo di procedura: «Concorso di progettazione», «Appalto integrato», «Progetto per privato»), oggetto, committente con CUP/CIG, quattro coppie meta.
- **Sì:** oggetto su massimo due righe; dati mancanti come «_______________», non vuoti.
- **No:** colori di fondo nella testata; logo SCS su fondo colorato; marchio CQY più grande del logo SCS.
- Va dentro un contenitore `.scs-tavola` o `.scs-ctx`: le misure scalano con la larghezza (1px di riferimento = 1/1587 della larghezza).
