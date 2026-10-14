# Tavola

Tavola A3 orizzontale completa (420×297 mm) che compone Intestazione, Riquadro, Cartiglio e PieDiTavola.

- **Struttura:** `.scs-tavola` (rapporto 420/297, contenitore di scala) > `.scs-tavola-in` (margine `space-6`, cornice `rule`) > `.scs-band` · `.scs-hdr` · `.scs-body` (`.scs-spine`, `.scs-grid`, `.scs-cart`) · `.scs-foot` · `.scs-band`.
- **Il consumatore fornisce:** i dati di testata/cartiglio/piè e gli elaborati dei riquadri. La larghezza del contenitore decide la scala: a schermo qualsiasi; in stampa 420mm.
- **Stampa:** tema `light`, «Salva come PDF» formato A3 orizzontale, margini nulli, grafica di sfondo attiva (altrimenti dorso e bande spariscono). Elaborati ≥300 dpi.
- **Serie di tavole:** stessa testata e stesso piè su tutte; cambiano solo numero, titolo, riquadri.
