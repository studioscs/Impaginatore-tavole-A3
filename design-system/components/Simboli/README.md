# Simboli

Freccia nord e scala grafica per i riquadri con planimetrie; gli unici simboli ammessi nella tavola.

- **Struttura:** `.scs-north` (SVG cerchio + freccia, lettera N) in alto a destra; `.scs-scalebar` (quattro moduli alternati pieno/vuoto + tacche) in basso a sinistra. Entrambi assoluti dentro `.scs-cell-body` o `.scs-ph`.
- **Il consumatore fornisce:** rotazione della freccia se il disegno non è orientato a nord (`transform: rotate()` sull'SVG), valori delle tacche coerenti con la scala del riquadro.
- **Colore:** sempre `ink`. Mai colorati, mai sovrapposti al disegno: lasciare `space-3` di aria.
