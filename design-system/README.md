Sistema di impaginazione per presentare progetti a enti pubblici e committenti privati: tavole A3 di gara e di progetto, slide di presentazione, relazioni. Nasce da due segni: il marchio **SCS Studio Tecnico** (due fogli sfalsati, azzurro e verde, lettering nero) e il marchio **Certiquality UNI EN ISO 9001:2015** (blu istituzionale). Il primo dà l'identità, il secondo l'autorevolezza.

## Principi

- **Il progetto è il protagonista.** La grafica SCS incornicia, non decora: fondo `paper`, filetti `rule` e `line`, colore solo nelle fasce di servizio.
- **Spigolo vivo.** Tutto è a `radius-none`, come il segno SCS. `radius-xs` solo sul chip lettera del riquadro, `radius-round` solo sui simboli.
- **Due fogli, due colori.** `scs-blue` e `scs-green` compaiono sempre in coppia, azzurro prima e verde dopo (alto/sinistra → basso/destra), come nel marchio: banda bicolore, dorso, doppia cornice.
- **Blu = garanzia.** `cqy-navy` è riservato agli elementi di identificazione e controllo: dorso della tavola, numero tavola, numerazione riquadri, fascia pagina. Mai come fondo di disegni o testo lungo.
- **Una tavola si legge in tre passaggi:** oggetto (testata) → contenuti (riquadri A–F) → dati formali (cartiglio e piè).

## Voce e testi

- Italiano tecnico-istituzionale, terza persona impersonale. Niente slogan, niente punti esclamativi, niente emoji.
- Etichette di campo in MAIUSCOLO spaziato (`tav-label`), come la dicitura STUDIO TECNICO del logo: `COMMITTENTE`, `FASE`, `SCALA`, `REVISIONI`.
- Titoli in maiuscolo/minuscolo normale: «Riqualificazione del lungomare nord», non tutto maiuscolo.
- Usa la terminologia del Codice dei contratti (D.Lgs. 36/2023): «Progetto di fattibilità tecnico-economica», «Progetto esecutivo», CUP, CIG, RUP.
- Esempi reali di riga normativa e firma: «D.Lgs. 36/2023 · NTC 2018 · ISO 19650» · «Ing. Francesco Stortoni — Ord. Ing. Macerata n. 1407».
- Separatore tra dati in linea: punto mediano « · ». Date: «Maggio 2026» in testata, «05/26» in tabella revisioni.

## Colore

| Ruolo | Token | Regola |
|---|---|---|
| Fondo | `paper`, `paper-alt` | `paper` per tavola e slide; `paper-alt` solo per testata riquadri, piè, celle meta |
| Testo | `ink`, `ink-muted` | `ink` per titoli e valori; `ink-muted` per etichette e note |
| Struttura | `rule`, `line` | `rule` (2px) per i bordi principali, `line` (1px) per la griglia interna |
| Identità | `scs-blue`, `scs-green` | Solo campiture/cornici, mai testo su `paper` |
| Testo colorato | `blue-ink`, `green-ink` | Occhielli, scale, legende: le versioni leggibili dei pastelli |
| Istituzionale | `cqy-navy` + `on-navy` | Dorso, numeri, fascia pagina |
| Segnale | `cqy-cyan` | Solo punti: revisione corrente, stato attivo |

- `scs-grey` (#76787c) è il grigio esatto del logo ma su bianco vale 4.42:1: usalo per testo ≥24px o per grafica; per testo piccolo `ink-muted`.
- Massimo **un** colore di identità per elemento oltre a `cqy-navy`. Mai sfumature, ombre o trasparenze sui colori del marchio.
- Le tavole per la stampa usano sempre il tema `light` (Carta). Il tema `dark` (Schermo) è solo per proiezione di slide; i loghi restano su piastra bianca.
- Elaborati grafici (render, planimetrie): colori propri del progetto; la cornice SCS non li ritinge.

## Tipografia

Famiglia unica `sans`: **Helvetica Neue** (fallback Helvetica, Arial). Per testi in alfabeti non latini (giapponese, devanagari…) usa la famiglia `intl` (Noto Sans). Nessun file font incluso: Helvetica Neue è di sistema su macOS; su Windows ricade su Arial.

Gerarchia con due soli pesi: **700** per titoli, numeri ed etichette, **400/500** per testo e valori.

Scala tavola A3 (stampa 420×297 mm; px di riferimento su 1587px = pt × 4/3):

| Stile | Stampa | Uso |
|---|---|---|
| `tav-number` | 30 pt | Numero tavola |
| `tav-title` | 18 pt | Oggetto in testata |
| `tav-cell-title` | 9 pt maiuscolo | Titolo riquadro |
| `tav-value` | 9 pt | Valori di testata, cartiglio, piè |
| `tav-note` | 7.5 pt | Note, legende, scale |
| `tav-label` | 6 pt maiuscolo spaziato | Etichette di campo |

Scala presentazione (slide 1920×1080): `display` copertina, `headline` titolo di slide, `lead` testo d'apertura, `body` testo corrente.

## Anatomia della tavola A3

1. **Margine** `space-6` e **cornice** `rule` 2px.
2. **Banda bicolore** in alto e in basso: `scs-blue` | `scs-green`, spessore `stroke-frame`.
3. **Testata** (Intestazione): logo SCS su piastra bianca · occhiello `blue-ink` + oggetto `tav-title` + committente/CUP/CIG · griglia meta Fase/Data/Mandataria/Responsabile · marchio CQY.
4. **Dorso** `cqy-navy` con «Tavola NN» verticale in `on-navy` e filo bicolore sul lato destro.
5. **Riquadri** A–F su griglia 2×3 (varianti 1×3, 2×2, zona unica; un riquadro può occupare 2 colonne o 2 righe). Testata `paper-alt` con chip lettera `cqy-navy`, titolo `tav-cell-title`, scala in `blue-ink`.
6. **Cartiglio** a destra: N° tavola, titolo, scala, formato, revisioni (la corrente segnata con punto `cqy-cyan`), progettista.
7. **Piè**: Committente · Oggetto · RTP · Normativa · numero pagina su `cqy-navy`.

Planimetrie: sempre freccia nord in alto a destra e scala grafica in basso a sinistra del riquadro (componente Simboli).

## Marchi

- **SCS Studio Tecnico** (`assets/Logos/scs-studio-tecnico.webp`): sempre su fondo bianco, mai ricolorato, mai ridisegnato. Nella testata occupa la cella logo; in copertina di slide in alto a sinistra. Area di rispetto minima: l'altezza della lettera S intorno al segno.
- **Certiquality ISO 9001:2015** (`assets/Logos/cqy-iso-9001.png`): solo nella versione fornita, a colori, su bianco, senza alterarne proporzioni o testi; in testata tavola e in copertina, a destra. Mai più grande del logo SCS. Prima di usarlo in formati nuovi verificare il regolamento d'uso del marchio dell'ente certificatore.

## Iconografia e motivi

- Niente icone illustrative né emoji nelle tavole. I soli simboli sono tecnici: freccia nord, scala grafica, chip lettera.
- Motivo grafico ammesso: la **doppia cornice** (due rettangoli sfalsati `scs-blue` dietro, `scs-green` davanti, tratto `stroke-frame`, scarto `space-8`), da usare in copertina o come cornice di un'immagine chiave — una volta per pagina.
