# -*- coding: utf-8 -*-
"""Genera il foglio di calcolo Trave Composta acciaio-cls (nuova / rinforzo)."""
import sys
from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.comments import Comment
from openpyxl.workbook.defined_name import DefinedName
from openpyxl.worksheet.datavalidation import DataValidation
from openpyxl.formatting.rule import FormulaRule

OUT = sys.argv[1]
FONT = "Helvetica Neue"

F_IN = PatternFill("solid", fgColor="FFF2CC")       # input: giallo chiaro
F_CALC = PatternFill("solid", fgColor="F2F2F2")     # calcolo: grigio
F_OUT = PatternFill("solid", fgColor="DDEBF7")      # risultato chiave: azzurro
F_H1 = PatternFill("solid", fgColor="1F3864")
F_H2 = PatternFill("solid", fgColor="D9E1F2")
F_TH = PatternFill("solid", fgColor="BDD7EE")
thin = Side(style="thin", color="A6A6A6")
BRD = Border(left=thin, right=thin, top=thin, bottom=thin)

def fnt(bold=False, color="000000", size=10, italic=False):
    return Font(name=FONT, bold=bold, color=color, size=size, italic=italic)

wb = Workbook()
NAMES = {}

def defname(name, ws, cell):
    assert name.lower() not in NAMES, name
    NAMES[name.lower()] = f"'{ws.title}'!{cell}"
    ref = f"'{ws.title}'!${''.join(c for c in cell if c.isalpha())}${''.join(c for c in cell if c.isdigit())}"
    wb.defined_names[name] = DefinedName(name, attr_text=ref)


class SB:
    """Sheet builder: colonne A=Parametro, B=Simbolo, C=Valore, D=U.M., E=Istruzioni/Note."""
    def __init__(self, title, heading, subtitle, widths=(48, 13, 24, 10, 95)):
        self.ws = wb.create_sheet(title)
        ws = self.ws
        for i, w in enumerate(widths):
            ws.column_dimensions["ABCDEFGHIJKLMNOPQRS"[i]].width = w
        ws.sheet_view.showGridLines = False
        ws["A1"] = heading
        ws["A1"].font = fnt(True, "FFFFFF", 14)
        ws["A1"].fill = F_H1
        ws.merge_cells("A1:E1")
        ws.row_dimensions[1].height = 26
        ws["A2"] = subtitle
        ws["A2"].font = fnt(False, "404040", 9, True)
        ws["A2"].alignment = Alignment(wrap_text=True, vertical="top")
        ws.merge_cells("A2:E2")
        ws.row_dimensions[2].height = 42
        ws["A3"] = ("Legenda:  celle GIALLE = dati da inserire (testo blu)  |  GRIGIE = calcolo automatico (non modificare)  |  "
                    "AZZURRE = risultati chiave.  Passa il mouse sulla cella del valore per leggere il commento con le istruzioni.")
        ws["A3"].font = fnt(False, "7F6000", 9)
        ws.merge_cells("A3:E3")
        self.r = 5
        ws.freeze_panes = "A5"

    def section(self, text):
        ws = self.ws
        self.r += 1
        ws.cell(self.r, 1, text).font = fnt(True, "1F3864", 11)
        for c in range(1, 6):
            ws.cell(self.r, c).fill = F_H2
        self.r += 1
        for c, t in enumerate(["Parametro", "Simbolo", "Valore", "U.M.", "Istruzioni per la compilazione / formula / riferimento normativo"], 1):
            cl = ws.cell(self.r, c, t)
            cl.font = fnt(True, size=9)
            cl.fill = F_TH
            cl.border = BRD
        self.r += 1

    def text(self, txt, height=None, bold=False, color="000000"):
        ws = self.ws
        ws.cell(self.r, 1, txt).font = fnt(bold, color, 9)
        ws.cell(self.r, 1).alignment = Alignment(wrap_text=True, vertical="top")
        ws.merge_cells(start_row=self.r, start_column=1, end_row=self.r, end_column=5)
        if height:
            ws.row_dimensions[self.r].height = height
        self.r += 1

    def add(self, name, label, sym, val, unit="", note="", kind="calc", fmt="0.00", dv=None):
        ws = self.ws
        r = self.r
        ws.cell(r, 1, label).font = fnt(size=10)
        ws.cell(r, 2, sym).font = fnt(size=10, italic=True)
        c = ws.cell(r, 3, val)
        ws.cell(r, 4, unit).font = fnt(size=9)
        if note.startswith("="):
            note = "Formula: " + note[1:].strip()
        ws.cell(r, 5, note).font = fnt(size=9, color="404040")
        ws.cell(r, 5).alignment = Alignment(wrap_text=True, vertical="top")
        ws.cell(r, 1).alignment = Alignment(wrap_text=True, vertical="top")
        if kind == "in":
            c.fill = F_IN
            c.font = fnt(True, "0000FF")
        elif kind == "out":
            c.fill = F_OUT
            c.font = fnt(True)
        else:
            c.fill = F_CALC
            c.font = fnt()
        c.alignment = Alignment(horizontal="center", vertical="top", wrap_text=isinstance(val, str) and not val.startswith("="))
        ws.cell(r, 2).alignment = Alignment(vertical="top")
        ws.cell(r, 4).alignment = Alignment(vertical="top")
        if fmt and not isinstance(val, str) or (isinstance(val, str) and val.startswith("=")):
            if fmt:
                c.number_format = fmt
        for k in range(1, 6):
            ws.cell(r, k).border = BRD
        if note:
            head = "DATO DA INSERIRE.\n" if kind == "in" else ("RISULTATO.\n" if kind == "out" else "CALCOLO AUTOMATICO - non modificare.\n")
            cm = Comment(head + note, "Foglio trave composta")
            cm.width = 420
            cm.height = 160
            c.comment = cm
        if dv is not None:
            dv.add(c)
        lines = max(1, len(note) // 105 + 1)
        ws.row_dimensions[r].height = max(15, 13 * lines)
        if name:
            defname(name, ws, f"C{r}")
        self.r += 1
        return f"C{r}"

    def dv_list(self, items, prompt_title="Scelta", prompt=""):
        dv = DataValidation(type="list", formula1='"' + ",".join(items) + '"', allow_blank=False)
        dv.promptTitle = prompt_title[:32]
        dv.prompt = prompt[:250]
        dv.showInputMessage = True
        dv.showErrorMessage = True
        dv.errorTitle = "Valore non ammesso"
        dv.error = "Scegliere un valore dall'elenco."
        self.ws.add_data_validation(dv)
        return dv

    def dv_num(self, lo, hi, prompt=""):
        dv = DataValidation(type="decimal", operator="between", formula1=str(lo), formula2=str(hi))
        dv.prompt = prompt[:250]
        dv.promptTitle = "Valore ammesso"
        dv.showInputMessage = True
        dv.showErrorMessage = True
        dv.error = f"Valore fuori dal campo ammesso ({lo} - {hi})."
        self.ws.add_data_validation(dv)
        return dv


def esito(ratio_expr, applic="1"):
    return f'=IF(NOT({applic}),"N.A.",IF({ratio_expr}<=1,"VERIFICATA","NON VERIFICATA"))'

def add_cf(ws, rng):
    ws.conditional_formatting.add(rng, FormulaRule(formula=[f'ISNUMBER(SEARCH("NON",{rng.split(":")[0]}))'], fill=PatternFill("solid", fgColor="F8CBAD"), font=Font(name=FONT, bold=True, color="9C0006")))
    ws.conditional_formatting.add(rng, FormulaRule(formula=[f'ISNUMBER(SEARCH("N.A.",{rng.split(":")[0]}))'], fill=PatternFill("solid", fgColor="EDEDED"), font=Font(name=FONT, color="808080")))
    ws.conditional_formatting.add(rng, FormulaRule(formula=[f'OR(ISNUMBER(SEARCH("VERIFICATA",{rng.split(":")[0]})),{rng.split(":")[0]}="OK")'], fill=PatternFill("solid", fgColor="C6EFCE"), font=Font(name=FONT, bold=True, color="006100")))
    ws.conditional_formatting.add(rng, FormulaRule(formula=[f'ISNUMBER(SEARCH("ATTENZIONE",{rng.split(":")[0]}))'], fill=PatternFill("solid", fgColor="FFEB9C"), font=Font(name=FONT, bold=True, color="9C5700")))

# =====================================================================================
# 00 LEGGIMI
# =====================================================================================
ws0 = wb.active
ws0.title = "00_LEGGIMI"
ws0.sheet_view.showGridLines = False
ws0.column_dimensions["A"].width = 4
ws0.column_dimensions["B"].width = 150
leg = [
    ("T", "TRAVE COMPOSTA ACCIAIO-CALCESTRUZZO - VERIFICA SEZIONE E CONNESSIONI"),
    ("S", "Scenario 1 = trave NUOVA (soletta c.a. su lamiera grecata + pioli)   |   Scenario 2 = RINFORZO di trave c.a. esistente con trave in acciaio sottoposta, resa collaborante con barre filettate inghisate chimicamente + nodo di attacco al pilastro esistente"),
    ("H", "1. COME SI USA (ordine di compilazione)"),
    ("P", "1) 01_DATI_GENERALI: scegliere lo SCENARIO (1 o 2) nella cella gialla in alto; inserire luce, larghezza di influenza, carichi, coefficienti e materiali."),
    ("P", "2) 02_PROFILO_ACCIAIO: inserire le dimensioni del profilo (da sagomario del produttore: h, b, tf, tw, r). Ala superiore e inferiore possono essere diverse (profili saldati)."),
    ("P", "3a) Scenario 1 -> 02A_SOLETTA_LAMIERA (geometria soletta e lamiera, classe cls, armatura trasversale) e poi 05A_PIOLI (pioli Nelson)."),
    ("P", "3b) Scenario 2 -> 02B_TRAVE_CA_ESISTENTE (geometria trave esistente, resistenza cls da prove, FC, spessore malta) e poi 05B_INGHISAGGI_TRAVE e 06_NODO_PILASTRO."),
    ("P", "4) 03_SEZIONE_COMPOSTA e 04_VERIFICHE_TRAVE sono interamente automatici (sezione plastica, sezione elastica omogeneizzata, flessione, taglio, frecce, fase di getto)."),
    ("P", "5) 07_RIEPILOGO raccoglie tutte le verifiche con rapporto domanda/capacita' ed esito. Le verifiche non pertinenti allo scenario scelto risultano 'N.A.'."),
    ("P", "I fogli dello scenario non scelto possono restare compilati: non influenzano i risultati. Ogni cella valore ha un COMMENTO (passare il mouse) con istruzioni e formula; la colonna E ripete le istruzioni."),
    ("H", "2. LEGENDA COLORI"),
    ("P", "GIALLO con testo BLU = dato da inserire.   GRIGIO = calcolo automatico (protetto logicamente: non sovrascrivere).   AZZURRO = risultato chiave.   VERDE = verificato.   ROSSO = non verificato.   ARANCIO = attenzione / controllo da fare a mano."),
    ("P", "Unita': lunghezze in mm (salvo luce e larghezze di influenza in m), forze in kN, momenti in kNm, tensioni in MPa = N/mm2, carichi distribuiti in kN/m2 o kN/m."),
    ("H", "3. RIFERIMENTI NORMATIVI E TECNICI"),
    ("P", "D.M. 17/01/2018 (NTC 2018) cap. 4.2 (acciaio), 4.3 (strutture composte), cap. 8 (costruzioni esistenti: LC, FC) e Circolare n.7/2019 (C8.5.4, Tab. C8.5.IV)."),
    ("P", "UNI EN 1994-1-1 (EC4): 5.4.1.2 larghezza efficace, 6.2.1 resistenza plastica e connessione parziale, 6.6.1.2 grado minimo di connessione, 6.6.3.1 pioli in soletta piena, 6.6.4 pioli con lamiera grecata (kl, kt, Tab. 6.2), 6.6.5 dettagli costruttivi, 6.6.6 armatura trasversale."),
    ("P", "UNI EN 1992-4 (EC2 parte 4) per gli ancoranti post-installati chimici (rottura acciaio, cono cls, sfilamento/aderenza combinata, bordo, pry-out, interazione N-V) + Benestare Tecnico Europeo (ETA/EAD 330499) dell'ancorante scelto, da cui prendere tau_Rk, c_min, s_min, gamma_inst, k8, psi_sus."),
    ("P", "UNI EN 1993-1-1 / 1-8 per classificazione, taglio plastico e rifollamento delle lamiere forate. Dispensa 'Travi composte acciaio-calcestruzzo' (allegata dall'utente) per impostazione del metodo n e del calcolo plastico."),
    ("H", "4. IPOTESI E LIMITI (LEGGERE PRIMA DI USARE I RISULTATI)"),
    ("P", "- Trave semplicemente appoggiata, momento positivo (soletta compressa). Per schemi diversi inserire manualmente MEd e VEd nelle celle di 'override' (01_DATI_GENERALI); le frecce restano calcolate per trave appoggiata."),
    ("P", "- Resistenza plastica: blocco rettangolare 0,85*fck/gamma_c nel cls compresso, fyd nell'acciaio; cls teso trascurato; armature longitudinali della soletta/trave esistente trascurate (a favore di sicurezza). Valida solo per profilo di classe 1 o 2 (controllata nel foglio 03)."),
    ("P", "- Scenario 1: con lamiera a nervature trasversali il cls nelle nervature e' trascurato (EC4). Con trave non puntellata il peso di getto/lamiera/trave grava sulla sola trave in acciaio (verifica fase di getto e freccia di fase 1)."),
    ("P", "- Scenario 2: le barre inghisate NON sono connettori duttili ai sensi EC4 6.6.1.1: la connessione si progetta in campo elastico (criterio ELASTICA, consigliato) oppure per la piena resistenza plastica N_cf (criterio PLASTICA, richiede grado di connessione >= 1). Non e' ammessa la connessione parziale."),
    ("P", "- Scenario 2: in 02B si sceglie quali carichi gravano sulla composta. 'SOLO SUCCESSIVI' (trave non scaricata): i carichi gia' presenti restano sulla trave c.a. esistente, da verificare a parte; profilo, inghisaggi e nodo portano solo i carichi successivi. 'TUTTI' (trave puntellata/martinettata, o calcolo cautelativo): tutto sulla composta. La verifica plastica (criterio PLASTICA) usa comunque il MEd totale. Le frecce sono gli incrementi dovuti ai carichi sulla composta."),
    ("P", "- Ancoranti: metodo EN 1992-4 semplificato per file di ancoranti (aree proiettate per singolo ancorante con passo e distanze dal bordo), psi_g,Np = 1, psi_ec = 1 (a favore di sicurezza). Splitting trascurato (cls fessurato e armato): verificare con ETA. Fori nel piatto da riempire (filling washer / resina) perche' tutti gli ancoranti collaborino a taglio."),
    ("P", "- Non sono incluse: verifica della piastra di attacco al pilastro, delle saldature e del collegamento bullonato trave-piastra/mensola (EN 1993-1-8), instabilita' flesso-torsionale in fase di getto se l'ala non e' vincolata, vibrazioni, fuoco, fatica, verifica del pilastro esistente per il nuovo carico concentrato."),
    ("P", "- I valori di esempio inseriti nelle celle gialle sono solo dimostrativi (in particolare tau_Rk dell'ancorante): sostituirli con i dati di progetto e dell'ETA del prodotto."),
    ("P", "- Il foglio e' uno strumento di supporto: il progettista resta responsabile della verifica dei risultati e dell'applicazione delle norme vigenti."),
    ("H", "5. STRUTTURA DEI FOGLI"),
    ("P", "00_LEGGIMI | 01_DATI_GENERALI | 02_PROFILO_ACCIAIO | 02A_SOLETTA_LAMIERA (sc.1) | 02B_TRAVE_CA_ESISTENTE (sc.2) | 03_SEZIONE_COMPOSTA | 04_VERIFICHE_TRAVE | 05A_PIOLI (sc.1) | 05B_INGHISAGGI_TRAVE (sc.2) | 06_NODO_PILASTRO (sc.2) | 07_RIEPILOGO | DB (tabelle materiali)"),
]
rr = 1
for kind, t in leg:
    c = ws0.cell(rr, 2, t)
    c.alignment = Alignment(wrap_text=True, vertical="top")
    if kind == "T":
        c.font = fnt(True, "FFFFFF", 15); c.fill = F_H1; ws0.row_dimensions[rr].height = 28
    elif kind == "S":
        c.font = fnt(False, "1F3864", 10, True); ws0.row_dimensions[rr].height = 30
    elif kind == "H":
        rr += 1; c = ws0.cell(rr, 2, t); ws0.cell(rr - 1, 2).value = None
        c.font = fnt(True, "1F3864", 12); c.fill = F_H2
    else:
        c.font = fnt(size=10)
        ws0.row_dimensions[rr].height = max(15, 14 * (len(t) // 140 + 1))
    rr += 1

# =====================================================================================
# DB
# =====================================================================================
db = wb.create_sheet("DB")
db.sheet_view.showGridLines = False
for col, w in zip("ABCDEFGHIJKLMN", [14, 10, 10, 4, 12, 10, 10, 4, 10, 12, 4, 12, 12, 12]):
    db.column_dimensions[col].width = w
db["A1"] = "TABELLE DI SUPPORTO (non modificare salvo verifica)"; db["A1"].font = fnt(True, "1F3864", 12)
def table(r0, c0, head, rows, title, src):
    db.cell(r0 - 2, c0, title).font = fnt(True, size=10)
    db.cell(r0 - 1, c0, src).font = fnt(False, "7F7F7F", 8, True)
    for j, h in enumerate(head):
        cl = db.cell(r0, c0 + j, h); cl.font = fnt(True, size=9); cl.fill = F_TH; cl.border = BRD
    for i, row in enumerate(rows):
        for j, v in enumerate(row):
            cl = db.cell(r0 + 1 + i, c0 + j, v); cl.font = fnt(size=9); cl.border = BRD
    return r0 + 1, r0 + len(rows)
cls_rows = [("C16/20", 16, 20), ("C20/25", 20, 25), ("C25/30", 25, 30), ("C28/35", 28, 35), ("C30/37", 30, 37),
            ("C32/40", 32, 40), ("C35/45", 35, 45), ("C40/50", 40, 50), ("C45/55", 45, 55), ("C50/60", 50, 60)]
a, b = table(5, 1, ["Classe", "fck", "Rck"], cls_rows, "Classi calcestruzzo", "NTC 2018 Tab. 4.1.I")
def rng(col, a, b):
    return f"DB!${col}${a}:${col}${b}"
wb.defined_names["tb_cls"] = DefinedName("tb_cls", attr_text=rng("A", a, b))
wb.defined_names["tb_clsfck"] = DefinedName("tb_clsfck", attr_text=rng("B", a, b))
st_rows = [("S235", 235, 360), ("S275", 275, 430), ("S355", 355, 510), ("S420", 420, 520), ("S460", 460, 540)]
a, b = table(5, 5, ["Acciaio", "fyk", "ftk"], st_rows, "Acciaio da carpenteria (t<=40mm)", "NTC 2018 Tab. 11.3.IX")
wb.defined_names["tb_st"] = DefinedName("tb_st", attr_text=rng("E", a, b))
wb.defined_names["tb_stfy"] = DefinedName("tb_stfy", attr_text=rng("F", a, b))
wb.defined_names["tb_stfu"] = DefinedName("tb_stfu", attr_text=rng("G", a, b))
bar_rows = [("M10", 10, 58.0), ("M12", 12, 84.3), ("M16", 16, 157.0), ("M20", 20, 245.0), ("M24", 24, 353.0), ("M27", 27, 459.0), ("M30", 30, 561.0)]
a, b = table(5, 9, ["Barra", "d [mm]", "As [mm2]"], bar_rows, "Barre filettate", "Area resistente ISO 898-1")
wb.defined_names["tb_bar"] = DefinedName("tb_bar", attr_text=rng("I", a, b))
wb.defined_names["tb_bard"] = DefinedName("tb_bard", attr_text=rng("J", a, b))
wb.defined_names["tb_baras"] = DefinedName("tb_baras", attr_text=rng("K", a, b))
cl_rows = [("4.6", 400, 240), ("5.8", 500, 400), ("8.8", 800, 640), ("10.9", 1000, 900), ("A4-70", 700, 450), ("A4-80", 800, 600)]
a, b = table(5, 13, ["Classe", "fuk", "fyk"], cl_rows, "Classi barre/viti", "ISO 898-1 / ISO 3506-1")
wb.defined_names["tb_bcl"] = DefinedName("tb_bcl", attr_text=rng("M", a, b))
wb.defined_names["tb_bclfu"] = DefinedName("tb_bclfu", attr_text=rng("N", a, b))
wb.defined_names["tb_bclfy"] = DefinedName("tb_bclfy", attr_text=rng("O", a, b))
db.column_dimensions["O"].width = 10
kt_rows = [("1_<=1", 0.85, 0.75), ("1_>1", 1.00, 0.75), ("2_<=1", 0.70, 0.60), ("2_>1", 0.80, 0.60)]
a, b = table(20, 1, ["nr_t", "passante", "fori"], kt_rows, "kt,max pioli con lamiera trasversale",
             "EC4 Tab. 6.2 (d<=20 passante; d=19-22 fori). Valori t>1mm da verificare su testo norma")
wb.defined_names["tb_ktk"] = DefinedName("tb_ktk", attr_text=rng("A", a, b))
wb.defined_names["tb_ktp"] = DefinedName("tb_ktp", attr_text=rng("B", a, b))
wb.defined_names["tb_ktf"] = DefinedName("tb_ktf", attr_text=rng("C", a, b))
lc_rows = [("LC1", 1.35), ("LC2", 1.20), ("LC3", 1.00)]
a, b = table(20, 5, ["Livello", "FC"], lc_rows, "Fattore di confidenza", "Circolare 7/2019 Tab. C8.5.IV")
wb.defined_names["tb_lc"] = DefinedName("tb_lc", attr_text=rng("E", a, b))
wb.defined_names["tb_fc"] = DefinedName("tb_fc", attr_text=rng("F", a, b))

SC1 = "1 - NUOVA: soletta c.a. su lamiera grecata + pioli"
SC2 = "2 - RINFORZO: trave acciaio sotto trave c.a. esistente + inghisaggi"

# =====================================================================================
# 01 DATI GENERALI
# =====================================================================================
s = SB("01_DATI_GENERALI", "01 - DATI GENERALI, SCENARIO, CARICHI E MATERIALI",
       "Compilare per primo questo foglio. La scelta dello SCENARIO attiva i fogli pertinenti (sc.1: 02A + 05A; sc.2: 02B + 05B + 06). "
       "Carichi: inserire i carichi di superficie (kN/m2) che vengono moltiplicati per la larghezza di influenza, piu' eventuali carichi lineari (kN/m).")
s.section("DATI PROGETTO")
s.add(None, "Commessa", "", "Commessa 2026-xx", "", "Testo libero, compare solo come intestazione.", "in", None)
s.add(None, "Opera / elemento strutturale", "", "Trave T1 - piano primo", "", "Testo libero: identificativo della trave verificata.", "in", None)
s.add(None, "Progettista", "", "Ing. ...", "", "Testo libero.", "in", None)
s.add(None, "Revisione / data", "", "Rev.0 - 2026", "", "Testo libero.", "in", None)
s.section("SCENARIO DI CALCOLO")
dv = s.dv_list([SC1, SC2], "Scenario", "1 = trave nuova con soletta su lamiera grecata e pioli. 2 = rinforzo di trave c.a. esistente con profilo in acciaio sottoposto e barre inghisate.")
s.add("sc_txt", "SCENARIO (scegliere dal menu a tendina)", "", SC2, "",
      "Scegliere 1 per trave NUOVA in acciaio con soletta c.a. gettata su lamiera grecata e connessione a pioli (Nelson). "
      "Scegliere 2 per RINFORZO: profilo in acciaio posto sotto una trave c.a. esistente, collegato con barre filettate e ancorante chimico, con attacco al pilastro esistente tramite piastra inghisata.", "in", None, dv)
s.add("sc_n", "Scenario attivo (numero)", "Sc", '=IF(LEFT(sc_txt,1)="1",1,2)', "", "Calcolato dalla scelta precedente. Usato da tutte le formule per attivare/disattivare le verifiche.", "out", "0")
s.section("GEOMETRIA E SCHEMA STATICO")
s.add("lu_m", "Luce di calcolo della trave (appoggio-appoggio)", "L", 6.0, "m",
      "Distanza tra gli appoggi (assi pilastri o assi appoggi). Schema di trave semplicemente appoggiata: L e' anche la distanza tra punti di momento nullo Le per la larghezza efficace.", "in", "0.00", s.dv_num(0.5, 30))
s.add("lu_mm", "Luce in mm", "L", "=lu_m*1000", "mm", "Conversione automatica.", "calc", "0")
s.add("inf_m", "Larghezza di influenza del solaio sulla trave", "i", 4.0, "m",
      "Larghezza di solaio che scarica sulla trave (di norma = interasse tra le travi, oppure meta' luce dei travetti/lamiera per ciascun lato). Moltiplica tutti i carichi in kN/m2.", "in", "0.00", s.dv_num(0, 20))
dv = s.dv_list(["SI", "NO"], "Puntellatura", "SI = trave puntellata durante il getto (tutto il carico va sulla sezione composta). NO = trave non puntellata (fase 1 sulla sola trave in acciaio).")
s.add("prop_txt", "Scenario 1: trave puntellata durante il getto?", "", "NO", "",
      "Solo scenario 1. NO (consigliato a favore di sicurezza se non si e' certi): peso proprio acciaio + lamiera + getto fresco + carico di costruzione gravano sulla sola trave in acciaio (verifica fase di getto e freccia di fase 1). "
      "SI: la trave e' sostenuta da puntelli fino a maturazione del getto, tutti i carichi agiscono sulla sezione composta. Ignorato nello scenario 2.", "in", None, dv)
s.add("st_one", "Carichi di fase 1 sulla sola trave in acciaio (1=si, 0=no)", "", '=IF(AND(sc_n=1,prop_txt="NO"),1,0)', "",
      "Automatico. Vale 1 solo per scenario 1 non puntellato. Nello scenario 2 la ripartizione dei carichi e' governata dalla scelta 'Carichi che gravano sulla sezione composta' nel foglio 02B.", "calc", "0")
s.add("c_gone", "Quota di G1 sulla sezione composta", "", '=IF(sc_n=1,1-st_one,IF(ex_load="TUTTI",1,0))', "-",
      "Automatico. Sc.1: 0 se non puntellata (G1 su sola trave acciaio), 1 se puntellata. Sc.2: 1 solo con 'TUTTI' (02B), altrimenti G1 resta sulla trave esistente.", "calc", "0")
s.add("c_gtwo", "Quota di G2 sulla sezione composta", "", '=IF(sc_n=1,1,IF(OR(ex_load="TUTTI",ex_g2pre="NO"),1,0))', "-",
      "Automatico. Sc.2: G2 grava sulla composta se applicato dopo il rinforzo o con opzione 'TUTTI'.", "calc", "0")
s.section("CARICHI CARATTERISTICI")
s.add("gpre_k", "G1 'pre': permanenti strutturali presenti PRIMA della collaborazione", "g1,pre", 3.0, "kN/m2",
      "Sc.1: peso soletta (getto + lamiera) per m2, es. soletta 12cm su lamiera h55 circa 2,3-2,6 kN/m2. "
      "Sc.2: peso del solaio esistente (es. laterocemento 20+4 circa 3,0-3,2 kN/m2). Il peso proprio del profilo e della trave c.a. sono aggiunti automaticamente sotto.", "in", "0.00")
s.add("gpre_l", "G1 'pre' lineare aggiuntivo", "G1,x", 0.0, "kN/m",
      "Eventuali carichi lineari strutturali aggiuntivi gia' presenti (es. muratura portata). Lasciare 0 se assenti.", "in", "0.00")
s.add("qc_k", "Sc.1: carico di costruzione in fase di getto", "qc", 1.0, "kN/m2",
      "Carichi di costruzione/maestranze durante il getto (EN 1991-1-6 4.11.1: indicativamente 0,75 kN/m2 + 10% peso getto su area di lavoro; 1,0-1,5 kN/m2 e' un valore usuale). Usato solo per la verifica di fase di getto, scenario 1.", "in", "0.00")
s.add("gtwo_k", "G2: permanenti non strutturali", "g2", 2.0, "kN/m2",
      "Pavimenti, massetti, impianti, incidenza tramezzi (NTC 3.1.3). Sc.2: indicare in 02B se sono gia' presenti prima del rinforzo (cambia solo il calcolo delle frecce).", "in", "0.00")
s.add("gtwo_l", "G2 lineare aggiuntivo", "G2,x", 0.0, "kN/m", "Es. tamponature o parapetti gravanti direttamente sulla trave.", "in", "0.00")
s.add("q_k", "Q: variabile", "qk", 2.0, "kN/m2", "Carico variabile per categoria d'uso (NTC Tab. 3.1.II): A residenziale 2,0; B uffici 3,0; C1 3,0; ecc.", "in", "0.00")
s.add("psi_two", "Coefficiente psi2 del variabile", "psi2", 0.3, "-", "NTC Tab. 2.5.I (A residenziale 0,3; B uffici 0,3; C 0,6; E 0,8). Usato per la freccia quasi permanente (informativa).", "in", "0.00")
s.section("COEFFICIENTI PARZIALI SLU (NTC Tab. 2.6.I, STR)")
s.add("gam_gone", "gamma G1", "gG1", 1.3, "-", "Permanenti strutturali: 1,3 (NTC 2018 Tab. 2.6.I, A1-STR).", "in", "0.00")
s.add("gam_gtwo", "gamma G2", "gG2", 1.5, "-", "Permanenti non strutturali: 1,5 (1,3 se compiutamente definiti).", "in", "0.00")
s.add("gam_q", "gamma Q", "gQ", 1.5, "-", "Variabili: 1,5.", "in", "0.00")
s.section("CARICHI LINEARI RISULTANTI (kN/m)")
s.add("g_acc", "Peso proprio profilo in acciaio", "Ga", "=a_area*78.5/1000000", "kN/m", "= A_a [mm2] x 78,5 kN/m3. Automatico dal foglio 02.", "calc", "0.00")
s.add("g_rc", "Peso proprio anima trave c.a. esistente (solo sc.2)", "Grc", "=IF(sc_n=2,ex_bw*(ex_h-ex_hf)*25/1000000,0)", "kN/m", "= bw x (H - hf) x 25 kN/m3 (la soletta e' gia' compresa in G1 pre). Zero nello scenario 1.", "calc", "0.00")
s.add("ln_gone", "G1 totale lineare", "G1", "=gpre_k*inf_m+gpre_l+g_acc+g_rc", "kN/m", "= g1,pre x i + G1,x + Ga + Grc", "calc", "0.00")
s.add("ln_qc", "Carico di costruzione lineare (sc.1)", "Qc", "=IF(sc_n=1,qc_k*inf_m,0)", "kN/m", "= qc x i (solo scenario 1)", "calc", "0.00")
s.add("ln_gtwo", "G2 totale lineare", "G2", "=gtwo_k*inf_m+gtwo_l", "kN/m", "= g2 x i + G2,x", "calc", "0.00")
s.add("ln_q", "Q totale lineare", "Q", "=q_k*inf_m", "kN/m", "= qk x i", "calc", "0.00")
s.add("q_slu", "Carico di progetto SLU", "qEd", "=gam_gone*ln_gone+gam_gtwo*ln_gtwo+gam_q*ln_q", "kN/m", "= gG1 G1 + gG2 G2 + gQ Q (combinazione fondamentale, NTC 2.5.1)", "out", "0.00")
s.section("SOLLECITAZIONI DI PROGETTO")
s.add("m_auto", "MEd trave appoggiata (automatico)", "", "=q_slu*lu_m^2/8", "kNm", "= qEd L^2 / 8", "calc", "0.0")
s.add("v_auto", "VEd trave appoggiata (automatico)", "", "=q_slu*lu_m/2", "kN", "= qEd L / 2", "calc", "0.0")
s.add("m_man", "MEd manuale (override, 0 = usa automatico)", "", 0, "kNm",
      "Inserire un valore >0 SOLO se lo schema non e' di trave appoggiata (es. da modello FEM). Le componenti G1/G2/Q vengono scalate proporzionalmente.", "in", "0.0")
s.add("v_man", "VEd manuale (override, 0 = usa automatico)", "", 0, "kN", "Come sopra, per il taglio massimo all'appoggio (usato anche per il flusso di scorrimento e per il nodo al pilastro).", "in", "0.0")
s.add("m_ed", "MEd di progetto", "MEd", "=IF(m_man>0,m_man,m_auto)", "kNm", "Momento massimo positivo SLU in campata.", "out", "0.0")
s.add("v_ed", "VEd di progetto", "VEd", "=IF(v_man>0,v_man,v_auto)", "kN", "Taglio massimo SLU all'appoggio.", "out", "0.0")
s.add("k_m", "Fattore di scala momenti", "", "=IF(m_auto>0,m_ed/m_auto,1)", "-", "MEd / MEd,auto", "calc", "0.000")
s.add("k_v", "Fattore di scala tagli", "", "=IF(v_auto>0,v_ed/v_auto,1)", "-", "VEd / VEd,auto", "calc", "0.000")
s.add("m_gone", "Quota MEd da G1", "MEd,G1", "=gam_gone*ln_gone*lu_m^2/8*k_m", "kNm", "Usata nelle verifiche elastiche (fasi e viscosita').", "calc", "0.0")
s.add("m_gtwo", "Quota MEd da G2", "MEd,G2", "=gam_gtwo*ln_gtwo*lu_m^2/8*k_m", "kNm", "", "calc", "0.0")
s.add("m_q", "Quota MEd da Q", "MEd,Q", "=gam_q*ln_q*lu_m^2/8*k_m", "kNm", "", "calc", "0.0")
s.add("v_gone", "Quota VEd da G1", "VEd,G1", "=gam_gone*ln_gone*lu_m/2*k_v", "kN", "", "calc", "0.0")
s.add("v_gtwo", "Quota VEd da G2", "VEd,G2", "=gam_gtwo*ln_gtwo*lu_m/2*k_v", "kN", "", "calc", "0.0")
s.add("v_q", "Quota VEd da Q", "VEd,Q", "=gam_q*ln_q*lu_m/2*k_v", "kN", "", "calc", "0.0")
s.add("m_comp", "MEd agente sulla sezione composta", "MEd,comp", "=c_gone*m_gone+c_gtwo*m_gtwo+m_q", "kNm", "Parte del momento SLU applicata dopo la solidarizzazione (usata per tensioni elastiche e connessione elastica).", "out", "0.0")
s.add("v_comp", "VEd agente sulla sezione composta", "VEd,comp", "=c_gone*v_gone+c_gtwo*v_gtwo+v_q", "kN", "Taglio SLU applicato dopo la solidarizzazione (flusso di scorrimento, nodo al pilastro sc.2).", "out", "0.0")
s.add("m_cost", "Sc.1 MEd fase di getto (sola trave acciaio)", "MEd,c", "=IF(sc_n=1,(gam_gone*ln_gone+gam_q*ln_qc)*lu_m^2/8,0)", "kNm", "= (gG1 G1 + gQ Qc) L^2/8. Solo scenario 1.", "calc", "0.0")
s.add("v_cost", "Sc.1 VEd fase di getto", "VEd,c", "=IF(sc_n=1,(gam_gone*ln_gone+gam_q*ln_qc)*lu_m/2,0)", "kN", "", "calc", "0.0")
s.section("MATERIALI - ACCIAIO DA CARPENTERIA (profilo)")
dv = s.dv_list([r[0] for r in st_rows], "Acciaio", "Grado acciaio del profilo (NTC Tab. 11.3.IX, t<=40mm).")
s.add("st_gr", "Grado acciaio profilo", "", "S275", "", "Scegliere dal menu. Per acciai esistenti/diversi modificare la tabella nel foglio DB.", "in", None, dv)
s.add("fyk_a", "Tensione di snervamento", "fyk", "=INDEX(tb_stfy,MATCH(st_gr,tb_st,0))", "MPa", "Da tabella DB.", "calc", "0")
s.add("ftk_a", "Tensione di rottura", "ftk", "=INDEX(tb_stfu,MATCH(st_gr,tb_st,0))", "MPa", "Da tabella DB.", "calc", "0")
s.add("gam_mo", "gamma M0", "gM0", 1.05, "-", "NTC 4.2.4.1.1: resistenza delle sezioni = 1,05.", "in", "0.00")
s.add("gam_mt", "gamma M2", "gM2", 1.25, "-", "NTC 4.2.8.1.1: sezioni forate / unioni = 1,25 (rifollamento).", "in", "0.00")
s.add("e_acc", "Modulo elastico acciaio", "Ea", 210000, "MPa", "NTC 11.3.4.1: 210000 MPa.", "in", "0")
s.add("fyd_a", "Resistenza di progetto acciaio", "fyd", "=fyk_a/gam_mo", "MPa", "= fyk / gM0", "out", "0.0")
s.add("eps_a", "Coefficiente epsilon", "eps", "=SQRT(235/fyk_a)", "-", "= sqrt(235/fyk), per la classificazione delle sezioni.", "calc", "0.000")
s.section("MATERIALI - CALCESTRUZZO (valori attivi secondo lo scenario)")
s.add("fck_c", "fck attiva", "fck", "=IF(sc_n=1,new_fck,ex_fcm/ex_fc)", "MPa",
      "Sc.1: fck della classe scelta in 02A. Sc.2: fcm da prove / FC (NTC 8.5.4 e Circ. C8.5.4: per i materiali esistenti si usa il valore medio diviso FC). ", "out", "0.00")
s.add("ecm_c", "Modulo elastico secante", "Ecm", "=IF(sc_n=1,22000*((new_fck+8)/10)^0.3,22000*(ex_fcm/10)^0.3)", "MPa",
      "NTC 11.2.10.3: Ecm = 22000 (fcm/10)^0,3 con fcm = fck+8 (sc.1) o fcm da prove (sc.2).", "calc", "0")
s.add("gam_c", "gamma c", "gc", 1.5, "-", "NTC 4.1.2.1.1.1: 1,5.", "in", "0.00")
s.add("alf_cc", "Coefficiente lunga durata / blocco rettangolare", "acc", 0.85, "-",
      "Coefficiente applicato al cls compresso: 0,85 (NTC 4.1.2.1.1.1 / blocco plastico EC4 6.2.1.2: 0,85 fcd con fcd=fck/gc).", "in", "0.00")
s.add("fcd_c", "Resistenza di progetto cls (blocco plastico)", "fcd", "=alf_cc*fck_c/gam_c", "MPa", "= acc fck / gc", "out", "0.00")
s.add("n_zero", "Rapporto modulare breve termine", "n0", "=e_acc/ecm_c", "-", "= Ea / Ecm (carichi variabili, istantanei).", "calc", "0.00")
s.add("k_visc", "Moltiplicatore per carichi permanenti (viscosita')", "kL", 3.0, "-",
      "n_L = kL x n0. Dispensa (Ec' = Ecm/3 per lunga durata) -> 3. EC4 5.4.2.2(11) consente per edifici n = 2 n0 per tutti i carichi. Valori 2-3.", "in", "0.00")
s.add("n_lt", "Rapporto modulare lungo termine", "nL", "=n_zero*k_visc", "-", "= kL x n0 (carichi permanenti G1, G2).", "calc", "0.00")
s.section("MATERIALI - ARMATURA ORDINARIA (sc.1: armatura trasversale soletta)")
s.add("fyk_s", "Snervamento armatura B450C", "fyk", 450, "MPa", "NTC 11.3.2.1.", "in", "0")
s.add("gam_s", "gamma s", "gs", 1.15, "-", "NTC 4.1.2.1.1.3.", "in", "0.00")
s.add("fyd_s", "Resistenza di progetto armatura", "fyd,s", "=fyk_s/gam_s", "MPa", "", "calc", "0.0")
s.section("LIMITI DI DEFORMABILITA' (SLE)")
s.add("lim_tot", "Freccia totale massima: L /", "", 250, "-", "NTC Tab. 4.2.XII solai in generale: L/250 (freccia totale). Modificabile.", "in", "0")
s.add("lim_q", "Freccia dovuta ai variabili: L /", "", 300, "-", "NTC Tab. 4.2.XII solai in generale: L/300. Modificabile (es. L/350 se tramezzi fragili).", "in", "0")
ws1 = s.ws

# =====================================================================================
# 02 PROFILO ACCIAIO
# =====================================================================================
s = SB("02_PROFILO_ACCIAIO", "02 - PROFILO IN ACCIAIO: GEOMETRIA, PROPRIETA', RESISTENZE DELLA SOLA TRAVE",
       "Inserire le dimensioni dal sagomario (profili laminati IPE/HEA/HEB: ali uguali, r = raggio di raccordo) o di un profilo saldato (r = 0, ali eventualmente diverse). "
       "Il raccordo e' modellato con una fascia di altezza r e larghezza tw + 2(1-pi/4)r: l'area risulta esatta, l'inerzia approssimata (errore trascurabile).")
s.section("DIMENSIONI")
s.add("a_nome", "Designazione profilo", "", "HEB 200", "", "Testo libero (es. IPE 300, HEB 200, saldato 400x200x10x15).", "in", None)
s.add("a_h", "Altezza totale", "h", 200, "mm", "Altezza del profilo da sagomario. Esempi: IPE300 h=300; HEB200 h=200.", "in", "0.0")
s.add("a_bs", "Larghezza ala superiore", "b,sup", 200, "mm", "Ala a contatto con la soletta (sc.1) o con l'intradosso della trave c.a. (sc.2). Sc.2: deve essere <= larghezza trave esistente bw.", "in", "0.0")
s.add("a_ts", "Spessore ala superiore", "tf,sup", 15, "mm", "Es. HEB200 tf=15; IPE300 tf=10,7.", "in", "0.0")
s.add("a_bi", "Larghezza ala inferiore", "b,inf", 200, "mm", "Per laminati = b,sup.", "in", "0.0")
s.add("a_ti", "Spessore ala inferiore", "tf,inf", 15, "mm", "Per laminati = tf,sup.", "in", "0.0")
s.add("a_tw", "Spessore anima", "tw", 9, "mm", "Es. HEB200 tw=9; IPE300 tw=7,1.", "in", "0.0")
s.add("a_r", "Raggio di raccordo ala-anima", "r", 18, "mm", "Laminati: da sagomario (HEB200 r=18; IPE300 r=15). Profili saldati: 0 (o il lato del cordone, trascurabile).", "in", "0.0")
s.add("a_hw", "Altezza anima tra le ali", "hw", "=a_h-a_ts-a_ti", "mm", "= h - tf,sup - tf,inf", "calc", "0.0")
s.add("a_fw", "Larghezza aggiuntiva fascia raccordi", "", "=2*(1-PI()/4)*a_r", "mm", "= 2(1-pi/4) r (area dei due raccordi su un lato spalmata su altezza r)", "calc", "0.00")
r0 = s.r + 1
ws = s.ws
ws.cell(r0, 1, "Strati del profilo (y misurata dal lembo superiore del profilo)").font = fnt(True, "1F3864", 10)
hdr = ["Strato", "y_top [mm]", "t [mm]", "b [mm]", "A [mm2]", "y_g [mm]", "A*y_g", "I propria [mm4]"]
for j, h in enumerate(hdr):
    c = ws.cell(r0 + 1, 1 + j, h); c.font = fnt(True, size=9); c.fill = F_TH; c.border = BRD
# tabella acciaio nelle colonne A..H (E e' larga: va bene)
lay = [("Ala superiore", "0", "a_ts", "a_bs"),
       ("Raccordi superiori", "a_ts", "a_r", "a_tw+a_fw"),
       ("Anima", "a_ts+a_r", "a_hw-2*a_r", "a_tw"),
       ("Raccordi inferiori", "a_h-a_ti-a_r", "a_r", "a_tw+a_fw"),
       ("Ala inferiore", "a_h-a_ti", "a_ti", "a_bi")]
STL = []
for i, (nm, yt, t, b) in enumerate(lay):
    r = r0 + 2 + i
    vals = [nm, "=" + yt, "=" + t, "=" + b, f"=C{r}*D{r}", f"=B{r}+C{r}/2", f"=E{r}*F{r}", f"=D{r}*C{r}^3/12"]
    for j, v in enumerate(vals):
        c = ws.cell(r, 1 + j, v); c.font = fnt(size=9); c.fill = F_CALC; c.border = BRD
        if j > 0: c.number_format = "#,##0.0"
    STL.append((nm, yt, t, b))
rA, rB = r0 + 2, r0 + 6
s.r = rB + 2
s.section("PROPRIETA' ELASTICHE DEL PROFILO")
s.add("a_area", "Area", "Aa", f"=SUM(E{rA}:E{rB})", "mm2", "Somma aree strati. Controllo: confrontare con il sagomario (es. HEB200 = 7808 mm2; IPE300 = 5381 mm2).", "out", "#,##0")
s.add("a_yg", "Baricentro dal lembo superiore", "yg,a", f"=SUM(G{rA}:G{rB})/a_area", "mm", "= sum(A y)/A", "calc", "0.0")
s.add("a_iy", "Momento d'inerzia", "Ia", f"=SUM(H{rA}:H{rB})+SUMPRODUCT(E{rA}:E{rB},(F{rA}:F{rB}-a_yg)^2)", "mm4",
      "Steiner. Controllo col sagomario (HEB200 Iy = 5696 cm4; IPE300 = 8356 cm4): differenze < 1-2% dovute al modello dei raccordi.", "out", "#,##0")
s.add("a_wi", "Modulo resistente elastico lembo inferiore", "Wel,inf", "=a_iy/(a_h-a_yg)", "mm3", "", "calc", "#,##0")
s.add("a_ws", "Modulo resistente elastico lembo superiore", "Wel,sup", "=a_iy/a_yg", "mm3", "", "calc", "#,##0")
s.add("a_peso", "Peso a metro", "", "=a_area*78.5/1000000", "kN/m", "= A x 78,5 kN/m3", "calc", "0.000")
# plastico sola trave
s.section("RESISTENZA PLASTICA DELLA SOLA TRAVE IN ACCIAIO (asse neutro di uguale area)")
r0 = s.r
hdr = ["Strato", "y_top", "t", "b", "w=b*fyd [N/mm]", "g_top [N]", "g_bot [N]", "attraversa", "x cand.", "t sopra", "t sotto", "M comp [Nmm]", "M traz [Nmm]"]
ws.column_dimensions["F"].width = 14
for col in "GHIJKLM":
    ws.column_dimensions[col].width = 14
for j, h in enumerate(hdr):
    c = ws.cell(r0, 1 + j, h); c.font = fnt(True, size=9); c.fill = F_TH; c.border = BRD
pa = r0 + 1
for i in range(5):
    r = pa + i
    src = rA + i
    g_top = f"=-SUMPRODUCT(D{src}:D{src + 4 - i},C{src}:C{src + 4 - i})*fyd_a" if i == 0 else f"=G{r - 1}"
    row = [f"=A{src}", f"=B{src}", f"=C{src}", f"=D{src}", f"=D{r}*fyd_a", g_top, f"=F{r}+2*E{r}*C{r}",
           f"=IF(AND(F{r}<0,G{r}>=0),1,0)", f"=IFERROR(B{r}-F{r}/(2*E{r}),0)",
           f"=MAX(0,MIN(B{r}+C{r},a_xpl)-B{r})", f"=MAX(0,B{r}+C{r}-MAX(B{r},a_xpl))",
           f"=E{r}*J{r}*(a_xpl-B{r}-J{r}/2)", f"=E{r}*K{r}*(MAX(B{r},a_xpl)+K{r}/2-a_xpl)"]
    for j, v in enumerate(row):
        c = ws.cell(r, 1 + j, v); c.font = fnt(size=9); c.fill = F_CALC; c.border = BRD
        if j > 0: c.number_format = "#,##0.0"
pb = pa + 4
s.r = pb + 2
s.section("RISULTATI SOLA TRAVE")
s.add("a_xpl", "Asse neutro plastico dal lembo superiore", "xpl,a", f"=SUMPRODUCT(H{pa}:H{pb},I{pa}:I{pb})", "mm", "Posizione in cui l'area compressa = area tesa.", "calc", "0.0")
s.add("a_npl", "Resistenza plastica assiale", "Npl,a", "=a_area*fyd_a/1000", "kN", "= Aa fyd", "calc", "#,##0.0")
s.add("a_mpl", "Momento resistente plastico sola trave", "Mpl,a,Rd", f"=(SUM(L{pa}:L{pb})+SUM(M{pa}:M{pb}))/1000000", "kNm", "Somma dei momenti delle risultanti rispetto all'asse neutro plastico.", "out", "0.0")
s.add("a_wpl", "Modulo plastico equivalente", "Wpl", "=a_mpl*1000000/fyd_a", "mm3", "Controllo col sagomario (HEB200 Wpl = 642,5 cm3; IPE300 = 628,4 cm3).", "calc", "#,##0")
s.add("a_mel", "Momento resistente elastico sola trave", "Mel,a,Rd", "=MIN(a_wi,a_ws)*fyd_a/1000000", "kNm", "= Wel,min fyd (da usare se classe 3).", "calc", "0.0")
s.add("a_av", "Area resistente a taglio", "Av", "=MAX(a_area-a_bs*a_ts-a_bi*a_ti+(a_tw+2*a_r)*(a_ts+a_ti)/2,a_hw*a_tw)", "mm2",
      "EN 1993-1-1 6.2.6(3): Av = A - 2 b tf + (tw + 2r) tf >= eta hw tw (eta = 1 a favore di sicurezza; ali diverse: media).", "calc", "#,##0")
s.add("a_vpl", "Resistenza plastica a taglio", "Vpl,a,Rd", "=a_av*fyd_a/SQRT(3)/1000", "kN", "= Av fyd / sqrt(3) (NTC 4.2.4.1.2.4). Il taglio e' affidato alla sola trave in acciaio (EC4 6.2.2.2).", "out", "0.0")
s.add("a_hwt", "Snellezza anima hw/tw", "", "=a_hw/a_tw", "-", "Se > 72 eps (eta=1) occorre verifica all'imbozzamento per taglio (EN 1993-1-5).", "calc", "0.0")
s.add("a_hwok", "Controllo imbozzamento anima a taglio", "", '=IF(a_hwt<=72*eps_a,"OK","ATTENZIONE: verificare imbozzamento EN 1993-1-5")', "", "", "out", None)
s.add("a_cf", "Rapporto c/t ala superiore", "c/tf", "=(a_bs-a_tw-2*a_r)/2/a_ts", "-", "Sbalzo dell'ala compressa: c = (b - tw - 2r)/2.", "calc", "0.00")
s.add("a_clf", "Classe ala superiore compressa (trave sola)", "", "=IF(a_cf<=9*eps_a,1,IF(a_cf<=10*eps_a,2,IF(a_cf<=14*eps_a,3,4)))", "-",
      "EN 1993-1-1 Tab. 5.2: 9eps / 10eps / 14eps. In sezione composta l'ala connessa ai connettori con passo regolamentare puo' considerarsi di classe 1 (EC4 5.5.2(1)).", "out", "0")
s.add("a_clw", "Classe anima in flessione (trave sola)", "", "=IF((a_hw-2*a_r)/a_tw<=72*eps_a,1,IF((a_hw-2*a_r)/a_tw<=83*eps_a,2,IF((a_hw-2*a_r)/a_tw<=124*eps_a,3,4)))", "-", "EN 1993-1-1 Tab. 5.2 flessione: 72/83/124 eps.", "out", "0")
add_cf(ws, f"C{s.r - 4}:C{s.r - 4}")
ws2 = s.ws

# =====================================================================================
# 02A SOLETTA LAMIERA (sc.1)
# =====================================================================================
s = SB("02A_SOLETTA_LAMIERA", "02A - SCENARIO 1: SOLETTA IN C.A. SU LAMIERA GRECATA",
       "Compilare solo se SCENARIO = 1. Dati della lamiera dalla scheda del produttore (altezza nervatura hp, larghezza media b0, spessore t). "
       "Con nervature trasversali alla trave il cls nelle nervature non collabora (EC4 6.2.1.2): spessore utile hc = h - hp.")
s.section("CALCESTRUZZO DEL GETTO")
dv = s.dv_list([r[0] for r in cls_rows], "Classe cls", "Classe del calcestruzzo del getto di completamento (NTC Tab. 4.1.I).")
s.add("new_cl", "Classe calcestruzzo getto", "", "C25/30", "", "Scegliere dal menu. Per strutture composte NTC 4.3.1: classi da C20/25 a C60/75.", "in", None, dv)
s.add("new_fck", "Resistenza caratteristica cilindrica", "fck", "=INDEX(tb_clsfck,MATCH(new_cl,tb_cls,0))", "MPa", "Da tabella DB.", "calc", "0")
s.section("GEOMETRIA SOLETTA E LAMIERA")
dv = s.dv_list(["TRASVERSALE", "PARALLELA", "SOLETTA PIENA"], "Lamiera", "Orientamento delle nervature rispetto all'asse della trave.")
s.add("dk_or", "Orientamento nervature lamiera", "", "TRASVERSALE", "",
      "TRASVERSALE = nervature perpendicolari alla trave (caso usuale per travi principali): riduzione kt dei pioli. PARALLELA = nervature parallele alla trave: riduzione kl. "
      "SOLETTA PIENA = getto su casseri/predalle senza lamiera: nessuna riduzione (impostare hp = 0).", "in", None, dv)
s.add("sl_h", "Spessore totale soletta (estradosso - fondo lamiera)", "h", 120, "mm", "Altezza complessiva getto + nervatura.", "in", "0")
s.add("dk_hp", "Altezza nervatura lamiera", "hp", 55, "mm", "Dalla scheda lamiera (es. 55, 75 mm). 0 per soletta piena. EC4: formule kt valide per hp <= 85 mm.", "in", "0")
s.add("dk_b", "Larghezza media della nervatura (cls)", "b0", 150, "mm",
      "Larghezza media della nervatura riempita di cls (per lamiere a coda di rondine: larghezza minima). Dalla scheda lamiera.", "in", "0")
s.add("dk_t", "Spessore lamiera", "t", 0.8, "mm", "Dalla scheda lamiera (0,7 - 1,25 mm). Influenza kt,max.", "in", "0.00")
dv = s.dv_list(["SI", "NO"], "Saldatura", "SI = pioli saldati attraverso la lamiera. NO = lamiera forata e piolo saldato direttamente sull'ala.")
s.add("dk_thr", "Pioli saldati attraverso la lamiera?", "", "SI", "", "SI = saldatura passante (richiede d <= 20 mm, EC4 6.6.5.8(3)). NO = lamiera preforata.", "in", None, dv)
s.add("sl_hc", "Spessore utile cls sopra le nervature", "hc", "=sl_h-dk_hp", "mm", "= h - hp. Solo questa parte collabora a compressione.", "out", "0")
s.section("LARGHEZZA EFFICACE (EC4 5.4.1.2)")
s.add("sl_bl", "Distanza disponibile lato sinistro (dall'asse trave)", "b1", 2000, "mm",
      "Meta' interasse con la trave adiacente, oppure distanza dal bordo libero della soletta per trave di bordo.", "in", "0")
s.add("sl_br", "Distanza disponibile lato destro (dall'asse trave)", "b2", 2000, "mm", "Come sopra, lato opposto.", "in", "0")
s.add("sl_bz", "Interasse trasversale pioli estremi", "b0,st", "=(st_nr-1)*st_et", "mm", "Da foglio 05A: (nr - 1) x et. 0 con una fila di pioli.", "calc", "0")
s.add("sl_be1", "Larghezza efficace lato sinistro", "be1", "=MAX(0,MIN(lu_mm/8,sl_bl-sl_bz/2))", "mm", "= min(Le/8 ; b1 - b0/2) con Le = L (trave appoggiata).", "calc", "0")
s.add("sl_be2", "Larghezza efficace lato destro", "be2", "=MAX(0,MIN(lu_mm/8,sl_br-sl_bz/2))", "mm", "= min(Le/8 ; b2 - b0/2)", "calc", "0")
s.add("sl_beff", "Larghezza efficace soletta", "beff", "=sl_bz+sl_be1+sl_be2", "mm", "= b0 + be1 + be2 (in mezzeria). Usata per resistenza e rigidezza.", "out", "0")
s.section("FASE DI GETTO")
dv = s.dv_list(["SI", "NO"], "Vincolo ala", "L'ala superiore e' vincolata lateralmente dalla lamiera fissata durante il getto?")
s.add("sl_ltb", "Ala superiore vincolata dalla lamiera in fase di getto?", "", "SI", "",
      "SI se la lamiera e' fissata all'ala (chiodatura/saldatura) e funziona da ritegno: niente svergolamento in fase di getto. NO: il foglio segnala che occorre la verifica flesso-torsionale (EN 1993-1-1 6.3.2) a parte.", "in", None, dv)
s.section("ARMATURA TRASVERSALE DELLA SOLETTA (EC4 6.6.6)")
s.add("sl_ast", "Armatura trasversale che attraversa il piano a-a", "Asf/sf", 565, "mm2/m",
      "Somma delle barre trasversali (superiori + inferiori) per metro di trave che attraversano la superficie di scorrimento ai lati della trave. Es.: phi8/150 = 335 mm2/m; phi10/200 = 393; phi12/200 = 565 mm2/m. Il contributo della lamiera continua (EC4 6.6.6.4) e' trascurato.", "in", "0")
s.add("sl_cot", "cot(theta) bielle", "cot", 1.0, "-", "EC4 6.6.6.2 / EC2 6.2.4(4): ala compressa 1,0 <= cot <= 2,0. 1,0 e' a favore di sicurezza per l'armatura.", "in", "0.00", s.dv_num(1, 2))
ws2a = s.ws

# =====================================================================================
# 02B TRAVE CA ESISTENTE (sc.2)
# =====================================================================================
s = SB("02B_TRAVE_CA_ESISTENTE", "02B - SCENARIO 2: TRAVE IN C.A. ESISTENTE DA RINFORZARE",
       "Compilare solo se SCENARIO = 2. Geometria da rilievo, resistenza del cls da prove (carote/sclerometro-SonReb tarati) e Livello di Conoscenza secondo NTC cap. 8 / Circ. 2019. "
       "Il profilo in acciaio e' posto sotto l'intradosso della trave, con eventuale strato di malta/resina di allettamento.")
s.section("GEOMETRIA")
s.add("ex_bw", "Larghezza anima trave esistente", "bw", 300, "mm", "Larghezza dell'intradosso dove si fissa il profilo (da rilievo).", "in", "0")
s.add("ex_h", "Altezza totale trave (compreso solaio)", "H", 500, "mm", "Dall'intradosso trave all'estradosso della soletta (o della trave se a spessore/emergente senza soletta).", "in", "0")
s.add("ex_hf", "Spessore soletta collaborante", "hf", 40, "mm",
      "Spessore della soletta piena superiore (es. 40-50 mm nel laterocemento). 0 se assente: la sezione cls diventa rettangolare bw x H.", "in", "0")
s.add("ex_bl", "Distanza disponibile lato sinistro (dall'asse trave)", "b1", 2000, "mm", "Meta' interasse con la trave adiacente o distanza dal bordo del solaio.", "in", "0")
s.add("ex_br", "Distanza disponibile lato destro (dall'asse trave)", "b2", 2000, "mm", "Come sopra.", "in", "0")
s.add("ex_beff", "Larghezza efficace della soletta", "beff", "=IF(ex_hf>0,MIN(ex_bl+ex_br,ex_bw+MAX(0,MIN(lu_mm/8,ex_bl-ex_bw/2))+MAX(0,MIN(lu_mm/8,ex_br-ex_bw/2))),ex_bw)", "mm",
      "= bw + somma min(L/8 ; bi - bw/2) (EC4 5.4.1.2 con b0 = bw). Se hf = 0 vale bw.", "out", "0")
s.add("ex_gap", "Spessore malta/resina tra intradosso e ala del profilo", "tg", 10, "mm",
      "Strato di allettamento per regolarizzare l'intradosso. Se tg > d/2 della barra o malta < 30 MPa, la barra lavora con braccio di leva (EN 1992-4 6.2.2.3) e la resistenza a taglio acciaio si riduce molto.", "in", "0")
dv = s.dv_list(["SI", "NO"], "Malta", "Malta/resina di allettamento con resistenza a compressione >= 30 MPa?")
s.add("ex_grok", "Malta di allettamento con fc >= 30 MPa?", "", "SI", "", "Condizione EN 1992-4 6.2.2.3 per trascurare il braccio di leva (insieme a tg <= d/2).", "in", None, dv)
s.section("CALCESTRUZZO ESISTENTE (NTC 8.5.4 - Circ. C8.5.4)")
s.add("ex_fcm", "Resistenza media cilindrica da prove", "fcm", 20, "MPa",
      "Valore medio della resistenza cilindrica in situ (carote: fc,cil). Se si dispone di Rc cubica: fcm ~ 0,83 Rcm. Il valore e' diviso per FC.", "in", "0.0")
dv = s.dv_list(["LC1", "LC2", "LC3"], "Livello conoscenza", "LC1 FC=1,35; LC2 FC=1,20; LC3 FC=1,00")
s.add("ex_lc", "Livello di conoscenza", "LC", "LC2", "", "Circ. 7/2019 Tab. C8.5.IV: LC1 -> FC 1,35; LC2 -> 1,20; LC3 -> 1,00.", "in", None, dv)
s.add("ex_fc", "Fattore di confidenza", "FC", "=INDEX(tb_fc,MATCH(ex_lc,tb_lc,0))", "-", "Da tabella DB.", "calc", "0.00")
s.add("ex_fceq", "fck equivalente usata nei calcoli", "fcm/FC", "=ex_fcm/ex_fc", "MPa", "Resistenza usata per sezione composta e ancoranti (al posto di fck).", "out", "0.00")
dv = s.dv_list(["SI", "NO"], "G2 esistenti", "I carichi G2 sono gia' presenti prima del rinforzo?")
s.add("ex_g2pre", "I carichi G2 sono gia' presenti prima del rinforzo?", "", "SI", "",
      "SI: G2 e' gia' portato dalla trave c.a. esistente (non grava sulla composta salvo puntellamento/martinettatura). NO: G2 viene applicato dopo il rinforzo e grava sulla composta.", "in", None, dv)
dv = s.dv_list(["SOLO SUCCESSIVI", "TUTTI"], "Carichi su composta", "SOLO SUCCESSIVI = trave non scaricata. TUTTI = trave puntellata/martinettata prima del collegamento.")
s.add("ex_load", "Carichi che gravano sulla sezione composta", "", "SOLO SUCCESSIVI", "",
      "SOLO SUCCESSIVI: la trave NON viene scaricata prima del collegamento; i carichi gia' presenti (G1 e, se gia' presenti, G2) restano sulla trave c.a. esistente "
      "(da verificare a parte) e la composta (profilo + inghisaggi) porta solo i carichi applicati dopo. "
      "TUTTI: la trave viene puntellata/martinettata e poi rilasciata a collegamento avvenuto, oppure si vuole un calcolo cautelativo: tutto il carico su profilo e connessione.", "in", None, dv)
s.add("ex_chk", "Controllo ala profilo <= larghezza trave", "", '=IF(a_bs<=ex_bw,"OK","ATTENZIONE: ala del profilo piu\' larga della trave esistente")', "", "Il profilo deve appoggiare interamente sotto l'intradosso.", "out", None)
add_cf(s.ws, f"C{s.r - 1}:C{s.r - 1}")
s.text("NOTA: prima dell'intervento verificare con pacometro/saggi la posizione di staffe e ferri longitudinali inferiori per evitare di tranciarli con le forature; "
       "verificare a parte la trave c.a. esistente per i carichi presenti prima della messa in opera del profilo (se non si scarica/puntella la trave).", 30, False, "9C5700")
ws2b = s.ws

# =====================================================================================
# 03 SEZIONE COMPOSTA
# =====================================================================================
s = SB("03_SEZIONE_COMPOSTA", "03 - SEZIONE COMPOSTA: RESISTENZA PLASTICA E PROPRIETA' ELASTICHE OMOGENEIZZATE",
       "Foglio automatico. La sezione e' descritta a strati dall'estradosso: [1] cls superiore (soletta), [2] cls inferiore (anima trave c.a. nel sc.2; nervature lamiera, non collaboranti, nel sc.1), "
       "[3] malta (sc.2, non collaborante), [4-8] profilo in acciaio. L'asse neutro plastico si trova imponendo l'equilibrio alla traslazione (dispensa eq. 4.1-4.2, EC4 6.2.1.2).")
ws = s.ws
s.section("GEOMETRIA ATTIVA DEL CALCESTRUZZO")
s.add("cs_bone", "Larghezza strato cls superiore", "b1", "=IF(sc_n=1,sl_beff,ex_beff)", "mm", "Sc.1: beff soletta. Sc.2: beff soletta collaborante (o bw).", "calc", "0")
s.add("cs_tone", "Spessore strato cls superiore", "t1", "=IF(sc_n=1,sl_hc,ex_hf)", "mm", "Sc.1: hc = h - hp. Sc.2: hf.", "calc", "0")
s.add("cs_btwo", "Larghezza strato cls inferiore", "b2", "=IF(sc_n=1,0,ex_bw)", "mm", "Sc.1: 0 (nervature trascurate). Sc.2: bw.", "calc", "0")
s.add("cs_ttwo", "Spessore strato cls inferiore", "t2", "=IF(sc_n=1,dk_hp,ex_h-ex_hf)", "mm", "Sc.1: hp. Sc.2: H - hf.", "calc", "0")
s.add("cs_gap", "Spessore malta (non collaborante)", "tg", "=IF(sc_n=1,0,ex_gap)", "mm", "", "calc", "0")
s.add("cs_yst", "Profondita' estradosso profilo", "y,st", "=cs_tone+cs_ttwo+cs_gap", "mm", "Distanza dall'estradosso cls al lembo superiore del profilo.", "calc", "0.0")
s.add("cs_htot", "Altezza totale sezione composta", "Htot", "=cs_yst+a_h", "mm", "", "calc", "0.0")
s.section("CALCOLO PLASTICO A STRATI (N, mm)")
r0 = s.r
hdr = ["Strato", "y_top", "t", "b", "f [MPa]", "k (1 cls / 2 acc.)", "acciaio?", "w=b*f [N/mm]", "g_top [N]", "g_bot [N]", "attraversa", "x cand. [mm]", "t compresso", "t teso (acc.)", "M comp [Nmm]", "M traz [Nmm]", "C cls [N]"]
for j, h in enumerate(hdr):
    c = ws.cell(r0, 1 + j, h); c.font = fnt(True, size=9); c.fill = F_TH; c.border = BRD
for col in "FGHIJKLMNOPQ":
    ws.column_dimensions[col].width = 14
layers = [("[1] cls superiore", "0", "cs_tone", "cs_bone", "fcd_c", 1, 0),
          ("[2] cls inferiore", "cs_tone", "cs_ttwo", "cs_btwo", "fcd_c", 1, 0),
          ("[3] malta (no)", "cs_tone+cs_ttwo", "cs_gap", "0", "0", 1, 0)]
for i, (nm, yt, t, b) in enumerate(STL):
    layers.append(("[%d] %s" % (4 + i, nm), "cs_yst+" + yt, t, b, "fyd_a", 2, 1))
pa = r0 + 1
pb = pa + len(layers) - 1
for i, (nm, yt, t, b, f, k, isst) in enumerate(layers):
    r = pa + i
    gt = f"=-SUMPRODUCT(G{pa}:G{pb},H{pa}:H{pb},C{pa}:C{pb})" if i == 0 else f"=J{r - 1}"
    row = [nm, "=" + yt, "=" + t, "=" + b, "=" + f, k, isst, f"=D{r}*E{r}", gt, f"=I{r}+F{r}*H{r}*C{r}",
           f"=IF(AND(I{r}<0,J{r}>=0),1,0)", f"=IFERROR(B{r}-I{r}/(F{r}*H{r}),0)",
           f"=MAX(0,MIN(B{r}+C{r},cs_xpl)-B{r})", f"=G{r}*MAX(0,B{r}+C{r}-MAX(B{r},cs_xpl))",
           f"=H{r}*M{r}*(cs_xpl-B{r}-M{r}/2)", f"=H{r}*N{r}*(MAX(B{r},cs_xpl)+N{r}/2-cs_xpl)", f"=(1-G{r})*H{r}*M{r}"]
    for j, v in enumerate(row):
        c = ws.cell(r, 1 + j, v); c.font = fnt(size=9); c.fill = F_CALC; c.border = BRD
        if j > 0: c.number_format = "#,##0.0"
s.r = pb + 2
s.text("Metodo: g(y) = (compressioni sopra y) - (trazioni sotto y); per ogni strato si calcola g al lembo superiore e inferiore; l'asse neutro e' nello strato dove g passa da negativo a positivo "
       "(interpolazione lineare esatta per strati a larghezza costante). Il momento plastico e' la somma dei momenti delle risultanti rispetto all'asse neutro.", 28)
s.section("RISULTATI PLASTICI (SLU)")
s.add("cs_xpl", "Asse neutro plastico dall'estradosso", "xpl", f"=SUMPRODUCT(K{pa}:K{pb},L{pa}:L{pb})", "mm", "", "out", "0.0")
s.add("cs_where", "Posizione asse neutro plastico", "", '=IF(cs_xpl<=cs_tone,"nel cls superiore (soletta)",IF(cs_xpl<=cs_tone+cs_ttwo,"nel cls inferiore",IF(cs_xpl<=cs_yst,"nella malta",IF(cs_xpl<=cs_yst+a_ts,"nell\'ala superiore del profilo","nell\'anima del profilo"))))', "",
      "Se nell'anima, controllare la classe dell'anima (sotto).", "out", None)
s.add("cs_ncf", "Forza di scorrimento per connessione completa", "Ncf", f"=SUM(Q{pa}:Q{pb})/1000", "kN", "= risultante di compressione nel cls = min(Rc ; Ra) (dispensa eq. 7.5). Da trasferire con i connettori tra appoggio e mezzeria.", "out", "#,##0.0")
s.add("cs_rc", "Resistenza a compressione cls totale", "Rc", f"=SUMPRODUCT(1-G{pa}:G{pb},H{pa}:H{pb},C{pa}:C{pb})/1000", "kN", "= somma fcd x area cls collaborante", "calc", "#,##0.0")
s.add("cs_mpl", "Momento resistente plastico sezione composta", "Mpl,Rd", f"=(SUM(O{pa}:O{pb})+SUM(P{pa}:P{pb}))/1000000", "kNm", "Connessione completa (eta = 1).", "out", "0.0")
s.add("cs_ratio", "Rapporto Mpl,Rd / Mpl,a,Rd", "", "=cs_mpl/a_mpl", "-", "Incremento di resistenza dovuto alla collaborazione (dispensa: tipicamente 1,25-2).", "calc", "0.00")
s.section("CLASSE DELLA SEZIONE IN CAMPO PLASTICO (EC4 5.5, EN 1993-1-1 Tab. 5.2)")
s.add("cs_cw", "Altezza netta anima", "c", "=a_hw-2*a_r", "mm", "", "calc", "0.0")
s.add("cs_alfa", "Quota compressa dell'anima", "alfa", "=MIN(1,MAX(0,(cs_xpl-(cs_yst+a_ts+a_r))/cs_cw))", "-", "0 se l'asse neutro e' sopra l'anima (anima tutta tesa).", "calc", "0.000")
s.add("cs_wlim", "Limite c/tw per classe 2", "", "=IF(cs_alfa=0,9999,IF(cs_alfa>0.5,456*eps_a/(13*cs_alfa-1),41.5*eps_a/cs_alfa))", "-", "EN 1993-1-1 Tab. 5.2: alfa>0,5 -> 456eps/(13alfa-1); alfa<=0,5 -> 41,5eps/alfa.", "calc", "0.0")
s.add("cs_wcl", "Anima di classe 1 o 2?", "", '=IF(cs_cw/a_tw<=cs_wlim,"OK","NON VERIFICATA: anima classe 3 - calcolo plastico non ammesso")', "", "", "out", None)
s.add("cs_fcl", "Ala superiore compressa", "", '=IF(cs_xpl<=cs_yst,"OK",IF(a_clf<=2,"OK","ATTENZIONE: ala classe "&a_clf&" - ammessa classe 1 solo se connettori con passo <= 22 tf eps (EC4 5.5.2, 6.6.5.5)"))', "", "", "out", None)
add_cf(ws, f"C{s.r - 2}:C{s.r - 1}")
s.add("cs_plok", "Calcolo plastico ammesso (1=si)", "", '=IF(cs_wcl="OK",1,0)', "", "", "calc", "0")

s.section("PROPRIETA' ELASTICHE OMOGENEIZZATE ALL'ACCIAIO (cls teso trascurato)")
s.text("Calcolo con due rapporti modulari: n0 (carichi di breve durata, Q) e nL (carichi permanenti). Asse neutro: soluzione dell'equazione di 2° grado nello strato di cls in cui cade "
       "(dispensa eq. 3.1-3.4), altrimenti sezione interamente reagente. S_int = momento statico dell'acciaio rispetto all'asse neutro = momento statico per il flusso di scorrimento all'interfaccia.", 30)
s.add("cs_aone", "Area strato cls superiore", "Ac1", "=cs_bone*cs_tone", "mm2", "", "calc", "#,##0")
s.add("cs_atwo", "Area strato cls inferiore", "Ac2", "=cs_btwo*cs_ttwo", "mm2", "", "calc", "#,##0")
s.add("cs_ys", "Profondita' baricentro acciaio", "ys", "=cs_yst+a_yg", "mm", "", "calc", "0.0")
for tag, nn, lab in (("s", "n_zero", "BREVE TERMINE (n0)"), ("l", "n_lt", "LUNGO TERMINE (nL)")):
    s.section("SEZIONE ELASTICA - " + lab)
    s.add(f"el{tag}_xa", "Asse neutro se nel cls superiore", "x1",
          f"=(-a_area+SQRT(a_area^2+2*(cs_bone/{nn})*a_area*cs_ys))/(cs_bone/{nn})", "mm", "(b1/2n) x^2 + Aa x - Aa ys = 0", "calc", "0.0")
    s.add(f"el{tag}_sz", "Termine noto strato inferiore", "S0", f"=cs_aone/{nn}*(cs_tone-cs_tone/2)-a_area*(cs_ys-cs_tone)", "mm3", "", "calc", "#,##0")
    s.add(f"el{tag}_bb", "Coefficiente B strato inferiore", "B", f"=cs_aone/{nn}+a_area", "mm2", "", "calc", "#,##0")
    s.add(f"el{tag}_u", "Asse neutro se nel cls inferiore (da t1)", "u",
          f"=IFERROR(IF(cs_btwo=0,-el{tag}_sz/el{tag}_bb,(-el{tag}_bb+SQRT(el{tag}_bb^2-4*cs_btwo/(2*{nn})*el{tag}_sz))/(2*cs_btwo/(2*{nn}))),1E9)", "mm", "(b2/2n) u^2 + B u + S0 = 0", "calc", "0.0")
    s.add(f"el{tag}_xc", "Asse neutro se sezione cls tutta compressa", "x3",
          f"=(cs_aone/{nn}*cs_tone/2+cs_atwo/{nn}*(cs_tone+cs_ttwo/2)+a_area*cs_ys)/(cs_aone/{nn}+cs_atwo/{nn}+a_area)", "mm", "", "calc", "0.0")
    s.add(f"el{tag}_x", "ASSE NEUTRO ELASTICO dall'estradosso", "x",
          f"=IF(el{tag}_xa<=cs_tone,el{tag}_xa,IF(AND(cs_btwo>0,el{tag}_u<=cs_ttwo),cs_tone+el{tag}_u,el{tag}_xc))", "mm", "", "out", "0.0")
    s.add(f"el{tag}_hone", "Altezza compressa strato superiore", "", f"=MIN(el{tag}_x,cs_tone)", "mm", "", "calc", "0.0")
    s.add(f"el{tag}_htwo", "Altezza compressa strato inferiore", "", f"=IF(cs_btwo>0,MAX(0,MIN(el{tag}_x-cs_tone,cs_ttwo)),0)", "mm", "", "calc", "0.0")
    s.add(f"el{tag}_j", "Momento d'inerzia omogeneizzato", "Jid",
          f"=a_iy+a_area*(cs_ys-el{tag}_x)^2+(cs_bone*el{tag}_hone^3/12+cs_bone*el{tag}_hone*(el{tag}_x-el{tag}_hone/2)^2)/{nn}"
          f"+(cs_btwo*el{tag}_htwo^3/12+cs_btwo*el{tag}_htwo*(el{tag}_x-cs_tone-el{tag}_htwo/2)^2)/{nn}", "mm4", "Dispensa eq. 3.2 / 3.4 generalizzate.", "out", "#,##0")
    s.add(f"el{tag}_sint", "Momento statico all'interfaccia", "S", f"=a_area*(cs_ys-el{tag}_x)", "mm3", "= Aa (ys - x) = momento statico della parte sopra l'interfaccia.", "calc", "#,##0")
    s.add(f"el{tag}_wi", "Modulo resistente lembo inferiore acciaio", "Wi", f"=el{tag}_j/(cs_htot-el{tag}_x)", "mm3", "", "calc", "#,##0")
    s.add(f"el{tag}_ratio", "Rapporto Jid / Ia", "", f"=el{tag}_j/a_iy", "-", "Incremento di rigidezza (dispensa: fino a ~3).", "calc", "0.00")
ws3 = s.ws

# =====================================================================================
# 04 VERIFICHE TRAVE
# =====================================================================================
s = SB("04_VERIFICHE_TRAVE", "04 - VERIFICHE DELLA TRAVE COMPOSTA (SLU flessione e taglio, fase di getto, tensioni elastiche, frecce SLE)",
       "Foglio automatico. Grado di connessione eta preso da 05A (sc.1, pioli) o 05B (sc.2, inghisaggi).")
ws = s.ws
s.section("SLU - FLESSIONE (resistenza plastica, EC4 6.2.1)")
s.add("vt_eta", "Grado di connessione", "eta", "=IF(sc_n=1,st_eta,an_eta)", "-", "Sc.1: da 05A. Sc.2: da 05B (con criterio PLASTICA; con ELASTICA e' solo indicativo).", "calc", "0.000")
s.add("vt_mrd", "Momento resistente di progetto", "MRd",
      "=IF(vt_eta>=1,cs_mpl,IF(sc_n=1,a_mpl+vt_eta*(cs_mpl-a_mpl),a_mpl))", "kNm",
      "eta>=1: Mpl,Rd. Sc.1 con eta<1 (pioli duttili): MRd = Mpl,a + eta (Mpl - Mpl,a) (EC4 6.2.1.3(5), dispensa eq. 7.8). "
      "Sc.2 con eta<1: connessione parziale NON ammessa con ancoranti non duttili -> si assume la sola trave in acciaio.", "out", "0.0")
s.add("vt_rm", "Rapporto MEd / MRd", "", "=IF(cs_plok=1,m_ed/vt_mrd,99)", "-", "Se la sezione non e' di classe 1-2 il calcolo plastico non e' ammesso (rapporto forzato a 99).", "out", "0.000")
s.add("vt_em", "Esito flessione plastica", "", esito("vt_rm", 'OR(sc_n=1,an_crit="PLASTICA")'), "", "Sc.2 con criterio ELASTICA: vale la verifica tensionale sotto.", "out", None)
s.add("vt_emin", "Sc.1: grado di connessione minimo", "eta,min", "=IF(lu_m<=25,MAX(0.4,1-(355/fyk_a)*(0.75-0.03*lu_m)),1)", "-",
      "EC4 6.6.1.2(1) ali uguali, Le<=25 m: eta >= 1 - (355/fy)(0,75 - 0,03 Le) e >= 0,4. Per ali disuguali vedere EC4 6.6.1.2(1) seconda formula.", "calc", "0.000")
s.add("vt_etaok", "Sc.1: verifica grado di connessione minimo", "", '=IF(sc_n<>1,"N.A.",IF(st_eta>=vt_emin,"VERIFICATA","NON VERIFICATA"))', "", "", "out", None)
s.section("SLU - TAGLIO (sola anima del profilo, EC4 6.2.2.2)")
s.add("vt_rv", "Rapporto VEd / Vpl,a,Rd", "", "=v_ed/a_vpl", "-", "Taglio interamente affidato al profilo.", "out", "0.000")
s.add("vt_ev", "Esito taglio", "", esito("vt_rv"), "", "", "out", None)
s.add("vt_mv", "Interazione M-V", "", '=IF(vt_rv<=0.5,"OK","ATTENZIONE: VEd>0,5 Vpl - ridurre fyd dell\'anima dove M e V sono contemporaneamente elevati (EC4 6.2.2.4)")', "",
      "Per trave appoggiata con carico uniforme M max e V max non coincidono: generalmente trascurabile.", "out", None)
s.section("SCORRIMENTO ALL'INTERFACCIA ACCIAIO-CALCESTRUZZO (metodo elastico, dispensa eq. 7.1-7.2)")
s.add("vt_flow", "Flusso di scorrimento all'appoggio", "vEd", "=((c_gone*v_gone+c_gtwo*v_gtwo)*ell_sint/ell_j+v_q*els_sint/els_j)*1000", "N/mm",
      "= VEd,G S(nL)/Jid(nL) + VEd,Q S(n0)/Jid(n0), con i soli carichi applicati alla sezione composta (Jourawski).", "out", "0.0")
s.add("vt_fel", "Forza di scorrimento elastica appoggio-mezzeria", "Fel", "=(c_gone*m_gone+c_gtwo*m_gtwo)*1000*ell_sint/ell_j+m_q*1000*els_sint/els_j", "kN",
      "= MEd S / Jid (risultante delle compressioni nel cls in mezzeria, dispensa eq. 7.2: V = M/z).", "out", "#,##0.0")
s.section("SC.1 - FASE DI GETTO (trave non puntellata, sola trave in acciaio)")
s.add("vt_rmc", "Rapporto MEd,c / Ma,Rd", "", "=IF(st_one=1,m_cost/IF(MAX(a_clf,a_clw)<=2,a_mpl,a_mel),0)", "-", "Ma,Rd = Mpl,a (classe 1-2) o Mel,a (classe 3).", "out", "0.000")
s.add("vt_emc", "Esito flessione fase di getto", "", esito("vt_rmc", "st_one=1"), "", "", "out", None)
s.add("vt_rvc", "Rapporto VEd,c / Vpl,a,Rd", "", "=IF(st_one=1,v_cost/a_vpl,0)", "-", "", "out", "0.000")
s.add("vt_evc", "Esito taglio fase di getto", "", esito("vt_rvc", "st_one=1"), "", "", "out", None)
s.add("vt_ltb", "Svergolamento in fase di getto", "", '=IF(st_one<>1,"N.A.",IF(sl_ltb="SI","OK","ATTENZIONE: verificare stabilita\' flesso-torsionale EN 1993-1-1 6.3.2"))', "", "", "out", None)
s.section("SLU - TENSIONI ELASTICHE (metodo n, per fasi)  - vincolante per sc.2 con criterio ELASTICA")
s.add("vt_sa", "Tensione lembo inferiore acciaio", "sigma,a",
      "=st_one*m_gone*1000000/a_wi+(c_gone*m_gone+c_gtwo*m_gtwo)*1000000*(cs_htot-ell_x)/ell_j+m_q*1000000*(cs_htot-els_x)/els_j", "MPa",
      "= M_G1/Wa (sc.1 non puntellata) + (quote di G1, G2 sulla composta) / Wi(nL) + M_Q / Wi(n0). Sc.2 'SOLO SUCCESSIVI': il profilo non e' sollecitato dai carichi gia' presenti.", "calc", "0.0")
s.add("vt_sc", "Tensione lembo superiore cls", "sigma,c",
      "=(c_gone*m_gone+c_gtwo*m_gtwo)*1000000*ell_x/(n_lt*ell_j)+m_q*1000000*els_x/(n_zero*els_j)", "MPa", "= M x / (n Jid) per ciascuna aliquota applicata alla composta (sc.2: escluse le tensioni gia' presenti nella trave esistente)", "calc", "0.00")
s.add("vt_rsa", "Rapporto sigma,a / fyd", "", "=vt_sa/fyd_a", "-", "", "out", "0.000")
s.add("vt_rsc", "Rapporto sigma,c / fcd", "", "=vt_sc/fcd_c", "-", "", "out", "0.000")
s.add("vt_esa", "Esito tensione acciaio", "", esito("vt_rsa", 'AND(sc_n=2,an_crit="ELASTICA")'), "", "Per sc.1 e sc.2-PLASTICA i valori sono informativi (vale la verifica plastica).", "out", None)
s.add("vt_esc", "Esito tensione calcestruzzo", "", esito("vt_rsc", 'AND(sc_n=2,an_crit="ELASTICA")'), "",
      "Sc.2: la tensione nel cls si somma a quella gia' presente per i carichi esistenti: verificare la trave c.a. a parte.", "out", None)
s.section("SLE - DEFORMABILITA' (trave appoggiata, carichi caratteristici)")
s.add("vt_fone", "Freccia da G1", "f,G1", "=IF(st_one=1,5*ln_gone*lu_mm^4/(384*e_acc*a_iy),c_gone*5*ln_gone*lu_mm^4/(384*e_acc*ell_j))", "mm",
      "Sc.1 non puntellata: sulla sola trave in acciaio (Ia); puntellata: sulla composta (nL). Sc.2: 0 se 'SOLO SUCCESSIVI' (gia' avvenuta), sulla composta se 'TUTTI'.", "calc", "0.00")
s.add("vt_ftwo", "Freccia da G2", "f,G2", "=c_gtwo*5*ln_gtwo*lu_mm^4/(384*e_acc*ell_j)", "mm", "Sulla composta con nL (0 se G2 gia' presente e non trasferito, sc.2).", "calc", "0.00")
s.add("vt_fq", "Freccia da Q", "f,Q", "=5*ln_q*lu_mm^4/(384*e_acc*els_j)", "mm", "Sulla composta con n0.", "calc", "0.00")
s.add("vt_ftot", "Freccia totale (incremento per sc.2)", "f,tot", "=vt_fone+vt_ftwo+vt_fq", "mm", "Non include contromonta.", "out", "0.00")
s.add("vt_rft", "Rapporto f,tot / (L/lim)", "", "=vt_ftot/(lu_mm/lim_tot)", "-", "", "out", "0.000")
s.add("vt_eft", "Esito freccia totale", "", esito("vt_rft"), "", "", "out", None)
s.add("vt_rfq", "Rapporto f,Q / (L/lim)", "", "=vt_fq/(lu_mm/lim_q)", "-", "", "out", "0.000")
s.add("vt_efq", "Esito freccia variabili", "", esito("vt_rfq"), "", "", "out", None)
s.add("vt_fqp", "Freccia quasi permanente (informativa)", "", "=vt_fone+vt_ftwo+psi_two*vt_fq", "mm", "G1 + G2 + psi2 Q", "calc", "0.00")
add_cf(ws, "C8:C80")
ws4 = s.ws

# =====================================================================================
# 05A PIOLI
# =====================================================================================
s = SB("05A_PIOLI", "05A - SCENARIO 1: CONNETTORI A PIOLO (Nelson) - resistenza, numero, grado di connessione, dettagli, armatura trasversale",
       "Compilare solo se SCENARIO = 1. Resistenza del piolo (NTC 4.3.4.3.1.2 / EC4 6.6.3.1, dispensa eq. 7.3-7.4) ridotta per lamiera grecata (EC4 6.6.4). "
       "Numero di pioli calcolato tra appoggio e mezzeria per la forza Ncf (connessione completa) con pioli duttili a passo costante.")
ws = s.ws
s.section("DATI PIOLI")
dv = s.dv_list(["13", "16", "19", "22", "25"], "Diametro", "Diametro gambo piolo [mm]. Con lamiera e saldatura passante d <= 20 mm.")
s.add("st_d", "Diametro gambo piolo", "d", 19, "mm", "Diametri commerciali 13-16-19-22-25 mm. Saldatura passante attraverso lamiera: d <= 20 mm (EC4 6.6.5.8(3)). EC4 6.6.3.1: 16 <= d <= 25.", "in", "0", dv)
s.add("st_h", "Altezza nominale piolo dopo la saldatura", "hsc", 100, "mm", "Altezza complessiva dopo saldatura (circa 5 mm meno della lunghezza nominale). Con lamiera: hsc >= hp + 2d.", "in", "0")
s.add("st_fu", "Resistenza a rottura acciaio piolo", "fu", 450, "MPa", "Tipicamente 450 MPa (S235J2+C450). Limitata a 500 MPa (450 con lamiera trasversale) nel calcolo.", "in", "0")
dv = s.dv_list(["1", "2"], "Pioli per fila", "Numero di pioli per nervatura/fila trasversale.")
s.add("st_nr", "Pioli per nervatura / per fila", "nr", 1, "-", "1 o 2 (EC4 kt per nr <= 2).", "in", "0", dv)
s.add("st_et", "Interasse trasversale tra pioli della stessa fila", "et", 0, "mm", "Solo se nr = 2: >= 4d (con lamiera) o 2,5d (soletta piena). Con nr = 1 lasciare 0.", "in", "0")
s.add("st_s", "Passo longitudinale delle file di pioli", "s", 150, "mm", "Con lamiera trasversale: multiplo del passo delle greche (es. 150, 300 mm). Limiti: s >= 5d; s <= min(6h ; 800).", "in", "0")
s.add("st_gv", "gamma V", "gV", 1.25, "-", "NTC 4.3.4.3.1.2 / EC4: 1,25.", "in", "0.00")
s.section("RESISTENZA DEL SINGOLO PIOLO")
s.add("st_fue", "fu di calcolo (limitata)", "fu,d", '=IF(AND(dk_or="TRASVERSALE",dk_hp>0),MIN(st_fu,450),MIN(st_fu,500))', "MPa", "EC4 6.6.3.1 (<=500) e 6.6.4.2(2) (<=450 con lamiera trasversale).", "calc", "0")
s.add("st_hd", "Rapporto hsc/d", "", "=st_h/st_d", "-", ">= 3 obbligatorio.", "calc", "0.00")
s.add("st_alfa", "Coefficiente alfa", "alfa", "=IF(st_hd>4,1,0.2*(st_hd+1))", "-", "= 0,2 (hsc/d + 1) per 3<=hsc/d<=4; = 1 per hsc/d > 4", "calc", "0.000")
s.add("st_pa", "Resistenza lato acciaio", "PRd,a", "=0.8*st_fue*PI()*st_d^2/4/st_gv/1000", "kN", "= 0,8 fu (pi d^2/4) / gV  (dispensa eq. 7.3)", "calc", "0.00")
s.add("st_pc", "Resistenza lato calcestruzzo", "PRd,c", "=0.29*st_alfa*st_d^2*SQRT(new_fck*ecm_c)/st_gv/1000", "kN", "= 0,29 alfa d^2 sqrt(fck Ecm) / gV  (dispensa eq. 7.4)", "calc", "0.00")
s.add("st_p", "Resistenza piolo in soletta piena", "PRd", "=MIN(st_pa,st_pc)", "kN", "", "out", "0.00")
s.section("RIDUZIONE PER LAMIERA GRECATA (EC4 6.6.4)")
s.add("st_hsce", "hsc usata nelle formule", "", "=MIN(st_h,dk_hp+75)", "mm", "EC4 6.6.4.1: hsc <= hp + 75 mm.", "calc", "0")
s.add("st_kl", "Fattore kl (nervature parallele)", "kl", "=IF(dk_hp>0,MIN(1,0.6*dk_b/dk_hp*(st_hsce/dk_hp-1)),1)", "-", "EC4 eq. 6.22: 0,6 (b0/hp)(hsc/hp - 1) <= 1", "calc", "0.000")
s.add("st_ktm", "kt,max", "", '=IF(dk_thr="SI",INDEX(tb_ktp,MATCH(st_nr&"_"&IF(dk_t<=1,"<=1",">1"),tb_ktk,0)),INDEX(tb_ktf,MATCH(st_nr&"_"&IF(dk_t<=1,"<=1",">1"),tb_ktk,0)))', "-", "EC4 Tab. 6.2 (foglio DB).", "calc", "0.00")
s.add("st_kt", "Fattore kt (nervature trasversali)", "kt", "=IF(dk_hp>0,MIN(st_ktm,0.7/SQRT(st_nr)*dk_b/dk_hp*(st_hsce/dk_hp-1)),1)", "-", "EC4 eq. 6.23: 0,7/sqrt(nr) (b0/hp)(hsc/hp - 1) <= kt,max", "calc", "0.000")
s.add("st_k", "Fattore di riduzione adottato", "k", '=IF(OR(dk_or="SOLETTA PIENA",dk_hp=0),1,IF(dk_or="TRASVERSALE",st_kt,st_kl))', "-", "", "calc", "0.000")
s.add("st_pr", "RESISTENZA DI PROGETTO DEL PIOLO", "PRd,red", "=st_k*st_p", "kN", "= k PRd", "out", "0.00")
s.section("NUMERO DI PIOLI E GRADO DI CONNESSIONE")
s.add("st_nf", "Pioli necessari per connessione completa (meta' luce)", "nf", "=cs_ncf/st_pr", "-", "= Ncf / PRd,red  (dispensa: np = V / PRd)", "out", "0.0")
s.add("st_np", "Pioli disposti tra appoggio e mezzeria", "n", "=st_nr*INT(lu_mm/2/st_s)", "-", "= nr x int(L/2 / s)", "out", "0")
s.add("st_eta", "Grado di connessione", "eta", "=st_np*st_pr/cs_ncf", "-", ">= 1 connessione completa; < 1 parziale (ammessa se >= eta,min, foglio 04).", "out", "0.000")
s.add("st_tot", "Pioli totali sulla trave (indicativo)", "", "=st_nr*(INT(lu_mm/st_s)+1)", "-", "", "calc", "0")
s.add("st_v", "Flusso di scorrimento elastico all'appoggio (info)", "vEd", "=vt_flow", "N/mm", "Da foglio 04. Con pioli duttili il passo costante e' ammesso (EC4 6.6.1.3(3)): valore informativo.", "calc", "0.0")
s.add("st_rv", "Rapporto flusso / capacita' locale (info)", "", "=st_v*st_s/(st_nr*st_pr*1000)", "-", "", "calc", "0.000")
s.section("DETTAGLI COSTRUTTIVI (EC4 6.6.5)")
chk = [
    ("st_c1", "hsc >= 3d", '=IF(sc_n<>1,"N.A.",IF(st_hd>=3,"OK","NON VERIFICATA"))', "EC4 6.6.5.7(1)"),
    ("st_c2", "hsc >= hp + 2d (con lamiera)", '=IF(sc_n<>1,"N.A.",IF(OR(dk_hp=0,st_h>=dk_hp+2*st_d),"OK","NON VERIFICATA"))', "EC4 6.6.5.8(1): il piolo deve sporgere almeno 2d sopra la lamiera."),
    ("st_c3", "d <= 20 mm con saldatura passante", '=IF(sc_n<>1,"N.A.",IF(OR(dk_hp=0,dk_thr="NO",st_d<=20),"OK","NON VERIFICATA"))', "EC4 6.6.5.8(3)"),
    ("st_c4", "d <= 2,5 tf (piolo non in asse anima)", '=IF(sc_n<>1,"N.A.",IF(st_d<=2.5*a_ts,"OK","ATTENZIONE: ammesso solo se il piolo e\' in asse all\'anima"))', "EC4 6.6.5.7(5)"),
    ("st_c5", "Passo s >= 5d", '=IF(sc_n<>1,"N.A.",IF(st_s>=5*st_d,"OK","NON VERIFICATA"))', "EC4 6.6.5.7(4)"),
    ("st_c6", "Passo s <= min(6 h ; 800 mm)", '=IF(sc_n<>1,"N.A.",IF(st_s<=MIN(6*sl_h,800),"OK","NON VERIFICATA"))', "EC4 6.6.5.5(3)"),
    ("st_c7", "Interasse trasversale et (nr=2)", '=IF(sc_n<>1,"N.A.",IF(OR(st_nr=1,st_et>=IF(dk_hp>0,4,2.5)*st_d),"OK","NON VERIFICATA"))', "EC4 6.6.5.7(4): 2,5d soletta piena, 4d altrimenti."),
    ("st_c8", "Distanza bordo piolo - bordo ala >= 20 mm", '=IF(sc_n<>1,"N.A.",IF((a_bs-st_et-st_d)/2>=20,"OK","NON VERIFICATA"))', "EC4 6.6.5.6(2)"),
    ("st_c9", "Ala classe 1 per vincolo connettori: s <= 22 tf eps", '=IF(sc_n<>1,"N.A.",IF(st_s<=22*a_ts*eps_a,"OK","ATTENZIONE: ala compressa non stabilizzata dai pioli"))', "EC4 6.6.5.5(2) soletta piena; con lamiera 15 tf eps."),
]
for nm, lab, f, ref in chk:
    s.add(nm, lab, "", f, "", ref, "out", None)
c_first = s.r - len(chk)
add_cf(ws, f"C{c_first}:C{s.r - 1}")
s.section("ARMATURA TRASVERSALE - SCORRIMENTO NELLA SOLETTA (EC4 6.6.6, EC2 6.2.4)")
s.add("tr_df", "Forza da trasferire su un lato (piano a-a)", "dF", "=MIN(1,st_eta)*cs_ncf*MAX(sl_be1,sl_be2)/sl_beff", "kN", "= eta Ncf x be,i / beff (lato piu' largo)", "calc", "0.0")
s.add("tr_v", "Tensione tangenziale sul piano a-a", "vEd", "=tr_df*1000/(sl_hc*lu_mm/2)", "MPa", "= dF / (hc x L/2); contributo lamiera trascurato (come nella dispensa).", "calc", "0.000")
s.add("tr_req", "Armatura trasversale necessaria", "Asf/sf,req", "=tr_v*sl_hc/(fyd_s*sl_cot)*1000", "mm2/m", "= vEd hc / (fyd cot theta)", "out", "0")
s.add("tr_r", "Rapporto richiesta / disposta", "", "=tr_req/sl_ast", "-", "", "out", "0.000")
s.add("tr_e", "Esito armatura trasversale", "", esito("tr_r", "sc_n=1"), "", "", "out", None)
s.add("tr_vmax", "Resistenza bielle compresse", "vRd,max", "=0.6*(1-new_fck/250)*fcd_c*sl_cot/(1+sl_cot^2)", "MPa", "= nu fcd sin(theta) cos(theta), nu = 0,6 (1 - fck/250)", "calc", "0.000")
s.add("tr_rb", "Rapporto vEd / vRd,max", "", "=tr_v/tr_vmax", "-", "", "out", "0.000")
s.add("tr_eb", "Esito bielle", "", esito("tr_rb", "sc_n=1"), "", "", "out", None)
add_cf(ws, f"C{s.r - 4}:C{s.r - 1}")
ws5a = s.ws

# =====================================================================================
# helper: blocco resistenze ancorante (usato in 05B e 06)
# =====================================================================================
def anchor_inputs(s, p, defaults):
    dvb = s.dv_list([r[0] for r in bar_rows], "Barra", "Diametro barra filettata.")
    s.add(f"{p}_m", "Barra filettata", "", defaults["m"], "", "Diametro nominale. Il diametro deve essere coperto dall'ETA dell'ancorante chimico.", "in", None, dvb)
    dvc = s.dv_list([r[0] for r in cl_rows], "Classe", "Classe di resistenza della barra.")
    s.add(f"{p}_cl", "Classe acciaio barra", "", defaults["cl"], "", "5.8 / 8.8 (zincate) o A4-70 inox. Classi con allungamento A >= 8% (duttili, k7 = 1).", "in", None, dvc)
    s.add(f"{p}_hef", "Profondita' di ancoraggio effettiva", "hef", defaults["hef"], "mm", "Da progetto, entro il campo ammesso dall'ETA (tipicamente 8d - 20d).", "in", "0")
    s.add(f"{p}_tau", "Tensione di aderenza caratteristica (cls fessurato)", "tauRk", defaults["tau"], "MPa",
          "DA ETA del prodotto: tau_Rk,cr per il diametro, la classe del cls (gia' moltiplicata per psi_c) e il campo di temperatura. Valore di esempio: SOSTITUIRE.", "in", "0.00")
    s.add(f"{p}_psus", "Fattore carichi permanenti", "psi_sus", 1.0, "-", "EN 1992-4 7.2.1.6: 1,0 se alfa_sus <= psi0_sus (da ETA); altrimenti psi_sus = psi0_sus + 1 - alfa_sus.", "in", "0.00")
    s.add(f"{p}_gin", "gamma inst (trazione)", "g,inst", 1.2, "-", "Da ETA (1,0 / 1,2 / 1,4). gamma_Mc = gamma_Mp = 1,5 gamma_inst a trazione; a taglio gamma_inst = 1,0.", "in", "0.00")
    s.add(f"{p}_k8", "Fattore pry-out", "k8", 2.0, "-", "Da ETA (di norma 2,0).", "in", "0.00")
    dvf = s.dv_list(["FESSURATO", "NON FESSURATO"], "Cls", "Stato del calcestruzzo nella zona di ancoraggio.")
    s.add(f"{p}_cr", "Stato del calcestruzzo", "", "FESSURATO", "", "FESSURATO a favore di sicurezza (zone tese/intradosso). NON FESSURATO solo se dimostrato (EN 1992-4 4.6).", "in", None, dvf)
    dvr = s.dv_list(["STANDARD", "ARMATURA DI BORDO"], "psi re,V", "STANDARD = 1,0. ARMATURA DI BORDO = 1,4 (barra di bordo >= 12 mm e staffe a passo <= 100 mm).")
    s.add(f"{p}_prv", "Fattore armatura di bordo (taglio)", "psi_re,V", "STANDARD", "", "EN 1992-4 7.2.2.5: STANDARD = 1,0; 1,4 se c'e' armatura di bordo phi>=12 con staffe a passo <= 100 mm.", "in", None, dvr)
    dvn = s.dv_list(["SI", "NO"], "Armatura fitta", "Armatura con passo < 150 mm (o < 100 per phi<=10) nella zona di ancoraggio?")
    s.add(f"{p}_rfn", "Armatura fitta nella zona di ancoraggio (shell spalling)?", "", "SI", "", "EN 1992-4 7.2.1.4(5): SI -> psi_re,N = 0,5 + hef/200 <= 1 (zone con staffe ravvicinate). NO -> 1,0.", "in", None, dvn)
    s.add(f"{p}_cmin", "Distanza minima dal bordo (ETA)", "cmin", defaults["cmin"], "mm", "Da ETA del prodotto per il diametro scelto.", "in", "0")
    s.add(f"{p}_smin", "Interasse minimo (ETA)", "smin", defaults["smin"], "mm", "Da ETA del prodotto per il diametro scelto.", "in", "0")
    s.add(f"{p}_d0", "Diametro foro nel piatto/ala in acciaio", "d0", defaults["d0"], "mm", "Foro di passaggio nell'acciaio (es. d + 2). I fori vanno riempiti (resina/rondella di riempimento) per far collaborare tutti gli ancoranti a taglio.", "in", "0")
    s.add(f"{p}_am", "Coefficiente vincolo per braccio di leva", "alfaM", 1.0, "-", "EN 1992-4 6.2.2.3: 1,0 rotazione libera (a favore di sicurezza); 2,0 piatto impedito di ruotare.", "in", "0.00")

def anchor_steel(s, p, gap, grok, nEd):
    s.section("RESISTENZE ACCIAIO DELLA BARRA (EN 1992-4 7.2.1.3, 7.2.2.3)")
    s.add(f"{p}_d", "Diametro barra", "d", f"=INDEX(tb_bard,MATCH({p}_m,tb_bar,0))", "mm", "", "calc", "0")
    s.add(f"{p}_as", "Area resistente", "As", f"=INDEX(tb_baras,MATCH({p}_m,tb_bar,0))", "mm2", "", "calc", "0.0")
    s.add(f"{p}_fu", "fuk barra", "fuk", f"=INDEX(tb_bclfu,MATCH({p}_cl,tb_bcl,0))", "MPa", "", "calc", "0")
    s.add(f"{p}_fy", "fyk barra", "fyk", f"=INDEX(tb_bclfy,MATCH({p}_cl,tb_bcl,0))", "MPa", "", "calc", "0")
    s.add(f"{p}_gmn", "gamma Ms,N", "", f"=MAX(1.4,1.2*{p}_fu/{p}_fy)", "-", "= 1,2 fuk/fyk >= 1,4", "calc", "0.000")
    s.add(f"{p}_gmv", "gamma Ms,V", "", f"=IF(AND({p}_fu<=800,{p}_fy/{p}_fu<=0.8),MAX(1.25,{p}_fu/{p}_fy),1.5)", "-", "= fuk/fyk >= 1,25 (fuk<=800 e fyk/fuk<=0,8), altrimenti 1,5", "calc", "0.000")
    s.add(f"{p}_nrs", "Resistenza a trazione acciaio", "NRd,s", f"={p}_as*{p}_fu/{p}_gmn/1000", "kN", "= As fuk / gMs,N", "out", "0.00")
    s.add(f"{p}_lev", "Braccio di leva presente? (1=si)", "", f"=IF(AND({gap}>0,OR({gap}>{p}_d/2,{grok}=\"NO\")),1,0)", "-", "Si' se spessore malta > d/2 oppure malta < 30 MPa.", "calc", "0")
    s.add(f"{p}_vso", "Resistenza a taglio senza braccio di leva", "VRk,s", f"=IF({p}_fu<=500,0.6,0.5)*{p}_as*{p}_fu/1000", "kN", "= k6 As fuk; k6 = 0,6 (fuk<=500) / 0,5 (500<fuk<=1000); k7 = 1 (acciaio duttile)", "calc", "0.00")
    s.add(f"{p}_mrk", "Momento caratteristico ridotto", "MRk,s", f"=1.2*PI()*(SQRT(4*{p}_as/PI()))^3/32*{p}_fu*MAX(0,1-{nEd}/{p}_nrs)/1000", "kNmm",
          "= 1,2 Wel fuk (1 - NEd/NRd,s), Wel con diametro equivalente all'area resistente", "calc", "0.00")
    s.add(f"{p}_vsm", "Resistenza a taglio con braccio di leva", "VRk,s,M", f"={p}_am*{p}_mrk/(0.5*{p}_d+{gap})", "kN", "= alfaM MRk,s / l ; l = 0,5 d + tg", "calc", "0.00")
    s.add(f"{p}_vrs", "Resistenza a taglio acciaio di progetto", "VRd,s", f"=IF({p}_lev=1,MIN({p}_vso,{p}_vsm),{p}_vso)/{p}_gmv", "kN", "", "out", "0.00")

def anchor_conc_consts(s, p):
    s.add(f"{p}_k1", "k1 cono (fessurato/non fess.)", "k1", f'=IF({p}_cr="FESSURATO",7.7,11)', "-", "EN 1992-4 7.2.1.4: 7,7 fessurato; 11,0 non fessurato.", "calc", "0.0")
    s.add(f"{p}_k9", "k9 rottura bordo", "k9", f'=IF({p}_cr="FESSURATO",1.7,2.4)', "-", "EN 1992-4 7.2.2.5: 1,7 fessurato; 2,4 non fessurato.", "calc", "0.0")
    s.add(f"{p}_gmc", "gamma Mc = gamma Mp (trazione)", "", f"=1.5*{p}_gin", "-", "= gc x gamma_inst", "calc", "0.00")
    s.add(f"{p}_prn", "psi re,N", "", f'=IF({p}_rfn="SI",MIN(1,0.5+{p}_hef/200),1)', "-", "", "calc", "0.000")
    s.add(f"{p}_ncz", "Resistenza cono singolo ancorante isolato", "N0Rk,c", f"={p}_k1*SQRT(fck_c)*{p}_hef^1.5/1000", "kN", "= k1 sqrt(fck) hef^1,5", "calc", "0.00")
    s.add(f"{p}_scr", "Interasse critico cono", "scr,N", f"=3*{p}_hef", "mm", "= 3 hef ; ccr,N = 1,5 hef", "calc", "0")
    s.add(f"{p}_npz", "Resistenza aderenza singolo ancorante isolato", "N0Rk,p", f"={p}_psus*{p}_tau*PI()*{p}_d*{p}_hef/1000", "kN", "= psi_sus tauRk pi d hef", "calc", "0.00")
    s.add(f"{p}_scp", "Interasse critico aderenza", "scr,Np", f"=MIN(3*{p}_hef,7.3*{p}_d*SQRT({p}_psus*{p}_tau))", "mm", "= 7,3 d sqrt(psi_sus tauRk) <= 3 hef ; ccr,Np = scr,Np/2", "calc", "0")

# =====================================================================================
# 05B INGHISAGGI TRAVE
# =====================================================================================
s = SB("05B_INGHISAGGI_TRAVE", "05B - SCENARIO 2: CONNESSIONE PROFILO / TRAVE C.A. CON BARRE FILETTATE INGHISATE (ancorante chimico)",
       "Compilare solo se SCENARIO = 2. Le barre attraversano l'ala superiore del profilo (fori riempiti) e sono inghisate nell'intradosso della trave esistente, su 1 o 2 file ai lati dell'anima. "
       "Verifiche per singolo ancorante secondo EN 1992-4 (acciaio, rifollamento ala, cono, aderenza, bordo laterale con taglio parallelo al bordo, pry-out, interazione N-V).")
ws = s.ws
s.section("CRITERIO DI PROGETTO DELLA CONNESSIONE")
dv = s.dv_list(["ELASTICA", "PLASTICA"], "Criterio", "ELASTICA consigliata per ancoranti non duttili.")
s.add("an_crit", "Criterio di progetto", "", "ELASTICA", "",
      "ELASTICA (consigliato): gli ancoranti trasferiscono la forza di scorrimento elastica F = MEd S/Jid e il flusso VEd S/Jid; la flessione si verifica a tensioni (foglio 04). "
      "PLASTICA: gli ancoranti devono trasferire la piena forza plastica Ncf (eta >= 1) e la flessione si verifica con Mpl,Rd. Connessione parziale non ammessa (ancoranti non duttili, EC4 6.6.1.1).", "in", None, dv)
s.section("DATI ANCORANTE")
anchor_inputs(s, "an", {"m": "M16", "cl": "8.8", "hef": 160, "tau": 7.0, "cmin": 80, "smin": 80, "d0": 18})
s.section("DISPOSIZIONE")
dv = s.dv_list(["1", "2"], "File", "Numero di barre per sezione trasversale.")
s.add("an_nr", "Barre per sezione trasversale (file)", "nr", 2, "-", "2 (consigliato): una barra per lato dell'anima. 1: barre sfalsate alternativamente ai due lati (indicare in et lo sfalsamento).", "in", "0", dv)
s.add("an_et", "Interasse trasversale tra le file", "et", 120, "mm", "Distanza tra le due barre di una sezione. Deve lasciare spazio a anima + raccordi + dado: et >= tw + 2r + 2d (controllo sotto).", "in", "0")
s.add("an_s", "Passo longitudinale", "s", 150, "mm", "Passo costante delle sezioni di ancoraggio lungo la trave. Vedere 'passo massimo ammissibile' nei risultati.", "in", "0")
s.add("an_kn", "Trazione di progetto / taglio di progetto", "kN", 0.1, "-",
      "Trazione associata per contrastare il distacco (uplift) del profilo: EC4 6.6.1.1(7) e dispensa par. 7 (1/10 del taglio). Aumentare se presenti carichi appesi al profilo.", "in", "0.00")
s.section("GEOMETRIA DERIVATA")
s.add("an_c", "Distanza dal bordo laterale della trave", "c1", "=(ex_bw-(an_nr-1)*an_et)/2", "mm", "= (bw - (nr-1) et)/2 (con nr = 1 le barre sfalsate hanno c1 = (bw - et)/2: inserire et).", "out", "0")
s.add("an_e2", "Distanza foro - bordo ala profilo", "e2", "=(a_bs-an_et)/2", "mm", "= (b,sup - et)/2", "calc", "0")
s.add("an_nh", "Ancoranti tra appoggio e mezzeria", "n", "=an_nr*INT(lu_mm/2/an_s)", "-", "", "out", "0")
s.section("AZIONI SUL SINGOLO ANCORANTE")
s.add("an_fel", "Forza di scorrimento elastica (mezzeria)", "Fel", "=vt_fel", "kN", "Da foglio 04: MEd,comp S / Jid (carichi applicati alla composta, vedi 02B 'Carichi che gravano sulla sezione composta').", "calc", "#,##0.0")
s.add("an_f", "Forza di scorrimento di progetto", "F", '=IF(an_crit="PLASTICA",cs_ncf,an_fel)', "kN", "PLASTICA: Ncf; ELASTICA: Fel", "out", "#,##0.0")
s.add("an_v", "Flusso di scorrimento all'appoggio", "vEd", '=IF(an_crit="PLASTICA",MAX(vt_flow,v_ed*1000*els_sint/els_j),vt_flow)', "N/mm", "ELASTICA: da foglio 04 (carichi sulla composta). PLASTICA: VEd totale S/Jid (n0), coerente con la verifica a Mpl.", "calc", "0.0")
s.add("an_vg", "Taglio per ancorante da forza globale", "VEd,1", "=an_f/an_nh", "kN", "= F / n", "calc", "0.00")
s.add("an_vl", "Taglio per ancorante da flusso locale all'appoggio", "VEd,2", "=an_v*an_s/an_nr/1000", "kN", "= vEd s / nr (distribuzione elastica: ancoranti non duttili)", "calc", "0.00")
s.add("an_ved", "TAGLIO DI PROGETTO PER ANCORANTE", "VEd", "=MAX(an_vg,an_vl)", "kN", "", "out", "0.00")
s.add("an_ned", "TRAZIONE DI PROGETTO PER ANCORANTE", "NEd", "=an_kn*an_ved", "kN", "= kN x VEd", "out", "0.00")
anchor_steel(s, "an", "ex_gap", "ex_grok", "an_ned")
s.add("an_kb", "k1 rifollamento", "k1", "=MAX(0,MIN(2.5,2.8*an_e2/an_d0-1.7))", "-", "EN 1993-1-8 Tab. 3.4: min(2,8 e2/d0 - 1,7 ; 2,5)", "calc", "0.000")
s.add("an_ab", "alfa b rifollamento", "ab", "=MAX(0,MIN(1,an_fu/ftk_a,an_s/(3*an_d0)-0.25))", "-", "= min(p1/3d0 - 1/4 ; fub/fu ; 1)", "calc", "0.000")
s.add("an_fb", "Rifollamento ala superiore del profilo", "Fb,Rd", "=an_kb*an_ab*ftk_a*an_d*a_ts/gam_mt/1000", "kN", "= k1 ab fu d tf / gM2", "out", "0.00")
s.section("RESISTENZE CALCESTRUZZO (EN 1992-4 7.2.1.4, 7.2.1.6, 7.2.2.4, 7.2.2.5)")
anchor_conc_consts(s, "an")
s.add("an_rcn", "Rapporto aree cono per ancorante", "Ac,N/A0c,N",
      "=MIN(an_s,an_scr)*IF(an_nr=2,MIN(an_c,an_scr/2)+MIN(an_et/2,an_scr/2),2*MIN(an_c,an_scr/2))/an_scr^2", "-",
      "Area proiettata di competenza del singolo ancorante: min(s;scr) x [min(c1;ccr) + min(et/2;ccr)] / scr^2", "calc", "0.000")
s.add("an_psn", "psi s,N", "", "=MIN(1,0.7+0.3*an_c/(an_scr/2))", "-", "= 0,7 + 0,3 c/ccr,N <= 1", "calc", "0.000")
s.add("an_nrkc", "Resistenza cono caratteristica", "NRk,c", "=an_ncz*an_rcn*an_psn*an_prn", "kN", "", "calc", "0.00")
s.add("an_nrc", "Resistenza cono di progetto", "NRd,c", "=an_nrkc/an_gmc", "kN", "", "out", "0.00")
s.add("an_rpn", "Rapporto aree aderenza per ancorante", "Ap,N/A0p,N",
      "=MIN(an_s,an_scp)*IF(an_nr=2,MIN(an_c,an_scp/2)+MIN(an_et/2,an_scp/2),2*MIN(an_c,an_scp/2))/an_scp^2", "-", "", "calc", "0.000")
s.add("an_psp", "psi s,Np", "", "=MIN(1,0.7+0.3*an_c/(an_scp/2))", "-", "", "calc", "0.000")
s.add("an_nrkp", "Resistenza aderenza/sfilamento caratteristica", "NRk,p", "=an_npz*an_rpn*an_psp*an_prn", "kN", "psi_g,Np = 1 (a favore di sicurezza)", "calc", "0.00")
s.add("an_nrp", "Resistenza aderenza di progetto", "NRd,p", "=an_nrkp/an_gmc", "kN", "", "out", "0.00")
s.add("an_lf", "Lunghezza efficace a taglio", "lf", "=IF(an_d<=24,MIN(an_hef,12*an_d),MIN(an_hef,MAX(8*an_d,300)))", "mm", "", "calc", "0")
s.add("an_vcz", "Resistenza bordo di base", "V0Rk,c",
      "=an_k9*an_d^(0.1*(an_lf/an_c)^0.5)*an_lf^(0.1*(an_d/an_c)^0.2)*SQRT(fck_c)*an_c^1.5/1000", "kN", "= k9 d^a lf^b sqrt(fck) c1^1,5 ; a = 0,1 (lf/c1)^0,5 ; b = 0,1 (d/c1)^0,2", "calc", "0.00")
s.add("an_rcv", "Rapporto aree bordo per ancorante", "Ac,V/A0c,V", "=MIN(an_s,3*an_c)*MIN(1.5*an_c,ex_h)/(4.5*an_c^2)", "-", "= min(s;3c1) x min(1,5c1;H) / 4,5 c1^2", "calc", "0.000")
s.add("an_phv", "psi h,V", "", "=MAX(1,SQRT(1.5*an_c/ex_h))", "-", "", "calc", "0.000")
s.add("an_vrc", "Resistenza bordo di progetto (taglio parallelo al bordo)", "VRd,c", "=an_vcz*an_rcv*an_phv*2*IF(an_prv=\"STANDARD\",1,1.4)/1.5", "kN",
      "psi_alfa,V = 2 (taglio parallelo al bordo laterale: sqrt(1/(cos^2 90 + (0,5 sin 90)^2))); psi_s,V = 1; gMc = 1,5", "out", "0.00")
s.add("an_vrcp", "Resistenza pry-out di progetto", "VRd,cp", "=an_k8*MIN(an_nrkc,an_nrkp)/1.5", "kN", "= k8 min(NRk,c ; NRk,p) / 1,5", "out", "0.00")
s.section("RISULTATI E VERIFICHE PER ANCORANTE")
s.add("an_vrd", "RESISTENZA A TAGLIO DELL'ANCORANTE", "VRd", "=MIN(an_vrs,an_fb,an_vrc,an_vrcp)", "kN", "Minimo tra acciaio, rifollamento ala, bordo, pry-out.", "out", "0.00")
s.add("an_vmode", "Meccanismo governante a taglio", "", '=IF(an_vrd=an_vrs,"acciaio barra",IF(an_vrd=an_fb,"rifollamento ala profilo",IF(an_vrd=an_vrc,"rottura bordo cls","pry-out")))', "", "", "out", None)
s.add("an_nrd", "RESISTENZA A TRAZIONE DELL'ANCORANTE", "NRd", "=MIN(an_nrs,an_nrc,an_nrp)", "kN", "", "out", "0.00")
s.add("an_eta", "Grado di connessione (forza plastica)", "eta", "=an_nh*an_vrd/cs_ncf", "-", "= n VRd / Ncf (con criterio PLASTICA deve essere >= 1).", "out", "0.000")
s.add("an_smax", "Passo massimo ammissibile", "s,max", "=an_nr*an_vrd*1000/MAX(an_v,an_f*1000/(lu_mm/2))", "mm", "= nr VRd / max(vEd ; F/(L/2)) : passo da non superare.", "out", "0")
s.add("an_rv", "Rapporto taglio VEd/VRd", "", "=an_ved/an_vrd", "-", "", "out", "0.000")
s.add("an_ev", "Esito taglio ancorante", "", esito("an_rv", "sc_n=2"), "", "", "out", None)
s.add("an_rn", "Rapporto trazione NEd/NRd", "", "=an_ned/an_nrd", "-", "", "out", "0.000")
s.add("an_en", "Esito trazione ancorante", "", esito("an_rn", "sc_n=2"), "", "", "out", None)
s.add("an_is", "Interazione acciaio (N/NRd,s)^2+(V/VRd,s)^2", "", "=(an_ned/an_nrs)^2+(an_ved/an_vrs)^2", "-", "EN 1992-4 Tab. 7.3 eq. 7.54", "out", "0.000")
s.add("an_eis", "Esito interazione acciaio", "", esito("an_is", "sc_n=2"), "", "", "out", None)
s.add("an_ic", "Interazione cls bN^1,5+bV^1,5", "", "=(an_ned/MIN(an_nrc,an_nrp))^1.5+(an_ved/MIN(an_vrc,an_vrcp))^1.5", "-", "EN 1992-4 Tab. 7.3 eq. 7.55 (rapporti massimi tra i modi di rottura del cls)", "out", "0.000")
s.add("an_eic", "Esito interazione calcestruzzo", "", esito("an_ic", "sc_n=2"), "", "", "out", None)
s.add("an_ee", "Esito grado di connessione (solo PLASTICA)", "", '=IF(OR(sc_n<>2,an_crit<>"PLASTICA"),"N.A.",IF(an_eta>=1,"VERIFICATA","NON VERIFICATA"))', "", "", "out", None)
s.section("CONTROLLI GEOMETRICI")
chk = [
    ("an_g1", "c1 >= cmin (ETA)", '=IF(sc_n<>2,"N.A.",IF(an_c>=an_cmin,"OK","NON VERIFICATA"))', "Distanza dal bordo laterale della trave."),
    ("an_g2", "s >= smin (ETA)", '=IF(sc_n<>2,"N.A.",IF(an_s>=an_smin,"OK","NON VERIFICATA"))', ""),
    ("an_g3", "et >= smin (ETA) se 2 file", '=IF(sc_n<>2,"N.A.",IF(OR(an_nr=1,an_et>=an_smin),"OK","NON VERIFICATA"))', ""),
    ("an_g4", "Spazio per dado: et >= tw + 2r + 2d", '=IF(sc_n<>2,"N.A.",IF(an_et>=a_tw+2*a_r+2*an_d,"OK","ATTENZIONE: spazio insufficiente per dado/chiave"))', "Indicativo: verificare con le dimensioni di dado e rondella."),
    ("an_g5", "e2 >= 1,2 d0 (EN 1993-1-8)", '=IF(sc_n<>2,"N.A.",IF(an_e2>=1.2*an_d0,"OK","NON VERIFICATA"))', ""),
    ("an_g6", "Spessore residuo: H - hef >= max(30 ; 2 d0)", '=IF(sc_n<>2,"N.A.",IF(ex_h-an_hef>=MAX(30,2*(an_d+4)),"OK","NON VERIFICATA"))', "EN 1992-4 / ETA hmin: verificare il valore esatto di hmin nell'ETA."),
    ("an_g7", "Passo s <= s,max", '=IF(sc_n<>2,"N.A.",IF(an_s<=an_smax,"OK","NON VERIFICATA"))', ""),
    ("an_g8", "Ala classe 1: s <= 22 tf eps", '=IF(sc_n<>2,"N.A.",IF(an_s<=22*a_ts*eps_a,"OK","ATTENZIONE: ala compressa non stabilizzata (rilevante se xpl nel profilo)"))', "EC4 6.6.5.5(2)"),
]
for nm, lab, f, ref in chk:
    s.add(nm, lab, "", f, "", ref, "out", None)
add_cf(ws, "C8:C140")
s.text("NOTE: (1) i fori nell'ala vanno riempiti (rondella di riempimento + resina) altrimenti solo una parte degli ancoranti lavora a taglio; (2) serraggio con chiave dinamometrica alla coppia ETA; "
       "(3) rilevare staffe e armature prima di forare; (4) se il meccanismo governante e' la rottura del bordo, aumentare c1 (ridurre et), il passo, hef, oppure prevedere ancoranti piu' piccoli e numerosi.", 42, False, "9C5700")
ws5b = s.ws

# =====================================================================================
# 06 NODO PILASTRO
# =====================================================================================
s = SB("06_NODO_PILASTRO", "06 - SCENARIO 2: ATTACCO DEL PROFILO AL PILASTRO ESISTENTE CON PIASTRA E ANCORANTI CHIMICI",
       "Compilare solo se SCENARIO = 2. Piastra verticale fissata sulla faccia del pilastro con nx x ny ancoranti; il profilo trasmette la reazione d'appoggio VEd con eccentricita' e rispetto alla faccia del pilastro. "
       "Trazione nella fila superiore = N/n + M/(z nx). Verifica di gruppo per cono/aderenza (fila tesa), bordo (fila verticale vicina al bordo, taglio parallelo), pry-out (gruppo).")
ws = s.ws
s.section("AZIONI")
s.add("nd_q", "Quota del taglio d'appoggio portata dal profilo", "", 1.0, "-", "1,0 = tutto il taglio di progetto passa dal profilo al pilastro (a favore di sicurezza). Ridurre solo se giustificato.", "in", "0.00")
s.add("nd_v", "Taglio di progetto sul nodo", "VEd", '=nd_q*IF(an_crit="PLASTICA",v_ed,v_comp)', "kN", "ELASTICA: taglio dei carichi applicati alla composta (VEd,comp, foglio 01). PLASTICA: VEd totale.", "out", "0.0")
s.add("nd_e", "Eccentricita' della reazione dalla faccia del pilastro", "e", 60, "mm",
      "Distanza tra la faccia del pilastro e la linea di trasmissione del taglio (asse bulloni del fazzoletto / centro appoggio su mensola + spessore piastra).", "in", "0")
s.add("nd_n", "Sforzo normale (trazione +)", "NEd", 0.0, "kN", "Eventuale tiro/catena trasmesso al pilastro. 0 se assente.", "in", "0.0")
s.add("nd_mom", "Momento sulla piastra", "MEd", "=nd_v*nd_e/1000", "kNm", "= VEd e", "calc", "0.00")
s.section("GEOMETRIA PILASTRO E PIASTRA")
s.add("nd_cb", "Larghezza faccia del pilastro (dove si fissa la piastra)", "bc", 350, "mm", "Dimensione orizzontale della faccia del pilastro.", "in", "0")
s.add("nd_cd", "Profondita' del pilastro (direzione ancoranti)", "Dc", 400, "mm", "Dimensione nella direzione dell'asse degli ancoranti.", "in", "0")
s.add("nd_ct", "Distanza fila superiore - bordo superiore cls", "c,top", 1000, "mm", "Distanza verso l'alto dal bordo libero piu' vicino (es. fine pilastro). Se il pilastro prosegue, inserire un valore grande (>= 1,5 hef).", "in", "0")
s.add("nd_bp", "Larghezza piastra", "bp", 240, "mm", "", "in", "0")
s.add("nd_hp", "Altezza piastra", "hp", 340, "mm", "", "in", "0")
s.add("nd_tp", "Spessore piastra", "tp", 15, "mm", "Usato per il rifollamento. La flessione della piastra va verificata a parte (EN 1993-1-8, T-stub).", "in", "0")
dv = s.dv_list([r[0] for r in st_rows], "Acciaio piastra", "Grado acciaio piastra.")
s.add("nd_pg", "Acciaio piastra", "", "S275", "", "", "in", None, dv)
s.add("nd_pfu", "ftk piastra", "fu", "=INDEX(tb_stfu,MATCH(nd_pg,tb_st,0))", "MPa", "", "calc", "0")
s.add("nd_ep", "Distanza foro - bordo piastra", "e1=e2", 45, "mm", ">= 1,2 d0 (EN 1993-1-8 Tab. 3.3).", "in", "0")
s.add("nd_gap", "Spessore malta tra piastra e pilastro", "tg", 5, "mm", "Strato di regolarizzazione (malta espansiva/resina). Se > d/2 o malta < 30 MPa -> braccio di leva.", "in", "0")
dv = s.dv_list(["SI", "NO"], "Malta", "Malta con fc >= 30 MPa?")
s.add("nd_grok", "Malta con fc >= 30 MPa?", "", "SI", "", "", "in", None, dv)
s.section("DATI ANCORANTI")
anchor_inputs(s, "nd", {"m": "M16", "cl": "8.8", "hef": 160, "tau": 7.0, "cmin": 80, "smin": 80, "d0": 18})
dv = s.dv_list(["1", "2", "3", "4"], "Colonne", "Numero di ancoranti in orizzontale.")
s.add("nd_nx", "Ancoranti per fila orizzontale", "nx", 2, "-", "Numero di colonne di ancoranti.", "in", "0", dv)
dv = s.dv_list(["1", "2", "3", "4", "5"], "File", "Numero di file orizzontali.")
s.add("nd_ny", "Numero di file orizzontali", "ny", 3, "-", "Numero di righe di ancoranti.", "in", "0", dv)
s.add("nd_sx", "Interasse orizzontale", "sx", 150, "mm", "", "in", "0")
s.add("nd_sy", "Interasse verticale", "sy", 120, "mm", "", "in", "0")
s.add("nd_zm", "Braccio interno z manuale (0 = automatico)", "z", 0, "mm",
      "Automatico (a favore di sicurezza): ny>=2 -> z = distanza fila superiore - fila inferiore (compressione assunta sulla fila inferiore); ny=1 -> z = 0,9 x distanza fila-bordo inferiore piastra.", "in", "0")
s.section("GEOMETRIA DERIVATA E AZIONI SUGLI ANCORANTI")
s.add("nd_cx", "Distanza dal bordo laterale del pilastro", "cx", "=(nd_cb-(nd_nx-1)*nd_sx)/2", "mm", "Piastra centrata sulla faccia: (bc - (nx-1) sx)/2", "out", "0")
s.add("nd_dt", "Distanza fila superiore - bordo inferiore piastra", "", "=nd_hp/2+(nd_ny-1)*nd_sy/2", "mm", "", "calc", "0")
s.add("nd_z", "Braccio interno", "z", "=IF(nd_zm>0,nd_zm,IF(nd_ny>=2,(nd_ny-1)*nd_sy,0.9*nd_dt))", "mm", "", "calc", "0")
s.add("nd_nt", "Numero totale ancoranti", "n", "=nd_nx*nd_ny", "-", "", "calc", "0")
s.add("nd_ned", "TRAZIONE ancorante piu' sollecitato (fila sup.)", "NEd", "=MAX(0,nd_n/nd_nt+nd_mom*1000/(nd_z*nd_nx))", "kN", "= N/n + M/(z nx)", "out", "0.00")
s.add("nd_ved", "TAGLIO per ancorante", "VEd", "=nd_v/nd_nt", "kN", "= VEd / n (fori riempiti)", "out", "0.00")
s.add("nd_nrow", "Trazione sulla fila superiore", "Nrow", "=nd_ned*nd_nx", "kN", "", "calc", "0.00")
anchor_steel(s, "nd", "nd_gap", "nd_grok", "nd_ned")
s.add("nd_kb", "k1 rifollamento piastra", "", "=MAX(0,MIN(2.5,2.8*nd_ep/nd_d0-1.7))", "-", "", "calc", "0.000")
s.add("nd_ab", "alfa b rifollamento piastra", "", "=MAX(0,MIN(1,nd_fu/nd_pfu,nd_ep/(3*nd_d0)))", "-", "Foro di estremita' nella direzione del carico (e1).", "calc", "0.000")
s.add("nd_fb", "Rifollamento piastra", "Fb,Rd", "=nd_kb*nd_ab*nd_pfu*nd_d*nd_tp/gam_mt/1000", "kN", "", "out", "0.00")
s.section("RESISTENZE CALCESTRUZZO (gruppo)")
anchor_conc_consts(s, "nd")
s.add("nd_acr", "Area cono fila tesa / A0", "Ac,N/A0", "=(2*MIN(nd_cx,1.5*nd_hef)+(nd_nx-1)*MIN(nd_sx,3*nd_hef))*(MIN(nd_ct,1.5*nd_hef)+1.5*nd_hef)/(9*nd_hef^2)", "-",
      "Fila superiore di nx ancoranti: [2 min(cx;ccr) + (nx-1) min(sx;scr)] x [min(ctop;ccr) + ccr] / scr^2", "calc", "0.000")
s.add("nd_psn", "psi s,N", "", "=MIN(1,0.7+0.3*MIN(nd_cx,nd_ct)/(1.5*nd_hef))", "-", "", "calc", "0.000")
s.add("nd_nrkc", "Cono fila tesa caratteristico", "NRk,c", "=nd_ncz*nd_acr*nd_psn*nd_prn", "kN", "", "calc", "0.00")
s.add("nd_nrc", "Cono fila tesa di progetto", "NRd,c", "=nd_nrkc/nd_gmc", "kN", "", "out", "0.00")
s.add("nd_apr", "Area aderenza fila tesa / A0", "Ap,N/A0", "=(2*MIN(nd_cx,nd_scp/2)+(nd_nx-1)*MIN(nd_sx,nd_scp))*(MIN(nd_ct,nd_scp/2)+nd_scp/2)/nd_scp^2", "-", "", "calc", "0.000")
s.add("nd_psp", "psi s,Np", "", "=MIN(1,0.7+0.3*MIN(nd_cx,nd_ct)/(nd_scp/2))", "-", "", "calc", "0.000")
s.add("nd_nrkp", "Aderenza fila tesa caratteristica", "NRk,p", "=nd_npz*nd_apr*nd_psp*nd_prn", "kN", "", "calc", "0.00")
s.add("nd_nrp", "Aderenza fila tesa di progetto", "NRd,p", "=nd_nrkp/nd_gmc", "kN", "", "out", "0.00")
s.add("nd_aca", "Area cono intero gruppo / A0 (per pry-out)", "", "=(2*MIN(nd_cx,1.5*nd_hef)+(nd_nx-1)*MIN(nd_sx,3*nd_hef))*(MIN(nd_ct,1.5*nd_hef)+(nd_ny-1)*MIN(nd_sy,3*nd_hef)+1.5*nd_hef)/(9*nd_hef^2)", "-", "", "calc", "0.000")
s.add("nd_apa", "Area aderenza intero gruppo / A0 (per pry-out)", "", "=(2*MIN(nd_cx,nd_scp/2)+(nd_nx-1)*MIN(nd_sx,nd_scp))*(MIN(nd_ct,nd_scp/2)+(nd_ny-1)*MIN(nd_sy,nd_scp)+nd_scp/2)/nd_scp^2", "-", "", "calc", "0.000")
s.add("nd_vrcp", "Pry-out gruppo di progetto", "VRd,cp", "=nd_k8*MIN(nd_ncz*nd_aca*nd_psn,nd_npz*nd_apa*nd_psp)*nd_prn/1.5", "kN", "= k8 min(NRk,c ; NRk,p)gruppo / 1,5", "out", "0.00")
s.add("nd_lf", "Lunghezza efficace a taglio", "lf", "=IF(nd_d<=24,MIN(nd_hef,12*nd_d),MIN(nd_hef,MAX(8*nd_d,300)))", "mm", "", "calc", "0")
s.add("nd_vcz", "Resistenza bordo di base", "V0Rk,c", "=nd_k9*nd_d^(0.1*(nd_lf/nd_cx)^0.5)*nd_lf^(0.1*(nd_d/nd_cx)^0.2)*SQRT(fck_c)*nd_cx^1.5/1000", "kN", "c1 = cx", "calc", "0.00")
s.add("nd_rcv", "Area bordo colonna di ancoranti / A0", "", "=(3*nd_cx+(nd_ny-1)*MIN(nd_sy,3*nd_cx))*MIN(1.5*nd_cx,nd_cd)/(4.5*nd_cx^2)", "-", "Colonna verticale di ny ancoranti vicina al bordo laterale.", "calc", "0.000")
s.add("nd_phv", "psi h,V", "", "=MAX(1,SQRT(1.5*nd_cx/nd_cd))", "-", "", "calc", "0.000")
s.add("nd_psv", "psi s,V", "", "=MIN(1,0.7+0.3*nd_ct/(1.5*nd_cx))", "-", "", "calc", "0.000")
s.add("nd_vrc", "Rottura bordo colonna di ancoranti (taglio verticale parallelo al bordo)", "VRd,c", "=nd_vcz*nd_rcv*nd_phv*nd_psv*2*IF(nd_prv=\"STANDARD\",1,1.4)/1.5", "kN", "psi_alfa,V = 2", "out", "0.00")
s.add("nd_vline", "Taglio sulla colonna di ancoranti vicina al bordo", "", "=nd_v/nd_nx", "kN", "", "calc", "0.00")
s.section("VERIFICHE NODO")
s.add("nd_r1", "Trazione acciaio NEd/NRd,s", "", "=nd_ned/nd_nrs", "-", "", "out", "0.000")
s.add("nd_r2", "Cono cls fila tesa Nrow/NRd,c", "", "=nd_nrow/nd_nrc", "-", "", "out", "0.000")
s.add("nd_r3", "Aderenza fila tesa Nrow/NRd,p", "", "=nd_nrow/nd_nrp", "-", "", "out", "0.000")
s.add("nd_r4", "Taglio acciaio VEd/VRd,s", "", "=nd_ved/nd_vrs", "-", "", "out", "0.000")
s.add("nd_r5", "Rifollamento piastra VEd/Fb,Rd", "", "=nd_ved/nd_fb", "-", "", "out", "0.000")
s.add("nd_r6", "Rottura bordo Vline/VRd,c", "", "=nd_vline/nd_vrc", "-", "", "out", "0.000")
s.add("nd_r7", "Pry-out VEd/VRd,cp", "", "=nd_v/nd_vrcp", "-", "", "out", "0.000")
s.add("nd_r8", "Interazione acciaio", "", "=nd_r1^2+nd_r4^2", "-", "(N/NRd,s)^2 + (V/VRd,s)^2 <= 1", "out", "0.000")
s.add("nd_r9", "Interazione cls", "", "=MAX(nd_r2,nd_r3)^1.5+MAX(nd_r6,nd_r7)^1.5", "-", "bN^1,5 + bV^1,5 <= 1", "out", "0.000")
s.add("nd_rmax", "RAPPORTO MASSIMO NODO", "", "=MAX(nd_r1,nd_r2,nd_r3,nd_r4,nd_r5,nd_r6,nd_r7,nd_r8,nd_r9)", "-", "", "out", "0.000")
s.add("nd_es", "ESITO NODO PILASTRO", "", esito("nd_rmax", "sc_n=2"), "", "", "out", None)
s.section("CONTROLLI GEOMETRICI")
chk = [
    ("nd_g1", "cx >= cmin (ETA)", '=IF(sc_n<>2,"N.A.",IF(nd_cx>=nd_cmin,"OK","NON VERIFICATA"))', ""),
    ("nd_g2", "sx, sy >= smin (ETA)", '=IF(sc_n<>2,"N.A.",IF(AND(OR(nd_nx=1,nd_sx>=nd_smin),OR(nd_ny=1,nd_sy>=nd_smin)),"OK","NON VERIFICATA"))', ""),
    ("nd_g3", "Piastra entro la faccia del pilastro", '=IF(sc_n<>2,"N.A.",IF(nd_bp<=nd_cb,"OK","NON VERIFICATA"))', ""),
    ("nd_g4", "Ancoranti entro la piastra con e >= 1,2 d0", '=IF(sc_n<>2,"N.A.",IF(AND((nd_nx-1)*nd_sx+2*nd_ep<=nd_bp,(nd_ny-1)*nd_sy+2*nd_ep<=nd_hp,nd_ep>=1.2*nd_d0),"OK","NON VERIFICATA"))', ""),
    ("nd_g5", "Spessore residuo pilastro Dc - hef >= max(30 ; 2 d0)", '=IF(sc_n<>2,"N.A.",IF(nd_cd-nd_hef>=MAX(30,2*(nd_d+4)),"OK","NON VERIFICATA"))', "Verificare hmin dell'ETA."),
]
for nm, lab, f, ref in chk:
    s.add(nm, lab, "", f, "", ref, "out", None)
add_cf(ws, "C8:C160")
s.text("NON INCLUSO: verifica della piastra a flessione (T-stub), delle saldature profilo-piastra o del collegamento bullonato fazzoletto/mensola (EN 1993-1-8), "
       "verifica locale del pilastro esistente e del nodo per la nuova azione concentrata.", 30, False, "9C5700")
ws6 = s.ws

# =====================================================================================
# 07 RIEPILOGO
# =====================================================================================
ws7 = wb.create_sheet("07_RIEPILOGO")
ws7.sheet_view.showGridLines = False
for col, w in zip("ABCDEF", [52, 36, 14, 14, 12, 20]):
    ws7.column_dimensions[col].width = w
ws7["A1"] = "07 - RIEPILOGO VERIFICHE"
ws7["A1"].font = fnt(True, "FFFFFF", 14); ws7["A1"].fill = F_H1; ws7.merge_cells("A1:F1")
ws7["A2"] = '="Scenario: "&sc_txt&"   |   L = "&TEXT(lu_m,"0.00")&" m   |   profilo "&a_nome&" "&st_gr'
ws7["A2"].font = fnt(True, "1F3864", 10); ws7.merge_cells("A2:F2")
ws7["A3"] = '="MEd = "&TEXT(m_ed,"0.0")&" kNm   VEd = "&TEXT(v_ed,"0.0")&" kN   Mpl,Rd = "&TEXT(cs_mpl,"0.0")&" kNm   Mpl,a,Rd = "&TEXT(a_mpl,"0.0")&" kNm   Jid(n0)/Ia = "&TEXT(els_ratio,"0.00")'
ws7["A3"].font = fnt(False, "404040", 9); ws7.merge_cells("A3:F3")
hdr = ["Verifica", "Riferimento", "Domanda", "Capacita'", "Rapporto", "Esito"]
for j, h in enumerate(hdr):
    c = ws7.cell(5, 1 + j, h); c.font = fnt(True, size=10); c.fill = F_TH; c.border = BRD
rows = [
    ("TRAVE COMPOSTA", None),
    ("Flessione SLU (plastica, con grado di connessione)", "EC4 6.2.1 / NTC 4.3.4", "m_ed", "vt_mrd", "vt_rm", "vt_em"),
    ("Grado di connessione minimo (sc.1)", "EC4 6.6.1.2", "st_eta", "vt_emin", None, "vt_etaok"),
    ("Classe sezione per calcolo plastico", "EN 1993-1-1 Tab. 5.2", None, None, None, "cs_wcl"),
    ("Taglio SLU (anima profilo)", "EC4 6.2.2.2 / NTC 4.2.4.1.2.4", "v_ed", "a_vpl", "vt_rv", "vt_ev"),
    ("Tensione acciaio (metodo n) - vincolante sc.2 ELASTICA", "metodo n per fasi", "vt_sa", "fyd_a", "vt_rsa", "vt_esa"),
    ("Tensione calcestruzzo (metodo n) - vincolante sc.2 ELASTICA", "metodo n per fasi", "vt_sc", "fcd_c", "vt_rsc", "vt_esc"),
    ("Freccia totale SLE", "NTC Tab. 4.2.XII", "vt_ftot", "lu_mm/lim_tot", "vt_rft", "vt_eft"),
    ("Freccia variabili SLE", "NTC Tab. 4.2.XII", "vt_fq", "lu_mm/lim_q", "vt_rfq", "vt_efq"),
    ("FASE DI GETTO (sc.1 non puntellata)", None),
    ("Flessione sola trave acciaio", "NTC 4.2.4.1.2", "m_cost", "a_mpl", "vt_rmc", "vt_emc"),
    ("Taglio sola trave acciaio", "NTC 4.2.4.1.2.4", "v_cost", "a_vpl", "vt_rvc", "vt_evc"),
    ("Stabilita' flesso-torsionale", "EN 1993-1-1 6.3.2", None, None, None, "vt_ltb"),
    ("PIOLI (sc.1)", None),
    ("Numero pioli / grado connessione (n disposti vs nf)", "EC4 6.6.1", "st_nf", "st_np", None, "vt_etaok"),
    ("Armatura trasversale soletta", "EC4 6.6.6", "tr_req", "sl_ast", "tr_r", "tr_e"),
    ("Bielle compresse soletta", "EC2 6.2.4", "tr_v", "tr_vmax", "tr_rb", "tr_eb"),
    ("Dettagli: hsc>=3d", "EC4 6.6.5.7", None, None, None, "st_c1"),
    ("Dettagli: hsc>=hp+2d", "EC4 6.6.5.8", None, None, None, "st_c2"),
    ("Dettagli: d<=20 con saldatura passante", "EC4 6.6.5.8", None, None, None, "st_c3"),
    ("Dettagli: passo minimo / massimo", "EC4 6.6.5.5 / 6.6.5.7", None, None, None, "st_c6"),
    ("INGHISAGGI TRAVE (sc.2)", None),
    ("Taglio per ancorante", "EN 1992-4", "an_ved", "an_vrd", "an_rv", "an_ev"),
    ("Trazione per ancorante", "EN 1992-4", "an_ned", "an_nrd", "an_rn", "an_en"),
    ("Interazione N-V acciaio", "EN 1992-4 eq. 7.54", None, None, "an_is", "an_eis"),
    ("Interazione N-V calcestruzzo", "EN 1992-4 eq. 7.55", None, None, "an_ic", "an_eic"),
    ("Grado di connessione >= 1 (solo criterio PLASTICA)", "EC4 6.6.1.1", None, None, "an_eta", "an_ee"),
    ("Passo ancoranti <= s,max", "-", "an_s", "an_smax", None, "an_g7"),
    ("Distanza dal bordo c1 >= cmin", "ETA", "an_cmin", "an_c", None, "an_g1"),
    ("NODO PILASTRO (sc.2)", None),
    ("Nodo profilo-pilastro (massimo tra tutte le verifiche)", "EN 1992-4 / EN 1993-1-8", None, None, "nd_rmax", "nd_es"),
    ("Distanza dal bordo cx >= cmin", "ETA", "nd_cmin", "nd_cx", None, "nd_g1"),
    ("Geometria ancoranti nella piastra", "EN 1993-1-8", None, None, None, "nd_g4"),
]
r = 6
for row in rows:
    if row[1] is None:
        c = ws7.cell(r, 1, row[0]); c.font = fnt(True, "1F3864", 10)
        for j in range(1, 7):
            ws7.cell(r, j).fill = F_H2
        r += 1
        continue
    lab, ref, dmd, cap, rat, es = row
    vals = [lab, ref, ("=" + dmd) if dmd else "", ("=" + cap) if cap else "", ("=" + rat) if rat else "", "=" + es]
    for j, v in enumerate(vals):
        c = ws7.cell(r, 1 + j, v); c.font = fnt(size=9); c.border = BRD
        if j in (2, 3, 4):
            c.number_format = "0.000" if j == 4 else "0.00"
            c.alignment = Alignment(horizontal="center")
        if j == 5:
            c.alignment = Alignment(horizontal="center"); c.font = fnt(True, size=9)
    r += 1
add_cf(ws7, f"F6:F{r}")
ws7.cell(r + 1, 1, "Esito globale (verifiche pertinenti allo scenario):").font = fnt(True, size=11)
ws7.cell(r + 1, 6, f'=IF(COUNTIF(F6:F{r - 1},"NON*")>0,"NON VERIFICATA",IF(COUNTIF(F6:F{r - 1},"ATTENZIONE*")>0,"ATTENZIONE","VERIFICATA"))').font = fnt(True, size=11)
add_cf(ws7, f"F{r + 1}:F{r + 1}")
ws7.cell(r + 2, 1, "Domanda/Capacita' in kN, kNm, MPa, mm secondo la riga. N.A. = non pertinente allo scenario scelto. ATTENZIONE = controllo da completare manualmente.").font = fnt(False, "7F7F7F", 8, True)
ws7.freeze_panes = "A6"

# ordine fogli
order = ["00_LEGGIMI", "01_DATI_GENERALI", "02_PROFILO_ACCIAIO", "02A_SOLETTA_LAMIERA", "02B_TRAVE_CA_ESISTENTE", "03_SEZIONE_COMPOSTA",
         "04_VERIFICHE_TRAVE", "05A_PIOLI", "05B_INGHISAGGI_TRAVE", "06_NODO_PILASTRO", "07_RIEPILOGO", "DB"]
wb._sheets = [wb[n] for n in order]
for n, col in (("02A_SOLETTA_LAMIERA", "548235"), ("05A_PIOLI", "548235"), ("02B_TRAVE_CA_ESISTENTE", "C55A11"),
               ("05B_INGHISAGGI_TRAVE", "C55A11"), ("06_NODO_PILASTRO", "C55A11"), ("07_RIEPILOGO", "1F3864"), ("01_DATI_GENERALI", "FFC000")):
    wb[n].sheet_properties.tabColor = col
for w in wb.worksheets:
    w.page_setup.orientation = "landscape"
    w.page_setup.paperSize = w.PAPERSIZE_A4
    w.page_setup.fitToWidth = 1
    w.page_setup.fitToHeight = 0
    w.sheet_properties.pageSetUpPr.fitToPage = True
wb.active = 1
wb.save(OUT)
print("saved", OUT, len(NAMES), "names")
