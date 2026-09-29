"""Génère Tarif_TTC_DR_GASMI.xlsx à partir de la liste des prix HT (PTH / PTG)."""
from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.comments import Comment
from openpyxl.worksheet.datavalidation import DataValidation

F = "Arial"
BLUE = Font(name=F, color="0000FF")
BLACK = Font(name=F)
BOLD = Font(name=F, bold=True)
GREEN = Font(name=F, color="008000")
HDR = Font(name=F, bold=True, color="FFFFFF")
HDR_FILL = PatternFill("solid", fgColor="1F4E78")
BLOCK_FILL = PatternFill("solid", fgColor="DDEBF7")
TOT_FILL = PatternFill("solid", fgColor="F2F2F2")
INPUT_FILL = PatternFill("solid", fgColor="FFFF00")
thin = Side(style="thin", color="BFBFBF")
BOX = Border(left=thin, right=thin, top=thin, bottom=thin)
DA = '#,##0 "DA";(#,##0 "DA");"-"'
PCT = "0.0%"

# (nom, [(composant, prix HT), ...], note éventuelle)
PROTHESES = [
    ("PTH DM sans ciment (tête INOX)", [("Tige sans ciment (Rmis)", 58000), ("Cotyle DM sans ciment", 68000), ("Insert DM", 25000), ("Tête en inox", 26000)], None),
    ("PTH DM sans ciment (tête Alumine)", [("Tige sans ciment (Rmis)", 58000), ("Cotyle DM sans ciment", 68000), ("Insert DM", 25000), ("Tête en alumine", 77000)], None),
    ("PTH DM cimentée (tête INOX)", [("Tige cimentée (Osteal)", 47000), ("Cotyle DM cimenté", 65000), ("Insert DM", 25000), ("Tête en acier", 26000), ("Ciment (une dose)", 7000)],
     "Ligne ciment ajoutée : absente de la liste d'origine alors que la version tête alumine l'inclut."),
    ("PTH DM cimentée (tête Alumine)", [("Tige cimentée (Osteal)", 47000), ("Cotyle DM cimenté", 65000), ("Insert DM", 25000), ("Tête en alumine", 77000), ("Ciment (une dose)", 7000)], None),
    ("PTH simple mobilité sans ciment céramique/céramique", [("Tige fémorale Rmis", 58000), ("Cotyle CERAFIT", 62000), ("Insert céramique", 78000), ("Tête céramique (alumine)", 77000)], None),
    ("PTH simple mobilité sans ciment", [("Tige fémorale Rmis", 58000), ("Cotyle CERAFIT", 62000), ("Insert PE", 24000), ("Tête en acier (inox)", 26000)], None),
    ("PTH cimentée", [("Tige cimentée Osteal", 47000), ("Cotyle PE à cimenter", 54000), ("Tête acier", 26000), ("Ciment (une dose)", 7000)],
     "Total d'origine =SUM(C48:C50) oubliait le ciment (127 000 au lieu de 134 000)."),
    ("PTG sans rotule", [("CFG fémur", 90000), ("Embase tibiale", 82000), ("Insert PE", 31000), ("Ciment (une dose)", 7000)],
     "Total d'origine =SUM(C56:C58) oubliait le ciment (203 000 au lieu de 210 000)."),
]
OPTIONS = [("Rotule", 12000), ("Rallonge tibiale", 40000), ("Vis", 7500), ("Dose de ciment", 7000)]

wb = Workbook()

# ---------- Paramètres ----------
p = wb.active
p.title = "Paramètres"
p["A1"] = "Paramètres de calcul"; p["A1"].font = Font(name=F, bold=True, size=14)
rows = [
    ("Taux de TVA", 0.19, "Taux normal algérien 19 %. À faire confirmer par le comptable : certains dispositifs médicaux relèvent du taux réduit 9 %."),
    ("Marge bénéficiaire sur prix de revient", 0.10, "Demandé par l'utilisateur : prix de revient + 10 %."),
]
for i, (lab, val, note) in enumerate(rows, start=3):
    p.cell(i, 1, lab).font = BOLD
    c = p.cell(i, 2, val); c.font = BLUE; c.fill = INPUT_FILL; c.number_format = PCT; c.border = BOX
    p.cell(i, 3, note).font = Font(name=F, italic=True, color="595959")
p["A6"] = "Cellules jaunes / texte bleu = valeurs à saisir. Tout le reste est calculé."
p["A6"].font = Font(name=F, italic=True)
p.column_dimensions["A"].width = 40; p.column_dimensions["B"].width = 12; p.column_dimensions["C"].width = 95
TVA = "Paramètres!$B$3"
MARGE = "Paramètres!$B$4"

# ---------- Tarif détaillé ----------
t = wb.create_sheet("Tarif détaillé")
t["A1"] = "Listing des prix DR GASMI — HT, TVA, TTC et prix de revient + marge"
t["A1"].font = Font(name=F, bold=True, size=14)
heads = ["Désignation", "Prix HT liste", "TVA", "Prix TTC liste", "Prix de revient (à saisir)",
         "HT = revient + marge", "TTC = revient + marge", "Marge réelle sur prix liste"]
widths = [52, 16, 14, 16, 20, 20, 20, 18]
for j, (h, w) in enumerate(zip(heads, widths), start=1):
    c = t.cell(3, j, h); c.font = HDR; c.fill = HDR_FILL; c.border = BOX
    c.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
    t.column_dimensions[c.column_letter].width = w
t.row_dimensions[3].height = 32
t.freeze_panes = "B4"

def line(r, label, ht, bold=False):
    t.cell(r, 1, label).font = BOLD if bold else BLACK
    c = t.cell(r, 2, ht); c.font = BLUE
    t.cell(r, 3, f"=B{r}*{TVA}")
    t.cell(r, 4, f"=B{r}+C{r}")
    e = t.cell(r, 5); e.fill = INPUT_FILL; e.font = BLUE
    t.cell(r, 6, f'=IF(E{r}="","",E{r}*(1+{MARGE}))')
    t.cell(r, 7, f'=IF(F{r}="","",F{r}*(1+{TVA}))')
    t.cell(r, 8, f'=IF(OR(E{r}="",E{r}=0),"",B{r}/E{r}-1)')
    for j in range(1, 9):
        t.cell(r, j).border = BOX
        if j in (2, 3, 4, 5, 6, 7): t.cell(r, j).number_format = DA
    t.cell(r, 8).number_format = PCT

totals = []  # (nom, ligne total)
r = 4
for nom, comps, note in PROTHESES:
    c = t.cell(r, 1, nom); c.font = BOLD
    for j in range(1, 9): t.cell(r, j).fill = BLOCK_FILL; t.cell(r, j).border = BOX
    if note: c.comment = Comment(note, "Analyse")
    r += 1; first = r
    for lab, ht in comps:
        line(r, "   " + lab, ht); r += 1
    last = r - 1
    t.cell(r, 1, "Prix de la prothèse").font = BOLD
    t.cell(r, 2, f"=SUM(B{first}:B{last})")
    t.cell(r, 3, f"=SUM(C{first}:C{last})")
    t.cell(r, 4, f"=SUM(D{first}:D{last})")
    t.cell(r, 5, f'=IF(COUNT(E{first}:E{last})<{last-first+1},"",SUM(E{first}:E{last}))')
    t.cell(r, 6, f'=IF(E{r}="","",E{r}*(1+{MARGE}))')
    t.cell(r, 7, f'=IF(F{r}="","",F{r}*(1+{TVA}))')
    t.cell(r, 8, f'=IF(OR(E{r}="",E{r}=0),"",B{r}/E{r}-1)')
    for j in range(1, 9):
        cc = t.cell(r, j); cc.fill = TOT_FILL; cc.border = BOX; cc.font = BOLD
        cc.number_format = PCT if j == 8 else DA
    totals.append((nom, r))
    r += 2

c = t.cell(r, 1, "Options : + rotule ou + rallonge tibiale avec sa vis"); c.font = BOLD
for j in range(1, 9): t.cell(r, j).fill = BLOCK_FILL; t.cell(r, j).border = BOX
r += 1
opt_rows = {}
for lab, ht in OPTIONS:
    line(r, "   " + lab, ht); opt_rows[lab] = r; r += 1
r += 1
t.cell(r, 1, "Prix de revient : ligne « Prix de la prothèse » calculée seulement quand tous les composants sont saisis.").font = Font(name=F, italic=True, color="595959")

# ---------- Récapitulatif ----------
s = wb.create_sheet("Récapitulatif")
s["A1"] = "Récapitulatif par prothèse"; s["A1"].font = Font(name=F, bold=True, size=14)
heads = ["Prothèse", "Prix HT liste", "TVA", "Prix TTC liste",
         "Option A : HT pour garder le même montant total (TTC = ancien HT)", "Baisse de HT (option A)",
         "Option B : HT = revient + marge", "Option B : TTC"]
widths = [52, 16, 14, 16, 26, 18, 18, 18]
for j, (h, w) in enumerate(zip(heads, widths), start=1):
    c = s.cell(3, j, h); c.font = HDR; c.fill = HDR_FILL; c.border = BOX
    c.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
    s.column_dimensions[c.column_letter].width = w
s.row_dimensions[3].height = 60
q = "'Tarif détaillé'!"
for i, (nom, tr) in enumerate(totals, start=4):
    s.cell(i, 1, nom).font = BLACK
    s.cell(i, 2, f"={q}B{tr}").font = GREEN
    s.cell(i, 3, f"={q}C{tr}").font = GREEN
    s.cell(i, 4, f"={q}D{tr}").font = GREEN
    s.cell(i, 5, f"=B{i}/(1+{TVA})")
    s.cell(i, 6, f"=B{i}-E{i}")
    s.cell(i, 7, f"={q}F{tr}").font = GREEN
    s.cell(i, 8, f"={q}G{tr}").font = GREEN
    for j in range(1, 9):
        s.cell(i, j).border = BOX
        if j > 1: s.cell(i, j).number_format = DA
recap_last = 3 + len(totals)
s.cell(recap_last + 2, 1, "Option A : le patient paie exactement l'ancien prix, TVA comprise ; le fournisseur absorbe la TVA (marge réduite).").font = Font(name=F, italic=True)
s.cell(recap_last + 3, 1, "Option B : prix construit sur le prix de revient + marge (onglet Paramètres), TVA ajoutée.").font = Font(name=F, italic=True)
s.freeze_panes = "B4"

# ---------- Paiement mixte ----------
m = wb.create_sheet("Paiement mixte")
m["A1"] = "Échéancier de règlement — chèque/virement + espèces (tout facturé)"
m["A1"].font = Font(name=F, bold=True, size=14)
m.column_dimensions["A"].width = 46; m.column_dimensions["B"].width = 50; m.column_dimensions["C"].width = 60
lab = lambda r, txt: m.cell(r, 1, txt).__setattr__("font", BOLD)
lab(3, "Prothèse")
m["B3"] = totals[0][0]; m["B3"].font = BLUE; m["B3"].fill = INPUT_FILL; m["B3"].border = BOX
dv = DataValidation(type="list", formula1=f"='Récapitulatif'!$A$4:$A${recap_last}", allow_blank=False)
m.add_data_validation(dv); dv.add("B3")
lab(4, "Base de prix")
m["B4"] = "Liste"; m["B4"].font = BLUE; m["B4"].fill = INPUT_FILL; m["B4"].border = BOX
dv2 = DataValidation(type="list", formula1='"Liste,TTC constant,Revient + marge"', allow_blank=False)
m.add_data_validation(dv2); dv2.add("B4")
m["C4"] = "Liste = HT liste + TVA · TTC constant = option A · Revient + marge = option B"
lab(5, "Options ajoutées (rotule, rallonge…) HT")
m["B5"] = 0; m["B5"].font = BLUE; m["B5"].fill = INPUT_FILL; m["B5"].number_format = DA; m["B5"].border = BOX

rr = "MATCH($B$3,'Récapitulatif'!$A$4:$A$%d,0)" % recap_last
col = lambda L: f"INDEX('Récapitulatif'!${L}$4:${L}${recap_last},{rr})"
lab(7, "Montant HT facturé")
m["B7"] = (f'=IF($B$4="Liste",{col("B")},IF($B$4="TTC constant",{col("E")},'
           f'IF({col("G")}="",NA(),{col("G")})))+$B$5')
m["C7"] = "Revient + marge : #N/A tant que les prix de revient ne sont pas saisis."
lab(8, "TVA"); m["B8"] = f"=B7*{TVA}"
lab(9, "Montant TTC facturé"); m["B9"] = "=B7+B8"
lab(11, "Réglé en espèces (avec reçu, porté sur la facture)")
m["B11"] = 50000; m["B11"].font = BLUE; m["B11"].fill = INPUT_FILL; m["B11"].border = BOX
m["C11"] = "Montant saisi = exemple de l'utilisateur (50 000 DA)."
lab(12, "Réglé par chèque / virement"); m["B12"] = "=B9-B11"
lab(13, "Part espèces / total"); m["B13"] = "=IF(B9=0,0,B11/B9)"
lab(14, "Contrôle : espèces + chèque = TTC facturé"); m["B14"] = '=IF(ABS(B11+B12-B9)<0.5,"OK","ERREUR")'
for rw in (7, 8, 9, 12):
    m[f"B{rw}"].number_format = DA; m[f"B{rw}"].border = BOX
m["B11"].number_format = DA
m["B13"].number_format = PCT
for rw in (9, 12): m[f"B{rw}"].font = BOLD
for rw in (3, 4, 5, 7, 11): m[f"C{rw}"].font = Font(name=F, italic=True, color="595959")
m["A16"] = ("Les deux règlements figurent sur la même facture TTC : la TVA porte sur la totalité du prix, "
            "quel que soit le mode de paiement.")
m["A16"].font = Font(name=F, italic=True)

# ---------- Notes ----------
n = wb.create_sheet("Notes d'analyse")
n.column_dimensions["A"].width = 120
notes = [
    "Anomalies relevées dans LISTING_DES_PRIX_DR_GASMI.xlsx",
    "1. PTH cimentée : total =SUM(C48:C50) n'inclut pas le ciment (C51). 127 000 → 134 000 DA HT avec ciment.",
    "2. PTG sans rotule : total =SUM(C56:C58) n'inclut pas le ciment (C59). 203 000 → 210 000 DA HT avec ciment.",
    "3. PTH DM cimentée tête INOX : pas de ligne ciment, alors que la version tête alumine en a une. Ligne ajoutée ici (+7 000 DA).",
    "4. Feuil2 et Feuil3 sont vides.",
    "",
    "Si l'une de ces différences est voulue (ciment fourni à part, par ex.), supprimer la ligne ciment dans « Tarif détaillé ».",
]
for i, txt in enumerate(notes, start=1):
    n.cell(i, 1, txt).font = BOLD if i == 1 else BLACK

for ws in wb.worksheets:
    for row in ws.iter_rows():
        for c in row:
            if c.font.name != F:
                f0 = c.font
                c.font = Font(name=F, bold=f0.bold, italic=f0.italic, color=f0.color, size=f0.sz)

wb.active = 1
wb.calculation.fullCalcOnLoad = True  # pas de valeurs en cache : Excel recalcule à l'ouverture
wb.save("/home/user/Claude/dr_gasmi/Tarif_TTC_DR_GASMI.xlsx")
print("ok")
