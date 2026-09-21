#!/usr/bin/env python3
"""
Regenerates cjenik-chokai.xlsx, the downloadable version of the price list.

The spreadsheet and the table in index.html must agree, so whenever a price
changes, edit SERVICES below, re-run this script, and mirror the same numbers
in index.html.

    python3 tools/build-cjenik-xlsx.py

Requires openpyxl.
"""

import datetime as dt
import os

from openpyxl import Workbook
from openpyxl.comments import Comment
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils import get_column_letter

# --------------------------------------------------------------------------
# Business data
# --------------------------------------------------------------------------

LEGAL_NAME = "CHOKAI, obrt za računalne i poslovne usluge, vl. Matko Stanković"
ADDRESS = "Petrova 84, 10123 Zagreb, Hrvatska"
OIB = "64389569095"
MBO = "98086448"
OBRTNICA = "21010143794"
EMAIL = "wellwho@gmail.com"
PHONE = "+385 98 738 460"

VALID_FROM = dt.date(2026, 9, 20)
REFERENCE_DATE = dt.date(2026, 9, 10)
VAT_RATE = 0.25

# (service, unit, net price, reference-date price incl. VAT)
SERVICES = [
    ("Rad po satu", "1 sat", 60.00, 75.00),
]

# --------------------------------------------------------------------------
# Styling
# --------------------------------------------------------------------------

FONT = "Arial"
NAVY = "1A2542"
ORANGE = "F28A15"

f_title = Font(name=FONT, size=16, bold=True, color=NAVY)
f_sub = Font(name=FONT, size=10, color="444444")
f_label = Font(name=FONT, size=10, bold=True, color=NAVY)
f_head = Font(name=FONT, size=10, bold=True, color="FFFFFF")
f_body = Font(name=FONT, size=11)
f_service = Font(name=FONT, size=11, bold=True)
f_input = Font(name=FONT, size=11, color="0000FF")      # hand-entered
f_calc = Font(name=FONT, size=11, color="000000")       # formula
f_note = Font(name=FONT, size=9, color="555555")
f_legend_hd = Font(name=FONT, size=10, bold=True, color=NAVY)

fill_head = PatternFill("solid", fgColor=NAVY)
fill_assume = PatternFill("solid", fgColor="FFFF00")    # key assumption
fill_ref = PatternFill("solid", fgColor="FDF0DE")       # reference column

thin = Side(style="thin", color="BBBBBB")
border = Border(left=thin, right=thin, top=thin, bottom=thin)

EUR = '#,##0.00\\ "€"'
PCT = '0.0%'
DATE = 'DD.MM.YYYY.'

wb = Workbook()
ws = wb.active
ws.title = "Cjenik"

# --------------------------------------------------------------------------
# Header block
# --------------------------------------------------------------------------

ws["A1"] = "CHOKAI — CJENIK USLUGA"
ws["A1"].font = f_title

for row, text in enumerate(
    [
        LEGAL_NAME,
        ADDRESS,
        f"OIB: {OIB}    MBO: {MBO}    Obrtnica: {OBRTNICA}",
        f"E-pošta: {EMAIL}    Telefon: {PHONE}",
    ],
    start=2,
):
    ws.cell(row=row, column=1, value=text).font = f_sub

ws["A7"] = "Cjenik vrijedi od:"
ws["A7"].font = f_label
ws["B7"] = VALID_FROM
ws["B7"].number_format = DATE
ws["B7"].font = f_input

ws["A8"] = "Referentni datum:"
ws["A8"].font = f_label
ws["B8"] = REFERENCE_DATE
ws["B8"].number_format = DATE
ws["B8"].font = f_input
ws["B8"].fill = fill_assume

ws["A9"] = "Stopa PDV-a:"
ws["A9"].font = f_label
ws["B9"] = VAT_RATE
ws["B9"].number_format = PCT
ws["B9"].font = f_input
ws["B9"].fill = fill_assume

# --------------------------------------------------------------------------
# Price table
# --------------------------------------------------------------------------

HEAD_ROW = 11
FIRST = HEAD_ROW + 1

headers = [
    "Usluga",
    "Jedinica",
    "Cijena bez PDV-a",
    "PDV",
    "Cijena s PDV-om",
    "Cijena na referentni datum (s PDV-om)",
    "Promjena (€)",
    "Promjena (%)",
]

for col, text in enumerate(headers, start=1):
    c = ws.cell(row=HEAD_ROW, column=col, value=text)
    c.font = f_head
    c.fill = fill_head
    c.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
    c.border = border

ws.row_dimensions[HEAD_ROW].height = 34

for i, (name, unit, net, ref) in enumerate(SERVICES):
    r = FIRST + i

    ws.cell(row=r, column=1, value=name).font = f_service
    ws.cell(row=r, column=2, value=unit).font = f_body

    c = ws.cell(row=r, column=3, value=net)            # hand-entered
    c.font, c.number_format = f_input, EUR

    c = ws.cell(row=r, column=4, value=f"=C{r}*$B$9")  # VAT
    c.font, c.number_format = f_calc, EUR

    c = ws.cell(row=r, column=5, value=f"=C{r}+D{r}")  # gross
    c.font, c.number_format = f_calc, EUR

    c = ws.cell(row=r, column=6, value=ref)            # hand-entered
    c.font, c.number_format = f_input, EUR
    c.fill = fill_ref
    c.comment = Comment(
        "Cijena koja je vrijedila na referentni datum (B8).\n\n"
        "Izvor: potvrdio vlasnik obrta 20.09.2026. — cijena je na "
        "10.09.2026. bila ista (60,00 € bez PDV-a / 75,00 € s PDV-om), "
        "pa od referentnog datuma nije bilo promjene.\n\n"
        "Kod buduće promjene cijene ovaj stupac zadržava cijenu s "
        "referentnog datuma; mijenja se samo stupac C.",
        "CHOKAI",
        width=320,
        height=130,
    )

    c = ws.cell(row=r, column=7, value=f"=E{r}-F{r}")
    c.font, c.number_format = f_calc, EUR

    # guarded so an empty or zero reference price cannot produce #DIV/0!
    c = ws.cell(row=r, column=8, value=f'=IF(F{r}=0,"",(E{r}-F{r})/F{r})')
    c.font, c.number_format = f_calc, PCT

    for col in range(1, 9):
        ws.cell(row=r, column=col).border = border

LAST = FIRST + len(SERVICES) - 1

# --------------------------------------------------------------------------
# Notes and editing legend
# --------------------------------------------------------------------------

row = LAST + 2

ws.cell(row=row, column=1, value="NAPOMENE").font = f_legend_hd
row += 1

for text in [
    "Cijene su izražene u eurima (EUR).",
    "Osnovne cijene ne uključuju PDV. Stopa PDV-a navedena je u ćeliji B9.",
    "Cijena na referentni datum (B8) istaknuta je radi usporedbe, u skladu s propisima o isticanju cijena.",
    "Za usluge koje nisu navedene u cjeniku ponuda se izrađuje na upit.",
]:
    ws.cell(row=row, column=1, value="•  " + text).font = f_note
    row += 1

row += 1
ws.cell(row=row, column=1, value="LEGENDA I UPUTE ZA UREĐIVANJE").font = f_legend_hd
row += 1

for text in [
    "Plavi tekst = ručni unos: naziv usluge, jedinica, cijena bez PDV-a (stupac C) i cijena na referentni datum (stupac F).",
    "Žuta ispuna = ključne pretpostavke: referentni datum (B8) i stopa PDV-a (B9). Promjena se odražava na cijeli cjenik.",
    "Crni tekst = izračun, ne uređivati: stupci D, E, G i H.",
    "Novi redak: kopirati postojeći redak kako bi se zadržale formule, zatim izmijeniti samo stupce A, B, C i F.",
    "Primjer unosa: Naziv usluge | 1 sat | 60,00 | (izračun) | (izračun) | 62,50 | (izračun) | (izračun)",
]:
    ws.cell(row=row, column=1, value="•  " + text).font = f_note
    row += 1

row += 1
ws.cell(
    row=row,
    column=1,
    value="Ovaj je dokument generiran iz tools/build-cjenik-xlsx.py. "
          "Kod izmjene cijena uskladiti i index.html.",
).font = f_note

# --------------------------------------------------------------------------
# Layout
# --------------------------------------------------------------------------

for col, width in {
    "A": 46, "B": 12, "C": 17, "D": 13, "E": 17, "F": 26, "G": 14, "H": 13
}.items():
    ws.column_dimensions[col].width = width

ws.freeze_panes = ws.cell(row=FIRST, column=1)
ws.sheet_view.showGridLines = False
ws.print_title_rows = f"{HEAD_ROW}:{HEAD_ROW}"
ws.page_setup.orientation = "landscape"
ws.page_setup.fitToWidth = 1
ws.sheet_properties.pageSetUpPr.fitToPage = True

out = os.path.join(os.path.dirname(os.path.dirname(os.path.abspath(__file__))),
                   "cjenik-chokai.xlsx")
wb.save(out)
print("wrote", out, "|", len(SERVICES), "service row(s), rows", FIRST, "-", LAST)
