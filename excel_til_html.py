"""
Konverterer VM2026_avansert_gruppetabeller_og_sluttspill.xlsx til HTML
for publisering via GitHub Pages.

Kjør: python excel_til_html.py
Utdata: docs/index.html
"""

import os
import re
from openpyxl import load_workbook
from openpyxl.utils import get_column_letter
from openpyxl.styles.fills import PatternFill

EXCEL_FIL = "VM2026_avansert_gruppetabeller_og_sluttspill.xlsx"
UTDATA_DIR = "docs"
UTDATA_FIL = os.path.join(UTDATA_DIR, "index.html")

# Ark som skal hoppes over (interne/tomme)
HOPP_OVER_ARK = set()

# Disse arkene vises alltid sist, i denne rekkefølgen
SIST = ["Spillere etter klubbland", "Klubber", "Nivå 2 og lavere", "Aldersfordeling", "Alder"]

# Maks kolonner å vise per ark (None = alle)
MAKS_KOLS = None


def rgb_fra_fill(fill):
    """Returnerer hex-fargekode fra en PatternFill, eller None."""
    if fill is None:
        return None
    if fill.fill_type not in (None, "none", "solid"):
        return None
    if fill.fgColor is None:
        return None
    c = fill.fgColor
    if c.type == "rgb":
        rgb = c.rgb
        if rgb and len(rgb) == 8 and rgb != "00000000" and rgb.upper() != "FFFFFFFF":
            return "#" + rgb[2:]  # fjern alpha-prefix
    return None


def rgb_fra_font_farge(color):
    """Returnerer hex fra font.color, eller None."""
    if color is None:
        return None
    if color.type == "rgb":
        rgb = color.rgb
        if rgb and len(rgb) == 8 and rgb != "00000000":
            return "#" + rgb[2:]
    return None


def escape(tekst):
    if tekst is None:
        return ""
    s = str(tekst)
    s = s.replace("&", "&amp;").replace("<", "&lt;").replace(">", "&gt;")
    return s


def formater_verdi(celle):
    """Returnerer celleverdien som string, med tallformatering."""
    v = celle.value
    if v is None:
        return ""
    if isinstance(v, float):
        if v == int(v):
            return str(int(v))
        return f"{v:.2f}".rstrip("0").rstrip(".")
    return str(v)


def bygg_tabell(ws, legg_til_anker=False):
    """Bygger HTML-tabell fra et ark. Med legg_til_anker=True returneres (html, seksjoner)."""
    # Finn brukt område
    min_rad = ws.min_row or 1
    max_rad = ws.max_row or 1
    min_kol = ws.min_column or 1
    max_kol = ws.max_column or 1

    if max_rad < min_rad or max_kol < min_kol:
        if legg_til_anker:
            return "<p><em>Tomt ark</em></p>", []
        return "<p><em>Tomt ark</em></p>"

    seksjon_nr = 0
    seksjoner = []

    # Bygg merged cell-kart: (rad, kol) → (rowspan, colspan) for master-celler
    # og sett med (rad, kol) for celler som skal hoppes over
    merged_info = {}   # (r, c) → {"rowspan": n, "colspan": n}
    skip_celler = set()

    for merge in ws.merged_cells.ranges:
        r1, c1, r2, c2 = merge.min_row, merge.min_col, merge.max_row, merge.max_col
        merged_info[(r1, c1)] = {
            "rowspan": r2 - r1 + 1,
            "colspan": c2 - c1 + 1,
        }
        for r in range(r1, r2 + 1):
            for c in range(c1, c2 + 1):
                if not (r == r1 and c == c1):
                    skip_celler.add((r, c))

    # Kolonnebredder (approksimert)
    col_widths = {}
    for i in range(min_kol, max_kol + 1):
        col_letter = get_column_letter(i)
        cd = ws.column_dimensions.get(col_letter)
        if cd and cd.width:
            col_widths[i] = max(30, int(cd.width * 7))
        else:
            col_widths[i] = 80

    html = ['<div class="tabell-wrapper"><table>']

    # Kolgroup for bredder
    html.append("<colgroup>")
    for i in range(min_kol, max_kol + 1):
        html.append(f'<col style="width:{col_widths.get(i, 80)}px">')
    html.append("</colgroup>")

    for r in range(min_rad, max_rad + 1):
        tr_id = ""
        if legg_til_anker:
            første_celle = ws.cell(row=r, column=min_kol)
            bg_første = rgb_fra_fill(første_celle.fill)
            val_første = formater_verdi(første_celle)
            if bg_første and bg_første.lstrip("#").upper() == "0F2044" and val_første:
                seksjon_nr += 1
                anchor_id = f"lagstat-seksjon-{seksjon_nr}"
                raw = val_første.split("—")[-1].strip() if "—" in val_første else val_første.strip()
                tittel = raw.split("   ")[0].strip()
                seksjoner.append((anchor_id, tittel))
                tr_id = f' id="{anchor_id}"'
        html.append(f"<tr{tr_id}>")
        for c in range(min_kol, max_kol + 1):
            if (r, c) in skip_celler:
                continue

            celle = ws.cell(row=r, column=c)
            verdi = formater_verdi(celle)

            stiler = []
            klasser = []

            # Bakgrunnsfarge
            bg = rgb_fra_fill(celle.fill)
            if bg:
                stiler.append(f"background:{bg}")

            # Font
            font = celle.font
            if font:
                if font.bold:
                    stiler.append("font-weight:bold")
                if font.italic:
                    stiler.append("font-style:italic")
                farge = rgb_fra_font_farge(font.color)
                if farge:
                    stiler.append(f"color:{farge}")
                if font.sz and font.sz != 10:
                    stiler.append(f"font-size:{int(font.sz)}pt")

            # Justering
            al = celle.alignment
            if al:
                if al.horizontal in ("center", "general") and verdi and all(
                    ch.isdigit() or ch in ".,%-+" for ch in verdi.replace(" ", "")
                ):
                    stiler.append("text-align:center")
                elif al.horizontal == "right":
                    stiler.append("text-align:right")
                elif al.horizontal == "center":
                    stiler.append("text-align:center")

            # Merged cell-attributter
            merge = merged_info.get((r, c), {})
            attrs = ""
            if merge.get("rowspan", 1) > 1:
                attrs += f' rowspan="{merge["rowspan"]}"'
            if merge.get("colspan", 1) > 1:
                attrs += f' colspan="{merge["colspan"]}"'

            stil_str = ";".join(stiler)
            if stil_str:
                attrs += f' style="{stil_str}"'

            tag = "td"
            html.append(f"<{tag}{attrs}>{escape(verdi)}</{tag}>")

        html.append("</tr>")

    html.append("</table></div>")
    if legg_til_anker:
        return "\n".join(html), seksjoner
    return "\n".join(html)


def generer_lagstat_nav(seksjoner):
    """Genererer en navigasjonsbar med lenker til seksjoner i Lagstatistikk."""
    if not seksjoner:
        return ""
    lenker = "".join(
        f'<a href="#{anchor_id}">{escape(tittel)}</a>'
        for anchor_id, tittel in seksjoner
    )
    return f'<nav class="lagstat-nav">{lenker}</nav>'


def main():
    os.makedirs(UTDATA_DIR, exist_ok=True)

    print(f"Leser {EXCEL_FIL} ...")
    wb = load_workbook(EXCEL_FIL, data_only=True)

    alle = [n for n in wb.sheetnames if n not in HOPP_OVER_ARK]
    sist_sett = set(SIST)
    ark_navn = [n for n in alle if n not in sist_sett] + [n for n in SIST if n in sist_sett and n in alle]
    print(f"  Fant {len(ark_navn)} ark: {', '.join(ark_navn)}")

    # Bygg innhold per ark
    ark_html = {}
    for navn in ark_navn:
        print(f"  Konverterer ark: {navn}")
        ws = wb[navn]
        if navn == "Lagstatistikk":
            tabell_html, seksjoner = bygg_tabell(ws, legg_til_anker=True)
            nav_html = generer_lagstat_nav(seksjoner)
            ark_html[navn] = nav_html + "\n" + tabell_html
        else:
            ark_html[navn] = bygg_tabell(ws)

    # Tab-knapper
    tab_knapper = []
    for i, navn in enumerate(ark_navn):
        aktiv = ' class="aktiv"' if i == 0 else ""
        safe_id = re.sub(r"[^a-zA-Z0-9_-]", "_", navn)
        tab_knapper.append(f'<button{aktiv} onclick="visArk(\'{safe_id}\')">{escape(navn)}</button>')

    # Ark-seksjoner
    ark_seksjoner = []
    for i, navn in enumerate(ark_navn):
        safe_id = re.sub(r"[^a-zA-Z0-9_-]", "_", navn)
        display = "block" if i == 0 else "none"
        ark_seksjoner.append(
            f'<div id="ark_{safe_id}" class="ark-innhold" style="display:{display}">\n'
            f'<h2>{escape(navn)}</h2>\n'
            f'{ark_html[navn]}\n'
            f'</div>'
        )

    tab_html = '\n'.join(tab_knapper)
    ark_html_blokk = '\n'.join(ark_seksjoner)
    tidspunkt = __import__('datetime').datetime.now().strftime("%Y-%m-%d")

    template_fil = os.path.join(UTDATA_DIR, "template.html")
    if not os.path.exists(template_fil):
        print(f"FEIL: {template_fil} mangler — kan ikke generere HTML.", file=__import__('sys').stderr)
        __import__('sys').exit(1)
    with open(template_fil, encoding="utf-8") as f:
        html = f.read()
    print(f"  Bruker template: {template_fil}")

    html = html.replace("<!-- TABS -->", tab_html)
    html = html.replace("<!-- ARK -->", ark_html_blokk)
    html = html.replace("<!-- TIDSPUNKT -->", tidspunkt)

    with open(UTDATA_FIL, "w", encoding="utf-8") as f:
        f.write(html)

    size_kb = os.path.getsize(UTDATA_FIL) // 1024
    print(f"\nFerdig! Skrevet til: {UTDATA_FIL} ({size_kb} KB)")
    print(f"  Ark konvertert: {len(ark_navn)}")


if __name__ == "__main__":
    main()
