#!/usr/bin/env python3
"""
add_heatmap_sheet.py
Lager arket "Heatmap" med ett minutt per celle (1–90 + tilleggstid),
fargekodet etter antall mål. Kilde: lagstatistikk_cache.json.
"""
import io, json, shutil, sys
from pathlib import Path
sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding="utf-8", errors="replace")
BASE_DIR = Path(__file__).parent

try:
    from openpyxl import load_workbook
    from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
except ImportError as e:
    sys.exit(f"Mangler pakke: {e}")

EXCEL_PATH  = BASE_DIR / "VM2026_avansert_gruppetabeller_og_sluttspill.xlsx"
STAT_CACHE  = BASE_DIR / "lagstatistikk_cache.json"
SHEET_NAME  = "Heatmap"

GOAL_TYPES  = {"Goal!", "Penalty Goal", "Own goal"}

# Fargegradient: 0 mål → hvit, 7 mål → navy
GRADIENT = [
    ("FFFFFF", "AAAAAA"),  # 0 — hvit, grå tekst
    ("DEEBF7", "1A1A2E"),  # 1 — svært lys blå
    ("BDD7EE", "1A1A2E"),  # 2 — lys blå
    ("5B9BD5", "FFFFFF"),  # 3 — mellomblå
    ("2171B5", "FFFFFF"),  # 4 — sterk blå
    ("1A3C6B", "FFFFFF"),  # 5 — mørk blå (designsystem)
    ("0F2044", "FFFFFF"),  # 6+ — navy (designsystem)
]

def _colors(count: int) -> tuple[str, str]:
    idx = min(count, len(GRADIENT) - 1)
    return GRADIENT[idx]


def _parse_min(s: str) -> int | None:
    if not s:
        return None
    s = str(s).strip().rstrip("'")
    if "'+" in s:
        parts = s.replace("'", "").split("+")
        try:
            return int(parts[0]) + int(parts[1])
        except Exception:
            return None
    try:
        return int(s)
    except Exception:
        return None


def tell_per_minutt() -> dict[int, int]:
    if not Path(STAT_CACHE).exists():
        sys.exit(f"FEIL: {STAT_CACHE} finnes ikke.")
    with open(STAT_CACHE, encoding="utf-8") as f:
        cache = json.load(f)
    counts: dict[int, int] = {}
    for mid, events in cache.items():
        for e in events:
            if e.get("TypeLocalized") not in GOAL_TYPES:
                continue
            m = _parse_min(e.get("MatchMinute"))
            if m and m >= 1:
                counts[m] = counts.get(m, 0) + 1
    return counts


def skriv_ark(counts: dict[int, int]) -> None:
    backup = Path(str(EXCEL_PATH) + ".bak")
    shutil.copy2(EXCEL_PATH, backup)
    wb = load_workbook(EXCEL_PATH)

    if SHEET_NAME in wb.sheetnames:
        del wb[SHEET_NAME]

    insert_after = "Scoringstidspunkt"
    if insert_after in wb.sheetnames:
        idx = wb.sheetnames.index(insert_after) + 1
    elif "Lagstatistikk" in wb.sheetnames:
        idx = wb.sheetnames.index("Lagstatistikk") + 1
    else:
        idx = len(wb.sheetnames)
    ws = wb.create_sheet(SHEET_NAME, idx)

    NAVY  = PatternFill("solid", fgColor="0F2044")
    BLUE  = PatternFill("solid", fgColor="1A3C6B")
    GRAY  = PatternFill("solid", fgColor="D9D9D9")
    thin  = Side(style="thin",   color="E2E8F0")
    thck  = Side(style="medium", color="0F2044")
    bot   = Border(bottom=thin)
    ctr   = Alignment(horizontal="center", vertical="center")
    lft   = Alignment(horizontal="left",   vertical="center")

    f_title = Font(name="Calibri", bold=True, size=13, color="FFFFFF")
    f_hdr   = Font(name="Calibri", bold=True, size=9,  color="FFFFFF")
    f_sub   = Font(name="Calibri", bold=True, size=8,  color="FFFFFF")
    f_gray  = Font(name="Calibri", size=8, color="888888")

    CELL_W = 4.5   # bredde per minutt-celle
    CELL_H = 20    # høyde per rad

    # ── Rad 1: Tittel ─────────────────────────────────────────────────────────
    ws.row_dimensions[1].height = 28
    ws.merge_cells("A1:L1")
    c = ws.cell(row=1, column=1, value="Scoringstidspunkt — Heatmap (mål per minutt)")
    c.font = f_title; c.fill = NAVY; c.alignment = lft

    # ── Rad 2: Kolonneoverskrifter (siffer: +1 ... +10) ───────────────────────
    ws.row_dimensions[2].height = 16
    ws.cell(row=2, column=1, value="")
    c = ws.cell(row=2, column=1)
    c.fill = NAVY

    for col_i in range(10):
        col = col_i + 2   # B=2 ... K=11
        c = ws.cell(row=2, column=col, value=col_i + 1)
        c.font = f_hdr; c.fill = BLUE; c.alignment = ctr

    # ── Rader 3–11: minutter 1–90 (9 rader × 10 kolonner) ────────────────────
    for row_i in range(9):
        row = row_i + 3
        ws.row_dimensions[row].height = CELL_H
        row_start = row_i * 10 + 1   # 1, 11, 21, ..., 81
        row_end   = row_start + 9    # 10, 20, ..., 90
        label = f"{row_start}–{row_end}"

        # Radlabel (kolonne A)
        lc = ws.cell(row=row, column=1, value=label)
        lc.font = f_hdr; lc.fill = NAVY; lc.alignment = ctr

        # Halvtids-separator: tykk kant under rad for minutt 41-50 (etter min 45)
        # Markeres med tykk kant under celle 41-50-raden (row_i=4, minutt 41-50)
        # Minutt 45 er col_i=4 i row_i=4
        for col_i in range(10):
            col    = col_i + 2
            minutt = row_start + col_i
            n      = counts.get(minutt, 0)
            bg, fg = _colors(n)

            cell = ws.cell(row=row, column=col, value=n if n > 0 else None)
            cell.fill      = PatternFill("solid", fgColor=bg)
            cell.font      = Font(name="Calibri", bold=(n > 0), size=10, color=fg)
            cell.alignment = ctr

            # Halvtidsmarkering: tykk kant til høyre for minutt 45 og til venstre for 46
            left_side  = thck if minutt == 46 else thin
            right_side = thck if minutt == 45 else thin
            top_side   = thin
            bot_side   = thin
            cell.border = Border(left=left_side, right=right_side,
                                 top=top_side, bottom=bot_side)

    # ── Rad 12: separator ─────────────────────────────────────────────────────
    ws.row_dimensions[12].height = 8
    for col in range(1, 12):
        c = ws.cell(row=12, column=col)
        c.fill = GRAY

    # ── Rad 13: tilleggstid-header ────────────────────────────────────────────
    ws.row_dimensions[13].height = 16
    lc = ws.cell(row=13, column=1, value="ET")
    lc.font = f_hdr; lc.fill = NAVY; lc.alignment = ctr

    for col_i in range(10):
        col = col_i + 2
        c = ws.cell(row=13, column=col, value=f"+{col_i + 1}")
        c.font = f_sub; c.fill = BLUE; c.alignment = ctr

    # ── Rad 14: tilleggstid-data (minutt 91–100) ──────────────────────────────
    ws.row_dimensions[14].height = CELL_H
    lc = ws.cell(row=14, column=1, value="90+")
    lc.font = f_hdr; lc.fill = NAVY; lc.alignment = ctr

    for col_i in range(10):
        col    = col_i + 2
        minutt = 91 + col_i
        n      = counts.get(minutt, 0)
        bg, fg = _colors(n)
        cell = ws.cell(row=14, column=col, value=n if n > 0 else None)
        cell.fill      = PatternFill("solid", fgColor=bg)
        cell.font      = Font(name="Calibri", bold=(n > 0), size=10, color=fg)
        cell.alignment = ctr
        cell.border    = Border(left=thin, right=thin, top=thin, bottom=thin)

    # ── Rad 16–17: Forklaring ─────────────────────────────────────────────────
    ws.row_dimensions[16].height = 14
    ws.cell(row=16, column=1, value="Fargeskala:").font = Font(name="Calibri", bold=True, size=9, color="1A1A2E")

    legend_items = [
        ("0 mål",  "FFFFFF", "AAAAAA"),
        ("1 mål",  "DEEBF7", "1A1A2E"),
        ("2 mål",  "BDD7EE", "1A1A2E"),
        ("3 mål",  "5B9BD5", "FFFFFF"),
        ("4 mål",  "2171B5", "FFFFFF"),
        ("5 mål",  "1A3C6B", "FFFFFF"),
        ("6+ mål", "0F2044", "FFFFFF"),
    ]
    ws.row_dimensions[17].height = 18
    for i, (label, bg, fg) in enumerate(legend_items):
        col = i + 2
        c = ws.cell(row=17, column=col, value=label)
        c.fill      = PatternFill("solid", fgColor=bg)
        c.font      = Font(name="Calibri", size=8, color=fg)
        c.alignment = ctr
        c.border    = Border(left=thin, right=thin, top=thin, bottom=thin)

    # ── Halvtidsmarkeringsforklaring ───────────────────────────────────────────
    ws.row_dimensions[19].height = 14
    note = ws.cell(row=19, column=1,
                   value="| Tykk strek markerer halvtidspausen (mellom minutt 45 og 46)")
    note.font = Font(name="Calibri", italic=True, size=8, color="6B7A99")

    # ── Kolonnebredder ────────────────────────────────────────────────────────
    ws.column_dimensions["A"].width = 7
    for col_i in range(10):
        letter = chr(ord("B") + col_i)
        ws.column_dimensions[letter].width = CELL_W
    ws.column_dimensions["L"].width = 3  # liten buffer-kolonne

    try:
        wb.save(EXCEL_PATH)
    except Exception as e:
        shutil.copy2(backup, EXCEL_PATH)
        sys.exit(f"FEIL — rullet tilbake: {e}")


def main():
    print("=" * 55)
    print("  VM 2026 — Scoringstidspunkt Heatmap")
    print("=" * 55)

    print("\n[1/2] Teller mål per minutt fra cache...")
    counts = tell_per_minutt()
    totalt = sum(counts.values())
    maks   = max(counts.values()) if counts else 0
    print(f"  {totalt} mål fordelt over {len(counts)} forskjellige minutter")
    print(f"  Maks i ett minutt: {maks} mål (minutt {max(counts, key=counts.get)})")

    print(f"\n[2/2] Skriver ark '{SHEET_NAME}' til Excel...")
    skriv_ark(counts)
    print("\nFerdig.")


if __name__ == "__main__":
    main()
