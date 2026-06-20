import json, sys
from collections import defaultdict
from openpyxl import load_workbook
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side

sys.stdout.reconfigure(encoding='utf-8', errors='replace')

from club_country_map import CLUB_COUNTRY
from top_division import TOP_DIVISION_CLUBS
from level_map import LEVEL_MAP, LAND_NO

with open('clubs_new.json', 'r', encoding='utf-8') as f:
    data = json.load(f)

country_total = defaultdict(int)
country_lvl   = defaultdict(lambda: defaultdict(int))  # country → level → count

for team, players in data.items():
    for player, club in players.items():
        if not club:
            continue
        club = club.strip()
        country = LAND_NO.get(CLUB_COUNTRY.get(club, 'Unknown'), CLUB_COUNTRY.get(club, 'Unknown'))
        country_total[country] += 1

        if club in TOP_DIVISION_CLUBS:
            country_lvl[country][1] += 1
        elif club in LEVEL_MAP:
            lvl = LEVEL_MAP[club][0]
            country_lvl[country][lvl] += 1

ranked = sorted(country_total.items(), key=lambda x: -x[1])

wb = load_workbook('VM2026_avansert_gruppetabeller_og_sluttspill.xlsx')

sheet_name = 'Spillere etter klubbland'
if sheet_name in wb.sheetnames:
    del wb[sheet_name]
ws = wb.create_sheet(sheet_name)
ws.auto_filter.ref = "A3:I3"

thin   = Side(style='thin', color='AAAAAA')
border = Border(left=thin, right=thin, top=thin, bottom=thin)
center = Alignment(horizontal='center', vertical='center')
left_al = Alignment(horizontal='left', vertical='center')

ws['A1'] = 'VM 2026 — Spillere etter klubbland'
ws['A1'].font = Font(name='Arial', bold=True, size=14, color='1F4E79')
ws.merge_cells('A1:I1')
ws['A1'].alignment = center

for col, header in enumerate(['Plass','Land','Spillere','Nivå 1','Nivå 2','Nivå 3','Nivå 4','Nivå 5','Nivå 6'], 1):
    cell = ws.cell(row=3, column=col, value=header)
    cell.font = Font(name='Arial', bold=True, size=11, color='FFFFFF')
    cell.fill = PatternFill('solid', start_color='1F4E79')
    cell.alignment = center
    cell.border = border

for i, (country, total) in enumerate(ranked, 1):
    row = i + 3
    lvls = country_lvl.get(country, {})
    fill = PatternFill('solid', start_color='D6E4F0') if i % 2 == 0 else PatternFill('solid', start_color='FFFFFF')
    for col, (val, al) in enumerate(zip(
        [i, country, total,
         lvls.get(1,0), lvls.get(2,0), lvls.get(3,0), lvls.get(4,0), lvls.get(5,0), lvls.get(6,0)],
        [center, left_al] + [center]*7
    ), 1):
        cell = ws.cell(row=row, column=col, value=val if val != 0 else '')
        cell.font = Font(name='Arial', size=11)
        cell.fill = fill
        cell.alignment = al
        cell.border = border
    # Always show numeric 0-values for totals and rank
    ws.cell(row=row, column=1).value = i
    ws.cell(row=row, column=3).value = total

total_row = len(ranked) + 4
for col, val, al in zip(range(1, 10),
    ['', 'Totalt'] + [f'=SUM({chr(64+c)}4:{chr(64+c)}{total_row-1})' for c in range(3, 10)],
    [center, left_al] + [center]*7):
    cell = ws.cell(row=total_row, column=col, value=val)
    cell.font = Font(name='Arial', bold=True, size=11)
    cell.fill = PatternFill('solid', start_color='E2EFDA')
    cell.alignment = al
    cell.border = border

ws.cell(row=total_row+2, column=1,
        value='Kilde: clubs_new.json — VM 2026 registrerte tropper').font = \
    Font(name='Arial', size=9, italic=True, color='888888')

for col, width in zip('ABCDEFGHI', [8, 22, 10, 10, 10, 10, 10, 10, 10]):
    ws.column_dimensions[col].width = width

wb.save('VM2026_avansert_gruppetabeller_og_sluttspill.xlsx')

total_spillere = sum(v for _, v in ranked)
total_klassifisert = sum(sum(lvls.values()) for lvls in country_lvl.values())
total_uklass = total_spillere - total_klassifisert

FORVENTET = 1248
sjekk = '✓' if total_spillere == FORVENTET else f'ADVARSEL: forventet {FORVENTET}'
print(f'\nLagret OK — {len(ranked)} land, {total_spillere} spillere ({sjekk})')
if total_uklass:
    print(f'  NB: {total_uklass} spillere uten nivåklassifisering')
print()
print(f'{"Land":<22} {"Tot":>5} {"L1":>5} {"L2":>5} {"L3":>5} {"L4":>5} {"L5":>5} {"Ukl":>5}')
print('-' * 65)
for country, total in ranked:
    lvls = country_lvl.get(country, {})
    klassifisert = sum(lvls.values())
    uklass = total - klassifisert
    ukl_str = str(uklass) if uklass else ''
    print(f'{country:<22} {total:>5} {lvls.get(1,0) or "":>5} {lvls.get(2,0) or "":>5} '
          f'{lvls.get(3,0) or "":>5} {lvls.get(4,0) or "":>5} {lvls.get(5,0) or "":>5} {ukl_str:>5}')
print('-' * 65)
tot_lvls = {}
for lvls in country_lvl.values():
    for k, v in lvls.items():
        tot_lvls[k] = tot_lvls.get(k, 0) + v
print(f'{"Totalt":<22} {total_spillere:>5} {tot_lvls.get(1,0):>5} {tot_lvls.get(2,0):>5} '
      f'{tot_lvls.get(3,0):>5} {tot_lvls.get(4,0):>5} {tot_lvls.get(5,0):>5} {total_uklass or "":>5}')
