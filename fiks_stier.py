"""
Erstatter hardkodede C:\Claude_dev\FotballVMClaude\... stier
med BASE_DIR / "filnavn" i alle Python-skript.
"""

import re
import os

MAPPE = r"C:\Claude_dev\FotballVMClaude"

MØNSTER = re.compile(r'r"C:\\Claude_dev\\FotballVMClaude\\([^"]+)"')

BASE_DIR_LINJE = 'BASE_DIR = Path(__file__).parent\n'
IMPORT_LINJE = 'from pathlib import Path\n'

endrede = []

for filnavn in os.listdir(MAPPE):
    if not filnavn.endswith(".py"):
        continue
    if filnavn in ("fiks_stier.py",):
        continue

    sti = os.path.join(MAPPE, filnavn)
    with open(sti, encoding="utf-8") as f:
        innhold = f.read()

    if 'C:\\Claude_dev\\FotballVMClaude\\' not in innhold:
        continue

    # Erstatt alle hardkodede stier
    nytt = MØNSTER.sub(lambda m: f'BASE_DIR / "{m.group(1)}"', innhold)

    # Legg til BASE_DIR-linje etter pathlib-import hvis ikke allerede der
    if 'BASE_DIR' not in nytt:
        continue  # ingen erstatninger ble gjort — hopp over

    linjer = nytt.splitlines(keepends=True)
    ny_linjer = []
    pathlib_funnet = False
    base_dir_lagt_til = False

    for i, linje in enumerate(linjer):
        ny_linjer.append(linje)

        # Legg til Path-import hvis mangler
        if not pathlib_funnet and linje.startswith("from pathlib import Path"):
            pathlib_funnet = True

        # Etter siste import-blokk: legg til BASE_DIR
        if (not base_dir_lagt_til and 'BASE_DIR' in nytt
                and linje.strip() == "" and i > 0
                and any(linjer[j].startswith(("import ", "from ")) for j in range(max(0, i-5), i))):
            if not any("BASE_DIR" in l for l in ny_linjer):
                if not pathlib_funnet:
                    ny_linjer.insert(-1, IMPORT_LINJE)
                    pathlib_funnet = True
                ny_linjer.insert(-1, BASE_DIR_LINJE)
                base_dir_lagt_til = True

    resultat = "".join(ny_linjer)

    # Siste sjanse: legg til øverst hvis ikke plassert ennå
    if not any("BASE_DIR = Path" in l for l in resultat.splitlines()):
        topp = []
        resten = resultat.splitlines(keepends=True)
        sett_inn_etter = 0
        for i, l in enumerate(resten):
            if l.startswith(("import ", "from ")):
                sett_inn_etter = i + 1
        resten.insert(sett_inn_etter, BASE_DIR_LINJE)
        if not any("from pathlib import Path" in l for l in resten):
            resten.insert(sett_inn_etter, IMPORT_LINJE)
        resultat = "".join(resten)

    with open(sti, "w", encoding="utf-8") as f:
        f.write(resultat)

    endrede.append(filnavn)
    print(f"  Fikset: {filnavn}")

print(f"\nFerdig — {len(endrede)} filer oppdatert.")
