#!/usr/bin/env python3
"""
oppdater.py
Daglig oppdatering av VM2026_avansert_gruppetabeller_og_sluttspill.xlsx.

Kjørerekkefølge:
  1. add_kamper_sheets.py           — kampresultater + gruppetabeller (inkl. tilskuere + stadion)
  2. add_sluttspill_sheet.py        — sluttspill-ark (32-delsfinale → Finale); alltid
  3–6. (hoppes over hvis ingen nye kamper siden sist)
  3. check_avvik.py --fix           — mål, assist, gule/røde kort; Toppscorere/Assists/Kort-ark
  4. add_alder_sheet.py             — oppdaterer Alder-ark; henter dato for ev. nye spillere
  5. add_birthdate_column.py        — fyll inn fødselsdatoer i gruppearkene (ev. nye spillere)
  6. add_minutter_column.py         — minutter spilt per spiller i gruppearkene
  7. add_lagstatistikk_sheet.py     — mål/nullere/straffespark/selvmål/skudd/formasjoner/bytter
  8. add_scoringstidspunkt_sheet.py — mål per 15-min intervall + søylediagram
  9. add_heatmap_sheet.py           — mål per minutt som fargekodet heatmap
 10. add_ballbesittelse_sheet.py    — ballbesittelse og skudd per kamp (leser wc2026_stats3.json)
 11. add_club_column.py             — fyll inn klubbnavn i gruppearkene (alltid — caches internt)

Flagg:
  --full          tvinger full gjeninnlesing av alle kamper
  --klubber       tvinger Klubber- og Spillere etter klubbland-ark
                  (kjøres automatisk hvis clubs_new.json er nyere enn Excel)
  --fodselsdatoer henter alle fødselsdatoer på nytt (fetch_alle_fodselsdatoer.py)

Automatiske triggere (mtime-sjekk):
  clubs_new.json, top_division.py eller level_map.py nyere → Klubber-ark + Spillere etter klubbland
  players.json nyere    → Aldersfordeling
  top_division.py eller level_map.py nyere → Nivå 2 og lavere

NB: Alle 1248 spillere har fødselsdato og klubb (100% dekningsgrad, per 2026-06-18).
    Tropp-endringer: Livramento->Chalobah (ENG), Garbett->Rogerson (NZL),
    Masharipov->Jiyanov (UZB), Issaoui->Slimane (TUN), Oliveros->Sanabria (URU).
    clubs_new.json: bruker players.json-lagnavn som nøkler; 471 unike klubber etter
    standardisering (fix_clubs_names.py + standardiser_klubber.py, 2026-06-18).
    Kort-ark: rangeres nå røde kort øverst, deretter gule — med gull/sølv/bronse
    som Toppscorere/Assists (check_avvik.py skriv_kort_sheet, 2026-06-18).
    Toppscorere/Assists: viser kun #, Spiller, Lag, Mål, Assist (Gule/Røde fjernet 2026-06-20).
    Skudd-seksjon: hentes fra FIFA /topseasonplayerstatistics via
    TotalAttempts + AttemptsOnTarget per spiller, aggregert per lag. Gratis, ingen cache.
    Formasjoner + spillerbytter (2026-06-19): henholdsvis /live/football/{mid} (cachet
    i live_cache.json) og TypeLocalized="Substitution" fra lagstatistikk_cache.json.
    Ved nye tropp-endringer: oppdater players.json, kjør fetch_alle_fodselsdatoer.py,
    og legg til ny spiller i clubs_new.json + club_country_map.py manuelt.
"""
import io, json, subprocess, sys, time
from datetime import datetime
from pathlib import Path
BASE_DIR = Path(__file__).parent

sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding="utf-8", errors="replace")

PYTHON          = sys.executable
CACHE_PATH      = Path(BASE_DIR / "kamper_resultater.json")
CLUBS_JSON      = Path(BASE_DIR / "clubs_new.json")
PLAYERS_JSON    = Path(BASE_DIR / "players.json")
TOP_DIVISION_PY = Path(BASE_DIR / "top_division.py")
LEVEL_MAP_PY    = Path(BASE_DIR / "level_map.py")
EXCEL_PATH      = Path(BASE_DIR / "VM2026_avansert_gruppetabeller_og_sluttspill.xlsx")


def count_spilte() -> int:
    if not CACHE_PATH.exists():
        return 0
    with open(CACHE_PATH, encoding="utf-8") as f:
        cache = json.load(f)
    return sum(1 for g in cache.values() for k in g if k.get("spilt"))


def kjør_steg(navn, cmd, stopp_ved_feil=True) -> bool:
    print(f"\n{'─'*55}")
    print(f"  {navn}")
    print(f"{'─'*55}")
    start  = time.time()
    result = subprocess.run(cmd, capture_output=False, text=False)
    elapsed = time.time() - start
    ok = result.returncode == 0
    print(f"\n  [{'✓ OK' if ok else '✗ FEIL'}]  ({elapsed:.1f}s)")
    if not ok and stopp_ved_feil:
        print("  Stopper — rett feilen og kjør oppdater.py på nytt.")
        sys.exit(1)
    return ok


def _nyere_enn_excel(*filer) -> bool:
    if not EXCEL_PATH.exists():
        return False
    excel_mtime = EXCEL_PATH.stat().st_mtime
    return any(f.exists() and f.stat().st_mtime > excel_mtime for f in filer)


def main():
    full          = "--full"          in sys.argv
    klubber       = "--klubber"       in sys.argv
    fodselsdatoer = "--fodselsdatoer" in sys.argv

    print("=" * 55)
    print(f"  VM 2026 — Daglig oppdatering")
    print(f"  {datetime.now().strftime('%d.%m.%Y  %H:%M')}")
    if full:
        print("  Modus: FULL (tvinger ny innlesing av alle kamper)")
    if klubber:
        print("  Flagg:  --klubber (tvinger oppdatering av Klubber- og Spillere etter klubbland-ark)")
    if fodselsdatoer:
        print("  Flagg:  --fodselsdatoer (henter alle fødselsdatoer på nytt)")
    print("=" * 55)

    total_start = time.time()

    # ── Steg 0: Fødselsdatoer (--fodselsdatoer — må kjøre før alt annet) ─────
    if fodselsdatoer:
        kjør_steg("Hent alle fødselsdatoer", [PYTHON, "fetch_alle_fodselsdatoer.py"])

    # ── Steg 1: Kampresultater ────────────────────────────────────────────────
    spilte_før = count_spilte()

    kamper_cmd = [PYTHON, "add_kamper_sheets.py"] + (["--full"] if full else [])
    kjør_steg("Kampresultater + gruppetabeller", kamper_cmd)

    spilte_etter = count_spilte()
    nye_kamper   = spilte_etter - spilte_før

    # ── Steg 2: Sluttspill-ark (alltid) ─────────────────────────────────────
    kjør_steg("Sluttspill-ark", [PYTHON, "add_sluttspill_sheet.py"])

    # ── Steg 3–6: kun ved nye kamper ─────────────────────────────────────────
    if nye_kamper == 0 and not full:
        print(f"\n  Ingen nye kamper siden sist ({spilte_etter} totalt) — "
              f"hopper over statistikk-oppdatering.")
    else:
        if nye_kamper > 0:
            print(f"\n  {nye_kamper} ny{'e' if nye_kamper > 1 else ''} kamp{'er' if nye_kamper > 1 else ''} — oppdaterer statistikk...")

        kjør_steg("Mål, assist og kort",             [PYTHON, "check_avvik.py", "--fix"])
        kjør_steg("Fødselsdatoer + Alder-ark",       [PYTHON, "add_alder_sheet.py"])
        kjør_steg("Fødselsdato-kolonne i gruppeark", [PYTHON, "add_birthdate_column.py"])
        kjør_steg("Minutter spilt i gruppeark",      [PYTHON, "add_minutter_column.py"])

    # ── Steg 7: Lagstatistikk (alltid — caches internt) ──────────────────────
    kjør_steg("Lagstatistikk", [PYTHON, "add_lagstatistikk_sheet.py"])

    # ── Steg 8: Scoringstidspunkt (alltid — leser fra lagstatistikk_cache.json) ──
    kjør_steg("Scoringstidspunkt", [PYTHON, "add_scoringstidspunkt_sheet.py"])

    # ── Steg 9: Heatmap (alltid — leser fra lagstatistikk_cache.json) ────────
    kjør_steg("Heatmap", [PYTHON, "add_heatmap_sheet.py"])

    # ── Steg 10: Ballbesittelse (alltid — leser fra wc2026_possession.json) ─────
    kjør_steg("Ballbesittelse", [PYTHON, "add_ballbesittelse_sheet.py"])

    # ── Steg 11: Klubb-kolonne (alltid — fyller bare tomme celler) ────────────
    kjør_steg("Klubb-kolonne i gruppeark", [PYTHON, "add_club_column.py"])

    # ── Steg 10–11: Klubber- og Spillere etter klubbland-ark (ved endringer) ──────
    kjør_klubber = klubber or _nyere_enn_excel(CLUBS_JSON, TOP_DIVISION_PY, LEVEL_MAP_PY)
    if kjør_klubber:
        årsak = "--klubber" if klubber else "clubs_new.json er nyere enn Excel"
        print(f"\n  Oppdaterer klubb-ark ({årsak}) ...")
        kjør_steg("Klubber-ark",         [PYTHON, "add_clubs_sheet.py"])
        kjør_steg("Spillere etter klubbland", [PYTHON, "add_country_sheet.py"])
    else:
        print("\n  clubs_new.json ikke endret siden sist — hopper over klubb-ark.")

    # ── Steg 12: Aldersfordeling (ved endringer i players.json) ──────────────
    if _nyere_enn_excel(PLAYERS_JSON):
        print("\n  players.json er nyere enn Excel — oppdaterer Aldersfordeling ...")
        kjør_steg("Aldersfordeling", [PYTHON, "add_stats_sheets.py"])
    else:
        print("\n  players.json ikke endret siden sist — hopper over Aldersfordeling.")

    # ── Steg 13: Nivå 2 og lavere (ved endringer i nivådata) ─────────────────
    if _nyere_enn_excel(TOP_DIVISION_PY, LEVEL_MAP_PY):
        endret = ", ".join(
            f.name for f in (TOP_DIVISION_PY, LEVEL_MAP_PY)
            if f.exists() and f.stat().st_mtime > EXCEL_PATH.stat().st_mtime
        )
        print(f"\n  {endret} er nyere enn Excel — oppdaterer Nivå 2 og lavere ...")
        kjør_steg("Nivå 2 og lavere", [PYTHON, "lag_nivå_tabell.py"])
    else:
        print("\n  Nivådata ikke endret siden sist — hopper over Nivå 2 og lavere.")

    elapsed = time.time() - total_start
    print(f"\n{'='*55}")
    print(f"  Ferdig ({elapsed:.0f}s)")
    print("=" * 55)


if __name__ == "__main__":
    main()
