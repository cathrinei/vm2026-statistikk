#!/usr/bin/env python3
"""
oppdater.py
Daglig oppdatering av VM2026_avansert_gruppetabeller_og_sluttspill.xlsx.

Kjørerekkefølge:
  1. add_kamper_sheets.py           — kampresultater + gruppetabeller (inkl. tilskuere + stadion)
  2. add_sluttspill_sheet.py        — sluttspill-ark (16-delsfinale → Finale); alltid
  3–6. (hoppes over hvis ingen nye kamper siden sist)
  3. check_avvik.py --fix           — mål, assist, gule/røde kort; Scorere og assists/Kort-ark
  4. add_alder_sheet.py             — oppdaterer Alder-ark; henter dato for ev. nye spillere
  5. add_birthdate_column.py        — fyll inn fødselsdatoer i gruppearkene (ev. nye spillere)
  6. add_minutter_column.py         — minutter spilt per spiller i gruppearkene
  7. add_lagstatistikk_sheet.py     — mål/nullere/straffespark/selvmål/skudd/formasjoner/bytter
  8. add_scoringstidspunkt_sheet.py — mål per 15-min intervall + søylediagram
  9. add_heatmap_sheet.py           — mål per minutt som fargekodet heatmap
 10. add_ballbesittelse_sheet.py    — ballbesittelse og skudd per kamp (leser wc2026_stats3.json)
 10b. add_stadion_sheet.py         — tilskuere og beleggsgrad per kamp (leser tilskuere_cache.json)
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
    Rangering i Kort-ark: røde øverst, deretter gule — med gull/sølv/bronse (check_avvik.py skriv_kort_sheet, 2026-06-18).
    Scorere og assists: viser #, Spiller, Lag, Mål, Assist — erstatter Toppscorere+Assists (2026-06-21).
    Skudd-seksjon: hentes fra FIFA /topseasonplayerstatistics via
    TotalAttempts + AttemptsOnTarget per spiller, aggregert per lag. Gratis, ingen cache.
    Formasjoner + spillerbytter (2026-06-19): henholdsvis /live/football/{mid} (cachet
    i live_cache.json) og TypeLocalized="Substitution" fra lagstatistikk_cache.json.
    Ved nye tropp-endringer: oppdater players.json, kjør fetch_alle_fodselsdatoer.py,
    og legg til ny spiller i clubs_new.json + club_country_map.py manuelt.
"""
import io, json, os, subprocess, sys, time
from datetime import datetime
from pathlib import Path
BASE_DIR = Path(__file__).parent

sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding="utf-8", errors="replace")

PYTHON          = sys.executable
CACHE_PATH           = Path(BASE_DIR / "kamper_resultater.json")
SLUTTSPILL_CACHE     = Path(BASE_DIR / "sluttspill_cache.json")
CLUBS_JSON           = Path(BASE_DIR / "clubs_new.json")
PLAYERS_JSON         = Path(BASE_DIR / "players.json")
TOP_DIVISION_PY      = Path(BASE_DIR / "top_division.py")
LEVEL_MAP_PY         = Path(BASE_DIR / "level_map.py")
EXCEL_PATH           = Path(BASE_DIR / "VM2026_avansert_gruppetabeller_og_sluttspill.xlsx")


def log(melding: str, level: str = "INFO"):
    ts = datetime.now().strftime("%H:%M:%S")
    print(f"[{ts}] [{level}] {melding}", flush=True)


def skriv_github_summary(linjer: list[str]):
    summary_fil = os.environ.get("GITHUB_STEP_SUMMARY")
    if not summary_fil:
        return
    with open(summary_fil, "a", encoding="utf-8") as f:
        f.write("\n".join(linjer) + "\n")


def count_spilte() -> int:
    if not CACHE_PATH.exists():
        return 0
    with open(CACHE_PATH, encoding="utf-8") as f:
        cache = json.load(f)
    return sum(1 for g in cache.values() for k in g if k.get("spilt"))


def count_spilte_sluttspill() -> int:
    if not SLUTTSPILL_CACHE.exists():
        return 0
    with open(SLUTTSPILL_CACHE, encoding="utf-8") as f:
        cache = json.load(f)
    return sum(1 for runde in cache.values() for k in runde if k.get("spilt"))


def kjør_steg(navn, cmd, stopp_ved_feil=True) -> str:
    print(f"\n{'─'*55}")
    print(f"  {navn}")
    print(f"{'─'*55}")
    start = time.time()
    linjer = []
    with subprocess.Popen(cmd, stdout=subprocess.PIPE, stderr=subprocess.STDOUT,
                          text=True, encoding="utf-8", errors="replace") as p:
        for linje in p.stdout:
            print(linje, end="", flush=True)
            linjer.append(linje)
        p.wait()
    elapsed = time.time() - start
    ok = p.returncode == 0
    print(f"\n  [{'✓ OK' if ok else '✗ FEIL'}]  ({elapsed:.1f}s)")
    if not ok and stopp_ved_feil:
        log(f"Steg feilet: {navn} — stopper.", level="ERROR")
        sys.exit(1)
    return "".join(linjer)


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
    spilte_gruppe_før      = count_spilte()
    spilte_sluttspill_før  = count_spilte_sluttspill()

    kamper_cmd = [PYTHON, "add_kamper_sheets.py"] + (["--full"] if full else [])
    kamper_output = kjør_steg("Kampresultater + gruppetabeller", kamper_cmd)

    # ── Steg 2: Sluttspill-ark (alltid) ─────────────────────────────────────
    kjør_steg("Sluttspill-ark", [PYTHON, "add_sluttspill_sheet.py"])

    spilte_gruppe_etter     = count_spilte()
    spilte_sluttspill_etter = count_spilte_sluttspill()
    nye_kamper              = spilte_gruppe_etter - spilte_gruppe_før
    nye_sluttspill          = spilte_sluttspill_etter - spilte_sluttspill_før
    spilte_etter            = spilte_gruppe_etter  # for summary

    # ── Steg 3: check_avvik — ved nye kamper ELLER i sluttspillfasen ────────
    # Kjøres alltid når det finnes spilte sluttspill-kamper, slik at gruppe-arkene
    # alltid viser totalstatistikk (gruppe + sluttspill), ikke bare gruppespillmål.
    if nye_kamper > 0 or nye_sluttspill > 0 or spilte_sluttspill_etter > 0 or full:
        antall = nye_kamper + nye_sluttspill
        if antall > 0:
            print(f"\n  {antall} ny{'e' if antall != 1 else ''} kamp{'er' if antall != 1 else ''} "
                  f"({nye_kamper} gruppe, {nye_sluttspill} sluttspill) — oppdaterer statistikk...")
        else:
            print(f"\n  Sluttspillfase aktiv — oppdaterer totalstatistikk i gruppe-ark...")
        kjør_steg("Mål, assist og kort", [PYTHON, "check_avvik.py", "--fix"])
    else:
        print(f"\n  Ingen nye kamper siden sist — hopper over check_avvik.")

    # ── Steg 4–6: kun ved nye gruppekamper (eller full) ─────────────────────
    # Steg 6 (minutter) kjøres også i sluttspillfasen — spillere akkumulerer
    # minutter gjennom hele VM, ikke bare i gruppespillet.
    if nye_kamper > 0 or full:
        kjør_steg("Fødselsdatoer + Alder-ark",       [PYTHON, "add_alder_sheet.py"])
        kjør_steg("Fødselsdato-kolonne i gruppeark", [PYTHON, "add_birthdate_column.py"])
        kjør_steg("Minutter spilt i gruppeark",      [PYTHON, "add_minutter_column.py"])
    elif spilte_sluttspill_etter > 0:
        kjør_steg("Minutter spilt i gruppeark",      [PYTHON, "add_minutter_column.py"])
    elif nye_sluttspill == 0:
        print(f"\n  Ingen nye gruppe-kamper — hopper over alder/fødselsdato/minutter.")

    # ── Steg 7: Lagstatistikk (alltid — caches internt) ──────────────────────
    kjør_steg("Lagstatistikk", [PYTHON, "add_lagstatistikk_sheet.py"])

    # ── Steg 7b: Sluttspill-spillere (alltid — henter fra cacher) ────────────
    kjør_steg("Sluttspill – spillerstatistikk", [PYTHON, "add_sluttspill_spillere_sheet.py"])

    # ── Steg 8: Scoringstidspunkt (alltid — leser fra lagstatistikk_cache.json) ──
    kjør_steg("Scoringstidspunkt", [PYTHON, "add_scoringstidspunkt_sheet.py"])

    # ── Steg 9: Heatmap (alltid — leser fra lagstatistikk_cache.json) ────────
    kjør_steg("Heatmap", [PYTHON, "add_heatmap_sheet.py"])

    # ── Steg 10: Ballbesittelse (alltid — leser fra wc2026_possession.json) ─────
    kjør_steg("Ballbesittelse", [PYTHON, "add_ballbesittelse_sheet.py"])

    # ── Steg 10b: Stadionoversikt (alltid — leser fra tilskuere_cache.json) ───
    kjør_steg("Stadionoversikt", [PYTHON, "add_stadion_sheet.py"])

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

    endrede_grupper = next(
        (l.strip().split("Endringer i: ", 1)[1] for l in kamper_output.splitlines() if "Endringer i:" in l),
        None
    )
    summary = [
        f"## VM 2026 oppdatering {datetime.now().strftime('%Y-%m-%d')}",
        f"- **Trigger:** `{os.environ.get('GITHUB_EVENT_NAME', 'lokal')}`",
        f"- **Kjøretid:** {elapsed:.0f}s",
        f"- **Kamper totalt:** {spilte_etter}",
    ]
    if nye_kamper > 0:
        summary.append(f"- **Nye kamper:** {nye_kamper}")
    if endrede_grupper:
        summary.append(f"- **Oppdaterte grupper:** {endrede_grupper}")
    elif nye_kamper == 0:
        summary.append("- Ingen nye kamper siden sist")
    skriv_github_summary(summary)


if __name__ == "__main__":
    main()
