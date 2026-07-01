#!/usr/bin/env python3
"""
valider_maltotal.py
Sjekker at beregnet måltotal per spilt kamp stemmer med FIFAs offisielle
sluttresultat (score_h + score_a i kamper_resultater.json/sluttspill_cache.json).

Måltotalen beregnes fra avvik_timeline_cache.json som summen av:
  - spillermål   (Type 0, med reell scoreendring)
  - selvmål      (Type 34)
  - straffemål i vanlig/ekstra spilletid (Type 41, med reell scoreendring)

Straffemål etter ekstraomganger (straffesparkkonkurranse) — Type 41 UTEN
reell scoreendring, der IdSubPlayer er motstanderens keeper — telles ikke
med, siden straffekonkurranse ikke regnes som mål i fotballstatistikk.

Kjøres av oppdater.py rett etter check_avvik.py, slik at avvik i rådata
(f.eks. sluttspillkamper som mangler fra kamp-oppslaget, eller feilklassifiserte
eventtyper) fanges opp før videre Excel-oppdatering.

Exit code 1 hvis ett eller flere avvik funnet.
"""
import io
import json
import sys
from pathlib import Path

sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding="utf-8", errors="replace")

BASE_DIR       = Path(__file__).parent
KAMPER_FIL     = BASE_DIR / "kamper_resultater.json"
SLUTTSPILL_FIL = BASE_DIR / "sluttspill_cache.json"
TIMELINE_FIL   = BASE_DIR / "avvik_timeline_cache.json"

MÅL_TYPER    = (0, 41)
SELVMÅL_TYPE = 34


def _last(fil: Path) -> dict:
    if not fil.exists():
        return {}
    with open(fil, encoding="utf-8") as f:
        return json.load(f)


def valider() -> list[tuple]:
    timeline       = _last(TIMELINE_FIL)
    kamper_raw     = _last(KAMPER_FIL)
    sluttspill_raw = _last(SLUTTSPILL_FIL)

    kamp_for_id = {}
    for liste in list(kamper_raw.values()) + list(sluttspill_raw.values()):
        for k in liste:
            if k.get("spilt") and k.get("dato"):
                kamp_for_id[k["id"]] = k

    avvik = []
    for matchid, kamp in kamp_for_id.items():
        events = timeline.get(matchid)
        if not events:
            continue  # timeline ikke hentet ennå — sjekkes neste kjøring

        fifa_totalt = (kamp.get("score_h") or 0) + (kamp.get("score_a") or 0)

        spillermål, selvmål, straffe_vanlig, straffe_konkurranse = 0, 0, 0, 0
        prev_h, prev_a = 0, 0
        for e in events:
            t = e.get("Type")
            cur_h = e["HomeGoals"] if e.get("HomeGoals") is not None else prev_h
            cur_a = e["AwayGoals"] if e.get("AwayGoals") is not None else prev_a
            økt = cur_h > prev_h or cur_a > prev_a
            if t == 0 and økt:
                spillermål += 1
            elif t == SELVMÅL_TYPE:
                selvmål += 1
            elif t == 41 and økt:
                straffe_vanlig += 1
            elif t == 41 and not økt:
                straffe_konkurranse += 1
            prev_h, prev_a = cur_h, cur_a

        beregnet = spillermål + selvmål + straffe_vanlig
        if beregnet != fifa_totalt:
            avvik.append((matchid, kamp.get("hjemme"), kamp.get("borte"), fifa_totalt, beregnet,
                          spillermål, selvmål, straffe_vanlig, straffe_konkurranse))

    return avvik, len(kamp_for_id)


def main():
    print("=" * 55)
    print("  Validerer måltotal mot FIFA-sluttresultat")
    print("=" * 55)

    avvik, antall_kamper = valider()

    print(f"\nSjekket {antall_kamper} spilte kamper.")
    if not avvik:
        print("  Ingen avvik — måltotalen stemmer med FIFA i alle kamper.")
        sys.exit(0)

    print(f"\n  {len(avvik)} kamp(er) med avvik mellom beregnet måltotal og FIFA:\n")
    for mid, h, b, fifa, ber, sm, sg, sv, se in avvik:
        print(f"  Kamp {mid} ({h} vs {b}): FIFA={fifa}, beregnet={ber}")
        print(f"    spillermål={sm}  selvmål={sg}  straffe(vanlig)={sv}  straffe(konkurranse, ekskl.)={se}")
    print("\n  Sjekk avvik_timeline_cache.json manuelt mot fifa.com match centre for disse kampene.")
    sys.exit(1)


if __name__ == "__main__":
    main()
