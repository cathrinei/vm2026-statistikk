#!/usr/bin/env python3
"""
fix_clubs_names.py
Re-matcher clubs_new.json mot players.json ved å normalisere navn.
Håndterer: usynlige tegn, aksentforskjeller, navnerekkefølge,
Australia (kun etternavn), Japan (etternavn-fornavn).
Skriver ny clubs_new.json med players.json-navn som nøkler.
"""
import io, json, re, sys, unicodedata
sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding="utf-8", errors="replace")

with open(r"C:\Claude_dev\FotballVMClaude\clubs_new.json", encoding="utf-8") as f:
    clubs_raw = json.load(f)
with open(r"C:\Claude_dev\FotballVMClaude\players.json", encoding="utf-8") as f:
    players = json.load(f)

# ── Lagnavn: players.json → clubs_new.json ────────────────────────────────────
PJSON_TO_CLUBS = {
    "Mexico":                   "Mexico",
    "South Africa":             "South Africa",
    "South Korea":              "Republic of Korea",
    "Korea Republic":           "Republic of Korea",
    "Canada":                   "Canada",
    "Bosnia and Herzegovina":   "Bosnia and Herzegovina",
    "Qatar":                    "Qatar",
    "Switzerland":              "Switzerland",
    "Brazil":                   "Brazil",
    "Morocco":                  "Morocco",
    "Haiti":                    "Haiti",
    "Scotland":                 "Scotland",
    "United States":            "United States of America",
    "USA":                      "United States of America",
    "Paraguay":                 "Paraguay",
    "Australia":                "Australia",
    "Turkey":                   "Türkiye",
    "Türkiye":                  "Türkiye",
    "Germany":                  "Germany",
    "Curacao":                  "Curaçao",
    "Curaçao":                  "Curaçao",
    "Ivory Coast":              "Côte d'Ivoire",
    "Côte d'Ivoire":            "Côte d'Ivoire",
    "Elfenbenskysten":          "Côte d'Ivoire",
    "Ecuador":                  "Ecuador",
    "Sweden":                   "Sweden",
    "Japan":                    "Japan",
    "New Zealand":              "New Zealand",
    "Tunisia":                  "Tunisia",
    "Spain":                    "Spain",
    "Cape Verde":               "Cabo Verde",
    "Cabo Verde":               "Cabo Verde",
    "Kapp Verde":               "Cabo Verde",
    "Saudi Arabia":             "Saudi Arabia",
    "Uruguay":                  "Uruguay",
    "Netherlands":              "Netherlands",
    "Belgium":                  "Belgium",
    "Egypt":                    "Egypt",
    "Jordan":                   "Jordan",
    "Norway":                   "Norway",
    "Norge":                    "Norway",
    "France":                   "France",
    "Senegal":                  "Senegal",
    "Colombia":                 "Colombia",
    "Argentina":                "Argentina",
    "Panama":                   "Panama",
    "England":                  "England",
    "Algeria":                  "Algeria",
    "Croatia":                  "Croatia",
    "Iran":                     "Islamic Republic of Iran",
    "IR Iran":                  "Islamic Republic of Iran",
    "Iraq":                     "Iraq",
    "Portugal":                 "Portugal",
    "Ghana":                    "Ghana",
    "DR Congo":                 "Democratic Republic of the Congo",
    "DR Kongo":                 "Democratic Republic of the Congo",
    "Uzbekistan":               "Uzbekistan",
    "Usbekistan":               "Uzbekistan",
    "Austria":                  "Austria",
    "Østerrike":                "Austria",
    "Sør-Korea":                "Republic of Korea",
    "Tsjekkia":                 "Czechia",
    "Tyrkia":                   "Türkiye",
    "Sverige":                  "Sweden",
    "Frankrike":                "France",
    "Algerie":                  "Algeria",
    "Nederland":                "Netherlands",
    "Belgia":                   "Belgium",
    "Marokko":                  "Morocco",
    "Skottland":                "Scotland",
    "Brasil":                   "Brazil",
    "Irak":                     "Iraq",
    "Saudi-Arabia":             "Saudi Arabia",
    "Kroatia":                  "Croatia",
    "Elfenbenskysten":          "Côte d'Ivoire",
    "Bosnia-Hercegovina":       "Bosnia and Herzegovina",
    "Sveits":                   "Switzerland",
    "Colombia":                 "Colombia",
    "Kapp Verde":               "Cabo Verde",
}

# ── Normalisering ─────────────────────────────────────────────────────────────
_SPECIAL = str.maketrans({
    "ø":"o","Ø":"o","æ":"ae","Æ":"ae","å":"a","Å":"a",
    "ß":"ss","ð":"d","þ":"th","ı":"i",
    "ü":"ue","Ü":"ue","ö":"oe","Ö":"oe","ä":"ae","Ä":"ae",
})
_JR = re.compile(r"\bjr\.?$", re.I)

def norm(s):
    def _fix(c):
        cat = unicodedata.category(c)
        if cat == "Cf": return ""        # usynlige tegn: slett
        if cat == "Zs" and c != " ": return " "  # ideografiske mellomrom: normaliser
        return c
    s = "".join(_fix(c) for c in (s or ""))
    s = s.translate(_SPECIAL)
    s = unicodedata.normalize("NFD", s)
    s = "".join(c for c in s if unicodedata.category(c) != "Mn")
    s = re.sub(r"[-]", " ", s)
    s = re.sub(r"['''.·]", "", s)
    s = " ".join(s.lower().split())
    return _JR.sub("junior", s)

def match_score(a: str, b: str) -> int:
    if a == b: return 3
    pa, pb = a.split(), b.split()
    if sorted(pa) == sorted(pb): return 2
    sa, sb = set(pa), set(pb)
    if len(sa) >= 2 and len(sb) >= 2:
        diff = abs(len(sa) - len(sb))
        if diff <= 1 and (sa <= sb or sb <= sa): return 1
    if pa and pb and pa[-1] == pb[-1] and len(pa) == len(pb): return 1
    return 0

def reverse_name(s: str) -> str:
    """'Tanaka Ao' -> 'Ao Tanaka' og omvendt."""
    parts = s.split()
    if len(parts) == 2:
        return f"{parts[1]} {parts[0]}"
    return s

def last_name(s: str) -> str:
    parts = s.split()
    return parts[-1] if parts else s

# ── Matching per lag ──────────────────────────────────────────────────────────
new_clubs: dict[str, dict[str, str]] = {}
total_matched = total_unmatched = 0
unmatched_list: list[str] = []

for gruppe, lag_dict in players.items():
    for pjlag, spillere in lag_dict.items():
        clubs_key = PJSON_TO_CLUBS.get(pjlag)
        if not clubs_key or clubs_key not in clubs_raw:
            # Prøv fuzzy-søk
            clubs_key = next(
                (k for k in clubs_raw if norm(pjlag) in norm(k) or norm(k) in norm(pjlag)),
                None
            )
        if not clubs_key:
            new_clubs[pjlag] = {s["name"]: "" for s in spillere}
            total_unmatched += len(spillere)
            unmatched_list.append(f"LAG IKKE FUNNET: {pjlag}")
            continue

        raw_tropp = clubs_raw[clubs_key]  # {råname: klubb}

        # Bygg normaliserte oppslag — direkte, reversert, etternavn-only
        cn_entries: list[tuple[str, str, str]] = []  # (norm_versjon, orig_name, klubb)
        for rn, klubb in raw_tropp.items():
            rn_clean = "".join(c for c in rn
                               if unicodedata.category(c) not in ("Cf","Zs") or c == " ").strip()
            n = norm(rn_clean)
            cn_entries.append((n, rn_clean, klubb))
            # Reversert (Japan, Australia surname-first)
            rev = norm(reverse_name(rn_clean))
            if rev != n:
                cn_entries.append((rev, rn_clean, klubb))
            # Etternavn-only (Australia)
            ln = last_name(n)
            if ln != n and len(ln) > 3:
                cn_entries.append((ln, rn_clean, klubb))

        new_clubs[pjlag] = {}
        for s in spillere:
            pjname = s["name"]
            pn = norm(pjname)
            pn_rev = norm(reverse_name(pjname))
            pn_ln  = last_name(pn)

            best_sc, best_klubb = 0, ""
            for cn, _, klubb in cn_entries:
                for candidate in (pn, pn_rev):
                    sc = match_score(candidate, cn)
                    # Etternavn-only matcher gir maks score 1
                    if sc == 0 and len(pn_ln) >= 3:
                        sc_ln = match_score(pn_ln, cn)
                        if sc_ln > 0:
                            sc = 1
                    if sc > best_sc:
                        best_sc, best_klubb = sc, klubb

            new_clubs[pjlag][pjname] = best_klubb
            if best_sc == 0:
                total_unmatched += 1
                unmatched_list.append(f"  UMATCHED: {pjlag} | {pjname}")
            else:
                total_matched += 1

# Skriv ny clubs_new.json
with open(r"C:\Claude_dev\FotballVMClaude\clubs_new.json", "w", encoding="utf-8") as f:
    json.dump(new_clubs, f, ensure_ascii=False, indent=2)

print(f"Matched:   {total_matched}")
print(f"Umatched:  {total_unmatched}")
print()
if unmatched_list:
    print("Gjenstående uten klubb:")
    for line in unmatched_list:
        print(line)
else:
    print("Alle spillere matchet!")
