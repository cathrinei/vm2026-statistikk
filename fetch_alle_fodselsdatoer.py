#!/usr/bin/env python3
"""
fetch_alle_fodselsdatoer.py
Henter fødselsdato for alle 1248 VM-spillere i én kjøring (48 API-kall).

Steg:
  1. Hent team-IDer fra kampkalenderen
  2. Hent tropp per lag via /teams/{id}/squad (BirthDate er innebygget i responsen)
  3. Match FIFA-spillernavn mot players.json per lag
  4. Oppdater players.json og player_alder.json
  5. Kjør add_birthdate_column.py for å fylle inn i Excel

Trygt å kjøre flere ganger — eksisterende datoer overskrives ikke.
"""
import io, json, re, subprocess, sys, time, unicodedata
from pathlib import Path

sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding="utf-8", errors="replace")

try:
    import requests
except ImportError:
    sys.exit("Mangler requests-pakke")

_FIFA_BASE   = "https://api.fifa.com/api/v3"
_FIFA_COMP   = "17"
_FIFA_SEASON = "285023"
_HEADERS     = {
    "User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36",
    "Accept": "application/json",
}

BASE_DIR    = Path(r"C:\Claude_dev\FotballVMClaude")
ALDER_CACHE = BASE_DIR / "player_alder.json"
PLAYERS_JSON = BASE_DIR / "players.json"
PYTHON      = sys.executable

# Engelske FIFA-navn → norske (brukes bare for visning, ikke som nøkler i oppslag)
NORSK = {
    "Mexico": "Mexico", "South Africa": "Sør-Afrika", "South Korea": "Sør-Korea",
    "Korea Republic": "Sør-Korea", "Czechia": "Tsjekkia", "Czech Republic": "Tsjekkia",
    "Canada": "Canada", "Bosnia and Herzegovina": "Bosnia-Hercegovina",
    "Qatar": "Qatar", "Switzerland": "Sveits",
    "Brazil": "Brasil", "Morocco": "Marokko", "Haiti": "Haiti", "Scotland": "Skottland",
    "United States": "USA", "USA": "USA", "Paraguay": "Paraguay",
    "Australia": "Australia", "Turkey": "Tyrkia", "Türkiye": "Tyrkia",
    "Germany": "Tyskland", "Curacao": "Curaçao", "Curaçao": "Curaçao",
    "Ivory Coast": "Elfenbenskysten", "Côte d'Ivoire": "Elfenbenskysten",
    "Ecuador": "Ecuador", "Sweden": "Sverige", "Japan": "Japan",
    "New Zealand": "New Zealand", "Tunisia": "Tunisia",
    "Spain": "Spania", "Cape Verde": "Kapp Verde", "Cabo Verde": "Kapp Verde",
    "Saudi Arabia": "Saudi-Arabia", "Uruguay": "Uruguay",
    "Netherlands": "Nederland", "Belgium": "Belgia", "Egypt": "Egypt", "Jordan": "Jordan",
    "Norway": "Norge", "France": "Frankrike", "Senegal": "Senegal", "Colombia": "Colombia",
    "Argentina": "Argentina", "Panama": "Panama", "England": "England", "Algeria": "Algerie",
    "Croatia": "Kroatia", "Iran": "Iran", "IR Iran": "Iran",
    "Iraq": "Irak", "Portugal": "Portugal",
    "Ghana": "Ghana", "DR Congo": "DR Kongo", "Congo DR": "DR Kongo",
    "Uzbekistan": "Usbekistan", "Austria": "Østerrike",
}

# ── Navnenormalisering ────────────────────────────────────────────────────────

_SPECIAL = str.maketrans({
    "ø": "o", "Ø": "o", "æ": "ae", "Æ": "ae", "å": "a", "Å": "a",
    "ß": "ss", "ð": "d", "þ": "th", "ı": "i",
})
_SPECIAL_DE = str.maketrans({
    "ü": "ue", "Ü": "ue", "ö": "oe", "Ö": "oe", "ä": "ae", "Ä": "ae",
    "ø": "o",  "Ø": "o",  "æ": "ae", "Æ": "ae", "å": "a",  "Å": "a",
    "ß": "ss", "ð": "d",  "þ": "th", "ı": "i",
})
_JR_RE = re.compile(r"\bjr\.?$", re.IGNORECASE)


def _norm(s: str, special=_SPECIAL) -> str:
    s = (s or "").translate(special)
    s = unicodedata.normalize("NFD", s)
    s = "".join(c for c in s if unicodedata.category(c) != "Mn")
    s = re.sub(r"-", " ", s)           # bindestrek → mellomrom (Ati-Zigi → Ati Zigi)
    s = re.sub(r"[''.]", "", s)        # apostrof/punktum → fjernes
    s = " ".join(s.lower().split())
    return _JR_RE.sub("junior", s)


def _title(s: str) -> str:
    """Konverterer 'Raul RANGEL' til 'Raul Rangel'."""
    return " ".join(w.capitalize() for w in (s or "").split())


# ── Aliases: players.json-norm → FIFA-squad-norm ──────────────────────────────
# Kun for spillere der navnene er genuine ulike (ikke varianter som fanges av _norm)

def _a(excel_name: str, fifa_name: str) -> tuple[str, str]:
    return (_norm(excel_name), _norm(fifa_name))


ALIASES: dict[str, str] = dict([
    # Brazil — FIFA: "NEYMAR JR" → norm "neymar junior"
    _a("Neymar",                    "Neymar Junior"),
    # DR Congo — k/ck-variant
    _a("Meschak Elia",              "Meschack Elia"),
    # Egypt — sammenslåtte vs separerte ord
    _a("Mohamed El Shenawy",        "Mohamed Elshenawy"),
    _a("Mohamed Abdelmonem",        "Mohamed Abdelmoneim"),
    # Iraq — arabisk romanisering
    _a("Manaf Younis",              "Munaf Younus"),
    _a("Ahmed Maknzi",              "Ahmed Maknazi"),
    _a("Zaid Ismail",               "Zaid Ismael"),
    # Jordan — arabiske navn har ulik romanisering mellom kildene
    _a("Mohammad Abu Hashish",      "Mohammad Abuhasheesh"),
    _a("Mohammad Abu Zrayq",        "Mohammad Abuzraiq"),
    _a("Musa Al-Taamari",           "Mousa Altamari"),
    _a("Odeh Al-Fakhouri",          "Odeh Fakhoury"),
    _a("Nour Bani Attiah",          "Nour Baniateyah"),
    _a("Mo Abualnadi",              "Mohammad Abualnadi"),
    _a("Salim Obaid",               "Saleem Obaid"),
    _a("Abdallah Al-Fakhouri",      "Abdallah Alfakhori"),
    _a("Ihsan Haddad",              "Ehsan Haddad"),
    _a("Mohannad Abu Taha",         "Mohannad Abutaha"),
    _a("Mohammad Al-Dawoud",        "Mohammad Aldaoud"),
    # Morocco — helt ulikt navn
    _a("Munir Mohamedi",            "Munir El Kajoui"),
    # Paraguay — kallenavn
    _a("Kaku",                      "Alejandro Romero Gamarra"),
    # Qatar — ulik etternavn
    _a("Tahsin Jamshid",            "Tahsin Mohammed"),
    # Tunisia — FIFA inkluderer prefiks som mangler i players.json
    _a("Hadj Mahmoud",              "Mohamed Hadj Mahmoud"),
    # Uzbekistan — X/H romanisering og forkortet fornavn
    _a("Odiljon Hamrobekov",        "Odiljon Xamrobekov"),
    _a("Avazbek Ulmasaliev",        "Avazbek Ulmasaliyev"),
    # Iran — ShortName er forkortet etternavn / iransk statsborgerskap-navn
    _a("Hossein Kanaanizadegan",    "Hossein Kanaani"),
    _a("Dennis Eckert",             "Dennis Dargahi"),
    # Cape Verde — Roberto Lopes spiller under kallenavnet Pico
    _a("Roberto Lopes",             "Pico Lopes"),
    # Egypt — Nabil Emad registrert under kallenavnet Donga i FIFA
    _a("Nabil Emad",                "Nabil Donga"),
    # Saudi Arabia — ulik romanisering i FIFA vs players.json
    _a("Firas Al-Buraikan",         "Feras Albrikan"),
    _a("Hassan Kadesh",             "Hassan Kadish"),
    _a("Alaa Al-Hejji",             "Ala Alhajji"),
    # Jordan — to offisielle navn for samme spiller
    _a("Mohammad Taha",             "Mohammad Abughoush"),
])


# FIFA-navn → players.json-navn der de avviker
_FIFA_TO_PJSON: dict[str, str] = {
    "Congo DR": "DR Congo",   # FIFA: "Congo DR", players.json: "DR Congo"
}


# ── Steg 1: Hent alle team-IDer fra kampkalenderen ───────────────────────────

def hent_team_ider() -> dict[str, str]:
    """Returnerer {engelskLagnavn: team_id} for alle 48 VM-lag.
    Bruker players.json-navn som nøkkel; faller tilbake på FIFA-navn."""
    r = requests.get(
        f"{_FIFA_BASE}/calendar/matches"
        f"?idCompetition={_FIFA_COMP}&idSeason={_FIFA_SEASON}&count=200&language=en-GB",
        headers=_HEADERS, timeout=20
    )
    r.raise_for_status()

    teams: dict[str, str] = {}
    for m in r.json().get("Results", []):
        for side in ("Home", "Away"):
            obj = m.get(side) or {}
            tid = obj.get("IdTeam")
            if not tid:
                continue
            name_en = next(
                (n["Description"] for n in obj.get("TeamName", [])
                 if n.get("Locale") == "en-GB"), ""
            )
            if not name_en:
                continue
            # Bruk players.json-navn hvis FIFA-navnet avviker
            key = _FIFA_TO_PJSON.get(name_en, name_en)
            if key not in teams:
                teams[key] = str(tid)
    return teams


# ── Steg 2: Hent tropp per lag ────────────────────────────────────────────────

def hent_tropp(team_id: str) -> list[dict]:
    """Henter spillerliste fra /teams/{id}/squad. Returnerer [{pid, name, birthdate}]."""
    r = requests.get(
        f"{_FIFA_BASE}/teams/{team_id}/squad"
        f"?idCompetition={_FIFA_COMP}&idSeason={_FIFA_SEASON}&language=en-GB",
        headers=_HEADERS, timeout=15
    )
    if r.status_code != 200:
        return []

    spillere = []
    for p in r.json().get("Players", []):
        pid = p.get("IdPlayer")
        if not pid:
            continue
        # PlayerName = lokalisert liste (fullt navn); ShortName = direkte streng (kallenavn/kortform)
        full_raw  = next(
            (n["Description"] for n in p.get("PlayerName", [])
             if n.get("Locale") == "en-GB"), ""
        )
        short_raw = p.get("ShortName") or ""
        name       = _title(full_raw or short_raw)
        short_name = _title(short_raw)
        bd         = (p.get("BirthDate") or "")[:10]
        spillere.append({"pid": str(pid), "name": name, "short_name": short_name, "birthdate": bd})
    return spillere


# ── Steg 3: Match FIFA-spillere mot players.json per lag ─────────────────────

def _match_score(a: str, b: str) -> int:
    """0=ingen match, 1=svak, 2=rekkefølge, 3=eksakt."""
    if a == b:
        return 3
    parts_a = a.split()
    parts_b = b.split()
    # Eksakt samme tokens i annen rekkefølge
    if sorted(parts_a) == sorted(parts_b):
        return 2
    # Siste token likt, samme antall tokens
    if parts_a and parts_b and parts_a[-1] == parts_b[-1] and len(parts_a) == len(parts_b):
        return 1
    # Delmengde-matching: én liste er delmengde av den andre (maks 1 token forskjell)
    set_a, set_b = set(parts_a), set(parts_b)
    if len(set_a) >= 2 and len(set_b) >= 2:
        diff = abs(len(set_a) - len(set_b))
        if diff <= 1 and (set_a <= set_b or set_b <= set_a):
            return 1
    return 0


def match_spillere(
    tropp_fifa: list[dict],
    tropp_pjson: list[dict],
) -> dict[int, str]:
    """Returnerer {pjson_index: birthdate}. Bruker greedy matching (sterkest treff først)."""
    fifa_entries = [
        {
            "pid":      p["pid"],
            "name":     p["name"],
            "bd":       p["birthdate"],
            "norm":     _norm(p["name"]),
            "norm_de":  _norm(p["name"], _SPECIAL_DE),
            "norm_sn":  _norm(p.get("short_name", "")),   # ShortName (kortform/kallenavn)
        }
        for p in tropp_fifa
    ]

    pj_entries = [
        {
            "idx":     i,
            "norm":    _norm(s.get("name", "")),
            "norm_de": _norm(s.get("name", ""), _SPECIAL_DE),
        }
        for i, s in enumerate(tropp_pjson)
    ]

    # Bygg matchingsmatrise — prøv både direkte norm og via ALIASES
    matches: list[tuple[int, int, int]] = []
    for fi, fe in enumerate(fifa_entries):
        for pi, pe in enumerate(pj_entries):
            # FIFA-navn: prøv direkte og de-normalisert
            # pj-navn: prøv direkte, de-normalisert, og alias
            pj_norms = {pe["norm"], pe["norm_de"], ALIASES.get(pe["norm"], ""), ALIASES.get(pe["norm_de"], "")}
            pj_norms.discard("")

            score = 0
            for pjn in pj_norms:
                for fn in (fe["norm"], fe["norm_de"], fe["norm_sn"]):
                    if fn:
                        score = max(score, _match_score(fn, pjn))
            if score > 0:
                matches.append((score, pi, fi))

    matches.sort(reverse=True)

    used_pj   = set()
    used_fifa = set()
    result: dict[int, str] = {}

    for score, pi, fi in matches:
        if pi in used_pj or fi in used_fifa:
            continue
        bd = fifa_entries[fi]["bd"]
        if bd:
            result[pi] = bd
        used_pj.add(pi)
        used_fifa.add(fi)

    return result


# ── Steg 4: Oppdater players.json og alder-cache ─────────────────────────────

def oppdater_players_json(
    team_data: dict[str, list[dict]],
    alder_cache: dict,
) -> tuple[int, int]:
    """Fyller inn manglende birthdates. Returnerer (nye, fremdeles_mangler)."""
    with open(PLAYERS_JSON, encoding="utf-8") as f:
        data = json.load(f)

    # Bygg en reverse-lookup fra norsk → engelsk for lag som er lagret på norsk i players.json
    # (f.eks. "Sør-Korea" → "Korea Republic", "Tsjekkia" → "Czechia")
    norsk_to_en: dict[str, str] = {}
    for en, no in NORSK.items():
        if no not in norsk_to_en:
            norsk_to_en[no] = en

    nye = fremdeles_mangler = allerede_hadde = total_spillere = 0

    for gruppe, lag in data.items():
        for lagnavn, spillere in lag.items():
            total_spillere += len(spillere)
            # Prøv direkte nøkkel, deretter engelsk variant av norsk navn
            fifa_tropp = team_data.get(lagnavn) or team_data.get(norsk_to_en.get(lagnavn, ""), [])
            if not fifa_tropp:
                for s in spillere:
                    if s.get("birthdate"):
                        allerede_hadde += 1
                    else:
                        fremdeles_mangler += 1
                continue

            matched = match_spillere(fifa_tropp, spillere)

            for i, s in enumerate(spillere):
                if s.get("birthdate"):
                    allerede_hadde += 1
                    continue
                if i in matched:
                    s["birthdate"] = matched[i]
                    nye += 1
                else:
                    fremdeles_mangler += 1

    with open(PLAYERS_JSON, "w", encoding="utf-8") as f:
        json.dump(data, f, ensure_ascii=False, indent=2)

    # Oppdater alder-cache med alle nye PIDs (add_alder_sheet.py bruker den videre)
    for tropp in team_data.values():
        for p in tropp:
            pid = p["pid"]
            if pid not in alder_cache and p["birthdate"]:
                alder_cache[pid] = {
                    "name":      p["name"],
                    "birthdate": p["birthdate"] + "T00:00:00Z",
                }

    with open(ALDER_CACHE, "w", encoding="utf-8") as f:
        json.dump(alder_cache, f, ensure_ascii=False, indent=2)

    return nye, fremdeles_mangler, total_spillere


# ── Steg 5: Logg gjenværende spillere uten dato ──────────────────────────────

def logg_umatched() -> None:
    with open(PLAYERS_JSON, encoding="utf-8") as f:
        data = json.load(f)

    print(f"\n  {'Lag':<25} {'Navn'}")
    print(f"  {'-'*25} {'-'*35}")
    for gruppe, lag in data.items():
        for lagnavn, spillere in lag.items():
            for s in spillere:
                if not s.get("birthdate"):
                    print(f"  {lagnavn:<25} {s['name']}")


# ── Hovedprogram ──────────────────────────────────────────────────────────────

def main():
    print("=" * 62)
    print("  VM 2026 — Innhenting av alle fødselsdatoer")
    print("=" * 62)

    # Les eksisterende alder-cache
    alder_cache: dict = {}
    if ALDER_CACHE.exists():
        with open(ALDER_CACHE, encoding="utf-8") as f:
            alder_cache = json.load(f)
    print(f"\n  {len(alder_cache)} spillere allerede i alder-cache")

    # Steg 1: Team-IDer (engelske navn som nøkler)
    print("\n[1/4] Henter team-IDer fra kampkalenderen...")
    team_ider = hent_team_ider()
    print(f"  {len(team_ider)} lag funnet")

    # Steg 2: Tropper (48 API-kall)
    print(f"\n[2/4] Henter tropper fra FIFA API ({len(team_ider)} kall)...")
    team_data: dict[str, list[dict]] = {}
    for i, (lagnavn_en, team_id) in enumerate(sorted(team_ider.items()), 1):
        tropp = hent_tropp(team_id)
        team_data[lagnavn_en] = tropp
        display = NORSK.get(lagnavn_en, lagnavn_en)
        status  = f"{len(tropp)} sp." if tropp else "FEIL"
        print(f"  [{i:2d}/{len(team_ider)}] {display:<22} {status}")
        time.sleep(0.15)

    total_fifa = sum(len(t) for t in team_data.values())
    print(f"\n  Totalt {total_fifa} spillere hentet fra FIFA API")

    # Steg 3: Oppdater players.json + alder-cache
    FORVENTET = 1248
    print("\n[3/4] Matcher mot players.json og oppdaterer...")
    nye, mangler, totalt = oppdater_players_json(team_data, alder_cache)
    sjekk = "✓" if totalt == FORVENTET else f"ADVARSEL: forventet {FORVENTET}"
    print(f"  → {nye} nye fødselsdatoer lagt inn")
    print(f"  → {mangler} spillere fremdeles uten dato")
    print(f"  → {totalt} spillere totalt i players.json ({sjekk})")

    if mangler > 0:
        logg_umatched()

    # Steg 4: Oppdater Excel
    print("\n[4/4] Oppdaterer Excel-filen...")
    r = subprocess.run([PYTHON, str(BASE_DIR / "add_birthdate_column.py")],
                       capture_output=False, text=False)
    if r.returncode != 0:
        print("  ADVARSEL: add_birthdate_column.py feilet")

    print(f"\n{'='*62}")
    print(f"  Ferdig.  {nye} nye datoer lagt inn  |  {mangler} gjenstår uten dato  |  {totalt}/{FORVENTET} spillere")
    print("=" * 62)


if __name__ == "__main__":
    main()
