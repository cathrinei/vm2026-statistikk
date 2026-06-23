#!/usr/bin/env python3
"""
standardiser_klubber.py
Normaliserer klubbnavn i clubs_new.json til én kanonisk form per klubb.
"""
import io, json, sys
sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding="utf-8", errors="replace")
from pathlib import Path
BASE_DIR = Path(__file__).parent

CANONICAL = {
    # Tyske klubber
    "Bayern Munich":                 "Bayern München",

    # Al-klubber: konsekvent bindestrek
    "Al Hilal":                     "Al-Hilal",
    "Al Nassr":                     "Al-Nassr",
    "Al Qadsiah":                   "Al-Qadsiah",
    "Al Arabi":                     "Al-Arabi",
    "Al Gharafa":                   "Al-Gharafa",
    "Al Wakrah":                    "Al-Wakrah",
    "Al Ain":                       "Al-Ain",
    "Al Talaba":                    "Al-Talaba",

    # Tyrkiske klubber: drop juridisk suffiks, bruk aksent
    "Galatasaray A.Ş.":             "Galatasaray",
    "Galatasaray S.K.":             "Galatasaray",
    "Galatasaray SK":               "Galatasaray",
    "Fenerbahçe A.Ş.":             "Fenerbahçe",
    "Fenerbahce SK":                "Fenerbahçe",
    "Beşiktaş A.Ş.":               "Beşiktaş",
    "Beşiktaş J.K.":               "Beşiktaş",
    "Besiktas":                     "Beşiktaş",
    "Trabzonspor Kulübü":           "Trabzonspor",
    "Çaykur Rizespor A.Ş.":        "Çaykur Rizespor",
    "İstanbul Başakşehir":          "Istanbul Basaksehir",

    # Nederlandske
    "PSV":                          "PSV Eindhoven",
    "Ajax Amsterdam":               "Ajax",
    "NEC Nijmegen":                 "Nijmegen",
    "KRC Genk":                     "Genk",

    # Spanske
    "Athletic Bilbao":              "Athletic Club",
    "Real Ovideo":                  "Real Oviedo",
    "FC Barcelona":                 "Barcelona",
    "FC Sevilla":                   "Sevilla",
    "Castellon":                    "Castellón",
    "Villarreal CF":                "Villarreal",

    # Tyske
    "TSG Hoffenheim":               "Hoffenheim",
    "TSG 1899 Hoffenheim":          "Hoffenheim",
    "Mainz 05":                     "Mainz",
    "VfB Stuttgart":                "Stuttgart",
    "Borussia Monchengladbach":     "Borussia Mönchengladbach",
    "Hamburger SV":                 "Hamburg",
    "Eintracht Frankfurt":          "Frankfurt",
    "FC ⁠Augsburg":            "Augsburg",   # usynlig tegn i opprinnelig

    # Franske
    "Olympique Lyonnais":           "Lyon",
    "Olympique Marseille":          "Marseille",
    "Olympique de Marseille":       "Marseille",
    "OGC Nice":                     "Nice",
    "AJ Auxerre":                   "Auxerre",
    "Angers SCO":                   "Angers",
    "Le Havre AC":                  "Le Havre",
    "RC Strasbourg":                "Strasbourg",
    "Saint-Etienne":                "Saint-Étienne",
    "Royale Union Saint-Gilloise":  "Union Saint-Gilloise",
    "Saint-Gilloise":               "Union Saint-Gilloise",
    "Royal Charleroi SC":           "Royal Charleroi",

    # Engelske
    "AFC Bournemouth":              "Bournemouth",
    "Brighton & Hove Albion":       "Brighton",
    "Newcastle United":             "Newcastle",
    "Nottingham Forest F.C.":       "Nottingham Forest",
    "Sunderland AFC":               "Sunderland",
    "West Ham United":              "West Ham",
    "Wolverhampton Wanderers":      "Wolverhampton",
    "Wrexham AFC":                  "Wrexham",

    # Portugisiske
    "FC Porto":                     "Porto",
    "SL Benfica":                   "Benfica",
    "SC Braga":                     "Braga",

    # Italienske
    "AS Monaco":                    "Monaco",
    "AS Roma":                      "Roma",
    "Atalanta BC":                  "Atalanta",
    "Sassuolo Calcio":              "Sassuolo",
    "US Sassuolo":                  "Sassuolo",

    # Tsjekkiske
    "FC Viktoria Plzen":            "Viktoria Plzeň",
    "Viktoria Plzen":               "Viktoria Plzeň",

    # Andre europeiske
    "Esteghlal FC":                 "Esteghlal",
    "Esperance":                    "Espérance de Tunis",
    "Panathinaikos F.C.":           "Panathinaikos",
    "Viking FK":                    "Viking",
    "FC Lugano":                    "Lugano",
    "FC Nordsjaelland":             "Nordsjaelland",
    "FC Copenhagen":                "Copenhagen",
    "Pafos FC":                     "Pafos",
    "AE Kifisias":                  "AE Kifisia",
    "HNK Rijeka":                   "Rijeka",
    "BSC Young Boys II":            "BSC Young Boys",

    # MLS / USA
    "Minnesota United FC":          "Minnesota United",
    "New York City FC":             "New York City",
    "Toronto FC":                   "Toronto",
    "FC Dallas":                    "Dallas",
    "San Diego FC":                 "San Diego",
    "Colorado Springs":             "Colorado Springs Switchbacks",

    # Sør-Amerika
    "UNAM Pumas":                   "Pumas UNAM",
}

with open(BASE_DIR / "clubs_new.json", encoding="utf-8") as f:
    c = json.load(f)

endringer = 0
for lag in c:
    for navn in c[lag]:
        gammel = c[lag][navn]
        ny = CANONICAL.get(gammel, gammel)
        if ny != gammel:
            c[lag][navn] = ny
            endringer += 1

with open(BASE_DIR / "clubs_new.json", "w", encoding="utf-8") as f:
    json.dump(c, f, ensure_ascii=False, indent=2)

print(f"Endringer: {endringer}")

# Tell unike klubbnavn etter
alle = sorted({k for lag in c.values() for k in lag.values() if k})
print(f"Unike klubber etter: {len(alle)}")
