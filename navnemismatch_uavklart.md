# clubs_new.json — status og endringslogg

Sist oppdatert: 2026-06-18. Alle 1248 spillere har nå klubb (100%).

---

## Navnemismatch-fix (2026-06-18)

`fix_clubs_names.py` reparerte den opprinnelige clubs_new.json som hadde utbredte navnemismatcher:

| Problem | Eksempel | Løsning |
|---|---|---|
| Australia: kun etternavn | "Ryan" → "Mathew Ryan" | Etternavn-only matching |
| Japan: etternavn-fornavn | "Tanaka Ao" → "Ao Tanaka" | Reversert to-ords navn |
| Norske lagnavn | "Sør-Korea" vs "Republic of Korea" | PJSON_TO_CLUBS-mapping |
| Ideografiske mellomrom | U+3000 i japanske navn | Zs-kategori normalisert til space |
| Sør-Korea: 19 av 26 manglet | — | Manuelt lagt inn fra kjent data |

Resultat etter fix + manuell patchrunde: **1248/1248 (100%)** har klubb.

---

## Klubbnavn-standardisering (2026-06-18)

`standardiser_klubber.py` normaliserte 142 oppføringer i clubs_new.json → 471 unike klubber.

Viktigste regler:
- Al-klubber: konsekvent bindestrek (`Al Hilal` → `Al-Hilal` osv.)
- Tyrkiske klubber: juridisk suffiks fjernet (`Galatasaray A.Ş.` → `Galatasaray`)
- Uten/med prefiks: `PSV` → `PSV Eindhoven`, `FC Barcelona` → `Barcelona`, `Ajax Amsterdam` → `Ajax`
- Franske: `Olympique de Marseille` → `Marseille`, `Olympique Lyonnais` → `Lyon`
- Tyske: `TSG Hoffenheim` → `Hoffenheim`, `VfB Stuttgart` → `Stuttgart`
- Britiske: `Brighton & Hove Albion` → `Brighton`, `Wolverhampton Wanderers` → `Wolverhampton`

Nye oppføringer lagt i `club_country_map.py`: Al-Arabi, Al-Gharafa, Al-Shamal, Al Fateh,
Angers, Beşiktaş, Boavista, Bochum, Borac Banja Luka, Colorado Springs Switchbacks,
El Paso Locomotive, Fagiano Okayama, Gimcheon Sangmu, Mafra, Seongnam, Sturm Graz,
Suwon, Ulsan HD, Vissel Kobe, Austin FC.

---

## Kort-ark rangering (2026-06-18)

`check_avvik.py → skriv_kort_sheet()`: rangeres nå som Toppscorere/Assists:
- Primær: røde kort (rode + gule_rode) synkende
- Sekundær: gule kort synkende
- Lik rang når begge verdier er identiske
- Gull/sølv/bronse-bakgrunn for topp 3 ranger

---

## Fødselsdato-aliaser (lagt inn i fetch_alle_fodselsdatoer.py)

| players.json | FIFA-API | Merknad |
|---|---|---|
| Nabil Emad | Nabil Donga | Egypt — Donga er kallenavn |
| Hossein Kanaanizadegan | Hossein Kanaani | Iran — ShortName i FIFA |
| Dennis Eckert | Dennis Dargahi | Iran — iransk statsborgerskap-navn |
| Roberto Lopes | Pico Lopes | Kapp Verde — Pico er kallenavn |
| Firas Al-Buraikan | Feras Albrikan | Saudi-Arabia — annen romanisering |
| Hassan Kadesh | Hassan Kadish | Saudi-Arabia — annen romanisering |
| Alaa Al-Hejji | Ala Alhajji | Saudi-Arabia — annen romanisering |
| Mohammad Taha | Mohammad Abughoush | Jordan — to offisielle navn |

---

## Tropp-endringer (bekreftet via FIFA-API trøyenummer)

| Lag | Nr | Ute | Inn | Fødselsdato ny |
|---|---|---|---|---|
| Tunisia | 25 | Nizar Issaoui | Anis Slimane | 2001-03-16 |
| England | 12 | Tino Livramento | Trevoh Chalobah | 1999-07-05 |
| New Zealand | 7 | Matthew Garbett | Logan Rogerson | 1998-05-28 |
| Uzbekistan | 10 | Jaloliddin Masharipov | Ruslanbek Jiyanov | 2001-06-05 |
| Uruguay | 25 | Agustín Oliveros | Juan Manuel Sanabria | 2000-03-29 |
