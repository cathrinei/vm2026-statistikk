# VM2026_avansert_gruppetabeller_og_sluttspill.xlsx — Innholdsoversikt

**Fil:** `VM2026_avansert_gruppetabeller_og_sluttspill.xlsx`  
**Sist oppdatert:** 2026-06-20  
**Antall ark:** 30

---

## Ark 1–12: Gruppe A–L

Hvert gruppeark inneholder tre seksjoner:

### Seksjon 1: Kampetabell (rad 1–8)
Viser alle 3 kamper i gruppen (inkl. fremtidige).

| Kolonne | Innhold |
|---------|---------|
| A | Runde |
| B | Dato |
| C | Hjemmelag |
| D | Resultat |
| E | Bortelag |
| F | Tilskuere |
| G | Stadion |
| H | Dommer |

Spilte kamper har mintgrønn bakgrunn. Resultatcellen viser målene i grønn bold.

### Seksjon 2: Gruppetabell (rad 10–15)
Automatisk beregnet fra kampresultatene.

| Kolonne | Innhold |
|---------|---------|
| A | Lag |
| B | Kamper spilt (K) |
| C | Vunnet (V) |
| D | Uavgjort (U) |
| E | Tapt (T) |
| F | Mål for (MF) |
| G | Mål mot (MM) |
| H | Målforskjell (MD) |
| I | Poeng (P) |

Sortert: poeng ↓ → målforskjell ↓ → mål for ↓

### Seksjon 3: Spillerliste (rad 1–115, kolonne L–U)
Fire lag per ark i fire blokker: rader 1–28, 30–57, 59–86, 88–115.
Hver blokk har en lagnavn-header, kolonneoverskrifter og 26 spillerrader.

| Kolonne | Innhold |
|---------|---------|
| L (12) | Nummer |
| M (13) | Navn |
| N (14) | Klubb |
| O (15) | Fødselsdato |
| P (16) | Posisjon |
| Q (17) | Mål |
| R (18) | Assist |
| S (19) | Gule kort |
| T (20) | Røde kort |
| U (21) | Minutter spilt |

Lagnavn-header: L–O merget (navy bold), P–U merget (trener, navy kursiv høyrestilt).  
Rader med mål eller assist har mintgrønn bakgrunn.

---

## Ark 13–18: Sluttspill

Seks ark for sluttspillrundene, opprettet av `add_sluttspill_sheet.py`:

| Ark | Innhold |
|-----|---------|
| 13 | 32-delsfinale (16 kamper) |
| 14 | 8-delsfinale (8 kamper) |
| 15 | Kvartfinaler (4 kamper) |
| 16 | Semifinaler (2 kamper) |
| 17 | Bronsefinale (1 kamp) |
| 18 | Finale (1 kamp) |

Kamper med ukjent lag vises som "TBA". Spilte kamper får mintgrønn bakgrunn.  
Kolonner: Dato, Hjemmelag, Resultat, Bortelag, Tilskuere, Stadion, Dommer.  
Cache: `sluttspill_cache.json`

---

## Ark 19: Toppscorere
Rangert etter antall mål. Topp 3 med gull/sølv/bronse.  
7 kolonner.

---

## Ark 14: Assists
Rangert etter antall målgivende pasninger. Topp 3 med gull/sølv/bronse.  
7 kolonner.

---

## Ark 15: Kort
Spillere med gule og røde kort, rangert røde øverst.  
Topp 3 med gull/sølv/bronse. 6 kolonner.

---

## Ark 16: Lagstatistikk
Statistikk per lag i 7 seksjoner (248 rader, 8 kolonner):

1. **Mål per kamp** — totalt + fordeling (spillermål / straffer / selvmål)
2. **Rene nullere** — antall kamper uten baklengsmål
3. **Straffespark** — tilkjent og scoret per lag
4. **Selvmål** — hvilke spillere som har scoret selvmål
5. **Skudd** — TotalAttempts og AttemptsOnTarget per lag med snitt og treff%
6. **Formasjoner** — primærformasjon og varianter per lag
7. **Spillerbytter** — antall bytter per kamp, snitt-minutt og tidligste bytte

Kilde: FIFA live-endepunkt + `topseasonplayerstatistics`-API. Cachet i `lagstatistikk_cache.json` og `live_cache.json`.

---

## Ark 17: Ballbesittelse
Kampstatistikk for alle spilte kamper — besittelse og skudd.  
14 kolonner, opptil 29 kamprader + forklaring.

Layout speilet rundt Score-kolonnen (H):

| Kolonne | Hjemmelag | — | Bortelag |
|---------|-----------|---|---------|
| A | Gruppe | | |
| B | Hjemme | N | Borte |
| C | Besittelse% | M | Besittelse% |
| D | Skudd totalt | L | Skudd totalt |
| E | Skudd på mål | K | Skudd på mål |
| F | Skudd utenfor | J | Skudd utenfor |
| G | Blokkert | I | Blokkert |
| H | **Score** | | |

Besittelse fargekodet: navy (≥65%) → lys blå (≤44%).  
Grønn bold = høyest verdi av de to lagene.

Kilde: `wc2026_stats.json` (oppdateres manuelt via claude.ai chat).

---

## Ark 18: Scoringstidspunkt
Mål fordelt på 6 intervaller à 15 minutter.  
Inkluderer søylediagram. 3 kolonner, 9 rader + diagram.

| Intervall | Mål | Andel |
|-----------|-----|-------|
| 1–15 | — | — |
| 16–30 | — | — |
| 31–45 | — | — |
| 46–60 | — | — |
| 61–75 | — | — |
| 76–90+ | — | — |

Kilde: `lagstatistikk_cache.json`.

---

## Ark 19: Heatmap
Mål per minutt (1–90 + tilleggstid 90+1..+10) som fargekodet rutenett.  
12 kolonner, 19 rader. Tykk strek markerer halvtidspausen (mellom min 45 og 46).

**Fargeskala:** hvit (0 mål) → navy (6+ mål)

| Farge | Betydning |
|-------|-----------|
| Hvit | 0 mål |
| Svært lys blå | 1 mål |
| Lys blå | 2 mål |
| Mellomblå | 3 mål |
| Sterk blå | 4 mål |
| Mørk blå | 5 mål |
| Navy | 6+ mål |

Kilde: `lagstatistikk_cache.json`.

---

## Ark 20: Spillere etter klubbland
Spillere gruppert etter klubbland med fordeling per liganiviå.  
9 kolonner.

---

## Ark 21: Klubber
Alle VM-klubber rangert etter antall VM-spillere.  
457 unike klubber.

| Kolonne | Innhold |
|---------|---------|
| A | Rang |
| B | Klubb |
| C | Land |
| D | Antall spillere |
| E | Nasjoner representert |

---

## Ark 22: Nivå 2 og lavere
Spillere som spiller i ikke-toppserien i sitt hjemland.  
Genereres av `lag_nivå_tabell.py` (kjøres automatisk ved endringer i `top_division.py`/`level_map.py`).  
6 kolonner.

---

## Ark 23: Aldersfordeling
Snittalder, yngste og eldste spiller per lag (alle 48 lag).  
9 kolonner, 50 rader.

---

## Ark 24: Alder
10 yngste og 10 eldste spillere i hele turneringen.  
6 kolonner, 27 rader. Topp 3 med gull/sølv/bronse.

---

## Oppdatering

```
python oppdater.py                   # normal daglig kjøring
python oppdater.py --full            # tving full gjeninnlesing av alle kamper
python excel_til_html.py             # regenerer docs/index.html fra docs/template.html
python generer_offensiv_analyse.py   # regenerer docs/offensiv_analyse.html fra Lagstatistikk-arket
```

Ballbesittelse-arket oppdateres ikke automatisk — `wc2026_stats.json` må hentes manuelt.

---

## HTML-utdata

### docs/index.html
Genereres av `excel_til_html.py` fra malen `docs/template.html`.  
Viser alle ark som navigerbare faner med søk.  
Inneholder lenke til `offensiv_analyse.html` øverst til høyre i headeren.

### docs/offensiv_analyse.html
Genereres av `generer_offensiv_analyse.py` direkte fra **Lagstatistikk**-arket.  
Viser offensiv analyse med fire metrikk-kort og et interaktivt boblediagram.

**Boblediagrammet** plotter skudd/kamp (x) mot mål/kamp (y) med boblestørrelse = treff%.  
Utvalg: alltid de 12 beste scorende lagene + de 3 med flest skudd og 0 mål — slik at
diagrammet forblir lesbart gjennom hele turneringen selv om mange lag akkumulerer mål.

| Farge | Betydning |
|-------|-----------|
| Blå | 2+ kamper spilt, har scoret |
| Grønn | 1 kamp spilt, har scoret |
| Rød | 0 mål — flest skudd (paradoks-lag) |
| Grå | 0 mål — øvrige |

Kjøresekvens i `daglig_oppdatering.yml`:  
`oppdater.py` → `excel_til_html.py` → `generer_offensiv_analyse.py` → commit + push
