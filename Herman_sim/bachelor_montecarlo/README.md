# Monte Carlo-simulering – multimodal transport Karmøy → Beograd

## Hva er dette?

Python-kode tilhørende bacheloroppgaven *"Analyse av multimodale 
transportruter under operasjonell usikkerhet"* (TLOG3030, NTNU vår 2026)
av Martin Olsøy Bråten og Herman Ose.

Skriptene leser inputene fra Excel-modellen (`transportmodell.xlsx`),
kjører N=10 000 simuleringer per rute, og produserer figurene og 
tabellene i kapittel 7 og 8 av oppgaven.

GC-modellen følger Hanssen et al. (2012) og Janić (2007), utvidet med 
pålitelighetspremie i tråd med Andersson et al. (2017):
GC = C_direkte + C_terminal + α · E[t] + β · σ(t)

der α = VFTTS = 330 kr/t, β = RR · α = 264 kr/t (Halse et al. 2019).

## Innhold

| Fil | Beskrivelse | Brukes i |
|-----|-------------|----------|
| `transportmodell.xlsx` | Excel-modell med rute-data og stokastiske parametere | Alle skript |
| `monte_carlo.py` | Hovedsimulering: forventet GC, ledetidsfordelinger | Kapittel 7 |
| `scenarioanalyse.py` | Block-train-scenario for R1 og R3 | Kapittel 8.3 |
| `tornado_analyse.py` | Parametrisk sensitivitetsanalyse (±20 %) | Kapittel 8.1 |
| `requirements.txt` | Python-pakker som kreves | – |

## Komme i gang

### 1. Sett opp miljø (én gang)

Åpne mappen i terminal eller VS Code:

```bash
python -m venv venv
# Mac/Linux:
source venv/bin/activate
# Windows:
venv\Scripts\activate

pip install -r requirements.txt
```

### 2. Kjør simuleringen

```bash
python monte_carlo.py            # hovedanalyse (kapittel 7)
python scenarioanalyse.py        # block-train (kapittel 8.3)
python tornado_analyse.py        # sensitivitet (kapittel 8.1)
```

### 3. Tilpass parametere

Åpne `transportmodell.xlsx`, gå til arket `Stokastiske_parametere`, 
endre verdier i **gule celler** (CV, triangulær min/mode/max, headway, 
disrupsjons-P). Lagre og kjør på nytt.

## Kommandolinje-flagg (monte_carlo.py)

```bash
python monte_carlo.py --n 50000           # flere iterasjoner
python monte_carlo.py --seed 7            # annen tilfeldighet
python monte_carlo.py --no-disrupt        # uten disrupsjons-haler
python monte_carlo.py --excel min.xlsx    # annen Excel-fil
python monte_carlo.py --out ./figurer     # lagre figurer i annen mappe
```

## Hva Excel-arket inneholder

| Ark | Hva | Endre? |
|-----|-----|--------|
| Ruter | Segmenter pr rute (deterministiske mean-verdier) | Bare hvis ny rute |
| Kostnadsmodeller | Original GC-tabell | Nei (referanse) |
| Charts | Original sammendrag | Nei (referanse) |
| **Stokastiske_parametere** | **CV, triangulær, headway, | **Ja, gule celler** |
| **Kostnadsparametere_MC** | **VFTTS, β/α, kostnadssatser** | **Ja, gule celler** |

## Sentrale parametere

CV-verdiene representerer plausible størrelsesordener basert på 
litteratur og er ikke primærdata for de spesifikke korridorene — se 
diskusjon i kapittel 6.7 i oppgaven.

| Parameter | Verdi | Karakteristikk |
|-----------|-------|----------------|
| Sjø RoRo CV | 0,12 | Short-sea med faste avganger, lav variasjon |
| Sjø havgående CV | 0,15 | Lengre strekning, mer væravhengig |
| Jernbane CV | 0,20 | Internasjonal intermodal, høyere variasjon |
| Lastebil CV | 0,10 | Lavest variasjon, fleksibelt rutevalg |
| Terminaltid jernbane | tri(3, 6, 12) t | Wiegmans & Behdani (2018) |
| Headway sjø | 68 t | Sea-Cargo AS (2025), 2–3 avganger/uke |
| Headway jernbane | varierer per segment | Operatørenes rutetabeller |
| Grense Norge–EU | tri(0,25, 0,5, 1) t | EØS-avtalen, forenklet toll |
| Grense EU–Serbia | tri(2, 3, 5) t | EU Delegation to Serbia (2019) |
| VFTTS α | 330 kr/t | Halse et al. (2019): 13,6 kr/tonn-time × 24 tonn |
| RR | 0,8 | Halse et al. (2019) |

## Tolkning av outputs

- **E[t], σ(t)**: Forventet ledetid og standardavvik. σ(t) er hjertet 
  i pålitelighetsanalysen.
- **P10/P50/P90**: 10 %, 50 % (median), 90 % persentiler. P90 er 
  "verste-90 %"-tid.
- **Pålitelighetspremie β·σ(t)**: Hvor mye den ekstra usikkerheten 
  koster i kr (Andersson et al. 2017).
- **GC m/påli.**: Total GC inkludert pålitelighetspremie. Dette er 
  det beslutningstakeren bør sammenligne på.

## Om

Bacheloroppgave TLOG3030, NTNU vår 2026.
Forfattere: Martin Olsøy Bråten og Herman Ose.
Veileder: Steffen Jaap Skotvoll Bakker.