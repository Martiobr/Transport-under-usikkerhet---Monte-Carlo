# Monte Carlo-simulering – multimodal transport Karmøy → Beograd

## Hva er dette?

Et Python-skript som leser inputene fra Excel-modellen (`model_mc_ready.xlsx`),
kjører N=10 000 simuleringer per rute, og produserer:

- En tabell med E[t], σ(t), P10/P50/P90 for ledetid og GC
- 4 figurer: ledetid-fordeling, GC-fordeling, ledetid-boxplot, og GC risk-return
- En CSV-fil (`mc_results_summary.csv`) med oppsummering

GC-modellen følger Hanssen et al. (2012) og Janić (2007):

```
GC = C_direkte + C_terminal + α · E[t] + β · σ(t)
```

der α = VFTTS = 330 kr/t, β = RR · α = 264 kr/t (Halse et al. 2019).

## Komme i gang i VS Code

### 1. Sett opp miljø (én gang)

Åpne mappen i VS Code, åpne integrert terminal, og kjør:

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
python monte_carlo.py
```

### 3. Tilpass parametere

Åpne `model_mc_ready.xlsx`, gå til arket `Stokastiske_parametere`, endre verdier
i **gule celler** (CV, triangulær min/mode/max, headway, disrupsjons-P).
Lagre. Kjør Python-skriptet på nytt.

## Kommandolinje-flagg

```bash
python monte_carlo.py --n 50000           # flere iterasjoner
python monte_carlo.py --seed 7            # annen tilfeldighet
python monte_carlo.py --no-disrupt        # uten disrupsjons-haler (sammenlign)
python monte_carlo.py --excel min.xlsx    # annen Excel-fil
python monte_carlo.py --out ./figurer     # lagre figurer i annen mappe
```

## Filstruktur

```
monte_carlo_transport/
├── monte_carlo.py            ← hovedskriptet
├── model_mc_ready.xlsx       ← input (Excel-modellen)
├── requirements.txt
├── README.md
└── (etter kjøring:)
    ├── mc_results_summary.csv
    ├── fig_ledetid_fordeling.png
    ├── fig_gc_fordeling.png
    ├── fig_ledetid_boxplot.png
    └── fig_gc_risk_return.png
```

## Hva Excel-arket inneholder

| Ark | Hva | Endre? |
|-----|-----|--------|
| Ruter | Segmenter pr rute (deterministiske mean-verdier) | Bare hvis ny rute |
| Kostnadsmodeller | Original GC-tabell | Nei (referanse) |
| Charts | Original sammendrag | Nei (referanse) |
| **Stokastiske_parametere** | **CV, triangulær, headway, disrupsjon** | **Ja, gule celler** |
| **Kostnadsparametere_MC** | **VFTTS, β/α, kostnadssatser** | **Ja, gule celler** |

## Hvor parametrene kommer fra

| Parameter | Verdi | Kilde |
|-----------|-------|-------|
| Sjø CV | 0,12 | Andersson et al. (2017); bransjenorm RoRo |
| Jernbane CV | 0,20 | Demiridis & Pyrgidis (2022); Demir et al. (2016) |
| Lastebil CV | 0,08 | de Jong (2014); Janić (2007) |
| Terminaltid jernbane | tri(3, 6, 12) t | Wiegmans & Behdani (2018); Mandouri et al. (2023) |
| Headway sjø | 68 t (=2·mean) | Sea-Cargo AS (2025); 2-3 avganger/uke |
| Headway jernbane | 48 t | METRANS (2024); 7-10 avganger/uke |
| Grense Norge–EU | tri(0,25, 0,5, 1) t | EØS-avtalen, forenklet toll |
| Grense EU–Serbia | tri(2, 3, 5) t | EU Delegation to Serbia (2019) |
| P(disrupsjon) jernbane | 0,08 pr segment | Christopher (2016); modell-antakelse |
| VFTTS α | 330 kr/t | Halse et al. (2019); 13,6 kr/tonn-time × 24 tonn |
| RR | 0,8 | Halse et al. (2019), Tabell 5.1 |

## Tolkning av outputs

- **E[t], σ(t)**: Forventet ledetid og standardavvik. σ(t) er hjertet i pålitelighetsanalysen.
- **P10/P50/P90**: 10 %, 50 % (median), 90 % persentiler. P90 er "verste-90 %"-tid (best-case).
- **Pålitelighetspremie β·σ(t)**: Hvor mye den ekstra usikkerheten koster i kr (Andersson et al. 2017).
- **GC m/påli.**: Total GC inkludert pålitelighetspremie. Dette er det beslutningstakeren bør sammenligne på.

## Veien videre

Naturlige utvidelser av modellen for diskusjonskapittelet:

1. **Sensitivitetsanalyse** – plott E[GC] som funksjon av VFTTS eller CV
2. **Korrelerte forsinkelser** – samme værhendelse rammer flere segmenter
3. **Stokastiske kostnader** – la drivstoffpris/RoRo-rate variere
4. **Service-level constraint** – minste pålitelighet (eks. P95 < 14 dager)
