"""
monte_carlo.py
==============
Monte Carlo-simulering av multimodale transportruter Karmøy -> Beograd.

Leser inputs direkte fra Excel-modellen (model_mc_ready.xlsx) og simulerer
ledetid + generaliserte kostnader (GC) for hver rute under operasjonell
usikkerhet.

Bruk:
    python monte_carlo.py                       # standard Excel-fil
    python monte_carlo.py --excel min_fil.xlsx  # annen fil
    python monte_carlo.py --no-disrupt          # uten disrupsjons-haler
    python monte_carlo.py --n 50000             # flere iterasjoner

Forfatter: Bachelor-team, 2026
"""

import argparse
from pathlib import Path

import numpy as np
import pandas as pd
import matplotlib.pyplot as plt

# ----------------------------------------------------------------------------
# 1. INNLESING FRA EXCEL
# ----------------------------------------------------------------------------

def read_segments(excel_path: Path) -> pd.DataFrame:
    """
    Leser Ruter-arket og returnerer ren DataFrame med ett segment pr rad.
    Filtrerer bort header-rader og 'Totalt'-rader.
    """
    df = pd.read_excel(excel_path, sheet_name='Ruter', header=None)

    # Finn header-rader (de som har 'Segment' i kolonne 2 / index 2)
    # Originale headers: rad 0, 9, 20, 30, 38 (0-indeksert)
    # Alle rader med RuteID utfylt = segment-rader
    df.columns = [
        'RuteID', 'Header', 'Segment', 'Segmenttype', 'Mode',
        'Fra', 'Til', 'Distanse_km', 'Hastighet_kmt', 'Transporttid_t',
        'Terminaltid_t', 'Ventetid_t', 'Segmenttid_t', 'Omlastninger',
        'Grenseforsinkelse_t', 'Dager', 'Routescanner', 'Extra'
    ]

    # Behold kun rader hvor RuteID er R1..R5 og Segment er et tall
    df = df[df['RuteID'].astype(str).str.match(r'^R\d+$', na=False)].copy()
    df = df[pd.to_numeric(df['Segment'], errors='coerce').notna()].copy()

    # Type-cast
    numeric_cols = ['Segment', 'Distanse_km', 'Hastighet_kmt', 'Transporttid_t',
                    'Terminaltid_t', 'Ventetid_t', 'Segmenttid_t',
                    'Omlastninger', 'Grenseforsinkelse_t']
    for c in numeric_cols:
        df[c] = pd.to_numeric(df[c], errors='coerce')

    df['Mode'] = df['Mode'].astype(str).str.strip()
    df = df.reset_index(drop=True)
    return df[['RuteID', 'Segment', 'Segmenttype', 'Mode', 'Fra', 'Til',
               'Distanse_km', 'Hastighet_kmt', 'Transporttid_t',
               'Terminaltid_t', 'Ventetid_t', 'Omlastninger',
               'Grenseforsinkelse_t']]


def read_stochastic_params(excel_path: Path) -> dict:
    """
    Leser Stokastiske_parametere-arket. Returnerer dict med 5 sub-tabeller.
    Robust mot endringer i antall rader: leser til neste section-header.
    """
    raw = pd.read_excel(excel_path, sheet_name='Stokastiske_parametere',
                        header=None)

    def find_row(text_start):
        for i, val in enumerate(raw.iloc[:, 0]):
            if isinstance(val, str) and val.startswith(text_start):
                return i
        return None

    def section_nrows(start_idx):
        """Tell datarader fra start_idx til neste section-header eller tom rad."""
        n = 0
        for i in range(start_idx, len(raw)):
            val = raw.iloc[i, 0]
            if val is None or (isinstance(val, float) and pd.isna(val)):
                break
            if isinstance(val, str) and (val.startswith('1.') or val.startswith('2.')
                                          or val.startswith('3.') or val.startswith('4.')
                                          or val.startswith('5.') or val.startswith('6.')):
                break
            n += 1
        return n

    # Seksjon 1: CV per modus
    s1_start = find_row('1. Transporttid')
    n1 = section_nrows(s1_start + 2)  # +2 for section-header og kolonne-header
    cv = pd.read_excel(excel_path, sheet_name='Stokastiske_parametere',
                       header=s1_start + 1, nrows=n1, usecols='A:C')
    cv.columns = ['Mode', 'Fordeling', 'CV']
    cv = cv.dropna(subset=['Mode']).set_index('Mode')

    # Seksjon 2: Terminaltid triangulær
    s2_start = find_row('2. Terminaltid')
    n2 = section_nrows(s2_start + 2)
    tri = pd.read_excel(excel_path, sheet_name='Stokastiske_parametere',
                        header=s2_start + 1, nrows=n2, usecols='A:D')
    tri.columns = ['Mode', 'Min', 'Mode_val', 'Max']
    tri = tri.dropna(subset=['Mode']).set_index('Mode')

    # Seksjon 3: Headway (har en ekstra forklaringslinje, så +3 i stedet for +2)
    s3_start = find_row('3. Ventetid')
    n3 = section_nrows(s3_start + 3)
    hw = pd.read_excel(excel_path, sheet_name='Stokastiske_parametere',
                       header=s3_start + 2, nrows=n3, usecols='A:B')
    hw.columns = ['Mode', 'Headway']
    hw = hw.dropna(subset=['Mode']).set_index('Mode')

    # Seksjon 4: Grenseforsinkelse
    s4_start = find_row('4. Grenseforsinkelse')
    n4 = section_nrows(s4_start + 2)
    border = pd.read_excel(excel_path, sheet_name='Stokastiske_parametere',
                           header=s4_start + 1, nrows=n4, usecols='A:D')
    border.columns = ['Grense', 'Min', 'Mode', 'Max']
    border = border.dropna(subset=['Grense']).set_index('Grense')

    # Seksjon 5: Disrupsjon (har forklaringslinje, +3)
    s5_start = find_row('5. Disrupsjons')
    n5 = section_nrows(s5_start + 3)
    disrupt = pd.read_excel(excel_path, sheet_name='Stokastiske_parametere',
                            header=s5_start + 2, nrows=n5, usecols='A:E')
    disrupt.columns = ['Mode', 'P', 'Min', 'Mode_val', 'Max']
    disrupt = disrupt.dropna(subset=['Mode']).set_index('Mode')

    # Seksjon 6: MC-innstillinger
    s6_start = find_row('6. Monte Carlo')
    n6 = section_nrows(s6_start + 2)
    mc = pd.read_excel(excel_path, sheet_name='Stokastiske_parametere',
                       header=s6_start + 1, nrows=n6, usecols='A:B')
    mc.columns = ['Parameter', 'Verdi']
    mc = mc.dropna(subset=['Parameter']).set_index('Parameter')['Verdi']

    return {
        'cv': cv, 'terminal_tri': tri, 'headway': hw,
        'border': border, 'disrupt': disrupt, 'mc': mc,
    }


def read_cost_params(excel_path: Path) -> dict:
    """Leser Kostnadsparametere_MC som dict."""
    df = pd.read_excel(excel_path, sheet_name='Kostnadsparametere_MC',
                       header=3, usecols='A:B')
    df.columns = ['Parameter', 'Verdi']
    df = df.dropna(subset=['Parameter'])
    return df.set_index('Parameter')['Verdi'].to_dict()


# ----------------------------------------------------------------------------
# 2. MONTE CARLO-SIMULERING
# ----------------------------------------------------------------------------

def simulate_segment_time(segment_row, params, rng, n, include_disrupt=True):
    """
    Trekker n stokastiske realisasjoner av segmenttid (timer).

    Komponenter:
      - Transporttid:     lognormal(mean=t_mean, CV=cv[mode])
      - Terminaltid:      triangular(min, mode, max) for mode-en
      - Ventetid:         uniform[0, headway[mode]]
      - Grenseforsinkelse: triangular for "EU_Serbia" hvis row har grenseforsinkelse > 0
      - Disrupsjon:       Bernoulli(P) * triangular for haleforsinkelse
    """
    mode = segment_row['Mode']

    # Normaliser jernbane-varianter til "Rail" for parameter-oppslag.
    # Dette lar Ruter-arket beholde rail1/rail2/rail3-betegnelser
    # mens Stokastiske_parametere kan ha bare én "Rail"-rad.
    mode_lookup = 'Rail' if mode.lower().startswith('rail') else mode

    # ---- Transporttid: lognormal med samme mean som deterministisk verdi ----
    # Hvis Excel-formel ikke har blitt beregnet (NaN), beregn selv:
    t_mean = float(segment_row['Transporttid_t']) if pd.notna(segment_row['Transporttid_t']) else \
             float(segment_row['Distanse_km']) / float(segment_row['Hastighet_kmt'])
    cv = float(params['cv'].loc[mode_lookup, 'CV']) if mode_lookup in params['cv'].index else 0.10
    if t_mean > 0 and cv > 0:
        # lognormal: hvis X = exp(mu + sigma*Z), så E[X] = exp(mu + sigma^2/2)
        # CV = sqrt(exp(sigma^2) - 1)  =>  sigma^2 = ln(1 + CV^2)
        sigma = np.sqrt(np.log(1 + cv**2))
        mu = np.log(t_mean) - sigma**2 / 2
        t_transport = rng.lognormal(mean=mu, sigma=sigma, size=n)
    else:
        t_transport = np.full(n, t_mean)

    # ---- Terminaltid: triangulær ----
    if mode_lookup in params['terminal_tri'].index:
        tri = params['terminal_tri'].loc[mode_lookup]
        t_terminal = rng.triangular(left=tri['Min'], mode=tri['Mode_val'],
                                    right=tri['Max'], size=n)
    else:
        t_terminal = np.full(n, float(segment_row['Terminaltid_t']))

    # ---- Ventetid: uniform[0, headway] ----
    if mode_lookup in params['headway'].index:
        hw = float(params['headway'].loc[mode_lookup, 'Headway'])
        t_wait = rng.uniform(0, hw, size=n) if hw > 0 else np.zeros(n)
    else:
        t_wait = np.full(n, float(segment_row['Ventetid_t']))

    # ---- Grenseforsinkelse: triangulær (EU_Serbia hvis registrert) ----
    border_t = float(segment_row['Grenseforsinkelse_t'])
    if border_t > 0:
        # Bruk EU_Serbia hvis > 1 time, ellers Norge_EU
        which = 'EU_Serbia' if border_t >= 1.5 else 'Norge_EU'
        if which in params['border'].index:
            b = params['border'].loc[which]
            t_border = rng.triangular(b['Min'], b['Mode'], b['Max'], size=n)
        else:
            t_border = np.full(n, border_t)
    else:
        t_border = np.zeros(n)

    # ---- Disrupsjon: P * forsinkelse ----
    if include_disrupt and mode_lookup in params['disrupt'].index:
        d = params['disrupt'].loc[mode_lookup]
        hits = rng.uniform(0, 1, size=n) < d['P']
        delays = rng.triangular(d['Min'], d['Mode_val'], d['Max'], size=n)
        t_disrupt = hits * delays
    else:
        t_disrupt = np.zeros(n)

    return t_transport + t_terminal + t_wait + t_border + t_disrupt


def simulate_route(route_id, segments_df, params, costs, rng, n,
                   include_disrupt=True):
    """
    Simulerer én rute n ganger. Returnerer dict med arrays for tid og kostnader.
    """
    route_segs = segments_df[segments_df['RuteID'] == route_id]
    n_segs = len(route_segs)

    # Sum segmenttider
    total_time = np.zeros(n)
    for _, seg in route_segs.iterrows():
        total_time += simulate_segment_time(seg, params, rng, n, include_disrupt)

    # ---- Kostnader ----
    # Direkte transportkostnader (deterministisk - prises ikke med variabel pris i denne modellen)
    direct = 0.0
    terminal_cost = 0.0
    load_tonn = costs['Lastemengde_tonn']

    for _, seg in route_segs.iterrows():
        mode = seg['Mode']
        dist = seg['Distanse_km']

        if mode in ('Sjø', 'Sjø_havgaaende'):
            direct += dist * costs['Sjø_kostnad_per_km']
        elif mode == 'Truck':
            direct += dist * costs['Vei_kostnad_per_tonn_km'] * load_tonn
        elif mode.lower().startswith('rail'):
            direct += dist * costs['Rail_kostnad_per_tonn_km'] * load_tonn

        # Terminalkostnad pr omlasting:
        # Vi behandler enhver omlasting > 0 som "complex" hvis det involverer mode-skift,
        # ellers "simple". Konservativt: bruk gjennomsnitt.
        n_omlast = seg['Omlastninger']
        if pd.notna(n_omlast) and n_omlast > 0:
            terminal_cost += n_omlast * costs['Terminal_complex']

    # GC = direkte + terminal + alfa * E[t] + beta * sigma(t)
    alpha = costs['VFTTS_alfa']
    beta = costs['Reliability_ratio_RR'] * alpha

    mean_t = total_time.mean()
    std_t = total_time.std()

    time_cost = alpha * total_time
    # GC pr realisasjon: bruk realisert tid for tidskostnad,
    # men beta*sigma er en aggregert kostnad pr rute (ikke pr realisasjon)
    gc_per_iter = direct + terminal_cost + time_cost
    gc_with_reliability = direct + terminal_cost + alpha * mean_t + beta * std_t

    return {
        'route_id': route_id,
        'time': total_time,
        'gc': gc_per_iter,
        'direct_cost': direct,
        'terminal_cost': terminal_cost,
        'mean_time': mean_t,
        'std_time': std_t,
        'gc_aggregated': gc_with_reliability,
        'reliability_premium': beta * std_t,
    }


# ----------------------------------------------------------------------------
# 3. RAPPORTERING
# ----------------------------------------------------------------------------

def print_summary(results):
    print("\n" + "=" * 88)
    print("MONTE CARLO-RESULTATER  (alle verdier i kr og timer)")
    print("=" * 88)

    rows = []
    for r in results:
        rows.append({
            'Rute': r['route_id'],
            'E[t] timer': r['mean_time'],
            'σ(t) timer': r['std_time'],
            't P10': np.percentile(r['time'], 10),
            't P50': np.percentile(r['time'], 50),
            't P90': np.percentile(r['time'], 90),
            'E[GC] kr': r['gc'].mean(),
            'σ(GC) kr': r['gc'].std(),
            'GC P10': np.percentile(r['gc'], 10),
            'GC P90': np.percentile(r['gc'], 90),
            'Pålit.premie β·σ(t) kr': r['reliability_premium'],
            'GC m/påli.': r['gc_aggregated'],
        })

    df = pd.DataFrame(rows).set_index('Rute')
    pd.set_option('display.float_format', lambda x: f'{x:,.1f}')
    pd.set_option('display.width', 200)
    pd.set_option('display.max_columns', 20)
    print(df.to_string())

    # Eksporter til CSV for videre bruk
    df.to_csv('mc_results_summary.csv')
    print("\n→ Lagret 'mc_results_summary.csv'")
    return df


def plot_results(results, out_dir=Path('.')):
    """Histogram av ledetid og GC pr rute, + sammenligningsplot."""
    out_dir.mkdir(exist_ok=True)
    colors = ['#1f77b4', '#ff7f0e', '#2ca02c', '#d62728', '#9467bd']

    # Plot 1: Ledetid-fordeling pr rute (histogrammer overlappet)
    fig, ax = plt.subplots(figsize=(10, 6))
    for i, r in enumerate(results):
        ax.hist(r['time'], bins=60, alpha=0.5, label=r['route_id'],
                color=colors[i], density=True)
    ax.set_xlabel('Total ledetid (timer)')
    ax.set_ylabel('Sannsynlighetstetthet')
    ax.set_title('Fordeling av total ledetid pr rute (Monte Carlo, N={})'
                 .format(len(results[0]['time'])))
    ax.legend()
    ax.grid(alpha=0.3)
    plt.tight_layout()
    fig.savefig(out_dir / 'fig_ledetid_fordeling.png', dpi=150)
    plt.close(fig)

    # Plot 2: GC-fordeling pr rute
    fig, ax = plt.subplots(figsize=(10, 6))
    for i, r in enumerate(results):
        ax.hist(r['gc'], bins=60, alpha=0.5, label=r['route_id'],
                color=colors[i], density=True)
    ax.set_xlabel('Generaliserte kostnader GC (kr)')
    ax.set_ylabel('Sannsynlighetstetthet')
    ax.set_title('Fordeling av GC pr rute')
    ax.legend()
    ax.grid(alpha=0.3)
    plt.tight_layout()
    fig.savefig(out_dir / 'fig_gc_fordeling.png', dpi=150)
    plt.close(fig)

    # Plot 3: Boxplot ledetid
    fig, ax = plt.subplots(figsize=(9, 6))
    data = [r['time'] for r in results]
    labels = [r['route_id'] for r in results]
    bp = ax.boxplot(data, tick_labels=labels, patch_artist=True,
                    showfliers=True, whis=(5, 95))
    for patch, color in zip(bp['boxes'], colors):
        patch.set_facecolor(color)
        patch.set_alpha(0.6)
    ax.set_ylabel('Total ledetid (timer)')
    ax.set_title('Sammenligning av ledetid (boks: Q1-Q3, whiskers: P5-P95)')
    ax.grid(alpha=0.3, axis='y')
    plt.tight_layout()
    fig.savefig(out_dir / 'fig_ledetid_boxplot.png', dpi=150)
    plt.close(fig)

    # Plot 4: Mean GC vs std GC scatter (risk-return-aktig plot)
    fig, ax = plt.subplots(figsize=(8, 6))
    for i, r in enumerate(results):
        ax.scatter(r['gc'].mean(), r['gc'].std(), s=200, color=colors[i],
                   label=r['route_id'], edgecolor='black', zorder=3)
        ax.annotate(r['route_id'],
                    (r['gc'].mean(), r['gc'].std()),
                    xytext=(5, 5), textcoords='offset points',
                    fontsize=11, fontweight='bold')
    ax.set_xlabel('Forventet GC (kr)')
    ax.set_ylabel('Standardavvik GC (kr)')
    ax.set_title('Trade-off: forventet kostnad vs. variasjon')
    ax.grid(alpha=0.3)
    plt.tight_layout()
    fig.savefig(out_dir / 'fig_gc_risk_return.png', dpi=150)
    plt.close(fig)

    print(f"→ Lagret 4 figurer i {out_dir.resolve()}")


# ----------------------------------------------------------------------------
# 4. MAIN
# ----------------------------------------------------------------------------

def main():
    parser = argparse.ArgumentParser()
    parser.add_argument('--excel', default='model_mc_ready_2.xlsx',
                        help='Path til Excel-modellen')
    parser.add_argument('--n', type=int, default=None,
                        help='Antall iterasjoner (overstyrer Excel)')
    parser.add_argument('--seed', type=int, default=None,
                        help='Random seed (overstyrer Excel)')
    parser.add_argument('--no-disrupt', action='store_true',
                        help='Skru av disrupsjons-haler')
    parser.add_argument('--out', default='.', help='Output-mappe for figurer')
    args = parser.parse_args()

    excel_path = Path(args.excel)
    if not excel_path.exists():
        raise FileNotFoundError(f"Finner ikke {excel_path}")

    print(f"Leser inputs fra: {excel_path.resolve()}")
    segments = read_segments(excel_path)
    params = read_stochastic_params(excel_path)
    costs = read_cost_params(excel_path)

    n = args.n or int(params['mc'].get('N_simuleringer', 10000))
    seed = args.seed or int(params['mc'].get('Random_seed', 42))
    include_disrupt = (
        not args.no_disrupt
        and bool(params['mc'].get('Inkluder_disrupsjoner', 1))
    )

    print(f"  Ruter funnet: {sorted(segments['RuteID'].unique())}")
    print(f"  Antall segmenter totalt: {len(segments)}")
    print(f"  N = {n:,}, seed = {seed}, disrupsjoner = {include_disrupt}")

    rng = np.random.default_rng(seed)
    results = []
    for route_id in sorted(segments['RuteID'].unique()):
        res = simulate_route(route_id, segments, params, costs, rng, n,
                             include_disrupt=include_disrupt)
        results.append(res)

    summary = print_summary(results)
    plot_results(results, Path(args.out))

    print("\nFerdig.")


if __name__ == '__main__':
    main()