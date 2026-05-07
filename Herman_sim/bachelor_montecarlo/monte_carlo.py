"""
monte_carlo.py
==============
Monte Carlo-simulering av multimodale transportruter Karmøy -> Beograd.

Leser inputs direkte fra Excel-modellen (model_mc_ready_2.xlsx) og simulerer
ledetid + generaliserte kostnader (GC) for hver rute under operasjonell
usikkerhet.

GC-formel:
    GC = C_direkte + C_terminal + alpha * E[t] + beta * sigma(t)
    der alpha = VFTTS, beta = RR * alpha (Halse et al. 2019)

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
from scipy.stats import gaussian_kde


# ----------------------------------------------------------------------------
# 1. INNLESING FRA EXCEL
# ----------------------------------------------------------------------------

def read_segments(excel_path: Path) -> pd.DataFrame:
    """
    Leser Ruter-arket og returnerer ren DataFrame med ett segment pr rad.
    Filtrerer bort header-rader og 'Totalt'-rader.
    """
    df = pd.read_excel(excel_path, sheet_name='Ruter', header=None)

    df.columns = [
        'RuteID', 'Header', 'Segment', 'Segmenttype', 'Mode',
        'Fra', 'Til', 'Distanse_km', 'Hastighet_kmt', 'Transporttid_t',
        'Terminaltid_t', 'Ventetid_t', 'Segmenttid_t', 'Omlastninger',
        'Grenseforsinkelse_t', 'Dager', 'Routescanner', 'Extra'
    ]

    # Behold kun rader hvor RuteID er R1..R5 og Segment er et tall
    df = df[df['RuteID'].astype(str).str.match(r'^R\d+$', na=False)].copy()
    df = df[pd.to_numeric(df['Segment'], errors='coerce').notna()].copy()

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
    Robust mot endringer i antall rader.
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
    n1 = section_nrows(s1_start + 2)
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

    # Seksjon 3: Headway
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

    # Seksjon 5: Disrupsjon
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
      - Grenseforsinkelse: triangular for "EU_Serbia"/"Norge_EU"
      - Disrupsjon:       Bernoulli(P) * triangular for haleforsinkelse
    """
    mode = segment_row['Mode']

    # Normaliser jernbane-varianter til "Rail" for parameter-oppslag.
    mode_lookup = 'Rail' if mode.lower().startswith('rail') else mode

    # ---- Transporttid: lognormal med samme mean som deterministisk verdi ----
    # Hvis Excel-formel ikke har blitt beregnet (NaN), beregn selv:
    t_mean = float(segment_row['Transporttid_t']) if pd.notna(segment_row['Transporttid_t']) else \
             float(segment_row['Distanse_km']) / float(segment_row['Hastighet_kmt'])
    cv = float(params['cv'].loc[mode_lookup, 'CV']) if mode_lookup in params['cv'].index else 0.10
    if t_mean > 0 and cv > 0:
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

    # ---- Grenseforsinkelse: triangulær ----
    border_t = float(segment_row['Grenseforsinkelse_t'])
    if border_t > 0:
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

    # Sum segmenttider
    total_time = np.zeros(n)
    for _, seg in route_segs.iterrows():
        total_time += simulate_segment_time(seg, params, rng, n, include_disrupt)

    # ---- Kostnader (deterministiske) ----
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

        n_omlast = seg['Omlastninger']
        if pd.notna(n_omlast) and n_omlast > 0:
            terminal_cost += n_omlast * costs['Terminal_complex']

    # GC = direkte + terminal + alpha * t + beta * sigma(t)
    alpha = costs['VFTTS_alfa']
    beta = costs['Reliability_ratio_RR'] * alpha

    mean_t = total_time.mean()
    std_t = total_time.std()

    time_cost = alpha * total_time
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
        'alpha': alpha,
        'beta': beta,
    }


# ----------------------------------------------------------------------------
# 3. RAPPORTERING (TABELL OG CSV)
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

    df.to_csv('mc_results_summary.csv')
    print("\n→ Lagret 'mc_results_summary.csv'")
    return df


# ----------------------------------------------------------------------------
# 4. FIGURER
# ----------------------------------------------------------------------------

def plot_results(results, out_dir=Path('.')):
    """Lag alle figurene og lagre dem i out_dir."""
    out_dir.mkdir(exist_ok=True)
    colors = ['#1f77b4', '#ff7f0e', '#2ca02c', '#d62728', '#9467bd']

    # ------------------------------------------------------------------
    # FIGUR 1: Ledetid-fordeling pr rute (KDE - glattet, i DAGER)
    # ------------------------------------------------------------------
    fig, ax = plt.subplots(figsize=(10, 6))
    all_times_days = np.concatenate([r['time'] / 24 for r in results])
    x_grid = np.linspace(all_times_days.min(), all_times_days.max(), 500)
    for i, r in enumerate(results):
        time_days = r['time'] / 24
        kde = gaussian_kde(time_days)
        ax.plot(x_grid, kde(x_grid), label=r['display_name'],
                color=colors[i], linewidth=2.5)
        ax.fill_between(x_grid, kde(x_grid), alpha=0.15, color=colors[i])
        # Vertikal linje på gjennomsnittet
        mean_days = r['mean_time'] / 24
        ax.axvline(mean_days, color=colors[i], linestyle='--',
                   linewidth=1.5, alpha=0.7)
    ax.set_xlabel('Total ledetid (dager)')
    ax.set_ylabel('Sannsynlighetstetthet')
    ax.set_title('Fordeling av total ledetid pr rute (Monte Carlo, N={})'
                 .format(len(results[0]['time'])))
    ax.legend()
    ax.grid(alpha=0.3)
    plt.tight_layout()
    fig.savefig(out_dir / 'fig_ledetid_fordeling.png', dpi=150)
    plt.close(fig)

    # ------------------------------------------------------------------
    # FIGUR 2: GC-fordeling pr rute (KDE - glattet)
    # ------------------------------------------------------------------
    fig, ax = plt.subplots(figsize=(10, 6))
    all_gc = np.concatenate([r['gc'] for r in results])
    x_grid = np.linspace(all_gc.min(), all_gc.max(), 500)
    for i, r in enumerate(results):
        kde = gaussian_kde(r['gc'])
        ax.plot(x_grid, kde(x_grid), label=r['display_name'],
                color=colors[i], linewidth=2.5)
        ax.fill_between(x_grid, kde(x_grid), alpha=0.15, color=colors[i])
    ax.set_xlabel('Generaliserte kostnader GC (kr)')
    ax.set_ylabel('Sannsynlighetstetthet')
    ax.set_title('Fordeling av GC pr rute')
    ax.legend()
    ax.grid(alpha=0.3)
    plt.tight_layout()
    fig.savefig(out_dir / 'fig_gc_fordeling.png', dpi=150)
    plt.close(fig)

    # ------------------------------------------------------------------
    # FIGUR 3: Boxplot av ledetid (i DAGER for slide 15)
    # ------------------------------------------------------------------
    fig, ax = plt.subplots(figsize=(10, 6))
    data_days = [r['time'] / 24 for r in results]  # konverter timer til dager
    labels = [r['display_name'] for r in results]
    bp = ax.boxplot(data_days, tick_labels=labels, patch_artist=True,
                    showfliers=True, whis=(5, 95))
    for patch, color in zip(bp['boxes'], colors):
        patch.set_facecolor(color)
        patch.set_alpha(0.6)
    ax.set_ylabel('Total ledetid (dager)')
    ax.set_title('Sammenligning av ledetid (boks: Q1-Q3, whiskers: P5-P95)')
    ax.grid(alpha=0.3, axis='y')
    ax.tick_params(axis='x', rotation=15)
    plt.tight_layout()
    fig.savefig(out_dir / 'fig_ledetid_boxplot.png', dpi=150)
    plt.close(fig)

    # ------------------------------------------------------------------
    # FIGUR 4: Risk-return scatter (E[GC] vs σ(GC))
    # ------------------------------------------------------------------
    fig, ax = plt.subplots(figsize=(9, 6))
    for i, r in enumerate(results):
        ax.scatter(r['gc'].mean(), r['gc'].std(), s=200, color=colors[i],
                   label=r['display_name'], edgecolor='black', zorder=3)
        short_label = r['display_name'].split(': ')[1] if ': ' in r['display_name'] else r['display_name']
        ax.annotate(short_label,
                    (r['gc'].mean(), r['gc'].std()),
                    xytext=(8, 5), textcoords='offset points',
                    fontsize=10, fontweight='bold')
    ax.set_xlabel('Forventet GC (kr)')
    ax.set_ylabel('Standardavvik GC (kr)')
    ax.set_title('Trade-off: forventet kostnad vs. variasjon')
    ax.grid(alpha=0.3)
    plt.tight_layout()
    fig.savefig(out_dir / 'fig_gc_risk_return.png', dpi=150)
    plt.close(fig)

    # ------------------------------------------------------------------
    # FIGUR 4b: Risk-return scatter for LEDETID (E[t] vs σ(t)) — i dager
    # ------------------------------------------------------------------
    fig, ax = plt.subplots(figsize=(9, 6))
    for i, r in enumerate(results):
        mean_days = r['mean_time'] / 24
        std_days = r['std_time'] / 24
        ax.scatter(mean_days, std_days, s=200, color=colors[i],
                   label=r['display_name'], edgecolor='black', zorder=3)
        short_label = r['display_name'].split(': ')[1] if ': ' in r['display_name'] else r['display_name']
        ax.annotate(short_label,
                    (mean_days, std_days),
                    xytext=(8, 5), textcoords='offset points',
                    fontsize=10, fontweight='bold')
    ax.set_xlabel('Forventet ledetid (dager)')
    ax.set_ylabel('Standardavvik ledetid (dager)')
    ax.set_title('Trade-off: forventet ledetid vs. variasjon')
    ax.grid(alpha=0.3)
    plt.tight_layout()
    fig.savefig(out_dir / 'fig_ledetid_risk_return.png', dpi=150)
    plt.close(fig)

    # ------------------------------------------------------------------
    # FIGUR 5: Enkel bar - bare forventet GC pr rute
    # ------------------------------------------------------------------
    fig, ax = plt.subplots(figsize=(10, 6))
    route_names = [r['display_name'] for r in results]
    mean_gcs = [r['gc'].mean() for r in results]
    bars = ax.bar(route_names, mean_gcs, color=colors, edgecolor='black')
    for bar, val in zip(bars, mean_gcs):
        ax.text(bar.get_x() + bar.get_width()/2, bar.get_height(),
                f'{val:,.0f} kr',
                ha='center', va='bottom', fontsize=10, fontweight='bold')
    ax.set_ylabel('Forventet GC (kr)')
    ax.set_title('Forventet generaliserte kostnader pr rute')
    ax.grid(alpha=0.3, axis='y')
    ax.set_ylim(bottom=min(mean_gcs) * 0.95, top=max(mean_gcs) * 1.05)
    ax.tick_params(axis='x', rotation=15)
    plt.tight_layout()
    fig.savefig(out_dir / 'fig_gc_mean.png', dpi=150)
    plt.close(fig)

    # ------------------------------------------------------------------
    # FIGUR 6: Stacked bar - GC brutt ned i komponenter (UTEN pålitelighet)
    # ------------------------------------------------------------------
    fig, ax = plt.subplots(figsize=(10, 6))
    direct_costs = [r['direct_cost'] for r in results]
    terminal_costs = [r['terminal_cost'] for r in results]
    time_costs = [r['alpha'] * r['mean_time'] for r in results]

    ax.bar(route_names, direct_costs, label='Direkte transport',
           color='#4472C4', edgecolor='black')
    bottom = list(direct_costs)
    ax.bar(route_names, terminal_costs, bottom=bottom,
           label='Terminalkostnad', color='#ED7D31', edgecolor='black')
    bottom = [b + t for b, t in zip(bottom, terminal_costs)]
    ax.bar(route_names, time_costs, bottom=bottom,
           label='Tidskostnad (α·E[t])', color='#A5A5A5', edgecolor='black')

    totals = [d + t + ti for d, t, ti in zip(direct_costs, terminal_costs, time_costs)]
    for i, total in enumerate(totals):
        ax.text(i, total, f'{total:,.0f}',
                ha='center', va='bottom', fontsize=10, fontweight='bold')

    ax.set_ylabel('Generaliserte kostnader (kr)')
    ax.set_title('GC pr rute — tradisjonell (uten pålitelighetspremie)')
    ax.legend(loc='lower right')
    ax.grid(alpha=0.3, axis='y')
    ax.tick_params(axis='x', rotation=15)
    plt.tight_layout()
    fig.savefig(out_dir / 'fig_gc_stacked_uten_palitelighet.png', dpi=150)
    plt.close(fig)

    # ------------------------------------------------------------------
    # FIGUR 7: Stacked bar - GC brutt ned i komponenter (MED pålitelighet)
    # ------------------------------------------------------------------
    fig, ax = plt.subplots(figsize=(10, 6))
    reliability_costs = [r['reliability_premium'] for r in results]

    ax.bar(route_names, direct_costs, label='Direkte transport',
           color='#4472C4', edgecolor='black')
    bottom = list(direct_costs)
    ax.bar(route_names, terminal_costs, bottom=bottom,
           label='Terminalkostnad', color='#ED7D31', edgecolor='black')
    bottom = [b + t for b, t in zip(bottom, terminal_costs)]
    ax.bar(route_names, time_costs, bottom=bottom,
           label='Tidskostnad (α·E[t])', color='#A5A5A5', edgecolor='black')
    bottom = [b + t for b, t in zip(bottom, time_costs)]
    ax.bar(route_names, reliability_costs, bottom=bottom,
           label='Pålitelighetspremie (β·σ(t))', color='#FFC000',
           edgecolor='black')

    totals = [d + t + ti + rl for d, t, ti, rl in
              zip(direct_costs, terminal_costs, time_costs, reliability_costs)]
    for i, total in enumerate(totals):
        ax.text(i, total, f'{total:,.0f}',
                ha='center', va='bottom', fontsize=10, fontweight='bold')

    ax.set_ylabel('Generaliserte kostnader (kr)')
    ax.set_title('GC pr rute — utvidet (med pålitelighetspremie)')
    ax.legend(loc='lower right')
    ax.grid(alpha=0.3, axis='y')
    ax.tick_params(axis='x', rotation=15)
    plt.tight_layout()
    fig.savefig(out_dir / 'fig_gc_stacked_med_palitelighet.png', dpi=150)
    plt.close(fig)

    print(f"→ Lagret 8 figurer i {out_dir.resolve()}")


# ----------------------------------------------------------------------------
# 5. MAIN
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

    # Navne-mapping for figurer (R-koden brukes fortsatt internt og i tabell)
    route_labels = {
        'R1': 'R1: Rostock',
        'R2': 'R2: Hamburg',
        'R3': 'R3: Rotterdam-München',
        'R4': 'R4: Rotterdam-Wien',
        'R5': 'R5: Bar (sjø)',
    }

    rng = np.random.default_rng(seed)
    results = []
    for route_id in sorted(segments['RuteID'].unique()):
        res = simulate_route(route_id, segments, params, costs, rng, n,
                             include_disrupt=include_disrupt)
        res['display_name'] = route_labels.get(route_id, route_id)
        results.append(res)

    print_summary(results)
    plot_results(results, Path(args.out))

    print("\nFerdig.")


if __name__ == '__main__':
    main()
