"""
tornado_analyse.py
==================
One-at-a-time (OAT) sensitivitetsanalyse: ±20 % på 5 parametere.
Genererer tornado-diagram med endring i E[GC] (inkl. pålitelighetspremie)
for alle 6 ruter (R1–R6, ingen block-train-versjoner).

Kjøring:
    python tornado_analyse.py                        # standard: transportmodell.xlsx
    python tornado_analyse.py --excel transportmodell.xlsx   # annen fil
    python tornado_analyse.py --n 5000               # færre iterasjoner
"""

import argparse
import numpy as np
import pandas as pd
import matplotlib.pyplot as plt
import matplotlib.patches as mpatches
from pathlib import Path

from monte_carlo import (
    read_segments, read_stochastic_params, read_cost_params,
    simulate_route,
)


RUTER   = ['R1', 'R2', 'R3', 'R4', 'R5', 'R6']
PERTURB = 0.20   # ±20 %
N_MC    = 10000
SEED    = 42

# Parametere som varieres — rekkefølge brukes i sortering
PARAMS = [
    'VFTTS',
    'Ventetid (headway)',
    'Terminalkost',
    'Rail punktlighet',
    'Veikostnad',
]

ROUTE_COLORS = {
    'R1': '#1f77b4', 'R2': '#ff7f0e', 'R3': '#2ca02c',
    'R4': '#d62728', 'R5': '#9467bd', 'R6': '#8c564b',
}


# ------------------------------------------------------------------------------
# Hjelpefunksjoner
# ------------------------------------------------------------------------------

def _copy_params(params):
    """Lag en kopi av params-dict der DataFrames kopieres (unngå mutasjon)."""
    return {k: (v.copy() if hasattr(v, 'copy') else v) for k, v in params.items()}


def _run_all_routes(segments, params, costs, n=N_MC, seed=SEED):
    """Kjør alle ruter og returner {rute_id: gc_aggregated}."""
    results = {}
    for rt in RUTER:
        rng = np.random.default_rng(seed)   # reset per rute for reproduserbarhet
        res = simulate_route(rt, segments, params, costs, rng, n,
                             include_disrupt=False)
        results[rt] = res['gc_aggregated']
    return results


def _perturbed_run(segments, params, costs, param_name, direction, n=N_MC, seed=SEED):
    """
    Returner {rute_id: gc_aggregated} med én parameter perturbert.
    direction: +1 (parameter øker) eller -1 (parameter reduseres).

    Rail punktlighet bruker invertert logikk:
        direction = +1 → CV * 0.8  (bedre punktlighet → lavere GC)
        direction = -1 → CV * 1.2  (dårligere punktlighet → høyere GC)
    """
    factor = 1 + direction * PERTURB
    p = _copy_params(params)
    c = dict(costs)
    s = segments.copy()

    if param_name == 'VFTTS':
        # alpha skaleres; beta = RR * alpha beregnes i simulate_route
        c['VFTTS_alfa'] = c['VFTTS_alfa'] * factor

    elif param_name == 'Ventetid (headway)':
        # Headway leses per segment fra Ruter-arket
        s['Headway_t'] = s['Headway_t'].astype(float) * factor

    elif param_name == 'Terminalkost':
        c['Terminal_complex'] = c['Terminal_complex'] * factor

    elif param_name == 'Rail punktlighet':
        # Invertert: lavere "punktlighet-parameter" = høyere CV = dårligere service
        #   direction = +1 → cv_factor = 0.8  (bedre punktlighet → lavere CV)
        #   direction = -1 → cv_factor = 1.2  (dårligere punktlighet → høyere CV)
        cv_factor = 1 - direction * PERTURB
        cv_df = p['cv'].copy()
        # Case-insensitivt oppslag — Excel kan ha 'Rail', 'rail', 'RAIL', osv.
        rail_key = next((k for k in cv_df.index if str(k).lower() == 'rail'), None)
        if rail_key is not None:
            cv_df.loc[rail_key, 'CV'] = cv_df.loc[rail_key, 'CV'] * cv_factor
        else:
            print(f"  ADVARSEL [{param_name}]: ingen rail-entry i CV-indeksen. "
                  f"Tilgjengelige: {list(cv_df.index)}")
        p['cv'] = cv_df

    elif param_name == 'Veikostnad':
        c['Vei_kostnad_per_tonn_km'] = c['Vei_kostnad_per_tonn_km'] * factor

    return _run_all_routes(s, p, c, n=n, seed=seed)


# ------------------------------------------------------------------------------
# Hoved-analyse
# ------------------------------------------------------------------------------

def compute_tornado(excel_path, n_mc=N_MC, seed=SEED):
    """
    Kjør OAT-analyse for alle 5 parametere og 6 ruter.
    Returnerer (df, baseline_dict).

    df har kolonner: Parameter, Rute, pct_lavere, pct_hoyere
      pct_lavere = % endring i GC når parameter er −20 %
      pct_hoyere = % endring i GC når parameter er +20 %
    """
    print("Leser Excel...")
    segments = read_segments(excel_path)
    params   = read_stochastic_params(excel_path)
    costs    = read_cost_params(excel_path)

    print(f"\nParams-nøkler:  {list(params.keys())}")
    print(f"CV-tabell:\n{params['cv']}")
    print(f"Costs-nøkler:   {sorted(costs.keys())}")
    print(f"VFTTS_alfa      = {costs['VFTTS_alfa']:.2f}")
    print(f"Terminal_complex= {costs['Terminal_complex']:.2f}")
    print(f"Vei_kr_tonn_km  = {costs['Vei_kostnad_per_tonn_km']:.4f}")
    print(f"Rail CV         = {params['cv'].loc['Rail', 'CV'] if 'Rail' in params['cv'].index else 'N/A'}")

    print(f"\nKjører baseline (N={n_mc:,}) ...")
    baseline = _run_all_routes(segments, params, costs, n=n_mc, seed=seed)
    print("Baseline GC (inkl. pålitelighet):")
    for rt, gc in baseline.items():
        print(f"  {rt}: {gc:,.0f} kr")

    rows = []
    for param in PARAMS:
        print(f"\n[{param}]")
        gc_low  = _perturbed_run(segments, params, costs, param, -1, n=n_mc, seed=seed)
        gc_high = _perturbed_run(segments, params, costs, param, +1, n=n_mc, seed=seed)
        for rt in RUTER:
            pct_low  = 100 * (gc_low[rt]  - baseline[rt]) / baseline[rt]
            pct_high = 100 * (gc_high[rt] - baseline[rt]) / baseline[rt]
            rows.append({
                'Parameter':  param,
                'Rute':       rt,
                'pct_lavere': pct_low,
                'pct_hoyere': pct_high,
            })
            print(f"  {rt}:  −20%→{pct_low:+.1f}%   +20%→{pct_high:+.1f}%")

    df = pd.DataFrame(rows)
    df.to_csv('tornado_results.csv', index=False)
    print("\n→ Lagret 'tornado_results.csv'")
    return df, baseline


# ------------------------------------------------------------------------------
# Tornado-diagram
# ------------------------------------------------------------------------------

def plot_tornado(df, out_dir=Path('.')):
    """
    Tornado-diagram:
      - Y-aksen: 5 parametere, sortert etter gjennomsnittlig absolutt swing (størst øverst)
      - X-aksen: % endring i E[GC] fra baseline
      - Per parameter: 6 mini-bars (én per rute), fargekoda
      - Solid fyll → parameter −20 %
      - Skravert (//) → parameter +20 %
    """
    # Sorter etter gjennomsnittlig absolutt swing (størst øverst = høyest y-posisjon)
    swing = df.groupby('Parameter').apply(
        lambda g: (g['pct_hoyere'].abs() + g['pct_lavere'].abs()).mean()
    ).sort_values(ascending=True)   # ascending=True siden vi inverterer y-aksen
    params_sorted = list(swing.index)

    n_params    = len(params_sorted)
    n_ruter     = len(RUTER)
    group_h     = 0.78
    bar_h       = group_h / n_ruter
    y_positions = np.arange(n_params)

    fig, ax = plt.subplots(figsize=(12, 7))

    for i, param in enumerate(params_sorted):
        for j, rute in enumerate(RUTER):
            row   = df[(df['Parameter'] == param) & (df['Rute'] == rute)].iloc[0]
            y_bar = y_positions[i] - group_h / 2 + bar_h * j + bar_h / 2
            color = ROUTE_COLORS[rute]

            pct_low  = row['pct_lavere']   # endring når parameter er −20 %
            pct_high = row['pct_hoyere']   # endring når parameter er +20 %

            if pct_low != 0:
                ax.barh(y_bar, pct_low, height=bar_h * 0.88,
                        color=color, edgecolor='black', linewidth=0.4, alpha=0.9)

            if pct_high != 0:
                ax.barh(y_bar, pct_high, height=bar_h * 0.88,
                        color=color, edgecolor='black', linewidth=0.4, alpha=0.9,
                        hatch='//')

    ax.axvline(0, color='black', linewidth=1.2)
    ax.set_yticks(y_positions)
    ax.set_yticklabels(params_sorted, fontsize=11)
    ax.invert_yaxis()   # størst swing øverst
    ax.set_xlabel('Endring i E[GC] (%) ved ±20 % parameterendring', fontsize=11)
    ax.set_title(
        'Sensitivitetsanalyse (OAT ±20 %): endring i forventet GC per rute\n'
        'Solid = parameter −20 %   |   Skravert = parameter +20 %',
        fontsize=11, pad=10,
    )
    ax.grid(axis='x', linewidth=0.4, alpha=0.4)

    # Legende: én patch per rute
    handles = [mpatches.Patch(color=ROUTE_COLORS[r], label=r) for r in RUTER]
    ax.legend(handles=handles, loc='upper left', fontsize=10,
              ncol=3, framealpha=0.95, title='Rute')

    plt.tight_layout()
    out_path = Path(out_dir) / 'fig_tornado.png'
    fig.savefig(out_path, dpi=150, bbox_inches='tight')
    plt.close(fig)
    print(f"→ Lagret '{out_path.name}'")


# ------------------------------------------------------------------------------
# Verifiseringsprint etter kjøring
# ------------------------------------------------------------------------------

def verify_results(df):
    """
    Printer en liten sjekk av forventede mønstre.
    Advarer hvis noe ser rart ut.
    """
    print("\n" + "=" * 60)
    print("VERIFISERING AV FORVENTEDE MØNSTRE")
    print("=" * 60)

    checks = [
        ("Rail punktlighet", ['R5', 'R6'], "~0 %",
         lambda pct: abs(pct) < 1.0),
        ("Veikostnad", ['R1', 'R2', 'R3', 'R4'], "liten effekt (<3 %)",
         lambda pct: abs(pct) < 3.0),
        ("Veikostnad", ['R6'], "stor effekt (>5 %)",
         lambda pct: abs(pct) > 5.0),
        ("VFTTS", RUTER, "merkbar effekt (>5 %)",
         lambda pct: abs(pct) > 5.0),
    ]

    for param, ruter, expectation, ok_fn in checks:
        for rt in ruter:
            row = df[(df['Parameter'] == param) & (df['Rute'] == rt)]
            if row.empty:
                continue
            pct = row.iloc[0]['pct_hoyere']
            status = "OK" if ok_fn(pct) else "ADVARSEL"
            print(f"  [{status}] {param} / {rt}: +20%→{pct:+.1f}%  (forventet: {expectation})")


# ------------------------------------------------------------------------------
# Main
# ------------------------------------------------------------------------------

def main():
    parser = argparse.ArgumentParser()
    parser.add_argument('--excel', type=Path, default=Path('transportmodell.xlsx'))
    parser.add_argument('--n', type=int, default=N_MC,
                        help='Antall MC-iterasjoner (lavere = raskere)')
    parser.add_argument('--out', type=Path, default=Path('.'))
    args = parser.parse_args()

    df, baseline = compute_tornado(args.excel, n_mc=args.n)
    verify_results(df)
    plot_tornado(df, args.out)
    print("\nFerdig.")


if __name__ == '__main__':
    main()
