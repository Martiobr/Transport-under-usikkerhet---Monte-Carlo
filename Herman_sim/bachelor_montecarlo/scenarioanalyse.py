"""
Scenarioanalyse og følsomhetsanalyse for bachelor-modellen.
Bygger på monte_carlo_v6.py — importerer alle nødvendige funksjoner.

Kjøres etter at hovedanalysen er kjørt:
    python3 scenarioanalyse.py --excel model_v6.xlsx
"""
import argparse
from pathlib import Path
import numpy as np
import pandas as pd
import matplotlib.pyplot as plt

from monte_carlo_v6 import (
    read_segments, read_stochastic_params, read_cost_params,
    simulate_route, simulate_segment_time
)


# ==============================================================================
# DEL 1: BREAK-EVEN-ANALYSE FOR R5 (SJØRUTE)
# ==============================================================================

def breakeven_sea_route(excel_path, n_mc=2000, seed=42,
                        rate_range=(3, 12), speed_range_knots=(14, 22),
                        n_grid=15):
    """
    Beregner E[GC] for R5 over et nett av sjøfraktrate (kr/km) og havgående
    sjøhastighet (knop). Identifiserer break-even-konturer mot øvrige ruter.
    
    Returnerer:
      rates:  np.array av kr/km-verdier (x-akse)
      speeds: np.array av knop-verdier (y-akse)  
      gc_r5:  2D-matrise med E[GC] for R5 i hvert punkt
      gc_others: dict med E[GC] for R1-R4, R6 (uavhengig av sjø-parametrene)
    """
    print("\n" + "=" * 70)
    print("BREAK-EVEN-ANALYSE FOR R5")
    print("=" * 70)
    
    segments = read_segments(excel_path)
    params = read_stochastic_params(excel_path)
    costs_base = read_cost_params(excel_path)
    
    # Først: kjør alle andre ruter med basisparametere (de er uavhengige av sjø-rate/hastighet)
    print("Kjører basis-simulering for R1-R4, R6...")
    rng = np.random.default_rng(seed)
    gc_others = {}
    for rt in ['R1', 'R2', 'R3', 'R4', 'R6']:
        if rt in segments['RuteID'].unique():
            res = simulate_route(rt, segments, params, costs_base, rng, n_mc,
                                 include_disrupt=False)
            gc_others[rt] = res['gc'].mean()
            print(f"  {rt}: E[GC] = {gc_others[rt]:,.0f} kr")
    
    # Sett opp rute-nett
    rates = np.linspace(rate_range[0], rate_range[1], n_grid)
    speeds_knots = np.linspace(speed_range_knots[0], speed_range_knots[1], n_grid)
    speeds_kmt = speeds_knots * 1.852  # konvertér til km/t
    
    gc_r5 = np.zeros((n_grid, n_grid))  # rows=speed, cols=rate
    
    print(f"\nKjører nett {n_grid}x{n_grid} = {n_grid*n_grid} kombinasjoner...")
    
    # Hent R5-segmentene
    r5_segs = segments[segments['RuteID'] == 'R5'].copy()
    
    for i, speed in enumerate(speeds_kmt):
        for j, rate in enumerate(rates):
            # Lag modifisert kostnadsdict
            costs_mod = dict(costs_base)
            costs_mod['Sjø_kostnad_per_km'] = rate
            
            # Lag modifisert segment-dataframe der havgående sjø får ny hastighet
            segs_mod = segments.copy()
            # Cast til float for å kunne sette desimaltall
            segs_mod['Hastighet_kmt'] = segs_mod['Hastighet_kmt'].astype(float)
            segs_mod['Transporttid_t'] = segs_mod['Transporttid_t'].astype(float)
            # Identifiser R5 seg 2 (Rotterdam-Bar, havgående)
            mask = (segs_mod['RuteID'] == 'R5') & (segs_mod['Segment'] == 2)
            segs_mod.loc[mask, 'Hastighet_kmt'] = speed
            # Oppdater transporttid for det segmentet
            for idx in segs_mod[mask].index:
                dist = segs_mod.loc[idx, 'Distanse_km']
                segs_mod.loc[idx, 'Transporttid_t'] = dist / speed
            
            rng = np.random.default_rng(seed)  # Reset for stabil sammenligning
            res = simulate_route('R5', segs_mod, params, costs_mod, rng, n_mc,
                                 include_disrupt=False)
            gc_r5[i, j] = res['gc'].mean()
    
    print("Ferdig med break-even-analysen.")
    return rates, speeds_knots, gc_r5, gc_others


def plot_breakeven(rates, speeds_knots, gc_r5, gc_others, out_dir=Path('.')):
    """
    Lager heatmap + 1D-plott for break-even-analysen.
    """
    out_dir = Path(out_dir)
    
    # ====== FIGUR 1: HEATMAP MED KONTURLINJER ======
    fig, ax = plt.subplots(figsize=(11, 7))
    
    # Heatmap av R5 GC
    im = ax.pcolormesh(rates, speeds_knots, gc_r5 / 1000,
                       cmap='RdYlGn_r', shading='auto')
    cbar = fig.colorbar(im, ax=ax, label='E[GC] for R5 (1000 kr)')
    
    # Konturlinjer for break-even mot andre ruter
    gc_r6 = gc_others.get('R6', None)
    gc_r2 = gc_others.get('R2', None)
    
    contour_levels = []
    contour_labels = []
    if gc_r6:
        contour_levels.append(gc_r6 / 1000)
        contour_labels.append(f'R5 = R6 ({gc_r6/1000:.0f}k kr)')
    if gc_r2:
        contour_levels.append(gc_r2 / 1000)
        contour_labels.append(f'R5 = R2 ({gc_r2/1000:.0f}k kr)')
    
    if contour_levels:
        cs = ax.contour(rates, speeds_knots, gc_r5 / 1000,
                        levels=sorted(contour_levels), colors='black',
                        linewidths=2.5, linestyles='--')
        ax.clabel(cs, inline=True, fontsize=10, fmt='%.0fk kr')
    
    # Marker nominell verdi (7 kr/km, 18 knop)
    ax.plot(7, 18, 'k*', markersize=20, markerfacecolor='yellow',
            markeredgewidth=2, label='Basisscenario\n(7 kr/km, 18 knop)')
    
    ax.set_xlabel('Sjøfraktrate (kr/km)', fontsize=11)
    ax.set_ylabel('Havgående hastighet (knop)', fontsize=11)
    ax.set_title('Break-even-analyse: E[GC] for R5 (Bar-ruten)\n'
                 'Konturlinjer viser hvor R5 matcher konkurrentene',
                 fontsize=12)
    ax.legend(loc='upper right', framealpha=0.95)
    plt.tight_layout()
    fig.savefig(out_dir / 'fig_breakeven_heatmap.png', dpi=150)
    plt.close(fig)
    print(f"→ Lagret 'fig_breakeven_heatmap.png'")
    
    # ====== FIGUR 2: TO 1D-PLOTT VED NOMINELL VERDI ======
    fig, axes = plt.subplots(1, 2, figsize=(13, 5))
    
    # Plot 1: GC vs rate ved 18 knop (nominell hastighet)
    nominal_speed_idx = np.argmin(np.abs(speeds_knots - 18))
    gc_vs_rate = gc_r5[nominal_speed_idx, :]
    
    ax = axes[0]
    ax.plot(rates, gc_vs_rate / 1000, 'b-', linewidth=2.5,
            label='R5 (Bar-ruten)')
    
    # Horisontale linjer for andre ruter
    colors_others = {'R1': '#1f77b4', 'R2': '#ff7f0e', 'R3': '#2ca02c',
                     'R4': '#d62728', 'R6': '#8c564b'}
    for rt, gc in gc_others.items():
        ax.axhline(gc / 1000, color=colors_others.get(rt, 'gray'),
                   linestyle=':', alpha=0.7, label=f'{rt}: {gc/1000:.0f}k kr')
    
    ax.axvline(7, color='gray', linestyle=':', alpha=0.5)
    ax.text(7.1, ax.get_ylim()[1] * 0.95, 'Basis: 7 kr/km',
            fontsize=9, color='gray')
    
    ax.set_xlabel('Sjøfraktrate (kr/km)')
    ax.set_ylabel('E[GC] (1000 kr)')
    ax.set_title(f'R5 GC vs sjøfraktrate (ved 18 knop)')
    ax.legend(loc='best', fontsize=9, framealpha=0.9)
    ax.grid(True, alpha=0.3)
    
    # Plot 2: GC vs hastighet ved 7 kr/km (nominell rate)
    nominal_rate_idx = np.argmin(np.abs(rates - 7))
    gc_vs_speed = gc_r5[:, nominal_rate_idx]
    
    ax = axes[1]
    ax.plot(speeds_knots, gc_vs_speed / 1000, 'b-', linewidth=2.5,
            label='R5 (Bar-ruten)')
    
    for rt, gc in gc_others.items():
        ax.axhline(gc / 1000, color=colors_others.get(rt, 'gray'),
                   linestyle=':', alpha=0.7, label=f'{rt}: {gc/1000:.0f}k kr')
    
    ax.axvline(18, color='gray', linestyle=':', alpha=0.5)
    ax.text(18.1, ax.get_ylim()[1] * 0.95, 'Basis: 18 knop',
            fontsize=9, color='gray')
    
    ax.set_xlabel('Havgående hastighet (knop)')
    ax.set_ylabel('E[GC] (1000 kr)')
    ax.set_title('R5 GC vs sjøhastighet (ved 7 kr/km)')
    ax.legend(loc='best', fontsize=9, framealpha=0.9)
    ax.grid(True, alpha=0.3)
    
    plt.tight_layout()
    fig.savefig(out_dir / 'fig_breakeven_1d.png', dpi=150)
    plt.close(fig)
    print(f"→ Lagret 'fig_breakeven_1d.png'")


# ==============================================================================
# DEL 2: BLOCK-TRAIN SCENARIO (R1 OG R3)
# ==============================================================================

def blocktrain_scenario(excel_path, n_mc=10000, seed=42):
    """
    Kjører tre scenarier for R1 og R3:
      - Basis:        Som nå (10 000 MC-iterasjoner med dagens parametere)
      - Tolkning 3:   Konkrete endringer (hovedscenario)
      - Tolkning 1:   Halver bare ventetid (robusthetssjekk)
    
    Returnerer dict med resultater per scenario per rute.
    """
    print("\n" + "=" * 70)
    print("BLOCK-TRAIN SCENARIO-ANALYSE")
    print("=" * 70)
    
    segments = read_segments(excel_path)
    params = read_stochastic_params(excel_path)
    costs = read_cost_params(excel_path)
    
    results = {'basis': {}, 'tolkning3': {}, 'tolkning1': {}}
    
    target_routes = ['R1', 'R3']
    
    # ----- Basis -----
    print("\n--- Scenario: BASIS (uendret) ---")
    for rt in target_routes:
        rng = np.random.default_rng(seed)
        res = simulate_route(rt, segments, params, costs, rng, n_mc,
                             include_disrupt=False)
        results['basis'][rt] = res
        print(f"  {rt}: E[GC] = {res['gc'].mean():,.0f} kr, "
              f"E[t] = {res['mean_time']/24:.1f} d, "
              f"σ(t) = {res['std_time']:.1f} t")
    
    # ----- Tolkning 3: Konkrete endringer -----
    print("\n--- Scenario: TOLKNING 3 (hovedscenario, konkrete endringer) ---")
    segs_t3 = segments.copy()
    # Cast til float for å unngå dtype-konflikt
    for col in ['Headway_t', 'Terminaltid_t', 'Ventetid_t', 'Transporttid_t',
                'Grenseforsinkelse_t', 'Omlastninger']:
        segs_t3[col] = segs_t3[col].astype(float)
    params_t3 = {k: v.copy() if hasattr(v, 'copy') else v for k, v in params.items()}
    
    # R1: Fjern Bratislava-overføring (segment 3 = truck)
    #     Headway alle rail 24->12 t
    #     Mellomterminal-tid: redusert via Terminaltid_t-kol (kan settes lavt)
    #     CV rail: 0.20 -> 0.15
    
    # Fjern R1 seg 3 (Bratislava terminal-overføring) ved å nullstille bidrag
    mask_r1_seg3 = (segs_t3['RuteID'] == 'R1') & (segs_t3['Segment'] == 3)
    segs_t3.loc[mask_r1_seg3, 'Transporttid_t'] = 0
    segs_t3.loc[mask_r1_seg3, 'Terminaltid_t'] = 0
    segs_t3.loc[mask_r1_seg3, 'Ventetid_t'] = 0
    segs_t3.loc[mask_r1_seg3, 'Headway_t'] = 0
    segs_t3.loc[mask_r1_seg3, 'Omlastninger'] = 0  # Fjerner terminalkostnad
    segs_t3.loc[mask_r1_seg3, 'Grenseforsinkelse_t'] = 0
    
    # R1 rail-segmenter: 24t -> 12t headway, og halver mellomterminal-tid
    r1_rail_mask = (segs_t3['RuteID'] == 'R1') & (segs_t3['Mode'].astype(str).str.lower().str.startswith('rail'))
    segs_t3.loc[r1_rail_mask, 'Headway_t'] = 12
    
    # Mellomterminaler (alle rail-segmenter unntatt siste) får halvert terminaltid
    r1_rail_rows = segs_t3[r1_rail_mask].index.tolist()
    for idx in r1_rail_rows[:-1]:  # Alle unntatt siste
        old_term = segs_t3.loc[idx, 'Terminaltid_t']
        segs_t3.loc[idx, 'Terminaltid_t'] = old_term / 3  # 6 -> 2 t
    
    # R3: Rotterdam-München 84->24, München-Lj 48->24, mellomterminal-tid halvert
    mask_r3_s2 = (segs_t3['RuteID'] == 'R3') & (segs_t3['Segment'] == 2)
    mask_r3_s3 = (segs_t3['RuteID'] == 'R3') & (segs_t3['Segment'] == 3)
    segs_t3.loc[mask_r3_s2, 'Headway_t'] = 24
    segs_t3.loc[mask_r3_s3, 'Headway_t'] = 24
    
    r3_rail_mask = (segs_t3['RuteID'] == 'R3') & (segs_t3['Mode'].astype(str).str.lower().str.startswith('rail'))
    r3_rail_rows = segs_t3[r3_rail_mask].index.tolist()
    for idx in r3_rail_rows[:-1]:
        old_term = segs_t3.loc[idx, 'Terminaltid_t']
        segs_t3.loc[idx, 'Terminaltid_t'] = old_term / 3
    
    # CV rail: 0.20 -> 0.15
    if 'Rail' in params_t3['cv'].index:
        params_t3['cv'].loc['Rail', 'CV'] = 0.15
    
    for rt in target_routes:
        rng = np.random.default_rng(seed)
        res = simulate_route(rt, segs_t3, params_t3, costs, rng, n_mc,
                             include_disrupt=False)
        results['tolkning3'][rt] = res
        basis_gc = results['basis'][rt]['gc'].mean()
        new_gc = res['gc'].mean()
        diff = new_gc - basis_gc
        print(f"  {rt}: E[GC] = {new_gc:,.0f} kr  ({diff:+,.0f} fra basis), "
              f"E[t] = {res['mean_time']/24:.1f} d, "
              f"σ(t) = {res['std_time']:.1f} t")
    
    # ----- Tolkning 1: Bare halver ventetid -----
    print("\n--- Scenario: TOLKNING 1 (robusthetssjekk: halver ventetid) ---")
    segs_t1 = segments.copy()
    segs_t1['Headway_t'] = segs_t1['Headway_t'].astype(float)
    
    # Halver headway for alle rail-segmenter i R1 og R3
    for rt in target_routes:
        rt_rail_mask = (segs_t1['RuteID'] == rt) & (segs_t1['Mode'].astype(str).str.lower().str.startswith('rail'))
        # Halver, ikke null - 0 betyr "ingen ventetid trekkes"
        old_hw = segs_t1.loc[rt_rail_mask, 'Headway_t']
        segs_t1.loc[rt_rail_mask, 'Headway_t'] = old_hw / 2
    
    for rt in target_routes:
        rng = np.random.default_rng(seed)
        res = simulate_route(rt, segs_t1, params, costs, rng, n_mc,
                             include_disrupt=False)
        results['tolkning1'][rt] = res
        basis_gc = results['basis'][rt]['gc'].mean()
        new_gc = res['gc'].mean()
        diff = new_gc - basis_gc
        print(f"  {rt}: E[GC] = {new_gc:,.0f} kr  ({diff:+,.0f} fra basis), "
              f"E[t] = {res['mean_time']/24:.1f} d, "
              f"σ(t) = {res['std_time']:.1f} t")
    
    return results


def plot_blocktrain(results, out_dir=Path('.')):
    """
    Lager søylediagram som viser basis vs Tolkning 3 vs Tolkning 1
    for R1 og R3, med pålitelighetspremie.
    """
    out_dir = Path(out_dir)
    
    routes = list(results['basis'].keys())
    scenarios = ['basis', 'tolkning3', 'tolkning1']
    scenario_labels = ['Basis (nå)', 'Block-train (hoved)', 'Halv ventetid (robust)']
    
    # Beregn komponenter for hver
    data = {sc: {'direct': [], 'terminal': [], 'time': [], 'reliability': [], 'total': []}
            for sc in scenarios}
    
    for sc in scenarios:
        for rt in routes:
            r = results[sc][rt]
            data[sc]['direct'].append(r['direct_cost'])
            data[sc]['terminal'].append(r['terminal_cost'])
            data[sc]['time'].append(r['alpha'] * r['mean_time'])
            data[sc]['reliability'].append(r['reliability_premium'])
            data[sc]['total'].append(r['gc_aggregated'])
    
    n_routes = len(routes)
    x = np.arange(n_routes)
    width = 0.27
    
    fig, ax = plt.subplots(figsize=(11, 6.5))
    
    colors_comp = {'direct': '#1f77b4', 'terminal': '#ff7f0e',
                   'time': '#2ca02c', 'reliability': '#d62728'}
    comp_labels = {'direct': 'Direkte transport', 'terminal': 'Terminalkostnad',
                   'time': 'Tidskostnad (α·E[t])', 'reliability': 'Pålitelighet (β·σ(t))'}
    
    for i, sc in enumerate(scenarios):
        offset = (i - 1) * width
        bottom = np.zeros(n_routes)
        for comp in ['direct', 'terminal', 'time', 'reliability']:
            vals = np.array(data[sc][comp])
            label = comp_labels[comp] if i == 0 else None  # Bare en gang i legend
            ax.bar(x + offset, vals, width, bottom=bottom,
                   color=colors_comp[comp], edgecolor='black',
                   linewidth=0.7, label=label)
            bottom += vals
        # Total over hver søyle
        for j, total in enumerate(data[sc]['total']):
            ax.text(j + offset, total + 1500, f'{total/1000:.0f}k',
                    ha='center', fontsize=8.5, fontweight='bold')
    
    # X-akse: gruppert per rute med scenario-labels
    ax.set_xticks(x)
    ax.set_xticklabels(routes, fontsize=11)
    
    # Sekundære labels under: scenario-navn ved hver søyle
    for i, sc_label in enumerate(scenario_labels):
        offset = (i - 1) * width
        for j in range(n_routes):
            ax.text(j + offset, -7500, sc_label, ha='center',
                    fontsize=7.5, rotation=20)
    
    ax.set_ylabel('GC med pålitelighetspremie (kr)', fontsize=11)
    ax.set_title('Block-train scenario for R1 og R3:\n'
                 'Sammenligning av basis, hovedscenario og robusthetssjekk',
                 fontsize=12)
    ax.legend(loc='upper right', framealpha=0.95, fontsize=9)
    ax.grid(True, axis='y', alpha=0.3)
    ax.set_axisbelow(True)
    
    # Gi plass til scenario-labels
    ax.set_ylim(bottom=-12000)
    
    plt.tight_layout()
    fig.savefig(out_dir / 'fig_blocktrain_scenario.png', dpi=150,
                bbox_inches='tight')
    plt.close(fig)
    print(f"→ Lagret 'fig_blocktrain_scenario.png'")


# ==============================================================================
# MAIN
# ==============================================================================

def main():
    parser = argparse.ArgumentParser()
    parser.add_argument('--excel', type=Path,
                        default=Path('model_v6.xlsx'))
    parser.add_argument('--n', type=int, default=2000,
                        help='MC-iterasjoner for break-even (lavere = raskere)')
    parser.add_argument('--n-block', type=int, default=10000,
                        help='MC-iterasjoner for block-train')
    parser.add_argument('--grid', type=int, default=15,
                        help='Nettstørrelse for break-even (15x15 = 225 punkter)')
    parser.add_argument('--out', type=Path, default=Path('.'))
    args = parser.parse_args()
    
    # ----- Break-even -----
    rates, speeds, gc_r5, gc_others = breakeven_sea_route(
        args.excel, n_mc=args.n, n_grid=args.grid
    )
    plot_breakeven(rates, speeds, gc_r5, gc_others, args.out)
    
    # ----- Block-train -----
    results_bt = blocktrain_scenario(args.excel, n_mc=args.n_block)
    plot_blocktrain(results_bt, args.out)
    
    print("\n" + "=" * 70)
    print("Scenarioanalyse ferdig.")
    print("=" * 70)


if __name__ == '__main__':
    main()
