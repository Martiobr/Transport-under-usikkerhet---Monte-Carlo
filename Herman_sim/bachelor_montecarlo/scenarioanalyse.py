"""
Scenarioanalyse og følsomhetsanalyse for bachelor-modellen.
Bygger på monte_carlo_v6.py — importerer alle nødvendige funksjoner.

Kjøres etter at hovedanalysen er kjørt:
    python3 scenarioanalyse.py --excel transportmodell.xlsx
"""
import argparse
from pathlib import Path
import numpy as np
import pandas as pd
import matplotlib.pyplot as plt

from monte_carlo import (
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
                   linestyle='--', linewidth=2.2, alpha=0.92, label=f'{rt}: {gc/1000:.0f}k kr')

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
                   linestyle='--', linewidth=2.2, alpha=0.92, label=f'{rt}: {gc/1000:.0f}k kr')

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
    
    results = {'basis': {}, 'tolkning3': {}, 'tolkning1': {}, 'alle_basis': {}}

    all_route_ids = sorted(segments['RuteID'].unique())
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
    
    # ----- Alle ruter ved basis (for sammenligning i figur) -----
    print("\n--- Basis for alle ruter (kontekst) ---")
    for rt in all_route_ids:
        rng = np.random.default_rng(seed)
        res = simulate_route(rt, segments, params, costs, rng, n_mc,
                             include_disrupt=False)
        results['alle_basis'][rt] = res
        print(f"  {rt}: E[GC] = {res['gc'].mean():,.0f} kr, "
              f"E[t] = {res['mean_time']/24:.1f} d")

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


def plot_blocktrain_ledetid(results, out_dir=Path('.')):
    """
    Grouped bar chart: E[t] i dager for R1 og R3 på tvers av tre scenarier.
    Feilstolper viser ±1 std. Prosentvis endring annoteres over.
    """
    out_dir = Path(out_dir)
    routes = list(results['basis'].keys())
    scenarios    = ['basis',      'tolkning3',          'tolkning1']
    sc_labels    = ['Basis (nå)', 'Block-train (hoved)', 'Halv ventetid (robust)']
    sc_colors    = ['#d4aaee',    '#2171b5',             '#6baed6']

    n_routes = len(routes)
    x = np.arange(n_routes)
    width = 0.27

    fig, ax = plt.subplots(figsize=(9, 6))

    for i, (sc, label, color) in enumerate(zip(scenarios, sc_labels, sc_colors)):
        offset = (i - 1) * width
        mean_d = np.array([results[sc][rt]['mean_time'] / 24 for rt in routes])
        std_d  = np.array([results[sc][rt]['std_time']  / 24 for rt in routes])
        ax.bar(x + offset, mean_d, width, label=label,
               color=color, edgecolor='black', linewidth=0.7)
        ax.errorbar(x + offset, mean_d, yerr=std_d,
                    fmt='none', color='black', capsize=4, linewidth=1)
        for j, (md, sd) in enumerate(zip(mean_d, std_d)):
            ax.text(j + offset, md + sd + 0.25, f'{md:.1f}d',
                    ha='center', fontsize=8.5, fontweight='bold')

    # Prosentvis endring tolkning3 vs basis
    for j, rt in enumerate(routes):
        basis_d = results['basis'][rt]['mean_time'] / 24
        new_d   = results['tolkning3'][rt]['mean_time'] / 24
        std_b   = results['basis'][rt]['std_time'] / 24
        pct = (new_d - basis_d) / basis_d * 100
        y_ann = basis_d + std_b + 2.2
        color_ann = '#2ca02c' if pct < 0 else '#d62728'
        ax.annotate(f'{pct:+.0f}%', xy=(j, y_ann), ha='center',
                    fontsize=11, color=color_ann, fontweight='bold')

    ax.set_xticks(x)
    ax.set_xticklabels(routes, fontsize=12)
    ax.set_ylabel('Forventet ledetid (dager)', fontsize=11)
    ax.set_title('Block-train scenario: Endring i ledetid for R1 og R3', fontsize=12)
    ax.legend(loc='upper right', framealpha=0.95)
    ax.grid(True, axis='y', alpha=0.3)
    ax.set_axisbelow(True)
    ax.set_ylim(bottom=0)
    plt.tight_layout()
    fig.savefig(out_dir / 'fig_blocktrain_ledetid.png', dpi=150)
    plt.close(fig)
    print("→ Lagret 'fig_blocktrain_ledetid.png'")


# Delte konstanter for sammenligningsfigurene
_ROUTE_COLORS = {
    'R1': '#1f77b4', 'R2': '#ff7f0e', 'R3': '#2ca02c',
    'R4': '#d62728', 'R5': '#9467bd', 'R6': '#8c564b',
}
_BLOCK_COLOR = '#FFD700'
_ROUTE_NAMES = {
    'R1': 'R1: Rostock', 'R2': 'R2: Hamburg', 'R3': 'R3: Rotterdam-München',
    'R4': 'R4: Rotterdam-Wien', 'R5': 'R5: Bar (sjø)', 'R6': 'R6: Rotterdam-truck',
}


def _build_rows(results, include_routes):
    """Bygg bar-data for de gitte rutene. Ruter i target_ids får par (basis + block-train)."""
    alle       = results['alle_basis']
    target_ids = set(results['basis'].keys())
    rows = []
    for rt in [r for r in sorted(alle.keys()) if r in include_routes]:
        if rt in target_ids:
            rows.append({
                'label': f'{rt}\nbasis',
                'gc': results['basis'][rt]['gc_aggregated'],
                'lt': results['basis'][rt]['mean_time'] / 24,
                'color': _ROUTE_COLORS.get(rt, '#aaaaaa'),
                'hatch': '',
            })
            rows.append({
                'label': f'{rt}\nblock-train',
                'gc': results['tolkning3'][rt]['gc_aggregated'],
                'lt': results['tolkning3'][rt]['mean_time'] / 24,
                'color': _BLOCK_COLOR,
                'hatch': '//',
            })
        else:
            rows.append({
                'label': rt,
                'gc': alle[rt]['gc_aggregated'],
                'lt': alle[rt]['mean_time'] / 24,
                'color': _ROUTE_COLORS.get(rt, '#aaaaaa'),
                'hatch': '',
            })
    return rows


def _plot_comparison(rows, val_key, ylabel, title, out_path, include_routes):
    """Render én sammenligningsfigur for enten GC eller ledetid."""
    from matplotlib.patches import Patch
    vals   = [r[val_key] for r in rows]
    is_gc  = val_key == 'gc'
    fmt_fn = (lambda v: f'{v:,.0f}') if is_gc else (lambda v: f'{v:.1f}d')

    fig, ax = plt.subplots(figsize=(10, 6))
    xs   = np.arange(len(rows))
    bars = ax.bar(xs, vals, color=[r['color'] for r in rows],
                  edgecolor='black', linewidth=0.8)
    for bar, row in zip(bars, rows):
        bar.set_hatch(row['hatch'])
    for k, val in enumerate(vals):
        ax.text(k, val + max(vals) * 0.012, fmt_fn(val),
                ha='center', va='bottom', fontsize=8.5, fontweight='bold', rotation=0)
    ax.set_xticks(xs)
    ax.set_xticklabels([r['label'] for r in rows], fontsize=9)
    ax.set_ylabel(ylabel, fontsize=11)
    ax.set_title(title, fontsize=12)
    ax.grid(True, axis='y', alpha=0.3)
    ax.set_axisbelow(True)
    ax.set_ylim(bottom=0, top=max(vals) * 1.25)

    # Én patch per rute med riktig rutefarge + gull for block-train
    legend_elements = [
        Patch(facecolor=_ROUTE_COLORS[rt], label=_ROUTE_NAMES.get(rt, rt))
        for rt in ['R1', 'R2', 'R3', 'R4', 'R5', 'R6']
        if rt in include_routes
    ]
    legend_elements.append(
        Patch(facecolor=_BLOCK_COLOR, hatch='//', edgecolor='black',
              label='Block-train (hoved)')
    )
    ax.legend(handles=legend_elements, loc='upper right', framealpha=0.95, fontsize=9)

    plt.tight_layout()
    fig.savefig(out_path, dpi=150)
    plt.close(fig)
    print(f"→ Lagret '{Path(out_path).name}'")


def plot_blocktrain_vs_ruter_gc(results, out_dir=Path('.')):
    """GC for alle ruter (R1–R6) med block-train-alternativ for R1 og R3."""
    include = {'R1', 'R2', 'R3', 'R4', 'R5', 'R6'}
    _plot_comparison(
        _build_rows(results, include), 'gc',
        'GC m/pålitelighetspremie (kr)', 'GC pr rute — alle ruter vs block-train (m/pålitelighetspremie)',
        Path(out_dir) / 'fig_blocktrain_vs_ruter_gc.png', include,
    )


def plot_blocktrain_vs_ruter_ledetid(results, out_dir=Path('.')):
    """Ledetid for alle ruter (R1–R6) med block-train-alternativ for R1 og R3."""
    include = {'R1', 'R2', 'R3', 'R4', 'R5', 'R6'}
    _plot_comparison(
        _build_rows(results, include), 'lt',
        'E[t] (dager)', 'Ledetid pr rute — alle ruter vs block-train',
        Path(out_dir) / 'fig_blocktrain_vs_ruter_ledetid.png', include,
    )


def plot_blocktrain_vs_jernbane_gc(results, out_dir=Path('.')):
    """GC for jernbaneruter R1–R4 med block-train-alternativ for R1 og R3."""
    include = {'R1', 'R2', 'R3', 'R4'}
    _plot_comparison(
        _build_rows(results, include), 'gc',
        'GC m/pålitelighetspremie (kr)', 'GC — jernbaneruter R1–R4 vs block-train (m/pålitelighetspremie)',
        Path(out_dir) / 'fig_blocktrain_vs_jernbane_gc.png', include,
    )


def plot_blocktrain_vs_jernbane_ledetid(results, out_dir=Path('.')):
    """Ledetid for jernbaneruter R1–R4 med block-train-alternativ for R1 og R3."""
    include = {'R1', 'R2', 'R3', 'R4'}
    _plot_comparison(
        _build_rows(results, include), 'lt',
        'E[t] (dager)', 'Ledetid — jernbaneruter R1–R4 vs block-train',
        Path(out_dir) / 'fig_blocktrain_vs_jernbane_ledetid.png', include,
    ) 


# ==============================================================================
# DEL 3: TRADE-OFF ANALYSE — R2, R3 BLOCK-TRAIN, R6
# ==============================================================================

def topptruter_analyse(bt_results):
    """
    Samler simuleringsresultater for de tre beste rutene:
      - R2  (Hamburg, jernbane basis)
      - R3  (Rotterdam-München, block-train tolkning 3)
      - R6  (Rotterdam-truck, dagens løsning)
    Returnerer liste av result-dicts med display_name satt.
    """
    return [
        {**bt_results['alle_basis']['R2'],
         'display_name': 'R2: Hamburg (jernbane)'},
        {**bt_results['tolkning3']['R3'],
         'display_name': 'R3: Block-train'},
        {**bt_results['alle_basis']['R6'],
         'display_name': 'R6: Rotterdam-truck'},
    ]


def plot_topptruter(top, out_dir=Path('.')):
    """
    Fire figurer for trade-off-analysen mellom R2, R3 block-train og R6:
      1. KDE-fordeling av ledetid
      2. KDE-fordeling av GC
      3. Risk-return scatter (ledetid og GC, 2-panel)
      4. GC-komponenter stacked bar
    """
    from scipy.stats import gaussian_kde
    out_dir = Path(out_dir)
    # R2=oransje, R3=grønn, R6=brun — samme farger som i hoveddataene
    colors = ['#ff7f0e', '#2ca02c', '#8c564b']

    # ------------------------------------------------------------------
    # FIGUR 1: KDE-fordeling av ledetid
    # ------------------------------------------------------------------
    fig, ax = plt.subplots(figsize=(10, 6))
    all_days = np.concatenate([r['time'] / 24 for r in top])
    x_grid = np.linspace(all_days.min(), all_days.max(), 500)
    for r, color in zip(top, colors):
        days = r['time'] / 24
        kde  = gaussian_kde(days)
        ax.plot(x_grid, kde(x_grid), color=color, linewidth=2.5,
                label=r['display_name'])
        ax.fill_between(x_grid, kde(x_grid), alpha=0.12, color=color)
        ax.axvline(days.mean(), color=color, linestyle='--',
                   linewidth=1.5, alpha=0.8)
        p10, p90 = np.percentile(days, [10, 90])
        ax.axvspan(p10, p90, alpha=0.06, color=color)
    ax.set_xlabel('Total ledetid (dager)', fontsize=11)
    ax.set_ylabel('Sannsynlighetstetthet', fontsize=11)
    ax.set_title('Fordeling av ledetid — R2, R3 block-train, R6', fontsize=12)
    ax.legend(fontsize=10, framealpha=0.95)
    ax.grid(alpha=0.3)
    plt.tight_layout()
    fig.savefig(out_dir / 'fig_topp3_ledetid_kde.png', dpi=150)
    plt.close(fig)
    print("→ Lagret 'fig_topp3_ledetid_kde.png'")

    # ------------------------------------------------------------------
    # FIGUR 2: KDE-fordeling av GC (med pålitelighetspremie)
    # ------------------------------------------------------------------
    fig, ax = plt.subplots(figsize=(10, 6))
    # Legg til pålitelighetspremien (fast tillegg) på fordelingen
    gc_with_rel = [r['gc'] + r['reliability_premium'] for r in top]
    all_gc = np.concatenate(gc_with_rel)
    x_grid = np.linspace(all_gc.min(), all_gc.max(), 500)
    for gc_arr, r, color in zip(gc_with_rel, top, colors):
        kde = gaussian_kde(gc_arr)
        ax.plot(x_grid, kde(x_grid), color=color, linewidth=2.5,
                label=r['display_name'])
        ax.fill_between(x_grid, kde(x_grid), alpha=0.12, color=color)
        ax.axvline(r['gc_aggregated'], color=color, linestyle='--',
                   linewidth=1.5, alpha=0.8)
    ax.set_xlabel('GC m/pålitelighetspremie (kr)', fontsize=11)
    ax.set_ylabel('Sannsynlighetstetthet', fontsize=11)
    ax.set_title('Fordeling av GC m/pålitelighetspremie — R2, R3 block-train, R6', fontsize=12)
    ax.legend(fontsize=10, framealpha=0.95)
    ax.grid(alpha=0.3)
    plt.tight_layout()
    fig.savefig(out_dir / 'fig_topp3_gc_kde.png', dpi=150)
    plt.close(fig)
    print("→ Lagret 'fig_topp3_gc_kde.png'")

    # ------------------------------------------------------------------
    # FIGUR 3: Risk-return scatter — 2-panel (ledetid + GC)
    # ------------------------------------------------------------------
    fig, axes = plt.subplots(1, 2, figsize=(13, 6))

    panels = [
        (axes[0],
         lambda r: r['mean_time'] / 24, lambda r: r['std_time'] / 24,
         'E[t] (dager)', 'σ(t) (dager)', 'Ledetid: forventet vs. variasjon'),
        (axes[1],
         lambda r: r['gc_aggregated'], lambda r: r['gc'].std(),
         'GC m/pålitelighetspremie (kr)', 'σ(GC) (kr)', 'GC m/pålitelighetspremie: forventet vs. variasjon'),
    ]
    for ax, x_fn, y_fn, xlabel, ylabel, title in panels:
        for r, color in zip(top, colors):
            xv, yv = x_fn(r), y_fn(r)
            ax.scatter(xv, yv, s=220, color=color, edgecolor='black',
                       zorder=3, label=r['display_name'])
            ax.annotate(r['display_name'], (xv, yv), xytext=(9, 5),
                        textcoords='offset points', fontsize=10, fontweight='bold')
        ax.set_xlabel(xlabel, fontsize=11)
        ax.set_ylabel(ylabel, fontsize=11)
        ax.set_title(title, fontsize=11)
        ax.legend(fontsize=9, framealpha=0.95)
        ax.grid(alpha=0.3)

    fig.suptitle('Risk-return analyse — R2, R3 block-train, R6', fontsize=12)
    plt.tight_layout()
    fig.savefig(out_dir / 'fig_topp3_risk_return.png', dpi=150)
    plt.close(fig)
    print("→ Lagret 'fig_topp3_risk_return.png'")

    # ------------------------------------------------------------------
    # FIGUR 4: GC-komponenter stacked bar
    # ------------------------------------------------------------------
    fig, ax = plt.subplots(figsize=(8, 6))
    route_names = [r['display_name'] for r in top]
    comp_colors = ['#4472C4', '#ED7D31', '#A5A5A5', '#FFC000']
    comp_labels = ['Direkte transport', 'Terminalkostnad',
                   'Tidskostnad (α·E[t])', 'Pålitelighetspremie (β·σ(t))']
    layers = [
        [r['direct_cost']            for r in top],
        [r['terminal_cost']          for r in top],
        [r['alpha'] * r['mean_time'] for r in top],
        [r['reliability_premium']    for r in top],
    ]
    bottom = np.zeros(len(top))
    for vals, color, label in zip(layers, comp_colors, comp_labels):
        vals_arr = np.array(vals)
        ax.bar(route_names, vals_arr, bottom=bottom,
               color=color, edgecolor='black', label=label)
        bottom += vals_arr
    for i, r in enumerate(top):
        total = r['gc_aggregated']
        ax.text(i, total + bottom.max() * 0.01, f'{total:,.0f} kr',
                ha='center', va='bottom', fontsize=10, fontweight='bold')
    ax.set_ylabel('GC (kr)', fontsize=11)
    ax.set_title('GC-komponenter — R2, R3 block-train, R6', fontsize=12)
    ax.legend(loc='upper right', framealpha=0.95, fontsize=9)
    ax.grid(alpha=0.3, axis='y')
    ax.set_axisbelow(True)
    ax.set_ylim(bottom=0, top=bottom.max() * 1.15)
    ax.tick_params(axis='x', rotation=10)
    plt.tight_layout()
    fig.savefig(out_dir / 'fig_topp3_gc_komponenter.png', dpi=150)
    plt.close(fig)
    print("→ Lagret 'fig_topp3_gc_komponenter.png'")


# ==============================================================================
# MAIN
# ==============================================================================

def main():
    parser = argparse.ArgumentParser()
    parser.add_argument('--excel', type=Path,
                        default=Path('transportmodell.xlsx'))
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
    plot_blocktrain_ledetid(results_bt, args.out)
    plot_blocktrain_vs_ruter_gc(results_bt, args.out)
    plot_blocktrain_vs_ruter_ledetid(results_bt, args.out)
    plot_blocktrain_vs_jernbane_gc(results_bt, args.out)
    plot_blocktrain_vs_jernbane_ledetid(results_bt, args.out)

    # ----- Topp 3 trade-off -----
    top3 = topptruter_analyse(results_bt)
    plot_topptruter(top3, args.out)

    print("\n" + "=" * 70)
    print("Scenarioanalyse ferdig.")
    print("=" * 70)


if __name__ == '__main__':
    main()
  