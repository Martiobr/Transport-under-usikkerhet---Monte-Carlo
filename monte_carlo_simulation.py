import numpy as np
import pandas as pd
import matplotlib.pyplot as plt
import matplotlib.patches as mpatches
import os

# ── Konfigurasjon ────────────────────────────────────────────────────────────
np.random.seed(42)
N                 = 10_000
DELIVERY_DEADLINE = 12          # dager
EXCEL_FILE        = "monte_carlo_data.xlsx"

# Ark-navn i Excel-filen
SHEET_MAP = {
    "Rute 1: Karmøy - Rotterdam - Tog - Balkan": "Rute 1 - Rotterdam",
    "Rute 2: Karmøy - Rostock - Tog - Balkan":   "Rute 2 - Rostock",
}

# ── Farger per rute (brukes konsekvent i alle plott) ─────────────────────────
ROUTE_COLORS = {
    "Rute 1: Karmøy - Rotterdam - Tog - Balkan": "#2E75B6",
    "Rute 2: Karmøy - Rostock - Tog - Balkan":   "#2E8653",
}

# ── Les data fra Excel ────────────────────────────────────────────────────────
def les_ruter_fra_excel(filepath: str) -> dict:
    """
    Leser triangulærfordelingsparametre (min, mode, maks) fra Excel.
    Returnerer dict:  rute_navn -> {segment -> (min, mode, maks)}
    """
    ruter = {}
    for rute_navn, ark_navn in SHEET_MAP.items():
        df = pd.read_excel(
            filepath,
            sheet_name=ark_navn,
            header=1,           # rad 2 er kolonneoverskrifter (0-indeksert → 1)
            usecols="A:D",      # Segment | Minimum | Mest sannsynlig | Maksimum
        )
        # Fjern sum-rad og tomme rader
        df.columns = ["Segment", "Minimum", "Mode", "Maksimum"]
        df = df.dropna(subset=["Minimum", "Mode", "Maksimum"])
        df = df[~df["Segment"].astype(str).str.upper().str.startswith("TOTAL")]

        segmenter = {}
        for _, row in df.iterrows():
            segmenter[str(row["Segment"]).strip()] = (
                float(row["Minimum"]),
                float(row["Mode"]),
                float(row["Maksimum"]),
            )
        ruter[rute_navn] = segmenter
        print(f"  Leste '{ark_navn}': {len(segmenter)} segmenter")
    return ruter


# ── Simulering ────────────────────────────────────────────────────────────────
def simuler_rute(rute_data: dict, n: int) -> np.ndarray:
    """Summer triangulært-fordelte stikkprøver for alle segmenter."""
    total = np.zeros(n)
    for minimum, mode, maksimum in rute_data.values():
        total += np.random.triangular(minimum, mode, maksimum, n)
    return total


# ── Kjør ─────────────────────────────────────────────────────────────────────
if not os.path.exists(EXCEL_FILE):
    raise FileNotFoundError(
        f"Finner ikke '{EXCEL_FILE}'.  "
        "Sørg for at Excel-malen ligger i samme mappe som dette scriptet."
    )

print(f"\nLeser ruteparametere fra '{EXCEL_FILE}' ...")
routes = les_ruter_fra_excel(EXCEL_FILE)

print(f"\nKjører Monte Carlo-simulering  (N = {N:,}) ...")
simulation_results = {}
summary_rows       = []

for rute_navn, rute_data in routes.items():
    total_tider = simuler_rute(rute_data, N)
    simulation_results[rute_navn] = total_tider

    summary_rows.append({
        "Rute":                        rute_navn,
        "Forventet ledetid (dager)":   round(np.mean(total_tider), 2),
        "Std.avvik":                   round(np.std(total_tider),  2),
        "P10":                         round(np.percentile(total_tider, 10), 2),
        "Median":                      round(np.percentile(total_tider, 50), 2),
        "P90":                         round(np.percentile(total_tider, 90), 2),
        "Sannsynlighet innen frist (%)":
            round(np.mean(total_tider <= DELIVERY_DEADLINE) * 100, 1),
    })

summary_df = pd.DataFrame(summary_rows).sort_values(
    by="Sannsynlighet innen frist (%)", ascending=False
)

print("\n" + "=" * 80)
print(summary_df.to_string(index=False))
print("=" * 80)


# ── Individuelle histogram per rute ──────────────────────────────────────────
for rute_navn, total_tider in simulation_results.items():
    farge   = ROUTE_COLORS[rute_navn]
    innen   = np.mean(total_tider <= DELIVERY_DEADLINE) * 100
    p10     = np.percentile(total_tider, 10)
    p50     = np.percentile(total_tider, 50)
    p90     = np.percentile(total_tider, 90)

    fig, ax = plt.subplots(figsize=(9, 5))
    ax.hist(total_tider, bins=60, color=farge, alpha=0.75, edgecolor="white",
            linewidth=0.4, label="Simulerte ledetider")

    ax.axvline(DELIVERY_DEADLINE, color="red",    linestyle="--", linewidth=1.8,
               label=f"Leveringsfrist = {DELIVERY_DEADLINE} d")
    ax.axvline(p10,               color="orange", linestyle=":",  linewidth=1.4,
               label=f"P10 = {p10:.1f} d")
    ax.axvline(p50,               color="black",  linestyle="-",  linewidth=1.4,
               label=f"Median = {p50:.1f} d")
    ax.axvline(p90,               color="purple", linestyle=":",  linewidth=1.4,
               label=f"P90 = {p90:.1f} d")

    ax.set_title(f"Fordeling av total ledetid\n{rute_navn}", fontsize=11, pad=12)
    ax.set_xlabel("Total ledetid (dager)", fontsize=10)
    ax.set_ylabel("Antall simuleringer",   fontsize=10)
    ax.legend(fontsize=9)

    tekst = f"P(ledetid ≤ {DELIVERY_DEADLINE} d) = {innen:.1f}%"
    ax.text(0.97, 0.95, tekst, transform=ax.transAxes,
            ha="right", va="top", fontsize=10,
            bbox=dict(boxstyle="round,pad=0.4", fc="white", alpha=0.8))

    plt.tight_layout()
    plt.show()


# ── Sammenlignende plot (begge ruter overlappet) ──────────────────────────────
fig, ax = plt.subplots(figsize=(10, 5))
patches = []
for rute_navn, total_tider in simulation_results.items():
    farge = ROUTE_COLORS[rute_navn]
    ax.hist(total_tider, bins=60, alpha=0.55, color=farge, edgecolor="white",
            linewidth=0.3)
    patches.append(mpatches.Patch(color=farge, alpha=0.75, label=rute_navn))

ax.axvline(DELIVERY_DEADLINE, color="red", linestyle="--", linewidth=2,
           label=f"Leveringsfrist = {DELIVERY_DEADLINE} d")

ax.set_title("Sammenligning av ledetidsfordelinger", fontsize=12, pad=12)
ax.set_xlabel("Total ledetid (dager)", fontsize=10)
ax.set_ylabel("Antall simuleringer",   fontsize=10)
ax.legend(handles=patches + [
    plt.Line2D([0], [0], color="red", linestyle="--", linewidth=2,
               label=f"Frist = {DELIVERY_DEADLINE} d")
], fontsize=9)
plt.tight_layout()
plt.show()
