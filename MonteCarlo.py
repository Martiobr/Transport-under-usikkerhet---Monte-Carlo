import numpy as np
import pandas as pd
import matplotlib.pyplot as plt

# Reproduserbare resultater
np.random.seed(42)

# Innstillinger
N = 10000
delivery_deadline = 12
input_file = "input_routes.xlsx"

# Les Excel
df = pd.read_excel(input_file, sheet_name="segmenter")

# Sjekk at nødvendige kolonner finnes
required_cols = ["Rute", "Segment", "Min", "Mode", "Max"]
if not all(col in df.columns for col in required_cols):
    raise ValueError(f"Excel-filen må inneholde kolonnene: {required_cols}")

simulation_results = {}
summary_rows = []

# Simuler hver rute
for route_name, group in df.groupby("Rute"):
    total_time = np.zeros(N)

    for _, row in group.iterrows():
        samples = np.random.triangular(row["Min"], row["Mode"], row["Max"], N)
        total_time += samples

    simulation_results[route_name] = total_time

    summary_rows.append({
        "Rute": route_name,
        "Forventet ledetid (dager)": round(np.mean(total_time), 2),
        "Std.avvik": round(np.std(total_time), 2),
        "P10": round(np.percentile(total_time, 10), 2),
        "Median": round(np.percentile(total_time, 50), 2),
        "P90": round(np.percentile(total_time, 90), 2),
        "Sannsynlighet innen frist (%)": round(np.mean(total_time <= delivery_deadline) * 100, 1)
    })

# Lag sammendragstabell
summary_df = pd.DataFrame(summary_rows)
summary_df = summary_df.sort_values(by="Sannsynlighet innen frist (%)", ascending=False)

print(summary_df.to_string(index=False))

# Lag histogram for hver rute
for route_name, total_time in simulation_results.items():
    plt.figure(figsize=(8, 5))
    plt.hist(total_time, bins=50, edgecolor="black")
    plt.axvline(delivery_deadline, linestyle="--", label=f"Frist = {delivery_deadline} dager")
    plt.title(f"Fordeling av total ledetid\n{route_name}")
    plt.xlabel("Total ledetid (dager)")
    plt.ylabel("Antall simuleringer")
    plt.legend()
    plt.tight_layout()
    plt.show()