import numpy as np
import pandas as pd
import matplotlib.pyplot as plt

# Reproduserbare resultater
np.random.seed(42)

# Innstillinger
N = 10000
delivery_deadline = 12
input_file = "data.xlsx"

# Les Excel-fil og hent bare ark som inneholder "Rute"
excel_file = pd.ExcelFile(input_file)
sheet_names = [sheet for sheet in excel_file.sheet_names if "Rute" in sheet]

simulation_results = {}
summary_rows = []

for sheet in sheet_names:
    # Leser arket med header på rad 2
    df = pd.read_excel(input_file, sheet_name=sheet, header=1)

    # Fjerner skjulte mellomrom i kolonnenavn
    df.columns = df.columns.str.strip()

    # Beholder bare relevante kolonner
    df = df[["Segment", "Min", "Mode", "Max"]]

    # Fjerner eventuelle tomme rader
    df = df.dropna(subset=["Min", "Mode", "Max"])

    # Simuler total ledetid for ruten
    total_time = np.zeros(N)

    for _, row in df.iterrows():
        samples = np.random.triangular(row["Min"], row["Mode"], row["Max"], N)
        total_time += samples

    simulation_results[sheet] = total_time

    # Oppsummeringsstatistikk
    summary_rows.append({
        "Rute": sheet,
        "Forventet ledetid (dager)": round(np.mean(total_time), 2),
        "Std.avvik": round(np.std(total_time), 2),
        "P10": round(np.percentile(total_time, 10), 2),
        "Median": round(np.percentile(total_time, 50), 2),
        "P90": round(np.percentile(total_time, 90), 2),
        "Sannsynlighet innen frist (%)": round(np.mean(total_time <= delivery_deadline) * 100, 1)
    })

# Lager sammendragstabell
summary_df = pd.DataFrame(summary_rows)
summary_df = summary_df.sort_values(by="Sannsynlighet innen frist (%)", ascending=False)

print("\nSammendrag av simuleringsresultater:\n")
print(summary_df.to_string(index=False))

# Lagrer sammendrag til Excel
summary_df.to_excel("resultat_sammendrag.xlsx", index=False)

# Plotter histogram for hver rute
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
    