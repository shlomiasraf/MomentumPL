import pandas as pd

def analyze_tight_vs_loose(file_path):
    xls = pd.ExcelFile(file_path)
    records = []

    for sheet_name in xls.sheet_names:
        df = pd.read_excel(file_path, sheet_name=sheet_name)
        if "W/D/L" not in df.columns or "Tight Schedule" not in df.columns:
            continue

        results = df["W/D/L"].tolist()
        tight_flags = df["Tight Schedule"].tolist()

        for i in range(2, len(results)):
            prev1, prev2, curr = results[i-1], results[i-2], results[i]
            tight = tight_flags[i]

            # Positive momentum check: WW → ?
            if prev1 == "W" and prev2 == "W":
                result = "Win" if curr == "W" else "Not Win"
                records.append({
                    "Team": sheet_name,
                    "Momentum": "Positive (WW)",
                    "Tight Schedule": tight,
                    "Outcome": result
                })

            # Negative momentum check: LL → ?
            if prev1 == "L" and prev2 == "L":
                result = "Loss" if curr == "L" else "Not Loss"
                records.append({
                    "Team": sheet_name,
                    "Momentum": "Negative (LL)",
                    "Tight Schedule": tight,
                    "Outcome": result
                })

    # Convert to DataFrame and print summary
    df_records = pd.DataFrame(records)
    summary = df_records.groupby(["Momentum", "Tight Schedule", "Outcome"]).size().reset_index(name="Count")

    print("\n📊 Momentum Analysis by Tight Schedule:\n")
    for _, row in summary.iterrows():
        print(f"{row['Momentum']} | Tight={row['Tight Schedule']} | {row['Outcome']}: {row['Count']}")

# Example usage:
file_path = "PLteamsData22:23_with_tight_schedule.xlsx"
analyze_tight_vs_loose(file_path)
