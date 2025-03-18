import pandas as pd

import pandas as pd

def detect_momentum_and_update(file_path):
    # Load the Excel file
    xls = pd.ExcelFile(file_path)

    # Dictionary to store updated data for each team
    updated_sheets = {}

    # Global counters
    total_wins_in_momentum = 0
    total_non_wins_in_momentum = 0
    total_losses_in_negative_momentum = 0
    total_non_losses_in_negative_momentum = 0

    # Iterate over each sheet (team) in the Excel file
    for sheet_name in xls.sheet_names:
        df = pd.read_excel(file_path, sheet_name=sheet_name)

        # Ensure the W/D/L column exists
        if "W/D/L" not in df.columns:
            print(f"Skipping {sheet_name}: 'W/D/L' column not found.")
            continue

        # Variables to track momentum for each team
        in_momentum = False
        in_negative_momentum = False
        win_streak = 0
        lose_streak = 0
        win_in_momentum = 0  # Wins while in positive momentum
        lose_in_negative_momentum = 0  # Losses while in negative momentum
        not_win_in_momentum = 0  # Games without a win while in momentum
        not_lose_in_negative_momentum = 0  # Games without a loss while in negative momentum

        # New columns to track momentum status
        momentum_column = []
        negative_momentum_column = []

        # Iterate over the W/D/L column
        for result in df["W/D/L"].dropna():
            # Check for positive momentum (winning streak)
            if result == "W":
                win_streak += 1
                if in_momentum:
                    win_in_momentum += 1
                if win_streak == 2:  # Entering momentum
                    in_momentum = True
            else:
                if in_momentum:
                    not_win_in_momentum += 1
                in_momentum = False  # Exiting momentum
                win_streak = 0

            # Check for negative momentum (losing streak)
            if result == "L":
                lose_streak += 1
                if in_negative_momentum:
                    lose_in_negative_momentum += 1
                if lose_streak == 2:  # Entering negative momentum
                    in_negative_momentum = True
            else:
                if in_negative_momentum:
                    not_lose_in_negative_momentum += 1
                in_negative_momentum = False  # Exiting negative momentum
                lose_streak = 0

            # Mark the row as part of positive/negative momentum or not
            momentum_column.append(in_momentum)
            negative_momentum_column.append(in_negative_momentum)

        # Add the momentum columns to the DataFrame
        df["In Momentum"] = momentum_column + [False] * (len(df) - len(momentum_column))  # Fill missing values
        df["In Negative Momentum"] = negative_momentum_column + [False] * (len(df) - len(negative_momentum_column))  # Fill missing values

        # Update global counters
        total_wins_in_momentum += win_in_momentum
        total_non_wins_in_momentum += not_win_in_momentum
        total_losses_in_negative_momentum += lose_in_negative_momentum
        total_non_losses_in_negative_momentum += not_lose_in_negative_momentum

        # Create summary rows
        summary_rows = pd.DataFrame({
            "W/D/L": [
                "Total Wins in Momentum",
                "Total Non-Wins in Momentum",
                "Total Losses in Negative Momentum",
                "Total Non-Losses in Negative Momentum"
            ],
            "In Momentum": [win_in_momentum, not_win_in_momentum, None, None],
            "In Negative Momentum": [None, None, lose_in_negative_momentum, not_lose_in_negative_momentum]
        })

        # Append the summary rows to the DataFrame
        df = pd.concat([df, summary_rows], ignore_index=True)

        # Store the updated sheet
        updated_sheets[sheet_name] = df

    # Save the updated Excel file
    output_file_path = file_path.replace(".xlsx", "_updated.xlsx")  # Save as a new file
    with pd.ExcelWriter(output_file_path, engine="xlsxwriter") as writer:
        for sheet_name, df in updated_sheets.items():
            df.to_excel(writer, sheet_name=sheet_name, index=False)

    # Print final global statistics
    print("\n✅ Final Summary Across All Teams:")
    print(f"Total Wins in Momentum: {total_wins_in_momentum}")
    print(f"Total Non-Wins in Momentum: {total_non_wins_in_momentum}")
    print(f"Total Losses in Negative Momentum: {total_losses_in_negative_momentum}")
    print(f"Total Non-Losses in Negative Momentum: {total_non_losses_in_negative_momentum}")
    print(f"\nUpdated file saved as: {output_file_path}")

# Example usage:
file_path = "PLteamsData19:20.xlsx"  # Update with the correct path
detect_momentum_and_update(file_path)
