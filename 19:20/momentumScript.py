import pandas as pd
import random

def compute_empirical_prob(sequence, k, target="W"):
    indices = []
    for i in range(k, len(sequence)):
        if all(sequence[i - j - 1] == target for j in range(k)):
            indices.append(i)
    if not indices:
        return None, 0
    return sum(1 for i in indices if sequence[i] == target) / len(indices), len(indices)

def simulate_bias(p, n, k, target="W", simulations=5000):
    empirical_probs = []
    choices = [target, "D" if target == "W" else "W", "L" if target != "L" else "D"]
    for _ in range(simulations):
        seq = [random.choices(choices, weights=[p, (1-p)/2, (1-p)/2])[0] for _ in range(n)]
        prob, _ = compute_empirical_prob(seq, k, target)
        if prob is not None:
            empirical_probs.append(prob)
    return sum(empirical_probs) / len(empirical_probs) if empirical_probs else 0

def correct_for_bias(empirical, p, n, k, target="W"):
    estimated_bias = simulate_bias(p, n, k, target)
    corrected = empirical + (p - estimated_bias)
    return corrected, estimated_bias

def analyze_momentum(file_path):
    xls = pd.ExcelFile(file_path)

    # Positive momentum
    total_wins_in_momentum = 0
    total_non_wins_in_momentum = 0
    total_corrected_wins_in_momentum = 0
    total_positive_opportunities = 0

    # Negative momentum
    total_losses_in_momentum = 0
    total_non_losses_in_momentum = 0
    total_corrected_losses_in_momentum = 0
    total_negative_opportunities = 0

    for sheet_name in xls.sheet_names:
        df = pd.read_excel(file_path, sheet_name=sheet_name)

        if "W/D/L" not in df.columns:
            continue

        results = df["W/D/L"].dropna().tolist()
        n = len(results)

        # --- Positive Momentum (after 2 wins) ---
        emp_win_prob, num_pos = compute_empirical_prob(results, k=2, target="W")
        if emp_win_prob is not None:
            p_w = results.count("W") / n if n > 0 else 0.5
            corrected_win_prob, est_win_bias = correct_for_bias(emp_win_prob, p_w, n, k=2, target="W")
            corrected_wins = corrected_win_prob * num_pos

            total_corrected_wins_in_momentum += corrected_wins
            total_positive_opportunities += num_pos

            # Count actual wins/non-wins
            wins, non_wins = 0, 0
            for i in range(2, n):
                if results[i-1] == "W" and results[i-2] == "W":
                    if results[i] == "W":
                        wins += 1
                    else:
                        non_wins += 1
            total_wins_in_momentum += wins
            total_non_wins_in_momentum += non_wins

            print(f"🔥 {sheet_name} — Positive Momentum:")
            print(f"    Empirical: {emp_win_prob:.3f}, Corrected: {corrected_win_prob:.3f}, Bias: {p_w - est_win_bias:.3f}")

        # --- Negative Momentum (after 2 losses) ---
        emp_loss_prob, num_neg = compute_empirical_prob(results, k=2, target="L")
        if emp_loss_prob is not None:
            p_l = results.count("L") / n if n > 0 else 0.5
            corrected_loss_prob, est_loss_bias = correct_for_bias(emp_loss_prob, p_l, n, k=2, target="L")
            corrected_losses = corrected_loss_prob * num_neg

            total_corrected_losses_in_momentum += corrected_losses
            total_negative_opportunities += num_neg

            # Count actual losses/non-losses
            losses, non_losses = 0, 0
            for i in range(2, n):
                if results[i-1] == "L" and results[i-2] == "L":
                    if results[i] == "L":
                        losses += 1
                    else:
                        non_losses += 1
            total_losses_in_momentum += losses
            total_non_losses_in_momentum += non_losses

            print(f"❄️ {sheet_name} — Negative Momentum:")
            print(f"    Empirical: {emp_loss_prob:.3f}, Corrected: {corrected_loss_prob:.3f}, Bias: {p_l - est_loss_bias:.3f}")

    # Final Summaries
    win_bias_total = total_corrected_wins_in_momentum - total_wins_in_momentum
    win_bias_avg = win_bias_total / total_positive_opportunities if total_positive_opportunities else 0
    loss_bias_total = total_corrected_losses_in_momentum - total_losses_in_momentum
    loss_bias_avg = loss_bias_total / total_negative_opportunities if total_negative_opportunities else 0

    print("\n✅ Final Summary — Positive Momentum:")
    print(f"Total Wins in Momentum: {total_wins_in_momentum}")
    print(f"Total Non-Wins in Momentum: {total_non_wins_in_momentum}")
    print(f"Positive Momentum Bias (Corrected - Actual Wins): {win_bias_total:.2f}")
    print(f"Average Bias Per Opportunity: {win_bias_avg:.3f}")

    print("\n🧊 Final Summary — Negative Momentum:")
    print(f"Total Losses in Negative Momentum: {total_losses_in_momentum}")
    print(f"Total Non-Losses in Negative Momentum: {total_non_losses_in_momentum}")
    print(f"Negative Momentum Bias (Corrected - Actual Losses): {loss_bias_total:.2f}")
    print(f"Average Bias Per Opportunity: {loss_bias_avg:.3f}")

    # Example usage:
file_path = "PLteamsData19:20.xlsx"  # Update with the correct path
analyze_momentum(file_path)