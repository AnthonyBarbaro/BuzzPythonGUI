import os
import pandas as pd

INPUT_DIR = "files"
OUTPUT_DIR = "output"
BRAND_NAME = "Turn"

def calculate_turn_inventory_cost(df: pd.DataFrame) -> pd.DataFrame:
    """
    Given a DataFrame with columns ['Brand','Product','Available','Cost'], 
    filter to BRAND_NAME and compute Total_Cost per product.
    Returns a DataFrame with ['Product','Available','Cost','Total_Cost'].
    """
    # Filter for brand (case-insensitive)
    mask = df['Brand'].astype(str).str.strip().str.lower() == BRAND_NAME.lower()
    turn_df = df.loc[mask, ['Product', 'Available', 'Cost']].copy()

    if turn_df.empty:
        return pd.DataFrame(columns=['Product','Available','Cost','Total_Cost'])

    # Coerce numeric
    turn_df['Available'] = pd.to_numeric(turn_df['Available'], errors='coerce').fillna(0).astype(int)
    turn_df['Cost']      = pd.to_numeric(turn_df['Cost'],      errors='coerce').fillna(0.0)

    # Compute
    turn_df['Total_Cost'] = turn_df['Available'] * turn_df['Cost']
    return turn_df

def process_all_files(input_dir: str, output_dir: str):
    # Ensure output folder exists
    os.makedirs(output_dir, exist_ok=True)

    for fname in os.listdir(input_dir):
        if not fname.lower().endswith(".csv"):
            continue

        path_in = os.path.join(input_dir, fname)
        try:
            df = pd.read_csv(path_in, usecols=['Brand','Product','Available','Cost'])
        except Exception as e:
            print(f"❌ Failed to read {fname}: {e}")
            continue

        summary_df = calculate_turn_inventory_cost(df)

        base, _ = os.path.splitext(fname)
        out_name = f"turn_cost_{base}.csv"
        path_out = os.path.join(output_dir, out_name)

        # Save even if empty (so you know it ran)
        summary_df.to_csv(path_out, index=False)
        total = summary_df['Total_Cost'].sum()
        print(f"✔ {fname} → {out_name}: {len(summary_df)} items, total cost ${total:.2f}")

if __name__ == "__main__":
    process_all_files(INPUT_DIR, OUTPUT_DIR)
