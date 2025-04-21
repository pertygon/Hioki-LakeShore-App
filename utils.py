# utils.py
import os

def save_results(df, filename, out_folder="results"):
    os.makedirs(out_folder, exist_ok=True)
    path = os.path.join(out_folder, filename)
    df.to_csv(path, index=False)
    print(f"[UTILS] Zapisano: {path}")
