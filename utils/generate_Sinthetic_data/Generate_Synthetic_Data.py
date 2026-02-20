import os
import numpy as np
import pandas as pd

# ------------------------------------------------------------
# BASE DIR = project root (relative path)
# ------------------------------------------------------------
BASE_DIR = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
DATA_FOLDER = os.path.join(BASE_DIR, "data", "Lab_results", "experiment1")

N_FILES = 10
N_POINTS = 200

os.makedirs(DATA_FOLDER, exist_ok=True)


def generate_dataset(n=200):
    t = np.linspace(0, 10, n)

    drift = t * 0.3
    clean_signal = np.sin(t) * 10 + drift

    noise = np.random.normal(0, 1.5, size=n)
    lower_signal = clean_signal - noise
    upper_signal = clean_signal + noise

    return pd.DataFrame({
        "Time (s)": t,
        "Clean Signal": clean_signal,
        "Lower Signal": lower_signal,
        "Upper Signal": upper_signal
    })


def generate_all():
    print(f"Generating synthetic data in:\n{DATA_FOLDER}\n")

    for i in range(1, N_FILES + 1):
        df = generate_dataset(N_POINTS)
        filename = f"experiment1_run_{i:02d}.xlsx"
        path = os.path.join(DATA_FOLDER, filename)
        df.to_excel(path, index=False)
        print(f"Created: {path}")

    print("\nDone!")


if __name__ == "__main__":
    generate_all()
