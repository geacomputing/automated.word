import os
import numpy as np
import pandas as pd
from pathlib import Path

# ------------------------------------------------------------
# BASE DIR = project root (relative path)
# ------------------------------------------------------------
# __file__      → full path of THIS Python file (generate_Sinthetic_data.py)
# resolve()     → converts it to an absolute path, resolving symlinks too
# parents[2]    → go 2 levels up:
#                 parents[0] = file directory (.../utils)
#                 parents[1] = project root /utils
#                 parents[2] = project root /automated.word  ← WE WANT THIS
#
# This ensures the script works correctly even if the project is moved
# to another computer, user, or directory.
BASE_DIR = Path(__file__).resolve().parents[2]

# Build the full path to the folder where synthetic data will be stored.
# Using os.path.join() ensures compatibility across operating systems.
DATA_FOLDER = os.path.join(BASE_DIR, "data", "Lab_results", "experiment1")

# Number of files to generate (experiment runs)
N_FILES = 3

# Number of sample points in each signal dataset
N_POINTS = 100

# Ensure target folder exists.
# exist_ok=True means: "do nothing if folder already exists".
os.makedirs(DATA_FOLDER, exist_ok=True)


def generate_dataset(n=200):
    """
    Generate a synthetic experimental dataset containing:
    - Time vector
    - Clean signal (sinusoidal + drift)
    - Lower signal (clean minus noise)
    - Upper signal (clean plus noise)

    Parameters
    ----------
    n : int
        Number of sample points in the generated signal.

    Returns
    -------
    pandas.DataFrame
        A DataFrame with 4 columns:
        'Time (s)', 'Clean Signal', 'Lower Signal', 'Upper Signal'
    """

    # Time vector from 0 to 10 seconds
    t = np.linspace(0, 10, n)

    # Add a linear drift to make the data more realistic (non-stationary)
    drift = t * 0.3

    # Clean sinusoidal signal, scaled and combined with drift
    clean_signal = np.sin(t) * 10 + drift

    # Gaussian noise, mean=0, stddev=1.5
    noise = np.random.normal(0, 1.5, size=n)

    # Lower/upper envelopes: simulate experimental variation bounds
    lower_signal = clean_signal - noise
    upper_signal = clean_signal + noise

    # Package into a DataFrame for easy export
    return pd.DataFrame({
        "Time (s)": t,
        "Clean Signal": clean_signal,
        "Lower Signal": lower_signal,
        "Upper Signal": upper_signal
    })


def generate_all():
    """
    Generate multiple synthetic experiment Excel files and save them
    into the experiment1 directory under data/Lab_results.
    """

    print(f"Generating synthetic data in:\n{DATA_FOLDER}\n")

    # Loop that produces N_FILES runs
    for i in range(1, N_FILES + 1):
        df = generate_dataset(N_POINTS)

        # Filename formatted as: experiment1_run_01.xlsx
        filename = f"experiment1_run_{i:02d}.xlsx"

        # Full path to the output file
        path = os.path.join(DATA_FOLDER, filename)

        # Export the DataFrame to Excel
        df.to_excel(path, index=False)

        print(f"Created: {path}")

    print("\nDone!")


if __name__ == "__main__":
    # This ensures the function runs only when executing the script directly,
    # not when imported as a module.
    generate_all()
