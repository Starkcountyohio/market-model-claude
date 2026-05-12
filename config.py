"""Project-wide paths.

The code lives in a git repo (cloned to e.g. C:\\Users\\<you>\\dev\\market-model-claude\\),
but the data lives in the shared OneDrive folder. We bridge the two with a single
environment variable, MARKET_MODEL_DATA, that each user sets once on their machine
to point at their local copy of the OneDrive 'Market Model Claude' folder.

Set it (PowerShell, one-time, current user):
    [Environment]::SetEnvironmentVariable("MARKET_MODEL_DATA",
        "C:\\Users\\<you>\\OneDrive - Stark County\\Desktop\\Market Model Claude",
        "User")
Then restart your shell / VS Code so the new env var is visible.

Verify:
    py -c "from config import SALES_CSV; print(SALES_CSV, SALES_CSV.exists())"
"""
import os
from pathlib import Path

_env = os.environ.get("MARKET_MODEL_DATA")
if not _env:
    raise EnvironmentError(
        "MARKET_MODEL_DATA environment variable is not set. "
        "Point it at your local OneDrive copy of the 'Market Model Claude' folder. "
        "See BETHANY_SETUP.md for click-by-click instructions."
    )

DATA_DIR = Path(_env)
if not DATA_DIR.exists():
    raise FileNotFoundError(
        f"MARKET_MODEL_DATA points at a non-existent path: {DATA_DIR}\n"
        "Make sure OneDrive has finished syncing the shared 'Market Model Claude' folder."
    )

# --- 2026 cycle ---
CYCLE_2026_DIR  = DATA_DIR / "2026 Market Model"
SALES_CSV       = CYCLE_2026_DIR / "Sales Jan 2021-Dec 2025.CSV"
MARKET_EXTRACT  = CYCLE_2026_DIR / "Market Sale Extract 5-12-26.html"
MKVAR_HTML      = CYCLE_2026_DIR / "MKVAR_2026-05-12T09-45-41.html"
MKMOD_HTML      = CYCLE_2026_DIR / "MKMOD_2026-05-12T09-48-22.html"
MKMODVAR_HTML   = CYCLE_2026_DIR / "MKMODVAR_2026-05-12T09-48-49.html"
MKCLUST_HTML    = CYCLE_2026_DIR / "MKCLUST_2026-05-12T09-47-48.html"
RCMODEL_HTML    = CYCLE_2026_DIR / "RCMODEL_2026-05-12T09-46-33.html"
RCTERM_V_HTML   = CYCLE_2026_DIR / "RCTERM_V_2026-05-12T09-47-22.html"

# --- Reference workbooks (root of shared folder) ---
FRAMEWORK_SFR    = DATA_DIR / "Model-Cluster-NGroup-NBHD Assignments.xlsx"
FRAMEWORK_CONDO  = DATA_DIR / "Framework Condo Model.xlsx"
TIME_ADJ_2025    = DATA_DIR / "Time adjustment for 2025 version.xlsx"

# --- Output workbooks land in the shared folder so both collaborators see them ---
OUT_DIR = DATA_DIR
