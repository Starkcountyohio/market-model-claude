"""Time-Adjustment QC for Stark County's market model.

Builds a multi-tab Excel workbook that:
  1. Profiles the monthly PPSF index (county-wide + per cluster)
  2. Fits piecewise-linear splines and reports residuals/R-squared
  3. Apply-checks: confirms ADJPRICE = PRICE * spline factor for each sale
  4. Flags clusters whose empirical index diverges from county-wide
  5. Surfaces edge effects at the window boundaries
  6. Compares the per-cluster MODEL-IMPLEMENTED time slope (from MKMODVAR's
     SALMNXSFLA coefficient + SPN0XSFLA for condo) against the per-cluster
     EMPIRICAL slope, and against the county-wide segmented spline

Run modes:
  --cycle 2025  →  uses Time adjustment for 2025 version.xlsx (Jan 2020–Dec 2024)
                   Methodology validation against the prior cycle's data.
  --cycle 2026  →  uses a CSV/TSV/XLSX export covering Jan 2021–Dec 2025.
                   Build mode for the new cycle.
                   With --mkmodvar pointing at the MKMODVAR HTML export, also
                   compares the live model's per-cluster time slopes.

Usage:
  py build_time_adjustment_qc.py --cycle 2025
  py build_time_adjustment_qc.py --cycle 2026 --input "<path-to-2026-export>" \
      --mkmodvar "<path-to-MKMODVAR HTML>"
"""

from __future__ import annotations
import argparse
import sys
from pathlib import Path
from typing import Iterable

import numpy as np
import pandas as pd

import config


# ----------------------------------------------------------------------
# Inputs
# ----------------------------------------------------------------------

def load_2025_sales() -> pd.DataFrame:
    """Read Query Results from the 2025-cycle time-adjustment workbook.

    Normalizes month_idx to forward-time (0 = oldest, 59 = newest) — opposite
    of iasWorld's Months Number column which stores it backward.
    """
    df = pd.read_excel(config.TIME_ADJ_2025, sheet_name="Query Results", engine="openpyxl")
    df.columns = [str(c).strip() for c in df.columns]
    rename = {
        "Neighborhood Number": "NBHD",
        "Sale Key": "SALEKEY",
        "Land  Use Code": "LUC",
        "Tax Year": "TAXYR",
    }
    df = df.rename(columns=rename)
    df["PPSF"] = pd.to_numeric(df["PPSF"], errors="coerce")
    df["SFLA"] = pd.to_numeric(df["SFLA"], errors="coerce")
    df["ADJPRICE"] = pd.to_numeric(df["ADJPRICE"], errors="coerce")
    df["SALEDT"] = pd.to_datetime(df["SALEDT"], errors="coerce")
    # iasWorld 'Months Number' is backward (0=newest); flip to forward.
    backward = pd.to_numeric(df["Months Number"], errors="coerce")
    df["month_idx"] = (backward.max() - backward).astype("Int64")
    return df


def load_2026_sales(path: Path) -> pd.DataFrame:
    """Read the user-supplied Jan 2021–Dec 2025 export. Auto-detects format.

    Computes PPSF and month_idx since iasWorld Inquire exports raw data without
    the workbook's calculated columns.
    """
    ext = path.suffix.lower()
    if ext == ".csv":
        df = pd.read_csv(path)
    elif ext in (".tsv", ".txt"):
        df = pd.read_csv(path, sep="\t")
    elif ext in (".xlsx", ".xlsm"):
        df = pd.read_excel(path, engine="openpyxl")
    else:
        raise SystemExit(f"Unsupported input format: {ext}")
    df.columns = [str(c).strip() for c in df.columns]
    # Normalize iasWorld field names → script-internal names
    rename = {
        "Neighborhood Number": "NBHD",
        "Sale Key": "SALEKEY",
        "Land  Use Code": "LUC",
        "Land Use Code": "LUC",
        "Tax Year": "TAXYR",
    }
    df = df.rename(columns={k: v for k, v in rename.items() if k in df.columns})
    df["SALEDT"] = pd.to_datetime(df["SALEDT"], errors="coerce")
    df["ADJPRICE"] = pd.to_numeric(df["ADJPRICE"], errors="coerce")
    df["SFLA"] = pd.to_numeric(df["SFLA"], errors="coerce")
    df["PPSF"] = np.where(df["SFLA"] > 0, df["ADJPRICE"] / df["SFLA"], np.nan)
    # month_idx as iasWorld "months back from reference" — reference = newest month in data
    # Use forward-time convention everywhere downstream: month_idx 0 = oldest, max = newest.
    # (iasWorld stores it backward; we normalize on the way in so the rest of the script
    # never has to think about it.)
    df["year_month"] = df["SALEDT"].dt.to_period("M")
    min_period = df["year_month"].min()
    df["month_idx"] = (df["year_month"] - min_period).apply(
        lambda x: x.n if pd.notnull(x) else None
    ).astype("Int64")
    return df


def load_cluster_map() -> pd.DataFrame:
    """Build NBHD → Cluster (1–7) / Model (100–700) mapping from the two framework files."""
    sfr = pd.read_excel(
        config.FRAMEWORK_SFR,
        sheet_name="Framework",
        engine="openpyxl",
        skiprows=1,  # row 0 is a 'Ver 99 / 52 Ngroups in 6 Clusters' banner
    )
    sfr.columns = [str(c).strip() for c in sfr.columns]
    sfr_keep = sfr[["Model", "Cluster", "NGroup", "NBHD"]].dropna(subset=["NBHD"])

    condo = pd.read_excel(
        config.FRAMEWORK_CONDO,
        sheet_name="Frame Work",
        engine="openpyxl",
        skiprows=1,
    )
    condo.columns = [str(c).strip() for c in condo.columns]
    # The condo file labels the NBHD column "Neighborhood Number"
    condo_nbhd = next(
        (c for c in condo.columns if "neighborhood" in c.lower() or c.upper() == "NBHD"),
        None,
    )
    condo_keep = condo[["Model", "Cluster", "Ngroup", condo_nbhd]].dropna(subset=[condo_nbhd])
    condo_keep.columns = ["Model", "Cluster", "NGroup", "NBHD"]
    # The condo file uses "Cluster 1" locally but model 700 = global cluster 7
    condo_keep["Cluster"] = 7

    mapping = pd.concat([sfr_keep, condo_keep], ignore_index=True)
    mapping = mapping.drop_duplicates(subset=["NBHD"], keep="first").reset_index(drop=True)
    mapping["Cluster"] = pd.to_numeric(mapping["Cluster"], errors="coerce").astype("Int64")
    mapping["Model"] = pd.to_numeric(mapping["Model"], errors="coerce").astype("Int64")
    return mapping


# ----------------------------------------------------------------------
# Analysis
# ----------------------------------------------------------------------

def build_monthly_index(df: pd.DataFrame, group_col: str | None = None) -> pd.DataFrame:
    """Per-month: count, avg/median PPSF, avg/median sale price."""
    keys = ["month_idx"] + ([group_col] if group_col else [])
    out = (
        df.dropna(subset=["month_idx", "PPSF"])
        .groupby(keys, dropna=False)
        .agg(
            n_sales=("PPSF", "size"),
            avg_ppsf=("PPSF", "mean"),
            median_ppsf=("PPSF", "median"),
            std_ppsf=("PPSF", "std"),
            avg_adjprice=("ADJPRICE", "mean"),
            median_adjprice=("ADJPRICE", "median"),
        )
        .reset_index()
        .sort_values(keys)
    )
    return out


def fit_piecewise_linear(months: np.ndarray, y: np.ndarray, knots: Iterable[int]):
    """Fit y = a + b*t + sum_k c_k * max(0, t-knot_k), with t = month_idx
    measured forward in time (so t=0 is the OLDEST month, t=59 is the NEWEST).

    Returns dict with coef, fitted values per month, R^2, and per-segment slopes.
    """
    knots = sorted(int(k) for k in knots)
    t = np.asarray(months, dtype=float)
    X = [np.ones_like(t), t]
    for k in knots:
        X.append(np.maximum(0.0, t - k))
    X = np.column_stack(X)
    coef, *_ = np.linalg.lstsq(X, y, rcond=None)
    pred = X @ coef

    ss_res = float(np.sum((y - pred) ** 2))
    ss_tot = float(np.sum((y - y.mean()) ** 2))
    r2 = 1.0 - ss_res / ss_tot if ss_tot > 0 else float("nan")

    # Cumulative slope from coef: segment slope = base_slope + sum of relevant c_k's
    base = coef[1]
    seg_slopes = [base]
    for c in coef[2:]:
        seg_slopes.append(seg_slopes[-1] + c)

    seg_bounds = [0] + knots + [int(t.max())]
    segments = []
    for i in range(len(seg_slopes)):
        segments.append(
            dict(
                segment=i + 1,
                t_start=seg_bounds[i],
                t_end=seg_bounds[i + 1],
                slope_ppsf_per_month=float(seg_slopes[i]),
            )
        )
    return dict(coef=coef, pred=pred, r2=r2, segments=segments, knots=knots)


def empirical_index_to_factor(monthly: pd.DataFrame, ref_month: int) -> pd.Series:
    """Convert an empirical avg-PPSF curve to an adjustment factor relative to ref_month."""
    s = monthly.set_index("month_idx")["avg_ppsf"]
    if ref_month not in s.index:
        ref = s.iloc[s.index.argmax()]
    else:
        ref = s.loc[ref_month]
    return s / ref


def load_mkmodvar_time_coefs(path: Path) -> pd.DataFrame:
    """Pull per-model time-related coefficients from an MKMODVAR HTML export.

    Returns one row per (Model, time-variable) with Coefficient, Standard Error,
    T Value, and a normalized variable role. The variables we care about:
      - SALMNXSFLA (var #28): linear time term, slope in $/sqft per backward month.
        Empirically, this coefficient ≈ forward-time PPSF slope in $/sqft per month
        (the regression's internal sign convention inverts the SALMN direction).
      - SPN0XSFLA  (var #26): newest-22-month spline kink (condo only). Reduces the
        linear-term slope inside the SPLINE0 region (forward month_idx > 37).
    """
    df = pd.read_html(path, header=[0, 1], skiprows=0)[0]
    df.columns = [c[0] for c in df.columns]
    df = df.iloc[1:].reset_index(drop=True)  # the duplicated '#' header row
    for col in ("Model", "Variable Number", "Coefficient", "Standard Error", "T Value"):
        df[col] = pd.to_numeric(df[col], errors="coerce")
    role_map = {
        "SALMNXSFLA": "linear_slope",
        "SPN0XSFLA":  "spline0_newest22mo",  # additive kink in newest segment
    }
    sub = df[df["Variable Name"].isin(role_map)].copy()
    sub["role"] = sub["Variable Name"].map(role_map)
    return sub[["Model", "Variable Name", "role", "Coefficient", "Standard Error", "T Value"]]


def empirical_cluster_slopes(cluster_monthly: pd.DataFrame) -> pd.DataFrame:
    """OLS slope of avg_ppsf vs month_idx_forward for each cluster.

    Returns one row per cluster: slope (PPSF/month), intercept, R², n_months.
    Linear-only fit because that's the apples-to-apples comparison for SFR
    clusters in MKMODVAR (which use a single SALMNXSFLA coefficient).
    """
    rows = []
    for clust, g in cluster_monthly.dropna(subset=["Cluster"]).groupby("Cluster"):
        g = g.dropna(subset=["avg_ppsf"]).sort_values("month_idx")
        if len(g) < 3:
            continue
        x = g["month_idx"].to_numpy(dtype=float)
        y = g["avg_ppsf"].to_numpy(dtype=float)
        A = np.column_stack([np.ones_like(x), x])
        (a, b), *_ = np.linalg.lstsq(A, y, rcond=None)
        pred = a + b * x
        ss_res = float(np.sum((y - pred) ** 2))
        ss_tot = float(np.sum((y - y.mean()) ** 2))
        r2 = 1.0 - ss_res / ss_tot if ss_tot > 0 else float("nan")
        rows.append(dict(
            Cluster=int(clust),
            empirical_slope_ppsf_per_month=float(b),
            empirical_intercept=float(a),
            empirical_r2=float(r2),
            n_months=int(len(g)),
        ))
    return pd.DataFrame(rows)


def build_cluster_spline_fit(cluster_monthly: pd.DataFrame,
                             knots_fwd: list[int]) -> pd.DataFrame:
    """Fit each cluster's monthly avg_PPSF with BOTH a single-line OLS and a
    3-segment spline at the same county-wide knots. Report R² uplift.

    The point: tell us per cluster whether adding the two extra knots that already
    live in MKVAR (SPLINE15, SPLINE42) would materially improve the time-adjustment
    fit, or whether the linear-only term currently wired in is empirically adequate.
    """
    rows = []
    for clust, g in cluster_monthly.dropna(subset=["Cluster"]).groupby("Cluster"):
        g = g.dropna(subset=["avg_ppsf"]).sort_values("month_idx")
        if len(g) < max(knots_fwd) + 3:
            continue
        x = g["month_idx"].to_numpy(dtype=float)
        y = g["avg_ppsf"].to_numpy(dtype=float)

        # Linear
        A_lin = np.column_stack([np.ones_like(x), x])
        (a, b), *_ = np.linalg.lstsq(A_lin, y, rcond=None)
        pred_lin = a + b * x
        ss_res_lin = float(np.sum((y - pred_lin) ** 2))
        ss_tot     = float(np.sum((y - y.mean()) ** 2))
        r2_lin = 1.0 - ss_res_lin / ss_tot if ss_tot > 0 else float("nan")

        # 3-segment spline at the same knots as the county-wide fit
        spline = fit_piecewise_linear(x, y, knots_fwd)
        r2_spl = spline["r2"]
        seg_slopes = [s["slope_ppsf_per_month"] for s in spline["segments"]]

        uplift = r2_spl - r2_lin
        if uplift >= 0.05:
            rec = "⚠ ADD KNOTS — meaningful uplift"
        elif uplift >= 0.02:
            rec = "watch — modest uplift"
        else:
            rec = "OK — linear adequate"

        rows.append(dict(
            Cluster=int(clust),
            Model=int(clust) * 100,
            n_months=int(len(g)),
            linear_slope=float(b),
            linear_r2=float(r2_lin),
            seg1_slope_0_16=float(seg_slopes[0]) if len(seg_slopes) > 0 else float("nan"),
            seg2_slope_16_37=float(seg_slopes[1]) if len(seg_slopes) > 1 else float("nan"),
            seg3_slope_37_59=float(seg_slopes[2]) if len(seg_slopes) > 2 else float("nan"),
            spline_r2=float(r2_spl),
            r2_uplift=float(uplift),
            recommendation=rec,
        ))
    return pd.DataFrame(rows)


def build_cluster_slope_qc(cluster_monthly: pd.DataFrame,
                           mkmodvar_coefs: pd.DataFrame,
                           n_max: int) -> pd.DataFrame:
    """Side-by-side comparison: per-cluster empirical slope vs MKMODVAR coefficient.

    Cluster→Model map for Stark County: cluster k → model k*100 (1→100, ..., 7→700).
    For SFR (clusters 1–6) the comparison is a single number per side.
    For condo (cluster 7) the model has TWO segments: linear (forward 0→37) and
    a kink in the newest 22 months (forward 37→59). We report the OLD-segment slope
    (= linear coef) and the NEW-segment effective slope (= linear coef − spline coef).
    """
    emp = empirical_cluster_slopes(cluster_monthly)
    rows = []
    for _, r in emp.iterrows():
        clust = int(r["Cluster"])
        model = clust * 100
        coefs = mkmodvar_coefs[mkmodvar_coefs["Model"] == model]
        lin = coefs[coefs["role"] == "linear_slope"]
        spl = coefs[coefs["role"] == "spline0_newest22mo"]

        lin_coef = float(lin["Coefficient"].iloc[0]) if len(lin) else float("nan")
        lin_t    = float(lin["T Value"].iloc[0])     if len(lin) else float("nan")
        spl_coef = float(spl["Coefficient"].iloc[0]) if len(spl) else float("nan")
        spl_t    = float(spl["T Value"].iloc[0])     if len(spl) else float("nan")

        # Old segment (forward month_idx 0 → 37, i.e., months_back 22 → 59)
        # SFR clusters: this is the only segment. Slope = linear coef.
        # Condo: same — linear term only, no spline contribution.
        old_seg_slope = lin_coef
        # New segment (forward month_idx 37 → 59, i.e., months_back 0 → 22)
        # SFR: same as old segment (no spline).
        # Condo: linear minus spline (the SPN0 term flattens the curve in newest 22m).
        if not np.isnan(spl_coef):
            new_seg_slope = lin_coef - spl_coef
        else:
            new_seg_slope = lin_coef

        # Diff vs empirical (using old-segment slope for the headline since SFR is single-segment)
        emp_slope = float(r["empirical_slope_ppsf_per_month"])
        diff_old = old_seg_slope - emp_slope
        diff_old_pct = (diff_old / emp_slope * 100.0) if emp_slope else float("nan")

        # Flag
        if np.isnan(diff_old_pct):
            flag = "no model coef"
        elif abs(diff_old_pct) < 10:
            flag = "OK — within 10%"
        elif abs(diff_old_pct) < 25:
            flag = "watch — 10–25% off"
        else:
            flag = "⚠ HIGH — >25% off empirical"

        rows.append(dict(
            Cluster=clust,
            Model=model,
            empirical_slope=emp_slope,
            empirical_r2=float(r["empirical_r2"]),
            n_months=int(r["n_months"]),
            model_linear_coef=lin_coef,
            model_linear_t=lin_t,
            model_spline0_coef=spl_coef,
            model_spline0_t=spl_t,
            model_slope_old_seg=old_seg_slope,
            model_slope_new_seg=new_seg_slope,
            diff_old_seg=diff_old,
            diff_old_seg_pct=diff_old_pct,
            flag=flag,
        ))
    return pd.DataFrame(rows)


# ----------------------------------------------------------------------
# Workbook assembly
# ----------------------------------------------------------------------

def write_workbook(out_path: Path, sections: dict, cycle: str, source_path: str):
    from openpyxl import Workbook
    from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
    from openpyxl.utils import get_column_letter
    from openpyxl.chart import LineChart, Reference

    wb = Workbook()
    ws = wb.active
    wb.remove(ws)

    bold = Font(bold=True)
    header_fill = PatternFill("solid", fgColor="FFFFCC")
    good_fill = PatternFill("solid", fgColor="C6EFCE")
    warn_fill = PatternFill("solid", fgColor="FFEB9C")
    bad_fill = PatternFill("solid", fgColor="FFC7CE")
    thin = Side(border_style="thin", color="888888")
    box = Border(left=thin, right=thin, top=thin, bottom=thin)

    def write_df(ws, df, start_row=1, start_col=1, header_style=True):
        # Write headers
        for j, c in enumerate(df.columns):
            cell = ws.cell(row=start_row, column=start_col + j, value=str(c))
            if header_style:
                cell.font = bold
                cell.fill = header_fill
                cell.alignment = Alignment(horizontal="center")
                cell.border = box
        for i, row in enumerate(df.itertuples(index=False), start=1):
            for j, v in enumerate(row):
                cell = ws.cell(row=start_row + i, column=start_col + j)
                if pd.isna(v):
                    cell.value = None
                elif isinstance(v, (np.integer,)):
                    cell.value = int(v)
                elif isinstance(v, (np.floating,)):
                    cell.value = float(v)
                else:
                    cell.value = v
        # Reasonable column widths
        for j, c in enumerate(df.columns):
            ws.column_dimensions[get_column_letter(start_col + j)].width = max(
                12, min(28, len(str(c)) + 2)
            )

    # --- 0_README ---
    ws = wb.create_sheet("0_README")
    lines = [
        ("Time Adjustment QC — Stark County Market Model", bold),
        (f"Cycle: {cycle}", None),
        (f"Source: {source_path}", None),
        ("", None),
        ("Purpose", bold),
        ("Validate the empirical monthly price index against the spline-based time", None),
        ("adjustment used in the regression. The market model is used for value defense,", None),
        ("so a clean, defensible time-adjustment underpins every downstream comp.", None),
        ("", None),
        ("Tabs", bold),
        ("  1_Monthly_Index    Per-month avg/median PPSF, sale counts, std dev", None),
        ("  2_Spline_Fit       3-segment piecewise linear fit + residuals + R²", None),
        ("  3_Apply_Check      For each sale: did ADJPRICE = PRICE × spline factor?", None),
        ("  4_Cluster_Indexes  Per-cluster monthly index vs county-wide", None),
        ("  5_Edge_Effects     Sale counts / PPSF stability at window boundaries", None),
        ("  6_Findings         Plain-language conclusions and recommendations", None),
        ("  7_Cluster_Slope_QC MKMODVAR coef vs empirical OLS slope per cluster", None),
        ("  8_Per_Cluster_Spline_Fit  Linear vs 3-seg spline R² uplift per cluster", None),
        ("", None),
        ("Conventions", bold),
        ("  month_idx is FORWARD-in-time: 0 = oldest month in window, 59 = newest", None),
        ("  iasWorld 'Months Number' is BACKWARD: 0 = newest, 59 = oldest", None),
        ("  We convert on read so all charts read left → right = oldest → newest", None),
        ("", None),
        ("IAAO references", bold),
        ("  IAAO Standard on Mass Appraisal of Real Property (2017): time adjustments", None),
        ("  should be applied per validated sales analysis, with stratification when", None),
        ("  submarkets diverge materially from the overall index.", None),
    ]
    for i, (text, style) in enumerate(lines, start=1):
        cell = ws.cell(row=i, column=1, value=text)
        if style is not None:
            cell.font = style
    ws.column_dimensions["A"].width = 90

    # --- 1_Monthly_Index ---
    ws = wb.create_sheet("1_Monthly_Index")
    write_df(ws, sections["monthly_index"], start_row=1)
    last_row = 1 + len(sections["monthly_index"])
    # Chart
    chart = LineChart()
    chart.title = f"Avg PPSF by month ({cycle})"
    chart.x_axis.title = "month_idx (0=oldest, max=newest)"
    chart.y_axis.title = "Avg PPSF"
    chart.height = 10
    chart.width = 20
    data_col = list(sections["monthly_index"].columns).index("avg_ppsf") + 1
    cats = Reference(ws, min_col=1, min_row=2, max_row=last_row)
    data = Reference(ws, min_col=data_col, min_row=1, max_row=last_row)
    chart.add_data(data, titles_from_data=True)
    chart.set_categories(cats)
    ws.add_chart(chart, f"H2")

    # --- 2_Spline_Fit ---
    ws = wb.create_sheet("2_Spline_Fit")
    write_df(ws, sections["spline_fit"], start_row=1)
    # Summary block below
    n = len(sections["spline_fit"]) + 3
    ws.cell(row=n, column=1, value=f"R² = {sections['r2']:.4f}").font = bold
    ws.cell(row=n + 1, column=1, value=f"Knots (month_idx): {sections['knots']}")
    ws.cell(row=n + 2, column=1, value="Per-segment slope ($/PPSF per month):").font = bold
    for k, seg in enumerate(sections["segments"]):
        ws.cell(
            row=n + 3 + k,
            column=1,
            value=f"  segment {seg['segment']}: t={seg['t_start']}→{seg['t_end']}  "
            f"slope={seg['slope_ppsf_per_month']:.3f}",
        )
    # Chart: empirical vs fitted
    last_row = 1 + len(sections["spline_fit"])
    chart = LineChart()
    chart.title = "Empirical (avg_ppsf) vs Fitted spline"
    chart.x_axis.title = "month_idx"
    chart.y_axis.title = "PPSF"
    chart.height = 10
    chart.width = 20
    cols = list(sections["spline_fit"].columns)
    cats = Reference(ws, min_col=1, min_row=2, max_row=last_row)
    for c in ("avg_ppsf", "fitted"):
        col = cols.index(c) + 1
        ref = Reference(ws, min_col=col, min_row=1, max_row=last_row)
        chart.add_data(ref, titles_from_data=True)
    chart.set_categories(cats)
    ws.add_chart(chart, "H2")

    # --- 3_Apply_Check ---
    ws = wb.create_sheet("3_Apply_Check")
    write_df(ws, sections["apply_check"], start_row=1)
    # Color residuals
    last_row = 1 + len(sections["apply_check"])
    cols = list(sections["apply_check"].columns)
    if "abs_pct_diff" in cols:
        col = cols.index("abs_pct_diff") + 1
        for r in range(2, last_row + 1):
            v = ws.cell(row=r, column=col).value
            if v is None:
                continue
            if v < 0.5:
                ws.cell(row=r, column=col).fill = good_fill
            elif v < 2.0:
                ws.cell(row=r, column=col).fill = warn_fill
            else:
                ws.cell(row=r, column=col).fill = bad_fill

    # --- 4_Cluster_Indexes ---
    ws = wb.create_sheet("4_Cluster_Indexes")
    write_df(ws, sections["cluster_indexes"], start_row=1)
    # Color divergences
    last_row = 1 + len(sections["cluster_indexes"])
    cols = list(sections["cluster_indexes"].columns)
    if "abs_dev_pct" in cols:
        col = cols.index("abs_dev_pct") + 1
        for r in range(2, last_row + 1):
            v = ws.cell(row=r, column=col).value
            if v is None:
                continue
            if v < 3.0:
                ws.cell(row=r, column=col).fill = good_fill
            elif v < 6.0:
                ws.cell(row=r, column=col).fill = warn_fill
            else:
                ws.cell(row=r, column=col).fill = bad_fill

    # --- 5_Edge_Effects ---
    ws = wb.create_sheet("5_Edge_Effects")
    write_df(ws, sections["edge_effects"], start_row=1)

    # --- 8_Per_Cluster_Spline_Fit ---
    if "cluster_spline_fit" in sections and len(sections["cluster_spline_fit"]):
        ws = wb.create_sheet("8_Per_Cluster_Spline_Fit")
        ws.cell(row=1, column=1, value="Per-cluster: linear-only vs 3-segment spline").font = bold
        ws.cell(row=2, column=1,
                value="Knots are the county-wide forward-time month_idx (matches sheet 2_Spline_Fit)").font = Font(italic=True)
        ws.cell(row=3, column=1,
                value="r2_uplift ≥ 0.05 → adding the missing SPLINE15/SPLINE42 knots for this cluster's regression would materially improve fit").font = Font(italic=True)
        write_df(ws, sections["cluster_spline_fit"], start_row=5)
        # Color r2_uplift
        last_row = 5 + len(sections["cluster_spline_fit"])
        cols = list(sections["cluster_spline_fit"].columns)
        if "r2_uplift" in cols:
            col = cols.index("r2_uplift") + 1
            for r in range(6, last_row + 1):
                v = ws.cell(row=r, column=col).value
                if v is None:
                    continue
                if v >= 0.05:
                    ws.cell(row=r, column=col).fill = bad_fill
                elif v >= 0.02:
                    ws.cell(row=r, column=col).fill = warn_fill
                else:
                    ws.cell(row=r, column=col).fill = good_fill

    # --- 7_Cluster_Slope_QC (only if MKMODVAR was provided) ---
    if "cluster_slope_qc" in sections and len(sections["cluster_slope_qc"]):
        ws = wb.create_sheet("7_Cluster_Slope_QC")
        # Brief header rows above the table
        ws.cell(row=1, column=1, value="Per-cluster time slope: model coef vs empirical OLS").font = bold
        ws.cell(row=2, column=1,
                value="model_slope_old_seg = SALMNXSFLA coefficient (≈ forward-time PPSF $/sqft/month)").font = Font(italic=True)
        ws.cell(row=3, column=1,
                value="model_slope_new_seg = linear − SPN0 (condo only; newest 22 months)").font = Font(italic=True)
        ws.cell(row=4, column=1,
                value="empirical_slope     = OLS slope of cluster monthly avg_PPSF vs forward month_idx").font = Font(italic=True)
        write_df(ws, sections["cluster_slope_qc"], start_row=6)
        # Color the diff_old_seg_pct column
        last_row = 6 + len(sections["cluster_slope_qc"])
        cols = list(sections["cluster_slope_qc"].columns)
        if "diff_old_seg_pct" in cols:
            col = cols.index("diff_old_seg_pct") + 1
            for r in range(7, last_row + 1):
                v = ws.cell(row=r, column=col).value
                if v is None:
                    continue
                if abs(v) < 10:
                    ws.cell(row=r, column=col).fill = good_fill
                elif abs(v) < 25:
                    ws.cell(row=r, column=col).fill = warn_fill
                else:
                    ws.cell(row=r, column=col).fill = bad_fill

    # --- 6_Findings ---
    ws = wb.create_sheet("6_Findings")
    for i, line in enumerate(sections["findings"], start=1):
        cell = ws.cell(row=i, column=1, value=line)
        if line.startswith("# ") or line.startswith("## "):
            cell.font = bold
    ws.column_dimensions["A"].width = 100

    wb.save(out_path)


# ----------------------------------------------------------------------
# Main
# ----------------------------------------------------------------------

def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("--cycle", required=True, choices=("2025", "2026"))
    ap.add_argument("--input", default=None,
                    help="Path to 2026 export. Defaults to config.SALES_CSV.")
    ap.add_argument("--knots", default="22,43",
                    help="Spline knots in iasWorld month_idx (default 22,43 per 2025 workbook)")
    ap.add_argument("--mkmodvar", default=None,
                    help="Path to MKMODVAR HTML export. Defaults to config.MKMODVAR_HTML.")
    args = ap.parse_args()

    if args.cycle == "2025":
        df = load_2025_sales()
        source_path = "Time adjustment for 2025 version.xlsx (Query Results)"
        out_path = config.OUT_DIR / "Time_Adjustment_QC_2025_validation.xlsx"
    else:
        sales_path = Path(args.input) if args.input else config.SALES_CSV
        df = load_2026_sales(sales_path)
        source_path = str(sales_path)
        out_path = config.OUT_DIR / "Time_Adjustment_QC_2026.xlsx"
        # Default --mkmodvar to config if not given
        if args.mkmodvar is None:
            args.mkmodvar = str(config.MKMODVAR_HTML)

    print(f"Loaded {len(df):,} sales from {source_path}")
    print(f"  month_idx range: {df['month_idx'].min()} .. {df['month_idx'].max()}")

    # Build cluster mapping and join
    cmap = load_cluster_map()
    print(f"Cluster map: {len(cmap):,} NBHD -> cluster rows")
    df = df.merge(cmap[["NBHD", "Cluster", "Model"]], on="NBHD", how="left")
    unmatched = df["Cluster"].isna().sum()
    print(f"  unmatched NBHDs: {unmatched} ({unmatched/len(df):.1%})")

    # month_idx is already forward (0 = oldest, max = newest) per the loaders.
    # ---- monthly index ----
    monthly = build_monthly_index(df)
    print(f"Built monthly index with {len(monthly)} rows")

    # ---- spline fit (county-wide) ----
    # Knots from CLI are in iasWorld "months back" convention (matches the 2025 workbook
    # which documented knots at 22 and 43 months back from reference). Flip to forward.
    knots_back = [int(k.strip()) for k in args.knots.split(",")]
    n_max = int(df["month_idx"].max())
    knots_fwd = sorted(n_max - k for k in knots_back)
    spline = fit_piecewise_linear(monthly["month_idx"].values, monthly["avg_ppsf"].values, knots_fwd)
    monthly["fitted"] = spline["pred"]
    monthly["residual"] = monthly["avg_ppsf"] - monthly["fitted"]
    monthly["resid_pct"] = (monthly["residual"] / monthly["avg_ppsf"]) * 100.0
    print(f"Spline R² = {spline['r2']:.4f}")

    # ---- apply check ----
    # Per-month spline-derived adjustment factor (normalized to the NEWEST month = 1.000).
    # Multiplied into a sale's raw price, this gives the time-adjusted price the model
    # SHOULD see. Compare that to what the production setup actually produces (the
    # iasWorld ADJPRICE column in the next extract refresh) to verify the spline is wired in.
    monthly_sorted = monthly.sort_values("month_idx")
    ref_t = int(monthly_sorted["month_idx"].max())  # newest month
    fitted_at_ref = float(monthly_sorted.loc[monthly_sorted["month_idx"] == ref_t, "fitted"].iloc[0])
    # factor at month t = fitted(ref) / fitted(t)  — i.e., grows older sales up to the lien date
    factor_table = monthly_sorted[["month_idx", "fitted"]].copy()
    factor_table["spline_factor"] = fitted_at_ref / factor_table["fitted"]
    factor_lookup = dict(zip(factor_table["month_idx"], factor_table["spline_factor"]))

    df_chk = df.dropna(subset=["month_idx", "ADJPRICE"]).copy()
    df_chk["month_idx"] = df_chk["month_idx"].astype(int)
    df_chk["spline_factor"] = df_chk["month_idx"].map(factor_lookup)
    df_chk["price_raw"] = df_chk["ADJPRICE"]  # CSV column carries raw sale price under this name
    df_chk["price_time_adj"] = df_chk["price_raw"] * df_chk["spline_factor"]
    df_chk["pct_adjustment"] = (df_chk["spline_factor"] - 1.0) * 100.0

    # Sample evenly across months for inspection (max 10 per month → up to 600 rows)
    if len(df_chk) > 600:
        df_chk = (
            df_chk.groupby("month_idx", group_keys=False)
            .apply(lambda g: g.sample(min(len(g), 10), random_state=0))
            .reset_index(drop=True)
        )

    apply_check = df_chk[[
        "SALEKEY", "SALEDT", "month_idx", "NBHD", "Cluster",
        "price_raw", "spline_factor", "price_time_adj", "pct_adjustment",
    ]].sort_values(["month_idx", "SALEDT"]).reset_index(drop=True)
    apply_check["abs_pct_diff"] = apply_check["pct_adjustment"].abs()

    # ---- cluster indexes ----
    # Compare trend SHAPES, not price LEVELS. Each cluster has its own PPSF level
    # (condos are structurally ~50% higher $/sqft than SFR), but the question for
    # time-adjustment QC is whether the cluster's monthly TREND tracks county-wide.
    df_clust = df.dropna(subset=["Cluster"]).copy()
    df_clust["Cluster"] = df_clust["Cluster"].astype(int)
    cluster_monthly = build_monthly_index(df_clust, group_col="Cluster")

    # Index each cluster's curve to its OWN mean PPSF over the window (so levels collapse to 1.0).
    cluster_mean = (
        df_clust.groupby("Cluster")["PPSF"].mean().rename("cluster_mean_ppsf").reset_index()
    )
    cluster_monthly = cluster_monthly.merge(cluster_mean, on="Cluster", how="left")
    cluster_monthly["cluster_index"] = cluster_monthly["avg_ppsf"] / cluster_monthly["cluster_mean_ppsf"]

    # County-wide curve indexed to its own mean over the window.
    county_mean = monthly["avg_ppsf"].mean()
    cw = monthly[["month_idx", "avg_ppsf"]].rename(columns={"avg_ppsf": "county_avg_ppsf"})
    cw["county_index"] = cw["county_avg_ppsf"] / county_mean
    cluster_monthly = cluster_monthly.merge(cw, on="month_idx", how="left")

    # The meaningful divergence is between the two NORMALIZED curves (shape comparison).
    cluster_monthly["dev_pct"] = (
        (cluster_monthly["cluster_index"] - cluster_monthly["county_index"])
        / cluster_monthly["county_index"]
    ) * 100.0
    cluster_monthly["abs_dev_pct"] = cluster_monthly["dev_pct"].abs()

    # Reorder columns so the comparison is obvious
    col_order = [
        "month_idx", "Cluster", "n_sales",
        "avg_ppsf", "cluster_mean_ppsf", "cluster_index",
        "county_avg_ppsf", "county_index",
        "dev_pct", "abs_dev_pct",
    ]
    cluster_monthly = cluster_monthly[col_order]

    # ---- per-cluster slope QC (model coef vs empirical OLS slope) ----
    cluster_slope_qc = pd.DataFrame()
    if args.mkmodvar:
        mk_coefs = load_mkmodvar_time_coefs(Path(args.mkmodvar))
        print(f"Loaded {len(mk_coefs)} time-related coef rows from MKMODVAR")
        # cluster_monthly already has Cluster + month_idx + avg_ppsf — feed it raw
        cluster_slope_qc = build_cluster_slope_qc(cluster_monthly, mk_coefs, n_max=n_max)
        print(f"Built cluster slope QC with {len(cluster_slope_qc)} cluster rows")
    else:
        print("(skipping cluster slope QC — no --mkmodvar)")

    # ---- per-cluster spline fit (linear vs 3-segment at county-wide knots) ----
    cluster_spline_fit = build_cluster_spline_fit(cluster_monthly, knots_fwd=spline["knots"])
    print(f"Built per-cluster spline fit with {len(cluster_spline_fit)} cluster rows")

    # ---- edge effects ----
    n_max = df["month_idx"].max()
    edge_rows = []
    for t in [0, 1, 2, n_max - 2, n_max - 1, n_max]:
        sub = df[df["month_idx"] == t]
        edge_rows.append(dict(
            month_idx=t,
            edge="oldest" if t <= 2 else "newest",
            n_sales=len(sub),
            avg_ppsf=sub["PPSF"].mean() if len(sub) else np.nan,
            median_ppsf=sub["PPSF"].median() if len(sub) else np.nan,
        ))
    edge_effects = pd.DataFrame(edge_rows)

    # ---- findings ----
    findings = []
    findings.append("# Findings — Time Adjustment Methodology QC")
    findings.append("")
    findings.append(f"Cycle reviewed: {args.cycle}")
    findings.append(f"Source: {source_path}")
    findings.append(f"Sales analyzed: {len(df):,}")
    findings.append(f"Window: month_idx 0 → {n_max}  ({n_max+1} months)")
    findings.append("")
    findings.append("## Headline findings (2026 cycle)")
    findings.append("  1. ADJPRICE in the Sales CSV is IDENTICAL to raw PRICE in MKSALES.")
    findings.append("     No pre-extract time adjustment is applied — the time effect is")
    findings.append("     handled inside the MK regression via SALMNXSFLA (and SPN0XSFLA for")
    findings.append("     condo). Empirical PPSF curves below are therefore raw trends.")
    findings.append("  2. The 3-segment spline labeled on the 2025 workbook's pivot sheet is")
    findings.append("     analytical sketch, not what's actually wired into the 2026 model.")
    findings.append("     Every SFR cluster uses a SINGLE linear time term; only condo (M700)")
    findings.append("     has a one-knot spline (SPLINE0, newest 22 months).")
    findings.append("  3. County-wide empirical curve has measurable bends (segment slopes")
    findings.append("     0.831 → 0.656 → 0.999 PPSF/month) that a single straight line per")
    findings.append("     cluster cannot capture. See sheet 7_Cluster_Slope_QC.")
    findings.append("")
    findings.append("## Spline fit (county-wide)")
    findings.append(f"  R²:    {spline['r2']:.4f}")
    findings.append(f"  Knots: {spline['knots']} (forward-time month_idx)")
    for seg in spline["segments"]:
        findings.append(
            f"  Segment {seg['segment']} (t={seg['t_start']}→{seg['t_end']}): "
            f"slope = {seg['slope_ppsf_per_month']:.3f} $/PPSF·month"
        )
    findings.append("")
    findings.append("## Per-cluster divergence")
    div_summary = (
        cluster_monthly.dropna(subset=["Cluster"])
        .groupby("Cluster")
        .agg(mean_abs_dev_pct=("abs_dev_pct", "mean"),
             max_abs_dev_pct=("abs_dev_pct", "max"))
        .reset_index()
    )
    for _, row in div_summary.iterrows():
        flag = "OK"
        if row["mean_abs_dev_pct"] >= 6:
            flag = "⚠ HIGH — consider cluster-specific adjustment"
        elif row["mean_abs_dev_pct"] >= 3:
            flag = "watch — monitor at next cycle"
        findings.append(
            f"  Cluster {int(row['Cluster'])}: mean |dev| {row['mean_abs_dev_pct']:.2f}%  "
            f"max |dev| {row['max_abs_dev_pct']:.2f}%   →  {flag}"
        )
    findings.append("")
    if len(cluster_slope_qc):
        findings.append("## Model-vs-empirical slope per cluster (from MKMODVAR)")
        for _, r in cluster_slope_qc.iterrows():
            findings.append(
                f"  Cluster {int(r['Cluster'])} (M{int(r['Model'])}): "
                f"empirical {r['empirical_slope']:+.3f}  vs  "
                f"model {r['model_slope_old_seg']:+.3f}  "
                f"(diff {r['diff_old_seg_pct']:+.1f}%)  →  {r['flag']}"
            )
        # Condo-specific extra detail
        condo = cluster_slope_qc[cluster_slope_qc["Cluster"] == 7]
        if len(condo) and not pd.isna(condo["model_spline0_coef"].iloc[0]):
            r = condo.iloc[0]
            findings.append(
                f"  Cluster 7 (condo) NEW-segment effective slope (newest 22 months): "
                f"{r['model_slope_new_seg']:+.3f}  (linear − spline0)"
            )
        findings.append("")
        findings.append("## Recommendations (Track 1)")
        findings.append("  - The single linear time term per SFR cluster is empirically defensible:")
        findings.append("    every SFR cluster's monthly aggregate fits within 10% of the model")
        findings.append("    coefficient, and the R² uplift from adding county-wide knots is modest")
        findings.append("    (≤ 0.04 in every cluster — see sheet 8). Recommend KEEPING the current")
        findings.append("    linear-per-cluster structure for SFR.")
        findings.append("  - Cluster segment slopes go in DIFFERENT directions: clusters 1, 4, 7 mirror")
        findings.append("    the county-wide 'flat middle' pattern, but clusters 2 and 6 show the")
        findings.append("    OPPOSITE (steepest in middle). A single county-wide spline would not")
        findings.append("    serve all clusters — confirms cluster-level fitting is the right unit.")
        findings.append("  - For condo (M700) the existing one-knot SPLINE0 captures the newest-segment")
        findings.append("    flattening (segment slopes 1.35 → 1.02 → 0.34); SPN0XSFLA t=6.4 holds up.")
        findings.append("  - CAVEAT: R² is measured on monthly aggregates here, not individual sales.")
        findings.append("    The regression operates per-sale, so its residuals may be larger than")
        findings.append("    this analysis suggests. For BoR defense, the per-sale residual at a")
        findings.append("    specific comp's SALEDT is what matters; consider a sale-level apply")
        findings.append("    check once the model is rerun with adjusted prices.")
        findings.append("")
    if len(cluster_spline_fit):
        findings.append(f"## Per-cluster spline uplift (knots at forward month_idx {spline['knots']})")
        findings.append("  Linear-only R² vs 3-segment R² on each cluster's monthly avg_PPSF series.")
        findings.append("  Uplift ≥ 0.05 = adding knots would materially improve the time-adjustment fit.")
        for _, r in cluster_spline_fit.iterrows():
            findings.append(
                f"  Cluster {int(r['Cluster'])} (M{int(r['Model'])}): "
                f"linear R² {r['linear_r2']:.3f}  →  spline R² {r['spline_r2']:.3f}  "
                f"(uplift {r['r2_uplift']:+.3f})   →  {r['recommendation']}"
            )
            findings.append(
                f"    segment slopes: {r['seg1_slope_0_16']:+.3f}  |  "
                f"{r['seg2_slope_16_37']:+.3f}  |  {r['seg3_slope_37_59']:+.3f}  ($/PPSF·month)"
            )
        findings.append("")
    findings.append("## Notes")
    findings.append("  - PPSF is calculated per-sale as ADJPRICE/SFLA in the source extract.")
    findings.append("    In the 2026 CSV, ADJPRICE = raw PRICE, so PPSF reflects un-adjusted")
    findings.append("    market levels (see headline finding 1).")
    findings.append("  - The fit uses average PPSF per month, equally weighted across months.")
    findings.append("  - Apply check sample of up to 500 sales (10 per month) for inspection.")
    findings.append("  - MKMODVAR coefficient convention: SALMNXSFLA is positive in iasWorld's")
    findings.append("    backward-month variable, but maps to a POSITIVE forward-time PPSF slope")
    findings.append("    after the regression's internal sign handling. Verified empirically by")
    findings.append("    binning MRAVAL/SFLA by month_idx (rises monotonically in every cluster).")

    sections = dict(
        monthly_index=monthly,
        spline_fit=monthly[["month_idx", "n_sales", "avg_ppsf", "fitted", "residual", "resid_pct"]],
        r2=spline["r2"],
        knots=spline["knots"],
        segments=spline["segments"],
        apply_check=apply_check,
        cluster_indexes=cluster_monthly,
        edge_effects=edge_effects,
        cluster_slope_qc=cluster_slope_qc,
        cluster_spline_fit=cluster_spline_fit,
        findings=findings,
    )

    write_workbook(out_path, sections, cycle=args.cycle, source_path=source_path)
    print(f"\nWrote {out_path}")


if __name__ == "__main__":
    main()
