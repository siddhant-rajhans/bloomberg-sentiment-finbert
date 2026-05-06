"""
05_factor_regression.py
Test Hypothesis H3: After controlling for Fama-French 3 factors (Market, SMB, HML),
the sentiment-sorted long-short portfolio produces statistically significant alpha
with Newey-West HAC standard errors.

We download FF3 daily factors from Ken French's library (free, allowed as
supplementary data per project rubric). The Bloomberg pull is the primary data;
FF3 is the supplementary risk model.

Outputs:
  - ../data/processed/ff3_factors.csv   (cached for reproducibility)
  - ../data/processed/ff3_regression.csv (regression results table)
  - ../figures/fig_ff3_alpha.png         (rolling 252-day alpha)
"""
from pathlib import Path
import io
import urllib.request
import zipfile
import numpy as np
import pandas as pd
import statsmodels.api as sm
import matplotlib.pyplot as plt
import matplotlib.dates as mdates

PROJECT  = Path(__file__).resolve().parent.parent
PROC_DIR = PROJECT / "data" / "processed"
FIG_DIR  = PROJECT / "figures"
FIG_DIR.mkdir(parents=True, exist_ok=True)

plt.rcParams.update({
    "figure.dpi": 120, "savefig.dpi": 200, "font.family": "DejaVu Sans",
    "axes.grid": True, "grid.alpha": 0.3,
    "axes.spines.top": False, "axes.spines.right": False,
})

FF3_URL = (
    "https://mba.tuck.dartmouth.edu/pages/faculty/ken.french/ftp/"
    "F-F_Research_Data_Factors_daily_CSV.zip"
)
FF3_CACHE = PROC_DIR / "ff3_factors.csv"


def fetch_ff3():
    if FF3_CACHE.exists():
        print(f"Loading cached FF3 from {FF3_CACHE}")
        df = pd.read_csv(FF3_CACHE, parse_dates=["date"])
        return df
    print(f"Downloading FF3 daily factors from Ken French's library...")
    req = urllib.request.Request(FF3_URL, headers={"User-Agent": "Mozilla/5.0"})
    with urllib.request.urlopen(req, timeout=60) as resp:
        raw = resp.read()
    with zipfile.ZipFile(io.BytesIO(raw)) as zf:
        csv_name = next(n for n in zf.namelist() if n.lower().endswith(".csv"))
        with zf.open(csv_name) as f:
            text = f.read().decode("latin-1")
    # FF3 file has a 4-line header preamble; data rows are YYYYMMDD format until
    # an "annual" section that we skip.
    lines = text.splitlines()
    start = next(i for i, l in enumerate(lines) if "Mkt-RF" in l and "SMB" in l and "HML" in l)
    rows = []
    for l in lines[start + 1:]:
        l = l.strip()
        if not l:
            break
        parts = [p.strip() for p in l.split(",")]
        if len(parts) != 5:
            break
        date_str = parts[0]
        if not (date_str.isdigit() and len(date_str) == 8):
            break  # reached annual section
        rows.append([
            pd.to_datetime(date_str, format="%Y%m%d"),
            float(parts[1]) / 100.0,
            float(parts[2]) / 100.0,
            float(parts[3]) / 100.0,
            float(parts[4]) / 100.0,
        ])
    df = pd.DataFrame(rows, columns=["date", "mkt_rf", "smb", "hml", "rf"])
    df.to_csv(FF3_CACHE, index=False)
    print(f"  cached to {FF3_CACHE}  ({len(df):,} rows, "
          f"{df['date'].min().date()} -> {df['date'].max().date()})")
    return df


def regress_with_nw(y, X, lags=5, name="strategy"):
    """OLS with Newey-West HAC standard errors."""
    X_const = sm.add_constant(X)
    model = sm.OLS(y, X_const, missing="drop").fit(
        cov_type="HAC", cov_kwds={"maxlags": lags}
    )
    return model


def main():
    rets = pd.read_csv(PROC_DIR / "strategy_returns.csv", parse_dates=["date"])
    rets = rets.set_index("date")
    print(f"Loaded strategy returns: {len(rets):,} rows, "
          f"{rets.index.min().date()} -> {rets.index.max().date()}")

    ff3 = fetch_ff3().set_index("date")
    df = rets.join(ff3, how="inner")
    print(f"Joined with FF3: {len(df):,} aligned days")

    # Excess returns: long-short is dollar-neutral, so doesn't need risk-free
    # subtraction (already a "spread"). Long-only universe and SPX subtract rf.
    df["ls_excess"]  = df["long_short_net"]
    df["ls_gross_excess"] = df["long_short_gross"]
    df["uni_excess"] = df["universe_ew"] - df["rf"]
    df["spx_excess"] = df["spx"]         - df["rf"]

    # Run FF3 regressions
    results = {}
    for label, dep_col in [
        ("L/S net (5bps)",   "ls_excess"),
        ("L/S gross",        "ls_gross_excess"),
        ("Universe EW",      "uni_excess"),
        ("S&P 500",          "spx_excess"),
    ]:
        m = regress_with_nw(
            df[dep_col],
            df[["mkt_rf", "smb", "hml"]],
            lags=5, name=label,
        )
        results[label] = m

    # Build summary table
    rows = []
    for label, m in results.items():
        p = m.params
        t = m.tvalues
        rows.append({
            "strategy":      label,
            "alpha_bps":     p["const"] * 1e4,
            "alpha_t":       t["const"],
            "alpha_p":       m.pvalues["const"],
            "alpha_ann_pct": p["const"] * 252 * 100,
            "beta_mkt":      p["mkt_rf"],
            "beta_mkt_t":    t["mkt_rf"],
            "beta_smb":      p["smb"],
            "beta_smb_t":    t["smb"],
            "beta_hml":      p["hml"],
            "beta_hml_t":    t["hml"],
            "r_squared":     m.rsquared,
            "n":             int(m.nobs),
        })
    out = pd.DataFrame(rows).round(4)
    out.to_csv(PROC_DIR / "ff3_regression.csv", index=False)
    print(f"\n=== FF3 regression results (Newey-West, 5 lags) ===")
    print(out.to_string(index=False))
    print(f"\nWrote {PROC_DIR/'ff3_regression.csv'}")

    # Pretty regression printouts for the report appendix
    appendix_path = PROC_DIR / "ff3_regression_full.txt"
    with open(appendix_path, "w", encoding="utf-8") as f:
        for label, m in results.items():
            f.write(f"\n{'='*70}\n{label}\n{'='*70}\n")
            f.write(m.summary().as_text())
            f.write("\n")
    print(f"Wrote full regression printouts to {appendix_path}")

    # Rolling 252-day alpha for the L/S net strategy
    print("Computing rolling 252-day alpha...")
    alpha_series = []
    window = 252
    df_ls = df[["ls_excess", "mkt_rf", "smb", "hml"]].dropna()
    for i in range(window, len(df_ls)):
        chunk = df_ls.iloc[i - window:i]
        try:
            m = regress_with_nw(
                chunk["ls_excess"],
                chunk[["mkt_rf", "smb", "hml"]],
                lags=5,
            )
            alpha_series.append({
                "date":  df_ls.index[i],
                "alpha_bps": m.params["const"] * 1e4,
                "alpha_t":   m.tvalues["const"],
            })
        except Exception:
            continue
    rolling = pd.DataFrame(alpha_series).set_index("date")
    rolling.to_csv(PROC_DIR / "ff3_rolling_alpha.csv")

    fig, (ax1, ax2) = plt.subplots(2, 1, figsize=(12, 7), sharex=True)
    ax1.plot(rolling.index, rolling["alpha_bps"], color="#1f78b4", lw=1.4)
    ax1.axhline(0, color="black", lw=0.7)
    ax1.set_ylabel("Daily alpha (bps)")
    ax1.set_title("Rolling 252-day Fama-French alpha — L/S strategy (net 5bps)")
    ax2.plot(rolling.index, rolling["alpha_t"], color="#33a02c", lw=1.4)
    ax2.axhline(0, color="black", lw=0.7)
    ax2.axhline(2,  color="red", lw=0.7, ls="--", alpha=0.6, label="t=+/-2")
    ax2.axhline(-2, color="red", lw=0.7, ls="--", alpha=0.6)
    ax2.legend(loc="upper left")
    ax2.set_ylabel("Newey-West t-stat")
    ax2.set_xlabel("")
    ax2.xaxis.set_major_locator(mdates.YearLocator())
    ax2.xaxis.set_major_formatter(mdates.DateFormatter("%Y"))
    fig.tight_layout()
    fig.savefig(FIG_DIR / "fig_ff3_alpha.png")
    plt.close(fig)
    print(f"Wrote {FIG_DIR/'fig_ff3_alpha.png'}")

    print("\nDone.")


if __name__ == "__main__":
    main()
