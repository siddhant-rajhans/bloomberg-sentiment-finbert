"""
03_ic_analysis.py
Test Hypothesis H1: Today's NEWS_SENTIMENT_DAILY_AVG predicts next-day returns.

We use the Information Coefficient (IC), the canonical signal-validation metric in
quantitative equity research:
    IC_t = SpearmanRank-correlation_{i in cross-section}(sentiment_{i,t}, ret_{i,t+1})

Outputs:
  - ../data/processed/ic_daily.csv      (date, ic_pearson, ic_spearman, n_stocks)
  - ../data/processed/ic_summary.csv    (overall + per-sector IC stats)
  - ../figures/fig_ic_timeseries.png    (daily IC + 60-day rolling)
  - ../figures/fig_ic_by_sector.png     (per-sector mean IC bar chart)
  - ../figures/fig_ic_cumulative.png    (cumulative IC curve)
"""
from pathlib import Path
import numpy as np
import pandas as pd
from scipy import stats
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


def daily_ic(df, signal_col, fwd_ret_col, min_n=5):
    """Cross-sectional rank IC per date."""
    out = []
    for date, g in df.groupby("date"):
        g = g[[signal_col, fwd_ret_col]].dropna()
        if len(g) < min_n:
            continue
        # guard against zero-variance edge cases
        if g[signal_col].nunique() < 2 or g[fwd_ret_col].nunique() < 2:
            continue
        r_p, _ = stats.pearsonr(g[signal_col],  g[fwd_ret_col])
        r_s, _ = stats.spearmanr(g[signal_col], g[fwd_ret_col])
        out.append((date, r_p, r_s, len(g)))
    return pd.DataFrame(out, columns=["date", "ic_pearson", "ic_spearman", "n_stocks"])


def main():
    panel = pd.read_csv(PROC_DIR / "panel.csv", parse_dates=["date"])
    panel = panel.sort_values(["ticker", "date"]).reset_index(drop=True)

    # forward returns at multiple horizons.
    # IMPORTANT: at row t, ret_tN = log_ret[t+1] + ... + log_ret[t+N]
    # i.e. STRICTLY forward, no same-day or past contamination.
    panel["ret_t1"]  = panel.groupby("ticker")["log_ret"].shift(-1)
    panel["ret_t5"]  = (
        panel.groupby("ticker")["log_ret"]
             .transform(lambda s: s.rolling(5).sum().shift(-5))
    )
    panel["ret_t21"] = (
        panel.groupby("ticker")["log_ret"]
             .transform(lambda s: s.rolling(21).sum().shift(-21))
    )
    # same-day IC for sanity (should be largest by construction since news drives moves)
    panel["ret_t0"] = panel["log_ret"]

    print(f"Panel rows: {len(panel):,}")
    print(f"After dropping rows w/o sentiment or fwd-return: "
          f"{panel.dropna(subset=['sentiment','ret_t1']).shape[0]:,}")

    # ---------- daily cross-sectional IC at multiple horizons ----------
    horizons = {
        "ret_t0":  "Same-day",
        "ret_t1":  "1-day fwd",
        "ret_t5":  "5-day fwd",
        "ret_t21": "21-day fwd",
    }
    ic_multi = {}
    for col, label in horizons.items():
        df = daily_ic(panel, "sentiment", col)
        df["date"] = pd.to_datetime(df["date"])
        df = df.sort_values("date").reset_index(drop=True)
        ic_multi[col] = df
        print(f"  {label:<12}  n_days={len(df):>5}  "
              f"mean_IC={df['ic_spearman'].mean():+.4f}  "
              f"t={df['ic_spearman'].mean()/(df['ic_spearman'].std()/np.sqrt(len(df))):+.2f}")

    ic = ic_multi["ret_t1"]  # primary 1-day result
    ic.to_csv(PROC_DIR / "ic_daily.csv", index=False)
    print(f"\nWrote {PROC_DIR/'ic_daily.csv'}  ({len(ic):,} daily ICs)")

    # ---------- aggregate IC stats (overall + by sector + by year) ----------
    def ic_stats(s):
        s = s.dropna()
        n = len(s)
        if n == 0:
            return pd.Series({"n": 0, "mean": np.nan, "std": np.nan,
                              "t_stat": np.nan, "icir_ann": np.nan})
        m, sd = s.mean(), s.std()
        return pd.Series({
            "n":        n,
            "mean":     m,
            "std":      sd,
            "t_stat":   m / (sd / np.sqrt(n)) if sd > 0 else np.nan,
            "icir_ann": (m / sd) * np.sqrt(252) if sd > 0 else np.nan,
        })

    overall = pd.DataFrame({
        "all_pearson":  ic_stats(ic["ic_pearson"]),
        "all_spearman": ic_stats(ic["ic_spearman"]),
    }).T

    # by sector — small sectors have <5 names so we drop the cross-sectional
    # filter to min_n=2 (Spearman well-defined on 2 obs but per-day noise high;
    # the time-series average is still meaningful given thousands of days).
    sec_rows = []
    for sec, g in panel.groupby("sector"):
        ic_sec = daily_ic(g, "sentiment", "ret_t1", min_n=2)
        s = ic_stats(ic_sec["ic_spearman"])
        s.name = f"{sec} (n_stocks={g['ticker'].nunique()})"
        sec_rows.append(s)
    by_sector = pd.DataFrame(sec_rows).round(4)

    # by year
    ic["year"] = ic["date"].dt.year
    by_year = ic.groupby("year")["ic_spearman"].apply(ic_stats).unstack().round(4)

    # save
    # multi-horizon table
    horizon_rows = []
    for col, label in horizons.items():
        s = ic_stats(ic_multi[col]["ic_spearman"])
        s.name = label
        horizon_rows.append(s)
    by_horizon = pd.DataFrame(horizon_rows).round(4)

    summary_path = PROC_DIR / "ic_summary.csv"
    with open(summary_path, "w") as f:
        f.write("# IC SUMMARY -- daily cross-sectional rank correlation\n")
        f.write("# Signal: NEWS_SENTIMENT_DAILY_AVG (today)\n")
        f.write("# Target: forward log return\n\n")
        f.write("## Overall (1-day fwd return)\n")
        overall.round(4).to_csv(f)
        f.write("\n## By holding horizon (Spearman)\n")
        by_horizon.to_csv(f)
        f.write("\n## By sector (Spearman, 1-day fwd)\n")
        by_sector.to_csv(f)
        f.write("\n## By year (Spearman, 1-day fwd)\n")
        by_year.to_csv(f)
    print(f"Wrote {summary_path}")

    print("\n=== Overall daily IC stats (1-day fwd) ===")
    print(overall.round(4).to_string())
    print("\n=== Spearman IC by horizon ===")
    print(by_horizon.to_string())
    print("\n=== Spearman IC by sector (1-day fwd) ===")
    print(by_sector.to_string())
    print("\n=== Spearman IC by year (1-day fwd) ===")
    print(by_year.to_string())

    # ---------- Figure: IC time series ----------
    fig, ax = plt.subplots(figsize=(12, 5))
    ax.plot(ic["date"], ic["ic_spearman"], color="#a6cee3", lw=0.5,
            alpha=0.6, label="Daily Spearman IC")
    rolling = ic["ic_spearman"].rolling(60, min_periods=20).mean()
    ax.plot(ic["date"], rolling, color="#1f78b4", lw=1.6,
            label="60-day rolling mean")
    ax.axhline(0, color="black", lw=0.7)
    ax.axhline(ic["ic_spearman"].mean(), color="#e31a1c", lw=1, ls="--",
               label=f"Full-sample mean = {ic['ic_spearman'].mean():.4f}")
    ax.set_title("Daily cross-sectional IC: sentiment(t) vs return(t+1)")
    ax.set_ylabel("Rank correlation")
    ax.legend(loc="upper left")
    ax.xaxis.set_major_locator(mdates.YearLocator())
    ax.xaxis.set_major_formatter(mdates.DateFormatter("%Y"))
    fig.tight_layout()
    fig.savefig(FIG_DIR / "fig_ic_timeseries.png")
    plt.close(fig)
    print(f"\nWrote {FIG_DIR/'fig_ic_timeseries.png'}")

    # ---------- Figure: IC by sector ----------
    fig, ax = plt.subplots(figsize=(9, 5))
    means = by_sector["mean"].sort_values()
    colors = ["#33a02c" if v > 0 else "#e31a1c" for v in means]
    means.plot(kind="barh", ax=ax, color=colors)
    ax.axvline(0, color="black", lw=0.8)
    ax.set_xlabel("Mean Spearman IC")
    ax.set_title("Mean daily IC by GICS sector")
    fig.tight_layout()
    fig.savefig(FIG_DIR / "fig_ic_by_sector.png")
    plt.close(fig)
    print(f"Wrote {FIG_DIR/'fig_ic_by_sector.png'}")

    # ---------- Figure: cumulative IC ----------
    fig, ax = plt.subplots(figsize=(12, 5))
    cum = ic["ic_spearman"].cumsum()
    ax.plot(ic["date"], cum, color="#1f78b4", lw=1.4)
    ax.axhline(0, color="black", lw=0.7)
    ax.set_title("Cumulative Spearman IC (sentiment vs next-day return)")
    ax.set_ylabel("Sum of daily ICs")
    ax.xaxis.set_major_locator(mdates.YearLocator())
    ax.xaxis.set_major_formatter(mdates.DateFormatter("%Y"))
    fig.tight_layout()
    fig.savefig(FIG_DIR / "fig_ic_cumulative.png")
    plt.close(fig)
    print(f"Wrote {FIG_DIR/'fig_ic_cumulative.png'}")


if __name__ == "__main__":
    main()
