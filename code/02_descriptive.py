"""
02_descriptive.py
Descriptive statistics + time-series figures for the report.

Inputs : ../data/processed/panel.csv, ../data/processed/benchmarks.csv
Outputs: ../figures/fig_universe_coverage.png
         ../figures/fig_sentiment_distribution.png
         ../figures/fig_sentiment_timeseries.png
         ../figures/fig_universe_returns.png
         ../data/processed/summary_stats.csv
"""
from pathlib import Path
import numpy as np
import pandas as pd
import matplotlib.pyplot as plt
import matplotlib.dates as mdates

PROJECT  = Path(__file__).resolve().parent.parent
PROC_DIR = PROJECT / "data" / "processed"
FIG_DIR  = PROJECT / "figures"
FIG_DIR.mkdir(parents=True, exist_ok=True)

plt.rcParams.update({
    "figure.dpi":    120,
    "savefig.dpi":   200,
    "font.family":   "DejaVu Sans",
    "axes.grid":     True,
    "grid.alpha":    0.3,
    "axes.spines.top":   False,
    "axes.spines.right": False,
})


def main():
    panel = pd.read_csv(PROC_DIR / "panel.csv", parse_dates=["date"])
    bench = pd.read_csv(PROC_DIR / "benchmarks.csv", parse_dates=["date"])
    print(f"Loaded panel: {len(panel):,} rows, {panel['ticker'].nunique()} tickers")

    # ---------- Summary stats per ticker ----------
    summary = panel.groupby("ticker").agg(
        sector       = ("sector", "first"),
        n_days       = ("date", "count"),
        first_date   = ("date", "min"),
        last_date    = ("date", "max"),
        mean_ret_bps = ("log_ret", lambda s: s.mean() * 1e4),
        vol_ann_pct  = ("log_ret", lambda s: s.std() * np.sqrt(252) * 100),
        avg_mcap_b   = ("mkt_cap", lambda s: s.mean() / 1e3),
        sent_mean    = ("sentiment", "mean"),
        sent_std     = ("sentiment", "std"),
        sent_cov_pct = ("sentiment", lambda s: s.notna().mean() * 100),
    ).round(3).sort_values("avg_mcap_b", ascending=False)
    summary.to_csv(PROC_DIR / "summary_stats.csv")
    print(f"\nTop 5 by avg market cap (Billions USD):")
    print(summary[["sector", "avg_mcap_b", "vol_ann_pct", "sent_mean"]].head().to_string())
    print(f"\nWrote {PROC_DIR / 'summary_stats.csv'}")

    # ---------- Figure 1: Sector coverage ----------
    fig, ax = plt.subplots(figsize=(10, 5))
    sec_count = (
        panel.drop_duplicates("ticker")
             .groupby("sector").size().sort_values()
    )
    sec_count.plot(kind="barh", ax=ax, color="#2c7fb8")
    ax.set_xlabel("Number of stocks")
    ax.set_title("Universe composition by GICS sector (30 large-caps)")
    fig.tight_layout()
    fig.savefig(FIG_DIR / "fig_universe_coverage.png")
    plt.close(fig)
    print(f"Wrote {FIG_DIR / 'fig_universe_coverage.png'}")

    # ---------- Figure 2: Sentiment distribution overall + by sector ----------
    fig, axes = plt.subplots(1, 2, figsize=(13, 5))
    panel["sentiment"].dropna().hist(
        bins=80, ax=axes[0], color="#2c7fb8", edgecolor="white"
    )
    axes[0].axvline(0, color="black", lw=0.7)
    axes[0].set_title(f"Daily sentiment, all tickers (n={panel['sentiment'].notna().sum():,})")
    axes[0].set_xlabel("NEWS_SENTIMENT_DAILY_AVG")
    axes[0].set_ylabel("Frequency")

    box_data = [
        panel.loc[panel.sector == s, "sentiment"].dropna().values
        for s in sorted(panel.sector.unique())
    ]
    axes[1].boxplot(
        box_data,
        labels=sorted(panel.sector.unique()),
        showfliers=False,
        patch_artist=True,
        boxprops=dict(facecolor="#a6cee3", edgecolor="black"),
    )
    axes[1].axhline(0, color="black", lw=0.7)
    axes[1].set_title("Sentiment distribution by sector")
    axes[1].set_ylabel("NEWS_SENTIMENT_DAILY_AVG")
    plt.setp(axes[1].get_xticklabels(), rotation=30, ha="right")
    fig.tight_layout()
    fig.savefig(FIG_DIR / "fig_sentiment_distribution.png")
    plt.close(fig)
    print(f"Wrote {FIG_DIR / 'fig_sentiment_distribution.png'}")

    # ---------- Figure 3: Time-series of universe-mean sentiment + VIX overlay ----------
    daily_sent = (
        panel.groupby("date")["sentiment"].mean().rolling(21, min_periods=10).mean()
    )
    fig, ax1 = plt.subplots(figsize=(12, 5))
    ax1.plot(daily_sent.index, daily_sent.values, color="#2c7fb8",
             lw=1.4, label="Univ. mean sentiment (21-day MA)")
    ax1.axhline(0, color="black", lw=0.7)
    ax1.set_ylabel("Mean sentiment", color="#2c7fb8")
    ax1.tick_params(axis="y", labelcolor="#2c7fb8")
    ax1.set_title("Universe-wide news sentiment vs. VIX (2018-2026)")
    ax2 = ax1.twinx()
    ax2.plot(bench["date"], bench["vix"], color="#d95f02", lw=0.9, alpha=0.7, label="VIX")
    ax2.set_ylabel("VIX", color="#d95f02")
    ax2.tick_params(axis="y", labelcolor="#d95f02")
    ax2.grid(False)

    for label, dt in [("COVID crash", "2020-03-20"),
                      ("Russia invasion", "2022-02-24"),
                      ("Banking crisis (SVB)", "2023-03-13")]:
        ax1.axvline(pd.Timestamp(dt), color="gray", lw=0.7, ls="--", alpha=0.7)
        ax1.text(pd.Timestamp(dt), ax1.get_ylim()[1] * 0.92, label,
                 rotation=90, va="top", ha="right", fontsize=8, color="gray")

    ax1.xaxis.set_major_locator(mdates.YearLocator())
    ax1.xaxis.set_major_formatter(mdates.DateFormatter("%Y"))
    fig.tight_layout()
    fig.savefig(FIG_DIR / "fig_sentiment_timeseries.png")
    plt.close(fig)
    print(f"Wrote {FIG_DIR / 'fig_sentiment_timeseries.png'}")

    # ---------- Figure 4: Cumulative returns of equal-weight universe vs SPX ----------
    panel_sorted = panel.sort_values(["ticker", "date"])
    daily_uni = (
        panel_sorted.groupby("date")["log_ret"].mean()
    )
    cum_uni = daily_uni.cumsum().apply(np.exp) - 1

    spx_ret = bench.set_index("date")["spx_ret"]
    cum_spx = spx_ret.cumsum().apply(lambda x: 0 if pd.isna(x) else x)
    cum_spx = (1 + spx_ret).cumprod() - 1

    fig, ax = plt.subplots(figsize=(12, 5))
    ax.plot(cum_uni.index, cum_uni.values * 100, label="Equal-wt universe (30 stocks)",
            color="#2c7fb8", lw=1.4)
    ax.plot(cum_spx.index, cum_spx.values * 100, label="S&P 500 (SPX)",
            color="#33a02c", lw=1.4, ls="--")
    ax.axhline(0, color="black", lw=0.7)
    ax.set_ylabel("Cumulative return (%)")
    ax.set_title("Cumulative return of universe vs S&P 500 (2018-2026)")
    ax.legend(loc="upper left")
    ax.xaxis.set_major_locator(mdates.YearLocator())
    ax.xaxis.set_major_formatter(mdates.DateFormatter("%Y"))
    fig.tight_layout()
    fig.savefig(FIG_DIR / "fig_universe_returns.png")
    plt.close(fig)
    print(f"Wrote {FIG_DIR / 'fig_universe_returns.png'}")

    print("\nAll descriptive figures + summary stats written.")


if __name__ == "__main__":
    main()
