"""
04_portfolio_backtest.py
Test Hypothesis H2: A long-short portfolio sorted on Bloomberg news sentiment
generates positive risk-adjusted returns net of transaction costs.

Design:
  - Rebalance every 5 trading days (the IC analysis showed the signal works at
    5-21 day horizons, not at 1-day).
  - On each rebalance date: rank stocks by sentiment_t (today). Long the top
    quintile (6 names, equal-weighted), short the bottom quintile (6 names).
    Hold for the next 5 trading days.
  - Charge a one-way 5 bps transaction cost on every trade (10 bps round-trip).
  - Compare to S&P 500 (SPX) and to the equal-weight long-only universe.

Outputs:
  - ../data/processed/strategy_returns.csv
  - ../data/processed/perf_summary.csv
  - ../figures/fig_equity_curves.png
  - ../figures/fig_drawdowns.png
  - ../figures/fig_quintile_returns.png
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
    "figure.dpi": 120, "savefig.dpi": 200, "font.family": "DejaVu Sans",
    "axes.grid": True, "grid.alpha": 0.3,
    "axes.spines.top": False, "axes.spines.right": False,
})

REBAL_FREQ_DAYS = 5
N_QUINTILES     = 5
COST_BPS_ONEWAY = 5.0   # 5 basis points per trade per side


def perf_stats(daily_ret, freq=252, name="Strategy"):
    """Compute annualised return, vol, Sharpe, Sortino, max DD, hit rate."""
    r = daily_ret.dropna()
    if len(r) < 30:
        return pd.Series(dtype=float, name=name)
    cum = (1 + r).cumprod()
    ann_ret = (cum.iloc[-1] ** (freq / len(r))) - 1
    ann_vol = r.std() * np.sqrt(freq)
    sharpe  = ann_ret / ann_vol if ann_vol > 0 else np.nan
    downside = r[r < 0]
    sortino = (ann_ret / (downside.std() * np.sqrt(freq))) if len(downside) and downside.std() > 0 else np.nan
    rolling_max = cum.cummax()
    dd = cum / rolling_max - 1
    max_dd = dd.min()
    hit = (r > 0).mean()
    return pd.Series({
        "n_days":      len(r),
        "ann_ret_pct": ann_ret * 100,
        "ann_vol_pct": ann_vol * 100,
        "sharpe":      sharpe,
        "sortino":     sortino,
        "max_dd_pct":  max_dd * 100,
        "hit_rate":    hit,
        "skew":        r.skew(),
        "kurt":        r.kurt(),
    }, name=name)


def main():
    panel = pd.read_csv(PROC_DIR / "panel.csv", parse_dates=["date"])
    bench = pd.read_csv(PROC_DIR / "benchmarks.csv", parse_dates=["date"])
    panel = panel.sort_values(["ticker", "date"]).reset_index(drop=True)

    # Build wide return matrix: rows=date, cols=ticker
    wide_ret = panel.pivot(index="date", columns="ticker", values="log_ret")
    wide_sent = panel.pivot(index="date", columns="ticker", values="sentiment")
    common_dates = wide_ret.index.intersection(wide_sent.index)
    wide_ret = wide_ret.loc[common_dates]
    wide_sent = wide_sent.loc[common_dates]

    # Convert log returns to simple returns for portfolio compounding
    wide_simple = np.exp(wide_ret) - 1

    # Build positions on every rebal date by ranking sentiment
    dates = wide_sent.index.sort_values()
    rebal_dates = dates[::REBAL_FREQ_DAYS]
    print(f"Total trading days: {len(dates)} | rebal every {REBAL_FREQ_DAYS} -> {len(rebal_dates)} rebals")

    # Position matrix (forward-filled between rebal dates)
    positions = pd.DataFrame(0.0, index=dates, columns=wide_sent.columns)

    for rd in rebal_dates:
        sent = wide_sent.loc[rd].dropna()
        if len(sent) < N_QUINTILES * 2:
            continue
        ranks = sent.rank(method="first")
        n = len(ranks)
        q_size = n // N_QUINTILES
        # bottom quintile = short; top quintile = long
        top    = ranks.nlargest(q_size).index
        bottom = ranks.nsmallest(q_size).index
        # equal weights, dollar-neutral, gross exposure 200%
        long_weight  = +1.0 / len(top)
        short_weight = -1.0 / len(bottom)
        positions.loc[rd, top] = long_weight
        positions.loc[rd, bottom] = short_weight

    # forward fill positions until next rebal
    positions = positions.replace(0, np.nan).ffill().fillna(0)

    # Strategy daily return = sum(position_{i,t-1} * simple_ret_{i,t})
    # We hold positions starting day after rebal (no look-ahead).
    pos_lag = positions.shift(1).fillna(0)
    strat_gross = (pos_lag * wide_simple).sum(axis=1)

    # Transaction costs: change in absolute position * cost_bps
    turnover = positions.diff().abs().sum(axis=1).fillna(0)
    trade_cost = turnover * (COST_BPS_ONEWAY / 1e4)
    strat_net = strat_gross - trade_cost.shift(1).fillna(0)

    # Equal-weight long-only universe (benchmark 1)
    universe_ew = wide_simple.mean(axis=1)

    # SPX (benchmark 2)
    spx = bench.set_index("date")["spx_ret"].reindex(dates)

    # Combine returns
    rets = pd.DataFrame({
        "long_short_gross": strat_gross,
        "long_short_net":   strat_net,
        "universe_ew":      universe_ew,
        "spx":              spx,
    }).dropna(how="all")
    rets.to_csv(PROC_DIR / "strategy_returns.csv")
    print(f"Wrote {PROC_DIR/'strategy_returns.csv'}  ({len(rets):,} rows)")

    # Performance stats
    stats = pd.concat([
        perf_stats(rets["long_short_gross"], name="L/S gross"),
        perf_stats(rets["long_short_net"],   name="L/S net of 5bps"),
        perf_stats(rets["universe_ew"],      name="Universe EW (long-only)"),
        perf_stats(rets["spx"],              name="S&P 500"),
    ], axis=1).T.round(3)
    stats.to_csv(PROC_DIR / "perf_summary.csv")
    print(f"\n=== Performance summary ===")
    print(stats.to_string())
    print(f"\nWrote {PROC_DIR/'perf_summary.csv'}")

    # ---------- Figure: Equity curves ----------
    cum = (1 + rets.fillna(0)).cumprod() - 1
    fig, ax = plt.subplots(figsize=(12, 6))
    palette = {"long_short_net": "#1f78b4", "long_short_gross": "#a6cee3",
               "universe_ew": "#33a02c", "spx": "#ff7f00"}
    labels  = {"long_short_net": "Long-Short (net 5bps)",
               "long_short_gross": "Long-Short (gross)",
               "universe_ew": "Universe equal-wt (long only)",
               "spx": "S&P 500"}
    for col in ["long_short_gross", "long_short_net", "universe_ew", "spx"]:
        ax.plot(cum.index, cum[col].values * 100,
                color=palette[col], lw=1.5, label=labels[col])
    ax.axhline(0, color="black", lw=0.7)
    ax.set_title("Equity curves: 5-day-rebalance L/S quintile vs benchmarks (2018-2026)")
    ax.set_ylabel("Cumulative return (%)")
    ax.legend(loc="upper left")
    ax.xaxis.set_major_locator(mdates.YearLocator())
    ax.xaxis.set_major_formatter(mdates.DateFormatter("%Y"))
    fig.tight_layout()
    fig.savefig(FIG_DIR / "fig_equity_curves.png")
    plt.close(fig)
    print(f"Wrote {FIG_DIR/'fig_equity_curves.png'}")

    # ---------- Figure: Drawdowns ----------
    fig, ax = plt.subplots(figsize=(12, 4))
    for col in ["long_short_net", "spx"]:
        c = (1 + rets[col].fillna(0)).cumprod()
        dd = (c / c.cummax() - 1) * 100
        ax.fill_between(dd.index, dd.values, 0,
                        alpha=0.4, label=labels[col], color=palette[col])
    ax.set_title("Drawdowns: L/S strategy (net) vs S&P 500")
    ax.set_ylabel("Drawdown (%)")
    ax.legend(loc="lower left")
    ax.xaxis.set_major_locator(mdates.YearLocator())
    ax.xaxis.set_major_formatter(mdates.DateFormatter("%Y"))
    fig.tight_layout()
    fig.savefig(FIG_DIR / "fig_drawdowns.png")
    plt.close(fig)
    print(f"Wrote {FIG_DIR/'fig_drawdowns.png'}")

    # ---------- Figure: Quintile-portfolio returns (monotonicity check) ----------
    # On each rebal day, compute the next-period (5d) return of each sentiment-sorted quintile
    q_returns = []
    for rd in rebal_dates[:-1]:
        sent = wide_sent.loc[rd].dropna()
        if len(sent) < N_QUINTILES * 2:
            continue
        try:
            quint_id = pd.qcut(sent, N_QUINTILES, labels=False, duplicates="drop")
        except ValueError:
            continue
        # 5-day forward simple return
        idx = dates.get_loc(rd)
        end_idx = min(idx + REBAL_FREQ_DAYS, len(dates) - 1)
        period_ret = wide_simple.iloc[idx + 1: end_idx + 1].sum()
        for q in range(N_QUINTILES):
            members = quint_id[quint_id == q].index
            if len(members) > 0:
                q_returns.append({
                    "rebal_date": rd, "quintile": q + 1,
                    "ret": period_ret.reindex(members).mean()
                })
    q_df = pd.DataFrame(q_returns)
    q_summary = q_df.groupby("quintile")["ret"].agg(
        ["mean", "std", "count"]
    )
    q_summary["t_stat"] = q_summary["mean"] / (q_summary["std"] / np.sqrt(q_summary["count"]))
    q_summary["mean_bps"] = q_summary["mean"] * 1e4
    print(f"\n=== Per-quintile 5-day returns (sorted by sentiment) ===")
    print(q_summary.round(3).to_string())

    fig, ax = plt.subplots(figsize=(8, 5))
    bars = ax.bar(q_summary.index, q_summary["mean_bps"],
                  color=["#e31a1c", "#fb9a99", "#cccccc", "#a6cee3", "#1f78b4"])
    ax.axhline(0, color="black", lw=0.8)
    ax.set_xlabel("Sentiment quintile (1 = most negative, 5 = most positive)")
    ax.set_ylabel("Avg 5-day return (bps)")
    ax.set_title("Quintile-sorted forward 5-day returns — monotonicity check")
    for i, (q, row) in enumerate(q_summary.iterrows()):
        ax.text(q, row["mean_bps"] + np.sign(row["mean_bps"]) * 5,
                f"t={row['t_stat']:+.1f}",
                ha="center", fontsize=9)
    fig.tight_layout()
    fig.savefig(FIG_DIR / "fig_quintile_returns.png")
    plt.close(fig)
    print(f"Wrote {FIG_DIR/'fig_quintile_returns.png'}")


if __name__ == "__main__":
    main()
