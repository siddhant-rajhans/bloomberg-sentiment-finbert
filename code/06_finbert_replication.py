"""
06_finbert_replication.py
Test Hypothesis H4: An open-source FinBERT pipeline applied to recent news
headlines produces sentiment scores that correlate with Bloomberg's proprietary
NEWS_SENTIMENT_DAILY_AVG signal.

We can't bulk-export Bloomberg headlines on a typical license, so we use a free
news source (yfinance, a Yahoo Finance wrapper) as supplementary data — explicitly
permitted by the project rubric. yfinance returns ~10 recent headlines per ticker,
so the replication is over the past ~30 days for our 30-stock universe.

Pipeline:
  1. Pull recent headlines via yfinance (free / supplementary).
  2. Score each headline with ProsusAI/finbert (3-class: positive/neutral/negative)
     and convert to a continuous score in [-1, +1].
  3. Aggregate to daily per-ticker mean.
  4. Spearman-correlate against Bloomberg's NEWS_SENTIMENT_DAILY_AVG on the
     same (date, ticker) pairs.

Outputs:
  - ../data/processed/finbert_headlines.csv
  - ../data/processed/finbert_vs_bloomberg.csv
  - ../figures/fig_finbert_vs_bloomberg.png
"""
from pathlib import Path
import warnings
import time
import numpy as np
import pandas as pd
from scipy import stats
import matplotlib.pyplot as plt

warnings.filterwarnings("ignore")

PROJECT  = Path(__file__).resolve().parent.parent
PROC_DIR = PROJECT / "data" / "processed"
FIG_DIR  = PROJECT / "figures"
FIG_DIR.mkdir(parents=True, exist_ok=True)

plt.rcParams.update({
    "figure.dpi": 120, "savefig.dpi": 200, "font.family": "DejaVu Sans",
    "axes.grid": True, "grid.alpha": 0.3,
    "axes.spines.top": False, "axes.spines.right": False,
})

TICKERS = [
    "AAPL", "MSFT", "GOOGL", "NVDA", "META", "AMZN", "TSLA",
    "JPM", "BAC", "GS", "WFC",
    "JNJ", "UNH", "PFE", "LLY", "MRK",
    "WMT", "PG", "KO", "PEP", "HD", "NKE", "DIS",
    "XOM", "CVX",
    "BA", "CAT", "GE",
    "T", "VZ",
]


def fetch_headlines():
    """Pull recent news headlines via yfinance for each ticker."""
    import yfinance as yf
    rows = []
    for t in TICKERS:
        try:
            news = yf.Ticker(t).news or []
        except Exception as e:
            print(f"  {t}: error {e}")
            continue
        for item in news:
            c = item.get("content", item)
            title = c.get("title")
            pub   = c.get("pubDate")
            summary = c.get("summary") or c.get("description") or ""
            provider = (c.get("provider") or {}).get("displayName", "?")
            if title and pub:
                rows.append({
                    "ticker":   t,
                    "title":    title,
                    "summary":  summary[:500],
                    "pubDate":  pub,
                    "provider": provider,
                })
        print(f"  {t:6s}  {len(news):3d} headlines")
    df = pd.DataFrame(rows)
    df["pubDate"] = pd.to_datetime(df["pubDate"], utc=True, errors="coerce")
    df["date"] = df["pubDate"].dt.tz_convert(None).dt.normalize()
    return df


def score_with_finbert(texts):
    """Run ProsusAI/finbert on a list of strings; return continuous score in [-1, 1]."""
    from transformers import AutoTokenizer, AutoModelForSequenceClassification
    import torch

    print("Loading FinBERT (ProsusAI/finbert)...")
    tok = AutoTokenizer.from_pretrained("ProsusAI/finbert")
    mdl = AutoModelForSequenceClassification.from_pretrained("ProsusAI/finbert")
    mdl.eval()

    # FinBERT label order: positive=0, negative=1, neutral=2 (per HF model card)
    # Score = P(positive) - P(negative)
    scores = []
    BATCH = 16
    with torch.no_grad():
        for i in range(0, len(texts), BATCH):
            batch = texts[i:i + BATCH]
            enc = tok(batch, padding=True, truncation=True, max_length=128,
                      return_tensors="pt")
            logits = mdl(**enc).logits
            probs = torch.softmax(logits, dim=-1).numpy()
            # ProsusAI/finbert label order: positive, negative, neutral
            score = probs[:, 0] - probs[:, 1]
            scores.extend(score.tolist())
            print(f"  scored {min(i+BATCH, len(texts))}/{len(texts)}")
    return np.array(scores)


def main():
    print("=== Step 1: pulling recent headlines via yfinance ===")
    headlines = fetch_headlines()
    print(f"\nTotal headlines: {len(headlines)}")
    if len(headlines) == 0:
        print("No headlines retrieved -- skipping FinBERT scoring")
        return
    print(f"Date range: {headlines['date'].min()} -> {headlines['date'].max()}")

    print("\n=== Step 2: scoring with FinBERT ===")
    t0 = time.time()
    headlines["finbert_score"] = score_with_finbert(
        (headlines["title"] + ". " + headlines["summary"]).tolist()
    )
    print(f"FinBERT scoring took {time.time() - t0:.1f}s")
    headlines.to_csv(PROC_DIR / "finbert_headlines.csv", index=False)
    print(f"\nWrote {PROC_DIR/'finbert_headlines.csv'}")

    # Sample output
    print("\n--- Sample scored headlines ---")
    print(headlines.sort_values("finbert_score").head(3)[["ticker","date","title","finbert_score"]].to_string(index=False))
    print("...")
    print(headlines.sort_values("finbert_score").tail(3)[["ticker","date","title","finbert_score"]].to_string(index=False))

    print("\n=== Step 3: daily aggregation ===")
    daily_finbert = (
        headlines.groupby(["date", "ticker"])["finbert_score"]
                 .agg(["mean", "count"])
                 .reset_index()
    )
    daily_finbert.columns = ["date", "ticker", "finbert_mean", "n_headlines"]
    print(f"Daily (date, ticker) cells: {len(daily_finbert)}")

    print("\n=== Step 4: comparing to Bloomberg ===")
    panel = pd.read_csv(PROC_DIR / "panel.csv", parse_dates=["date"])
    bb = panel[["date", "ticker", "sentiment"]].rename(columns={"sentiment": "bloomberg_sentiment"})
    merged = daily_finbert.merge(bb, on=["date", "ticker"], how="inner")
    merged.to_csv(PROC_DIR / "finbert_vs_bloomberg.csv", index=False)
    print(f"Merged rows: {len(merged)}")
    print(f"Wrote {PROC_DIR/'finbert_vs_bloomberg.csv'}")

    pair = merged.dropna(subset=["finbert_mean", "bloomberg_sentiment"])
    print(f"Valid (non-null) pairs: {len(pair)} of {len(merged)}")

    if len(pair) >= 5:
        rho_p, p_p = stats.pearsonr(pair["finbert_mean"], pair["bloomberg_sentiment"])
        rho_s, p_s = stats.spearmanr(pair["finbert_mean"], pair["bloomberg_sentiment"])
        print(f"\nPearson correlation:  {rho_p:+.3f}  (p = {p_p:.4f})")
        print(f"Spearman correlation: {rho_s:+.3f}  (p = {p_s:.4f})")
        # also reference statistics for the report
        with open(PROC_DIR / "finbert_correlation.txt", "w", encoding="utf-8") as f:
            f.write("FinBERT vs Bloomberg sentiment correlation\n")
            f.write(f"N (matched daily ticker pairs): {len(pair)}\n")
            f.write(f"Pearson  rho = {rho_p:+.4f}, p = {p_p:.4f}\n")
            f.write(f"Spearman rho = {rho_s:+.4f}, p = {p_s:.4f}\n")
            f.write(f"\nFinBERT mean = {pair['finbert_mean'].mean():+.4f}, std = {pair['finbert_mean'].std():.4f}\n")
            f.write(f"Bloomberg mean = {pair['bloomberg_sentiment'].mean():+.4f}, std = {pair['bloomberg_sentiment'].std():.4f}\n")
        merged = pair

        fig, axes = plt.subplots(1, 2, figsize=(13, 5))
        axes[0].scatter(merged["finbert_mean"], merged["bloomberg_sentiment"],
                        alpha=0.6, s=30, color="#1f78b4")
        axes[0].axhline(0, color="black", lw=0.5)
        axes[0].axvline(0, color="black", lw=0.5)
        # 45-degree fit line
        if merged["finbert_mean"].std() > 0:
            slope, intercept, *_ = stats.linregress(merged["finbert_mean"], merged["bloomberg_sentiment"])
            x = np.linspace(merged["finbert_mean"].min(), merged["finbert_mean"].max(), 50)
            axes[0].plot(x, slope * x + intercept, color="#e31a1c", lw=1.4,
                         label=f"OLS fit: y = {slope:+.3f} x + {intercept:+.3f}")
            axes[0].legend(loc="upper left")
        axes[0].set_xlabel("FinBERT daily mean score")
        axes[0].set_ylabel("Bloomberg NEWS_SENTIMENT_DAILY_AVG")
        axes[0].set_title(
            f"FinBERT vs Bloomberg sentiment\n"
            f"Spearman = {rho_s:+.3f} (p = {p_s:.4f})  |  N = {len(merged)}"
        )

        axes[1].hist(merged["finbert_mean"], bins=20, alpha=0.6,
                     color="#1f78b4", label="FinBERT")
        axes[1].hist(merged["bloomberg_sentiment"], bins=20, alpha=0.6,
                     color="#e31a1c", label="Bloomberg")
        axes[1].axvline(0, color="black", lw=0.7)
        axes[1].set_xlabel("Sentiment score")
        axes[1].set_ylabel("Count")
        axes[1].set_title("Score distributions: FinBERT vs Bloomberg")
        axes[1].legend()
        fig.tight_layout()
        fig.savefig(FIG_DIR / "fig_finbert_vs_bloomberg.png")
        plt.close(fig)
        print(f"Wrote {FIG_DIR/'fig_finbert_vs_bloomberg.png'}")
    else:
        print(f"WARN: only {len(pair)} valid pairs -- correlation skipped")

    print("\nDone.")


if __name__ == "__main__":
    main()
