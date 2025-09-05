# TopGainersShortingAlgo
I’m releasing my AlphaTrading setup to the community. It’s built for U.S. micro/small caps and focuses on shorting overbought movers with a simple, configurable rule set and an Excel (xlwings) UI.

Core strategy (default config):

Universe: Market cap < $500M, price $1–$20
Daily scan: Top gainers under the cap/price filters
Entry: RSI > 75 → short (if shares are borrowable)
Exit targets: +3% take-profit, 70% stop-loss (both configurable)
Risk guardrails: $1,000 max per position, max 10 concurrent positions
Extras: configurable filters (change %, total/avg value, price bands, RSI), and an “average-deficit” safety option
Live Excel dashboard shows current portfolio, P&L, and signals (via xlwings)

Results so far:
Personal live trading log shows ~94% win rate across recent sample sets.
Not financial advice; results will vary.

Tech stack:
Python + ib_insync talking to IBKR (TWS or IB Gateway)
Excel UI via xlwings
Runs locally; parameters editable in Excel

Get the code
GitHub (public): https://github.com/thangavasanth/TopGainersShortingAlgo

Contributions welcome! Ideas for add-ons: borrow-availability checks, borrow fee awareness, volatility filters, better position sizing, alternative signals (e.g., %VWAP, parabolic SAR), broker adapters, and backtesting module.

Prerequisites (summary)
IBKR account (margin enabled) and TWS or IB Gateway installed. 
Anaconda (Python) and Spyder (optional IDE)
Python packages: ib_insync, xlwings (plus their requirements).

Quick install (Windows/macOS)
# 1) Install Anaconda (then open "Anaconda Prompt" / Terminal)
# 2) Install libs in requirement.txt 
# 3) (Optional) Install xlwings Excel add-in for buttons/macros



IBKR setup
Open/upgrade to IBKR margin account (required for shorting). 
Install Trader Workstation (TWS) or IB Gateway and enable API access (recommended fixed API port; allow read/open orders). 
Verify you can log in and receive market data (paper trading first).

Project setup
Clone the repo → update Yaali_Algo_Trading.xlsm (caps, price range, RSI, TP/SL, max positions, per-trade cap).
Launch IB Gateway → run python run.py → Excel UI opens and streams signals/portfolio.
Start in paper mode; switch to live only after you’re comfortable.

# https://buymeacoffee.com/vasanththangasamy)

Call for collaborators:
If you have ideas to improve entries/exits, risk controls, or broker coverage, please fork & PR. I’m happy to merge good additions and credit contributors.

⚠️ Disclaimer
This project is shared for educational and research purposes only.
It is not financial advice and should not be considered an invitation to buy, sell, or short any security.
Trading, especially short-selling, involves substantial risk — including the risk of losing more than your initial capital. Past performance (such as win rates in testing) does not guarantee future results.
Always:
Use paper trading before going live.
Understand the risks of margin accounts and leverage.
Consult a licensed financial advisor if you are unsure.
By using this code, you agree that I take no responsibility for any financial losses, damages, or outcomes resulting from its use.

#excel #python #ibkr #trading #quant #xlwings #ibinsync #shortselling #opensource
