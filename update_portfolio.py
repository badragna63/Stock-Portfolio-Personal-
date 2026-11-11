import os, sqlite3
import yfinance as yf

folder = os.path.join(os.path.expanduser("~"), "portfolio-system")
db_path = os.path.join(folder, "portfolio.db")

conn = sqlite3.connect(db_path)
cur = conn.cursor()

cur.execute("""
    CREATE TABLE IF NOT EXISTS prices (
        date TEXT,
        ticker TEXT,
        close REAL,
        PRIMARY KEY (date, ticker)
    )
""")

cur.execute("""
    CREATE TABLE IF NOT EXISTS snapshots (
        date TEXT PRIMARY KEY,
        total_value REAL,
        cash REAL,
        cost_basis REAL,
        gain_loss REAL
    )
""")

cur.execute("""
    CREATE TABLE IF NOT EXISTS daily_quotes (
        date TEXT,
        ticker TEXT,
        price REAL,
        prev_close REAL,
        PRIMARY KEY (date, ticker)
    )
""")
cur.execute("CREATE TABLE IF NOT EXISTS watchlist (ticker TEXT PRIMARY KEY, notes TEXT, date_added TEXT)")

row = cur.execute("SELECT amount FROM cash ORDER BY date DESC LIMIT 1").fetchone()
cash = row[0] if row else 0.0

positions_value = 0.0
cost_basis = 0.0

holdings = cur.execute("SELECT ticker, shares, avg_cost FROM holdings").fetchall()

print("Fetching live prices...")
for ticker, shares, avg_cost in holdings:
    price = None
    t = yf.Ticker(ticker)
    
    try:
        price = float(t.fast_info["last_price"])
    except Exception:
        try:
            hist = t.history(period="5d")
            if not hist.empty:
                price = float(hist["Close"].iloc[-1])
        except Exception as e:
            print(f"❌ Could not get price for {ticker}: {e}")

    if price is None:
        price = avg_cost
        print(f"⚠️ Warning: Using average cost (${avg_cost}) for {ticker}")

    positions_value += price * shares
    cost_basis += avg_cost * shares
    
    cur.execute(
        "INSERT OR REPLACE INTO prices (date, ticker, close) VALUES (date('now','localtime'), ?, ?)",
        (ticker, price)
    )
    prev = None
    try:
        prev = float(t.fast_info["previous_close"])
    except Exception:
        pass
    cur.execute(
        "INSERT OR REPLACE INTO daily_quotes (date, ticker, price, prev_close) VALUES (date('now','localtime'), ?, ?, ?)",
        (ticker, price, prev)
    )
    print(f"✓ {ticker:<6} : ${price:.2f}")

try:
    sp_ticker = yf.Ticker("^GSPC")
    try:
        sp = float(sp_ticker.fast_info["last_price"])
    except Exception:
        sp = float(sp_ticker.history(period="5d")["Close"].iloc[-1])
    cur.execute(
        "INSERT OR REPLACE INTO prices (date, ticker, close) VALUES (date('now','localtime'), '^GSPC', ?)",
        (sp,)
    )
    print(f"✓ S&P 500: ${sp:.2f}")
except Exception as e:
    print("Could not update S&P 500 index -", e)

print("\nFetching Interested-list prices...")
try:
    watch = [r[0] for r in cur.execute("SELECT ticker FROM watchlist")]
except Exception:
    watch = []
held = {h[0] for h in holdings}
for wt in watch:
    if wt in held:
        continue
    tt = yf.Ticker(wt)
    wprice = wprev = None
    try:
        wprice = float(tt.fast_info["last_price"])
    except Exception:
        try:
            whist = tt.history(period="5d")
            if not whist.empty:
                wprice = float(whist["Close"].iloc[-1])
        except Exception:
            pass
    try:
        wprev = float(tt.fast_info["previous_close"])
    except Exception:
        pass
    if wprice is None:
        print(f"Could not get price for {wt} (watching)")
        continue
    cur.execute(
        "INSERT OR REPLACE INTO daily_quotes (date, ticker, price, prev_close) VALUES (date('now','localtime'), ?, ?, ?)",
        (wt, wprice, wprev)
    )
    print(f"  {wt:<6} : ${wprice:.2f}  (watching)")

total_value = positions_value + cash
gain_loss = positions_value - cost_basis

cur.execute(
    "INSERT OR REPLACE INTO snapshots (date, total_value, cash, cost_basis, gain_loss) VALUES (date('now','localtime'), ?, ?, ?, ?)",
    (total_value, cash, cost_basis, gain_loss)
)

conn.commit()
conn.close()

print("\n--- Daily Snapshot Saved ---")
print(f"Positions Value : ${positions_value:,.2f}")
print(f"Cash            : ${cash:,.2f}")
print(f"Total Value     : ${total_value:,.2f}")
print(f"Gain / Loss     : ${gain_loss:,.2f}")