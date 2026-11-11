"""
Shared helper code for your portfolio system.
You never need to open or edit this file. The other scripts use it.
"""
import os, sys, sqlite3

FOLDER = os.path.join(os.path.expanduser("~"), "portfolio-system")
DB_PATH = os.path.join(FOLDER, "portfolio.db")
DASHBOARD = os.path.join(FOLDER, "portfolio_dashboard.xlsx")
ADD_SHEET = "Add Stock"
FIRST_ROW = 6          # first input row on the Add Stock sheet
LAST_ROW = 60
INTERESTED_SHEET = "Interested"
INTERESTED_FIRST_ROW = 4
BUYING_POWER_CELL = "I5"   # on the Summary sheet


class Skip(Exception):
    """Raised when you type 'skip' to cancel the stock you are working on."""


_interactive = sys.stdin is not None


def can_ask():
    return _interactive


# ----------------------------------------------------------------- database
def connect():
    os.makedirs(FOLDER, exist_ok=True)
    conn = sqlite3.connect(DB_PATH)
    cur = conn.cursor()
    cur.execute("CREATE TABLE IF NOT EXISTS holdings (ticker TEXT PRIMARY KEY, shares REAL, avg_cost REAL, sector TEXT)")
    cur.execute("CREATE TABLE IF NOT EXISTS cash (date TEXT PRIMARY KEY, amount REAL)")
    cur.execute("CREATE TABLE IF NOT EXISTS watchlist (ticker TEXT PRIMARY KEY, notes TEXT, date_added TEXT)")
    cur.execute("CREATE TABLE IF NOT EXISTS pending_stocks (ticker TEXT PRIMARY KEY, shares REAL, avg_cost REAL, action TEXT)")
    cur.execute("CREATE TABLE IF NOT EXISTS meta (key TEXT PRIMARY KEY, value TEXT)")
    conn.commit()
    return conn


def get_meta(conn, key):
    row = conn.execute("SELECT value FROM meta WHERE key = ?", (key,)).fetchone()
    return row[0] if row else None


def set_meta(conn, key, value):
    conn.execute("INSERT OR REPLACE INTO meta (key, value) VALUES (?, ?)", (key, value))
    conn.commit()


def get_holding(conn, ticker):
    return conn.execute("SELECT shares, avg_cost FROM holdings WHERE ticker = ?", (ticker,)).fetchone()


def is_watched(conn, ticker):
    return conn.execute("SELECT 1 FROM watchlist WHERE ticker = ?", (ticker,)).fetchone() is not None


def save_position(conn, ticker, shares, avg_cost):
    conn.execute(
        "INSERT INTO holdings (ticker, shares, avg_cost) VALUES (?, ?, ?) "
        "ON CONFLICT(ticker) DO UPDATE SET shares = excluded.shares, avg_cost = excluded.avg_cost",
        (ticker, shares, avg_cost))
    if is_watched(conn, ticker):
        conn.execute("DELETE FROM watchlist WHERE ticker = ?", (ticker,))
        print(f"{ticker} moved from your Interested list into your holdings.")
    conn.commit()


def add_watch(conn, ticker):
    conn.execute("INSERT OR IGNORE INTO watchlist (ticker, notes, date_added) VALUES (?, NULL, date('now','localtime'))",
                 (ticker,))
    conn.commit()


def get_cash(conn):
    row = conn.execute("SELECT amount FROM cash ORDER BY date DESC LIMIT 1").fetchone()
    return row[0] if row else 0.0


def set_cash(conn, amount):
    # only the current balance is used anywhere, so keep a single row
    conn.execute("DELETE FROM cash")
    conn.execute("INSERT INTO cash (date, amount) VALUES (date('now','localtime'), ?)", (amount,))
    conn.commit()


def list_holdings(conn):
    rows = conn.execute("SELECT ticker, shares, avg_cost FROM holdings WHERE shares > 0 ORDER BY ticker").fetchall()
    print("\n--- Your holdings ---")
    if not rows:
        print("  (none yet)")
    for t, s, a in rows:
        print(f"  {t:<7} {s:>10g} shares   avg cost ${a:,.2f}")
    print(f"  Buying power (cash): ${get_cash(conn):,.2f}")
    watch = conn.execute("SELECT ticker FROM watchlist ORDER BY ticker").fetchall()
    print("\n--- Interested (watching, 0 shares) ---")
    print("  " + (", ".join(w[0] for w in watch) if watch else "(none yet)"))


# ------------------------------------------------------------------ asking
def to_num(value):
    if value is None:
        return None
    if isinstance(value, (int, float)):
        return float(value)
    text = str(value).replace("$", "").replace(",", "").strip()
    if not text:
        return None
    try:
        return float(text)
    except ValueError:
        return None


def ask(prompt):
    """Ask a question in the console. Returns None if nobody can answer (scheduled run)."""
    global _interactive
    if not _interactive:
        return None
    try:
        text = input(prompt)
    except (EOFError, RuntimeError, OSError):
        _interactive = False
        return None
    text = text.strip()
    if text.lower() in ("skip", "cancel", "q", "quit"):
        raise Skip()
    return text


def ask_number(prompt, default=None, allow_zero=False):
    while True:
        text = ask(prompt)
        if text is None:
            return None
        if text == "" and default is not None:
            return default
        n = to_num(text)
        if n is not None and (n > 0 or (allow_zero and n == 0)):
            return n
        print("  Please type a number like 2.5 (or type skip to cancel).")


def yes_no(prompt):
    while True:
        text = ask(prompt)
        if text is None:
            return None
        if text.lower() in ("y", "yes"):
            return True
        if text.lower() in ("n", "no"):
            return False
        print("  Please type y or n.")


def lookup_price(ticker):
    try:
        import yfinance as yf
        p = float(yf.Ticker(ticker).fast_info["last_price"])
        return p if p > 0 else None
    except Exception:
        return None


# ------------------------------------------------------------------- flows
def watch_flow(conn, ticker):
    """Put a stock on the Interested list (0 shares): it gets prices and news but is not a holding."""
    existing = get_holding(conn, ticker)
    if existing and existing[0] > 0:
        print(f"{ticker}: you already own {existing[0]:g} shares, so it is already tracked as a holding.")
        return "skipped"
    if is_watched(conn, ticker):
        print(f"{ticker} is already on your Interested list.")
        return "skipped"
    if can_ask():
        price = lookup_price(ticker)
        if price is None:
            print(f"I could not find a current price for {ticker}. Check the spelling.")
            ok = yes_no("Add it anyway? (y/n): ")
            if ok is None:
                return "pending"
            if not ok:
                raise Skip()
    add_watch(conn, ticker)
    print(f"Added {ticker} to your Interested list. It will get prices and news on the next refresh.")
    return "applied"


def add_flow(conn, ticker, shares, avg):
    existing = get_holding(conn, ticker)
    if existing and existing[0] <= 0:
        existing = None          # an old 0-share row counts as "not owned"

    # Nobody is there to answer questions (a scheduled run): only do the clear-cut cases.
    if not can_ask():
        if existing is None and shares is not None and shares == 0:
            return watch_flow(conn, ticker)
        price = lookup_price(ticker)
        if existing is None and shares and avg and price:
            save_position(conn, ticker, shares, avg)
            print(f"Added {ticker}: {shares:g} shares at ${avg:,.2f}")
            return "applied"
        print(f"{ticker}: waiting for your answers (run the refresh in Spyder).")
        return "pending"

    price = lookup_price(ticker)
    print(f"\n--- {ticker} ---")
    if price is None:
        print("I could not find a current price for this ticker. Check the spelling.")
        ok = yes_no("Add it anyway? (y/n): ")
        if ok is None:
            return "pending"
        if not ok:
            raise Skip()
    else:
        print(f"Current price: ${price:,.2f}")

    mode = "new"
    if existing:
        old_sh, old_avg = existing
        print(f"You already own {old_sh:g} shares at an average cost of ${old_avg:,.2f}.")
        while True:
            c = ask("Did you (1) buy more shares, (2) replace it with new totals, or (3) skip? [1/2/3]: ")
            if c is None:
                return "pending"
            if c == "1":
                mode = "add"
                break
            if c == "2":
                mode = "replace"
                break
            if c == "3":
                raise Skip()
            print("  Please type 1, 2, or 3.")

    default_avg = None
    if mode == "add":
        q_sh = "How many shares did you buy? "
        q_avg = "What price per share did you pay? " + (f"(press Enter for ${price:,.2f}) " if price else "")
        default_avg = price
    elif mode == "replace":
        q_sh = "What is your new TOTAL number of shares? "
        q_avg = "What is your new average cost per share? "
    else:
        q_sh = "How many shares do you own? (type 0 to just watch it) "
        q_avg = "What is your average cost per share? "

    if shares is None:
        shares = ask_number(q_sh, allow_zero=(mode == "new"))
        if shares is None:
            return "pending"
    else:
        print(f"Shares (from Excel): {shares:g}")

    if mode == "new" and shares == 0:
        return watch_flow(conn, ticker)

    if avg is None:
        avg = ask_number(q_avg, default=default_avg)
        if avg is None:
            return "pending"
    else:
        print(f"Price per share (from Excel): ${avg:,.2f}")

    lot_cost = None
    if mode == "add":
        total = old_sh + shares
        new_avg = (old_sh * old_avg + shares * avg) / total
        lot_cost = shares * avg
        print(f"\nNew position: {total:g} shares at an average of ${new_avg:,.2f}"
              f"   (was {old_sh:g} shares at ${old_avg:,.2f})")
    else:
        total, new_avg = shares, avg
        if mode == "new":
            lot_cost = shares * avg
        print(f"\nNew position: {total:g} shares at an average of ${new_avg:,.2f}")

    ok = yes_no("Save this? (y/n): ")
    if ok is None:
        return "pending"
    if not ok:
        raise Skip()
    save_position(conn, ticker, total, new_avg)
    print(f"Saved {ticker}.")

    if lot_cost:
        try:
            cash = get_cash(conn)
            take = yes_no(f"Subtract ${lot_cost:,.2f} from your buying power (now ${cash:,.2f})? "
                          "(y/n, say n if you already owned these): ")
            if take:
                set_cash(conn, cash - lot_cost)
                print(f"Buying power is now ${cash - lot_cost:,.2f}.")
        except Skip:
            pass
    return "applied"


def remove_flow(conn, ticker, sold):
    existing = get_holding(conn, ticker)
    if existing is None or existing[0] <= 0:
        if is_watched(conn, ticker):
            if not can_ask():
                return "pending"
            try:
                ok = yes_no(f"Remove {ticker} from your Interested list? (y/n): ")
            except Skip:
                return "skipped"
            if ok is None:
                return "pending"
            if ok:
                conn.execute("DELETE FROM watchlist WHERE ticker = ?", (ticker,))
                conn.commit()
                print(f"Removed {ticker} from your Interested list.")
                return "applied"
            return "skipped"
        print(f"{ticker} is not in your holdings or your Interested list, nothing to remove.")
        return "skipped"
    if not can_ask():
        print(f"{ticker}: waiting for you to confirm the sale (run the refresh in Spyder).")
        return "pending"

    owned, avg = existing
    price = lookup_price(ticker)
    print(f"\n--- {ticker} ---")
    print(f"You own {owned:g} shares at an average cost of ${avg:,.2f}.")
    if sold is None:
        sold = ask_number("How many shares did you sell? (press Enter to remove the whole position) ",
                          default=owned)
        if sold is None:
            return "pending"
    else:
        print(f"Shares sold (from Excel): {sold:g}")
    if sold >= owned:
        sold = owned
        print(f"This removes {ticker} from your portfolio completely.")
    else:
        print(f"After the sale you will own {owned - sold:g} shares (average cost stays ${avg:,.2f}).")

    ok = yes_no("Save this? (y/n): ")
    if ok is None:
        return "pending"
    if not ok:
        raise Skip()
    if sold >= owned:
        conn.execute("DELETE FROM holdings WHERE ticker = ?", (ticker,))
    else:
        conn.execute("UPDATE holdings SET shares = ? WHERE ticker = ?", (owned - sold, ticker))
    conn.commit()
    print(f"Saved {ticker}.")

    if price:
        try:
            proceeds = sold * price
            cash = get_cash(conn)
            add = yes_no(f"Add about ${proceeds:,.2f} (shares x current price) to your buying power "
                         f"(now ${cash:,.2f})? (y/n): ")
            if add:
                set_cash(conn, cash + proceeds)
                print(f"Buying power is now ${cash + proceeds:,.2f}.")
        except Skip:
            pass
    return "applied"


def process_row(conn, ticker, shares=None, avg=None, action=None):
    """Handle one stock. Returns 'applied', 'pending' (still needs answers) or 'skipped'."""
    ticker = (ticker or "").strip().upper()
    action = (action or "").strip().upper()
    if not ticker:
        return "skipped"
    try:
        if action in ("REMOVE", "SELL", "DELETE"):
            return remove_flow(conn, ticker, shares)
        if action in ("WATCH", "INTERESTED"):
            return watch_flow(conn, ticker)
        return add_flow(conn, ticker, shares, avg)
    except Skip:
        print(f"Skipped {ticker}.")
        return "skipped"


# ------------------------------------------------------------- Excel sync
def read_workbook_inputs(path):
    """Read what you typed in Excel: Add Stock rows, Interested notes, buying power.
    Returns None if the file can't be read (for example, it is open)."""
    from openpyxl import load_workbook
    try:
        wb = load_workbook(path, data_only=True)
    except Exception as e:
        print("Could not read the Excel file. Save and close it, then try again.", e)
        return None

    rows = []
    if ADD_SHEET in wb.sheetnames:
        ws = wb[ADD_SHEET]
        for r in range(FIRST_ROW, LAST_ROW + 1):
            t = ws.cell(row=r, column=1).value
            if t is None or not str(t).strip():
                continue
            action = ws.cell(row=r, column=4).value
            rows.append((str(t).strip().upper(),
                         to_num(ws.cell(row=r, column=2).value),
                         to_num(ws.cell(row=r, column=3).value),
                         str(action).strip().upper() if action else ""))

    notes = {}
    if INTERESTED_SHEET in wb.sheetnames:
        ws = wb[INTERESTED_SHEET]
        for r in range(INTERESTED_FIRST_ROW, INTERESTED_FIRST_ROW + 200):
            t = ws.cell(row=r, column=1).value
            if t is None or not str(t).strip():
                continue
            n = ws.cell(row=r, column=7).value
            notes[str(t).strip().upper()] = (str(n).strip() if n is not None and str(n).strip() else None)

    buying_power = None
    if "Summary" in wb.sheetnames:
        buying_power = to_num(wb["Summary"][BUYING_POWER_CELL].value)

    return {"rows": rows, "notes": notes, "buying_power": buying_power}


def sync_from_workbook():
    """Pick up what you typed in Excel and process it. Returns True if anything changed."""
    conn = connect()
    changed = False

    if os.path.exists(DASHBOARD):
        mtime = str(os.stat(DASHBOARD).st_mtime_ns)
        if get_meta(conn, "handled_mtime") != mtime:
            data = read_workbook_inputs(DASHBOARD)
            if data is not None:
                # 1) stocks typed on the Add Stock sheet
                for t, s, a, act in data["rows"]:
                    conn.execute("INSERT OR REPLACE INTO pending_stocks (ticker, shares, avg_cost, action) "
                                 "VALUES (?, ?, ?, ?)", (t, s, a, act))

                # 2) notes typed on the Interested sheet
                for t, note in data["notes"].items():
                    conn.execute("UPDATE watchlist SET notes = ? WHERE ticker = ?", (note, t))

                # 3) buying power typed on the Summary sheet (only if you actually changed it)
                bp = data["buying_power"]
                exported = get_meta(conn, "exported_cash")
                if bp is not None and exported is not None:
                    try:
                        if abs(bp - float(exported)) > 0.004:
                            set_cash(conn, bp)
                            print(f"Buying power updated from Excel: ${bp:,.2f}")
                            changed = True
                    except ValueError:
                        pass
                conn.commit()
                set_meta(conn, "handled_mtime", mtime)

    pending = conn.execute("SELECT ticker, shares, avg_cost, action FROM pending_stocks ORDER BY ticker").fetchall()
    if not pending:
        print("No new stocks waiting in Excel's 'Add Stock' sheet.")
        conn.close()
        return changed

    print(f"\nFound {len(pending)} stock(s) from Excel's 'Add Stock' sheet. Type 'skip' at any question to cancel one.")
    for t, s, a, act in pending:
        status = process_row(conn, t, s, a, act)
        if status in ("applied", "skipped"):
            conn.execute("DELETE FROM pending_stocks WHERE ticker = ?", (t,))
            conn.commit()
        if status == "applied":
            changed = True

    left = conn.execute("SELECT COUNT(*) FROM pending_stocks").fetchone()[0]
    if left:
        print(f"\n{left} stock(s) are still waiting for your answers. They stay on the Add Stock sheet.")
    conn.close()
    return changed
