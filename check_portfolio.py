import os, time, sqlite3

FOLDER = os.path.join(os.path.expanduser("~"), "portfolio-system")
DB = os.path.join(FOLDER, "portfolio.db")
XLSX = os.path.join(FOLDER, "portfolio_dashboard.xlsx")


def text_of(name):
    try:
        with open(os.path.join(FOLDER, name), encoding="utf-8") as f:
            return f.read()
    except Exception:
        return None


print("=" * 60)
print("PORTFOLIO CHECK  (this only reads things, it changes nothing)")
print("=" * 60)
print("Folder:", FOLDER)

# ---------- 1. are the files the newest versions? ----------
print("\n1) Are your scripts up to date?")
checks = [
    ("run_all_scripts.py", "sync_from_excel.py", "must include the sync_from_excel.py step"),
    ("portfolio_tools.py", "def sync_from_workbook", "replace with the newest portfolio_tools.py"),
    ("sync_from_excel.py", "sync_from_workbook", "replace with the newest sync_from_excel.py"),
    ("export_to_excel.py", "Interested", "replace with the newest export_to_excel.py"),
    ("update_portfolio.py", "daily_quotes", "replace with the newest update_portfolio.py"),
    ("add_stock_interactive.py", "portfolio_tools", "replace with the newest add_stock_interactive.py"),
]
for name, needle, fix in checks:
    t = text_of(name)
    if t is None:
        print(f"   MISSING  {name}  ({fix})")
    elif needle in t:
        print(f"   ok       {name}")
    else:
        print(f"   OLD      {name}  ({fix})")
ra = text_of("run_all_scripts.py") or ""
if '"add_holding.py"' in ra or '"add_holdings.py"' in ra:
    print("   WARNING  run_all_scripts.py still runs the add_holding(s) step. It re-applies the numbers typed in that file")
    print("            on every run, which can undo your updates. Use the newest run_all_scripts.py.")

# ---------- 2. what is actually saved in the database ----------
print("\n2) What is saved in your database (this is the real source of truth)")
if not os.path.exists(DB):
    print("   No database found at", DB)
else:
    conn = sqlite3.connect(DB)
    cur = conn.cursor()

    def rows(sql):
        try:
            return cur.execute(sql).fetchall()
        except sqlite3.OperationalError:
            return []

    for t, s, a in rows("SELECT ticker, shares, avg_cost FROM holdings ORDER BY ticker"):
        print(f"   {t:<7} {s:>10g} shares   avg cost ${a:,.2f}")
    cash = rows("SELECT amount FROM cash ORDER BY date DESC LIMIT 1")
    print("   Buying power (cash):", f"${cash[0][0]:,.2f}" if cash else "not set")
    w = rows("SELECT ticker FROM watchlist ORDER BY ticker")
    print("   Interested list:", ", ".join(x[0] for x in w) if w else "(empty)")
    pend = rows("SELECT ticker, shares, avg_cost, action FROM pending_stocks")
    if pend:
        print("   Waiting for your answers:", pend)

    # ---------- 3. the Excel file ----------
    print("\n3) The Excel file")
    if not os.path.exists(XLSX):
        print("   No dashboard yet at", XLSX)
    else:
        print("   Path:", XLSX)
        print("   Last saved:", time.ctime(os.path.getmtime(XLSX)))
        handled = rows("SELECT value FROM meta WHERE key = 'handled_mtime'")
        now_m = str(os.stat(XLSX).st_mtime_ns)
        if handled and handled[0][0] == now_m:
            print("   Changed since the last refresh? NO. So Python will not read anything from Excel yet.")
            print("   (If you typed something in Excel, save it with Ctrl+S, then run the refresh.)")
        else:
            print("   Changed since the last refresh? YES. The next refresh will read what you typed.")
        try:
            from openpyxl import load_workbook
            wb = load_workbook(XLSX, data_only=True)
            print("   Sheets:", ", ".join(wb.sheetnames))
            if "Add Stock" in wb.sheetnames:
                ws = wb["Add Stock"]
                typed = []
                for r in range(6, 31):
                    t = ws.cell(row=r, column=1).value
                    if t is not None and str(t).strip():
                        typed.append(tuple(ws.cell(row=r, column=c).value for c in range(1, 5)))
                print("   Typed on the Add Stock sheet:", typed if typed else "(nothing)")
            else:
                print("   There is no 'Add Stock' sheet, so this is an OLD dashboard file. Run the refresh once with the newest export_to_excel.py.")
            if "Summary" in wb.sheetnames:
                print("   Buying power cell in Excel:", wb["Summary"]["I5"].value)
            if "Holdings" in wb.sheetnames:
                ws = wb["Holdings"]
                print("   What Excel's Holdings sheet currently shows:")
                for r in range(4, 30):
                    t = ws.cell(row=r, column=1).value
                    if t is None or t == "Total":
                        break
                    print(f"      {t:<7} {ws.cell(row=r, column=2).value} shares")
        except PermissionError:
            print("   Could not read the file (is it open?). Close it and run this again.")
        except Exception as e:
            print("   Could not read the file:", e)
    conn.close()

print("\n" + "=" * 60)
print("HOW UPDATES WORK")
print("  These change your holdings:")
print("    - add_stock_interactive.py (answer the questions)")
print("    - the Add Stock sheet in Excel (then save, close, and run run_all_scripts.py)")
print("    - the yellow Buying Power cell on the Summary sheet")
print("  These do NOT change your holdings:")
print("    - typing over the numbers on the Holdings or Summary sheets (they are rebuilt from the database)")
print("    - editing the STOCKS list in add_holdings.py without running that file")
print("=" * 60)
