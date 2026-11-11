import os, sys, runpy
sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
import portfolio_tools as pt

HERE = os.path.dirname(os.path.abspath(__file__))

MENU = """
==========================================
        PORTFOLIO MANAGER
==========================================
 1) Add a stock (0 shares = just watch it), or update one I bought more of
 2) Sell shares or remove a stock
 3) Update my buying power (cash)
 4) Check Excel for stocks I typed in the 'Add Stock' sheet
 5) See my holdings and Interested list
 6) Refresh everything now (prices, news, Excel)
 7) Quit
"""


def refresh_everything():
    for name in ("run_all_scripts.py", "run_all.py"):
        path = os.path.join(HERE, name)
        if os.path.exists(path):
            runpy.run_path(path, run_name="__main__")
            return True
    print("Could not find run_all_scripts.py in this folder.")
    return False


changed = False
while True:
    print(MENU)
    choice = input("Type a number (1-7) and press Enter: ").strip()
    conn = pt.connect()
    try:
        if choice == "1":
            t = input("Ticker (like AAPL), or press Enter to go back: ").strip().upper()
            if t:
                status = pt.process_row(conn, t)
                changed = changed or status == "applied"

        elif choice == "2":
            t = input("Ticker you sold (like AAPL), or press Enter to go back: ").strip().upper()
            if t:
                status = pt.process_row(conn, t, action="REMOVE")
                changed = changed or status == "applied"

        elif choice == "3":
            print(f"Your saved buying power is ${pt.get_cash(conn):,.2f}.")
            try:
                amount = pt.ask_number("What is your buying power now? (type skip to cancel) ", allow_zero=True)
                if amount is not None:
                    pt.set_cash(conn, amount)
                    print(f"Buying power saved: ${amount:,.2f}")
                    changed = True
            except pt.Skip:
                print("Cancelled.")

        elif choice == "4":
            print("Tip: save the Excel file (Ctrl+S) first so your changes are included.")
            changed = pt.sync_from_workbook() or changed

        elif choice == "5":
            pt.list_holdings(conn)

        elif choice == "6":
            conn.close()
            if refresh_everything():
                changed = False
            continue

        elif choice == "7":
            conn.close()
            if changed:
                again = input("You made changes. Refresh Excel now so it shows them? (y/n): ").strip().lower()
                if again in ("y", "yes"):
                    refresh_everything()
            print("Goodbye!")
            break

        else:
            print("Please type a number from 1 to 7.")
    finally:
        try:
            conn.close()
        except Exception:
            pass
