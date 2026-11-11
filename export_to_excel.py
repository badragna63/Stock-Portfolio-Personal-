import os, sys, sqlite3
from datetime import datetime
from openpyxl import Workbook, load_workbook
from openpyxl.utils import get_column_letter
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.worksheet.datavalidation import DataValidation
from openpyxl.formatting.rule import CellIsRule
from openpyxl.chart import PieChart, LineChart, Reference
from openpyxl.chart.axis import DateAxis
from openpyxl.chart.label import DataLabelList

folder = os.path.join(os.path.expanduser("~"), "portfolio-system")
db_path = os.path.join(folder, "portfolio.db")
out_path = os.path.join(folder, "portfolio_dashboard.xlsx")
try:
    sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
except NameError:
    pass

# Sheets this script rebuilds from your database every time. Any OTHER sheet you add
# (like "My Notes") is kept exactly as you left it.
GENERATED = ["Summary", "Holdings", "Interested", "Add Stock", "News", "History"]

# =====================================================================
#  READ EVERYTHING FROM THE DATABASE
# =====================================================================
conn = sqlite3.connect(db_path)
cur = conn.cursor()
for ddl in (
    "CREATE TABLE IF NOT EXISTS holdings (ticker TEXT PRIMARY KEY, shares REAL, avg_cost REAL, sector TEXT)",
    "CREATE TABLE IF NOT EXISTS cash (date TEXT PRIMARY KEY, amount REAL)",
    "CREATE TABLE IF NOT EXISTS prices (date TEXT, ticker TEXT, close REAL, PRIMARY KEY (date, ticker))",
    "CREATE TABLE IF NOT EXISTS snapshots (date TEXT PRIMARY KEY, total_value REAL, cash REAL, cost_basis REAL, gain_loss REAL)",
    "CREATE TABLE IF NOT EXISTS daily_quotes (date TEXT, ticker TEXT, price REAL, prev_close REAL, PRIMARY KEY (date, ticker))",
    "CREATE TABLE IF NOT EXISTS watchlist (ticker TEXT PRIMARY KEY, notes TEXT, date_added TEXT)",
    "CREATE TABLE IF NOT EXISTS pending_stocks (ticker TEXT PRIMARY KEY, shares REAL, avg_cost REAL, action TEXT)",
    "CREATE TABLE IF NOT EXISTS news (id INTEGER PRIMARY KEY AUTOINCREMENT, date TEXT, ticker TEXT, headline TEXT, source TEXT, url TEXT, summary TEXT)",
):
    cur.execute(ddl)
conn.commit()


def latest_close(t):
    r = cur.execute("SELECT close FROM prices WHERE ticker = ? ORDER BY date DESC LIMIT 1", (t,)).fetchone()
    return r[0] if r else None


def latest_quote(t):
    r = cur.execute("SELECT price, prev_close FROM daily_quotes WHERE ticker = ? ORDER BY date DESC LIMIT 1", (t,)).fetchone()
    return r if r else (None, None)


def prev_close_from_quotes(t):
    r = cur.execute("SELECT prev_close FROM daily_quotes WHERE ticker = ? AND prev_close IS NOT NULL "
                    "ORDER BY date DESC LIMIT 1", (t,)).fetchone()
    return r[0] if r else None


def previous_day_close(t):
    r = cur.execute("SELECT close FROM prices WHERE ticker = ? ORDER BY date DESC LIMIT 1 OFFSET 1", (t,)).fetchone()
    return r[0] if r else None


holdings = []   # (ticker, shares, avg_cost, price, prev_close)
for t, sh, ac in cur.execute("SELECT ticker, shares, avg_cost FROM holdings WHERE shares > 0 ORDER BY ticker").fetchall():
    price = latest_close(t) or latest_quote(t)[0] or ac
    prev = prev_close_from_quotes(t) or previous_day_close(t) or price
    holdings.append((t, sh, ac, price, prev))

held = {h[0] for h in holdings}
watch = []      # (ticker, price, prev_close, date_added, notes)
for t, notes, added in cur.execute("SELECT ticker, notes, date_added FROM watchlist ORDER BY ticker").fetchall():
    if t in held:
        continue
    p, pv = latest_quote(t)
    if p is None:
        p = latest_close(t)
    watch.append((t, p, pv, added, notes))

row = cur.execute("SELECT amount FROM cash ORDER BY date DESC LIMIT 1").fetchone()
cash = row[0] if row else 0.0

history = []    # (date, total_value, gain_loss, sp500)
for d, total, gl in cur.execute("SELECT date, total_value, gain_loss FROM snapshots ORDER BY date").fetchall():
    try:
        dd = datetime.strptime(d, "%Y-%m-%d").date()
    except Exception:
        continue
    sp = cur.execute("SELECT close FROM prices WHERE ticker = '^GSPC' AND date = ?", (d,)).fetchone()
    history.append((dd, total, gl, sp[0] if sp else None))

try:
    pending = cur.execute("SELECT ticker, shares, avg_cost, action FROM pending_stocks ORDER BY ticker").fetchall()
except sqlite3.OperationalError:
    pending = []

NEWS_PER_STOCK = 1
news_rows = []
for t in [h[0] for h in holdings] + [w[0] for w in watch]:
    label = t + (" (watching)" if t not in held else "")
    for d, headline, source, url, summary in cur.execute(
            "SELECT date, headline, source, url, summary FROM news WHERE ticker = ? "
            "ORDER BY date DESC, id DESC LIMIT ?", (t, NEWS_PER_STOCK)).fetchall():
        news_rows.append((label, d, headline, source, url, summary))
conn.close()

# =====================================================================
#  STYLES
# =====================================================================
FONT = "Calibri"
NAVY = "1F3864"
base = Font(name=FONT, size=11)
bold = Font(name=FONT, size=11, bold=True)
title_font = Font(name=FONT, size=16, bold=True, color=NAVY)
banner_font = Font(name=FONT, size=20, bold=True, color="FFFFFF")
subtitle_font = Font(name=FONT, size=10, italic=True, color="595959")
section_font = Font(name=FONT, size=13, bold=True, color=NAVY)
header_font = Font(name=FONT, size=11, bold=True, color="FFFFFF")
header_fill = PatternFill("solid", fgColor=NAVY)
card_fill = PatternFill("solid", fgColor="F2F5FA")
input_fill = PatternFill("solid", fgColor="FFF9DB")
input_font = Font(name=FONT, size=11, color="0000FF")
link_font = Font(name=FONT, size=10, color="0563C1", underline="single")
GREEN = PatternFill(start_color="E2F0D9", end_color="E2F0D9", fill_type="solid")
RED = PatternFill(start_color="FBE3E3", end_color="FBE3E3", fill_type="solid")

thin = Side(style="thin", color="D9D9D9")
thin_border = Border(left=thin, right=thin, top=thin, bottom=thin)

USD = '$#,##0.00;($#,##0.00);"-"'
PCT = '0.00%;(0.00%);"-"'


def style_header(ws, row, ncols, first_col=1):
    for c in range(first_col, first_col + ncols):
        cell = ws.cell(row=row, column=c)
        cell.font = header_font
        cell.fill = header_fill
        cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)


def green_red(ws, rng):
    """Conditional formatting: green when above zero, red when below zero."""
    ws.conditional_formatting.add(rng, CellIsRule(operator="greaterThan", formula=["0"],
                                                  font=Font(color="375623", bold=True), fill=GREEN))
    ws.conditional_formatting.add(rng, CellIsRule(operator="lessThan", formula=["0"],
                                                  font=Font(color="9C0006", bold=True), fill=RED))


def section(ws, row, text, last_col):
    ws.cell(row=row, column=1, value=text).font = section_font
    for c in range(1, last_col + 1):
        ws.cell(row=row, column=c).border = Border(bottom=Side(style="medium", color=NAVY))


# =====================================================================
#  WORKBOOK (reuse the existing file so your own sheets are kept)
# =====================================================================
fresh = True
wb = None
if os.path.exists(out_path):
    try:
        wb = load_workbook(out_path)
        for _name in GENERATED:
            if _name in wb.sheetnames:
                del wb[_name]
        fresh = False
    except Exception as e:
        print("Could not reuse the existing workbook, building a new one:", e)
        wb = None
if wb is None:
    wb = Workbook()
    fresh = True

n = len(holdings)
first_h = 4
last_h = first_h + n - 1 if n > 0 else first_h
total_row = last_h + 1


def htot(col):
    """Reference to a total on the Holdings sheet (0 if you have no holdings)."""
    return f"=Holdings!{col}{total_row}" if n else 0


# =====================================================================
#  SUMMARY
# =====================================================================
ss = wb.create_sheet("Summary", 0)
ss.sheet_view.showGridLines = False
for col in range(1, 11):
    ss.column_dimensions[get_column_letter(col)].width = 15

ss.merge_cells("A1:J1")
ss["A1"] = "Portfolio Dashboard"
ss["A1"].font = banner_font
ss["A1"].fill = header_fill
ss["A1"].alignment = Alignment(horizontal="left", vertical="center", indent=1)
ss.row_dimensions[1].height = 36
for c in range(1, 11):
    ss.cell(row=1, column=c).fill = header_fill
ss["A2"] = "Updated: " + datetime.now().strftime("%Y-%m-%d %H:%M")
ss["A2"].font = subtitle_font


def card(c1, c2, label, value, vfmt, sub, sfmt, is_input=False):
    L1, L2 = get_column_letter(c1), get_column_letter(c2)
    for r in (4, 5, 6):
        ss.merge_cells(f"{L1}{r}:{L2}{r}")
        for c in range(c1, c2 + 1):
            ss.cell(row=r, column=c).fill = card_fill
            ss.cell(row=r, column=c).border = thin_border
    lab = ss.cell(row=4, column=c1, value=label)
    lab.font = Font(name=FONT, size=9, bold=True, color="595959")
    lab.alignment = Alignment(horizontal="center", vertical="center")
    val = ss.cell(row=5, column=c1, value=value)
    val.number_format = vfmt
    val.font = Font(name=FONT, size=18, bold=True, color=NAVY)
    val.alignment = Alignment(horizontal="center", vertical="center")
    sb = ss.cell(row=6, column=c1, value=sub)
    sb.number_format = sfmt
    sb.font = Font(name=FONT, size=10, color="595959")
    sb.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
    if is_input:
        for c in range(c1, c2 + 1):
            ss.cell(row=5, column=c).fill = input_fill
        val.font = Font(name=FONT, size=18, bold=True, color="0000FF")


ss.row_dimensions[4].height = 20
ss.row_dimensions[5].height = 34
ss.row_dimensions[6].height = 30

today_dollars = (f"=SUMPRODUCT(Holdings!B{first_h}:B{last_h},Holdings!E{first_h}:E{last_h}-Holdings!J{first_h}:J{last_h})"
                 if n else 0)
card(1, 2, "TOTAL VALUE", f"=I5+Holdings!F{total_row}" if n else "=I5", USD, "Stocks + buying power", "@")
card(3, 4, "TODAY'S CHANGE ($)", today_dollars, USD, "=IFERROR(C5/(A5-C5),0)", PCT)
card(5, 6, "ALL-TIME GAIN / LOSS", htot("G"), USD, htot("H"), PCT)
card(7, 8, "INVESTED (STOCKS)", htot("F"), USD, htot("D"), '"Cost basis: "$#,##0.00')
card(9, 10, "BUYING POWER (type here)", cash, USD, "Edit this number. It saves to Python on your next refresh.", "@", is_input=True)
green_red(ss, "C5")
green_red(ss, "C6")
green_red(ss, "E5")
green_red(ss, "E6")

row = 8
section(ss, row, "Today's Movers", 10)
mv_headers = ["Ticker", "Shares", "Price", "Prev Close", "Today ($/share)", "Today (%)",
              "Your Change Today ($)", "Market Value", "% of Portfolio"]
for i, h in enumerate(mv_headers, 1):
    ss.cell(row=row + 1, column=i, value=h)
style_header(ss, row + 1, len(mv_headers))
ss.row_dimensions[row + 1].height = 32

mv_first = row + 2
if n:
    for i in range(n):
        r = mv_first + i
        hr = first_h + i
        ss.cell(row=r, column=1, value=f"=Holdings!A{hr}").font = bold
        ss.cell(row=r, column=2, value=f"=Holdings!B{hr}")
        ss.cell(row=r, column=3, value=f"=Holdings!E{hr}")
        ss.cell(row=r, column=4, value=f"=Holdings!J{hr}")
        ss.cell(row=r, column=5, value=f"=C{r}-D{r}")
        ss.cell(row=r, column=6, value=f"=IFERROR(E{r}/D{r},0)")
        ss.cell(row=r, column=7, value=f"=B{r}*E{r}")
        ss.cell(row=r, column=8, value=f"=Holdings!F{hr}")
        ss.cell(row=r, column=9, value=f"=IFERROR(H{r}/$A$5,0)")
        for c in range(1, 10):
            cell = ss.cell(row=r, column=c)
            cell.border = thin_border
            if c > 1:
                cell.font = base
        ss.cell(row=r, column=2).number_format = '#,##0.####'
        for c in (3, 4, 5, 7, 8):
            ss.cell(row=r, column=c).number_format = USD
        for c in (6, 9):
            ss.cell(row=r, column=c).number_format = PCT
    mv_last = mv_first + n - 1
    tot = mv_last + 1
    ss.cell(row=tot, column=1, value="Total").font = bold
    ss.cell(row=tot, column=7, value=f"=SUM(G{mv_first}:G{mv_last})")
    ss.cell(row=tot, column=8, value=f"=SUM(H{mv_first}:H{mv_last})")
    ss.cell(row=tot, column=6, value=f"=IFERROR(G{tot}/(H{tot}-G{tot}),0)")
    ss.cell(row=tot, column=9, value=f"=IFERROR(H{tot}/$A$5,0)")
    for c in range(1, 10):
        ss.cell(row=tot, column=c).border = Border(top=Side(style="thin", color=NAVY))
        if c > 1:
            ss.cell(row=tot, column=c).font = bold
    ss.cell(row=tot, column=7).number_format = USD
    ss.cell(row=tot, column=8).number_format = USD
    ss.cell(row=tot, column=6).number_format = PCT
    ss.cell(row=tot, column=9).number_format = PCT
    green_red(ss, f"E{mv_first}:G{mv_last}")
    green_red(ss, f"F{tot}:G{tot}")

    r = tot + 2
    ss.cell(row=r, column=1, value="Top gainer today").font = bold
    ss.cell(row=r, column=2, value=f'=IFERROR(INDEX(A{mv_first}:A{mv_last},MATCH(MAX(F{mv_first}:F{mv_last}),F{mv_first}:F{mv_last},0)),"")').font = bold
    ss.cell(row=r, column=3, value=f'=IFERROR(MAX(F{mv_first}:F{mv_last}),"")').number_format = PCT
    ss.cell(row=r + 1, column=1, value="Top loser today").font = bold
    ss.cell(row=r + 1, column=2, value=f'=IFERROR(INDEX(A{mv_first}:A{mv_last},MATCH(MIN(F{mv_first}:F{mv_last}),F{mv_first}:F{mv_last},0)),"")').font = bold
    ss.cell(row=r + 1, column=3, value=f'=IFERROR(MIN(F{mv_first}:F{mv_last}),"")').number_format = PCT
    green_red(ss, f"C{r}:C{r + 1}")
    row = r + 3
else:
    ss.cell(row=mv_first, column=1, value="No holdings yet. Add one with add_stock_interactive.py or the Add Stock sheet.").font = subtitle_font
    row = mv_first + 3

# ---- allocation tables (these feed the pie charts) ----
section(ss, row, "Where your money is", 10)
for i, h in enumerate(["Name", "Value", "% of Portfolio"], 1):
    ss.cell(row=row + 1, column=i, value=h)
style_header(ss, row + 1, 3)
al_first = row + 2
for i in range(n):
    r = al_first + i
    ss.cell(row=r, column=1, value=f"=Holdings!A{first_h + i}").font = base
    ss.cell(row=r, column=2, value=f"=Holdings!F{first_h + i}").number_format = USD
    ss.cell(row=r, column=3, value=f"=IFERROR(B{r}/$A$5,0)").number_format = PCT
r = al_first + n
ss.cell(row=r, column=1, value="Buying power").font = base
ss.cell(row=r, column=2, value="=$I$5").number_format = USD
ss.cell(row=r, column=3, value=f"=IFERROR(B{r}/$A$5,0)").number_format = PCT
al_last = r
for rr in range(al_first, al_last + 1):
    for c in range(1, 4):
        ss.cell(row=rr, column=c).border = thin_border
        if c > 1:
            ss.cell(row=rr, column=c).font = base

r = al_last + 2
section(ss, r, "Invested vs. cash", 10)
for i, h in enumerate(["Name", "Value", "% of Portfolio"], 1):
    ss.cell(row=r + 1, column=i, value=h)
style_header(ss, r + 1, 3)
iv_first = r + 2
ss.cell(row=iv_first, column=1, value="Stocks").font = base
ss.cell(row=iv_first, column=2, value=f"=Holdings!F{total_row}" if n else 0).number_format = USD
ss.cell(row=iv_first, column=3, value=f"=IFERROR(B{iv_first}/$A$5,0)").number_format = PCT
ss.cell(row=iv_first + 1, column=1, value="Buying power").font = base
ss.cell(row=iv_first + 1, column=2, value="=$I$5").number_format = USD
ss.cell(row=iv_first + 1, column=3, value=f"=IFERROR(B{iv_first + 1}/$A$5,0)").number_format = PCT
for rr in (iv_first, iv_first + 1):
    for c in range(1, 4):
        ss.cell(row=rr, column=c).border = thin_border
        if c > 1:
            ss.cell(row=rr, column=c).font = base


def pie(title, first, last, anchor):
    p = PieChart()
    p.title = title
    p.add_data(Reference(ss, min_col=2, min_row=first, max_row=last), titles_from_data=False)
    p.set_categories(Reference(ss, min_col=1, min_row=first, max_row=last))
    p.dataLabels = DataLabelList()
    p.dataLabels.showPercent = True
    p.dataLabels.showVal = False
    p.dataLabels.showCatName = False
    p.dataLabels.showSerName = False
    p.dataLabels.showLeaderLines = False
    p.height, p.width = 8.5, 13
    ss.add_chart(p, anchor)


pie("Portfolio allocation (stocks + buying power)", al_first, al_last, "L3")
pie("Invested vs. buying power", iv_first, iv_first + 1, "L21")

# =====================================================================
#  HOLDINGS
# =====================================================================
hs = wb.create_sheet("Holdings")
hs["A1"] = "Holdings"
hs["A1"].font = title_font
headers = ["Ticker", "Shares", "Avg Cost", "Cost Basis", "Current Price",
           "Market Value", "Gain/Loss ($)", "Gain/Loss (%)", "Weight", "Prev Close"]
for i, h in enumerate(headers, 1):
    hs.cell(row=3, column=i, value=h)
style_header(hs, 3, len(headers))

for i, (ticker, shares, avg_cost, price, prev_price) in enumerate(holdings):
    r = first_h + i
    hs.cell(row=r, column=1, value=ticker).font = input_font
    hs.cell(row=r, column=2, value=shares).font = input_font
    hs.cell(row=r, column=3, value=avg_cost).font = input_font
    hs.cell(row=r, column=5, value=price).font = input_font
    hs.cell(row=r, column=10, value=prev_price if prev_price else price).font = input_font
    hs.cell(row=r, column=4, value=f'=IF(OR(B{r}="",C{r}=""),"",B{r}*C{r})')
    hs.cell(row=r, column=6, value=f'=IF(OR(B{r}="",E{r}=""),"",B{r}*E{r})')
    hs.cell(row=r, column=7, value=f'=IF(OR(F{r}="",D{r}=""),"",F{r}-D{r})')
    hs.cell(row=r, column=8, value=f'=IF(OR(F{r}="",D{r}=""),"",IFERROR(G{r}/D{r},0))')
    hs.cell(row=r, column=9, value=f'=IF(F{r}="","",IFERROR(F{r}/$F${total_row},0))')
    for c in (4, 6, 7, 8, 9):
        hs.cell(row=r, column=c).font = base
    hs.cell(row=r, column=2).number_format = '#,##0.####'
    for c in (3, 4, 5, 6, 7, 10):
        hs.cell(row=r, column=c).number_format = USD
    hs.cell(row=r, column=8).number_format = PCT
    hs.cell(row=r, column=9).number_format = PCT

tr = total_row
hs.cell(row=tr, column=1, value="Total").font = bold
if n:
    hs.cell(row=tr, column=2, value=f"=SUM(B{first_h}:B{last_h})")
    hs.cell(row=tr, column=4, value=f"=SUM(D{first_h}:D{last_h})")
    hs.cell(row=tr, column=6, value=f"=SUM(F{first_h}:F{last_h})")
    hs.cell(row=tr, column=7, value=f"=SUM(G{first_h}:G{last_h})")
    hs.cell(row=tr, column=8, value=f"=IFERROR(G{tr}/D{tr},0)")
    hs.cell(row=tr, column=9, value=f"=SUM(I{first_h}:I{last_h})")
hs.cell(row=tr, column=2).number_format = '#,##0.####'
hs.cell(row=tr, column=2).font = bold
for c in (4, 6, 7):
    hs.cell(row=tr, column=c).number_format = USD
    hs.cell(row=tr, column=c).font = bold
for c in (8, 9):
    hs.cell(row=tr, column=c).number_format = PCT
    hs.cell(row=tr, column=c).font = bold
if n:
    green_red(hs, f"G{first_h}:H{tr}")
for col, w in zip("ABCDEFGHIJ", [10, 10, 12, 14, 14, 14, 14, 14, 10, 12]):
    hs.column_dimensions[col].width = w

# =====================================================================
#  INTERESTED (watchlist: stocks you follow but own 0 shares of)
# =====================================================================
iw = wb.create_sheet("Interested")
iw["A1"] = "Interested (stocks I'm watching)"
iw["A1"].font = title_font
iw["A2"] = ("Add one with the Add Stock sheet (0 shares, or Action = WATCH). Type your research in the Notes column; "
            "notes are saved on your next refresh. These stocks also get news.")
iw["A2"].font = subtitle_font
for i, h in enumerate(["Ticker", "Price", "Prev Close", "Today ($)", "Today (%)", "Date Added", "Notes"], 1):
    iw.cell(row=3, column=i, value=h)
style_header(iw, 3, 7)
if watch:
    for i, (t, p, pv, added, notes) in enumerate(watch):
        r = 4 + i
        iw.cell(row=r, column=1, value=t).font = Font(name=FONT, size=11, bold=True, color="0000FF")
        iw.cell(row=r, column=2, value=p).font = input_font
        iw.cell(row=r, column=3, value=pv if pv else p).font = input_font
        iw.cell(row=r, column=4, value=f'=IF(OR(B{r}="",C{r}=""),"",B{r}-C{r})').font = base
        iw.cell(row=r, column=5, value=f'=IF(OR(B{r}="",C{r}=""),"",IFERROR((B{r}-C{r})/C{r},""))').font = base
        iw.cell(row=r, column=6, value=added).font = base
        nc = iw.cell(row=r, column=7, value=notes)
        nc.font = input_font
        nc.fill = input_fill
        nc.alignment = Alignment(wrap_text=True, vertical="top")
        for c in (2, 3, 4):
            iw.cell(row=r, column=c).number_format = USD
        iw.cell(row=r, column=5).number_format = PCT
        for c in range(1, 8):
            iw.cell(row=r, column=c).border = thin_border
    green_red(iw, f"D4:E{3 + len(watch)}")
else:
    iw["A4"] = "Nothing here yet. Add a stock with 0 shares to follow it without owning it."
    iw["A4"].font = subtitle_font
for col, w in zip("ABCDEFG", [12, 12, 12, 12, 12, 14, 60]):
    iw.column_dimensions[col].width = w

# =====================================================================
#  ADD STOCK
# =====================================================================
asx = wb.create_sheet("Add Stock", 2)
asx["A1"] = "Add, Update, or Remove a Stock"
asx["A1"].font = title_font
guide = [
    "1) Type a ticker below (like NVDA). Shares and Avg Cost are optional: leave them blank and Python will ask you.",
    "2) Shares = 0, or Action = WATCH, adds it to your Interested sheet. Action = REMOVE sells one (Shares = how many you sold, blank = all).",
    "3) Save and close this file, then run run_all_scripts.py or add_stock_interactive.py in Spyder and answer the questions.",
]
for i, text in enumerate(guide, start=2):
    asx.cell(row=i, column=1, value=text).font = subtitle_font
for i, h in enumerate(["Ticker", "Shares", "Avg Cost", "Action", "Note"], 1):
    asx.cell(row=5, column=i, value=h)
style_header(asx, 5, 5)
for r in range(6, 31):
    for c in range(1, 6):
        cell = asx.cell(row=r, column=c)
        cell.border = thin_border
        cell.font = input_font
    asx.cell(row=r, column=2).number_format = '#,##0.####'
    asx.cell(row=r, column=3).number_format = USD
for i, (p_t, p_s, p_a, p_act) in enumerate(pending):
    r = 6 + i
    asx.cell(row=r, column=1, value=p_t)
    asx.cell(row=r, column=2, value=p_s)
    asx.cell(row=r, column=3, value=p_a)
    asx.cell(row=r, column=4, value=p_act or None)
    asx.cell(row=r, column=5, value="Waiting for your answers in Spyder").font = subtitle_font
dv = DataValidation(type="list", formula1='"ADD,REMOVE,WATCH"', allow_blank=True)
asx.add_data_validation(dv)
dv.add("D6:D30")
for col, w in zip("ABCDE", [12, 12, 14, 14, 40]):
    asx.column_dimensions[col].width = w

# =====================================================================
#  NEWS
# =====================================================================
ns = wb.create_sheet("News")
ns["A1"] = "Recent news for your holdings and Interested stocks"
ns["A1"].font = title_font
for i, h in enumerate(["Ticker", "Date", "Headline", "Source", "Summary", "Link"], 1):
    ns.cell(row=3, column=i, value=h)
style_header(ns, 3, 6)
wrap = Alignment(wrap_text=True, vertical="top")
for i, (label, d, headline, source, url, summary) in enumerate(news_rows):
    r = 4 + i
    for c, v in enumerate([label, d, headline, source, summary], 1):
        cell = ns.cell(row=r, column=c, value=v)
        cell.font = base
        cell.alignment = wrap
    if url:
        link = ns.cell(row=r, column=6, value="Open")
        link.hyperlink = url
        link.font = link_font
        link.alignment = Alignment(vertical="top")
if not news_rows:
    ns["A4"] = "No news saved yet. It appears after morning_news.py runs."
    ns["A4"].font = subtitle_font
for col, w in zip("ABCDEF", [20, 12, 60, 18, 70, 8]):
    ns.column_dimensions[col].width = w

# =====================================================================
#  HISTORY
# =====================================================================
hist = wb.create_sheet("History")
hist.sheet_view.showGridLines = False
hist["A1"] = "History: how your money has changed over time"
hist["A1"].font = title_font
hist["A2"] = ("Every refresh saves one snapshot per day. 'Value' includes your buying power, so adding or withdrawing cash moves it. "
              "'Stock gain/loss' only tracks your stocks (price vs. what you paid).")
hist["A2"].font = subtitle_font

m = len(history)
snap_first = 13
snap_last = snap_first + m - 1 if m else snap_first

section(hist, 3, "Compared with earlier dates", 8)
cmp_headers = ["Period", "Compared with", "Value then", "Value now", "Value change ($)", "Value change (%)",
               "Stock gain/loss change ($)", "S&P 500 change (%)"]
for i, h in enumerate(cmp_headers, 1):
    hist.cell(row=4, column=i, value=h)
style_header(hist, 4, 8)
hist.row_dimensions[4].height = 32

if m:
    D = f"$A${snap_first}:$A${snap_last}"
    V = f"$B${snap_first}:$B${snap_last}"
    G = f"$G${snap_first}:$G${snap_last}"
    S = f"$H${snap_first}:$H${snap_last}"
    periods = [("Yesterday", 1), ("Last 7 days", 7), ("Last 30 days", 30), ("Since my first snapshot", None)]
    for i, (name, days) in enumerate(periods):
        r = 5 + i
        hist.cell(row=r, column=1, value=name).font = bold
        if days:
            hist.cell(row=r, column=2, value=f'=IFERROR(INDEX({D},MATCH(MAX({D})-{days},{D},1)),"not enough history")')
        else:
            hist.cell(row=r, column=2, value=f"=MIN({D})")
        hist.cell(row=r, column=3, value=f'=IFERROR(INDEX({V},MATCH(B{r},{D},0)),"")')
        hist.cell(row=r, column=4, value=f"=B{snap_last}")
        hist.cell(row=r, column=5, value=f'=IF(C{r}="","",D{r}-C{r})')
        hist.cell(row=r, column=6, value=f'=IF(C{r}="","",IFERROR(E{r}/C{r},""))')
        hist.cell(row=r, column=7, value=f'=IFERROR(INDEX({G},ROWS({G}))-INDEX({G},MATCH(B{r},{D},0)),"")')
        hist.cell(row=r, column=8, value=f'=IFERROR(INDEX({S},ROWS({S}))/INDEX({S},MATCH(B{r},{D},0))-1,"")')
        hist.cell(row=r, column=2).number_format = "yyyy-mm-dd"
        for c in (3, 4, 5, 7):
            hist.cell(row=r, column=c).number_format = USD
        for c in (6, 8):
            hist.cell(row=r, column=c).number_format = PCT
        for c in range(1, 9):
            hist.cell(row=r, column=c).border = thin_border
            if c > 1:
                hist.cell(row=r, column=c).font = base
            if c in (2,):
                hist.cell(row=r, column=c).alignment = Alignment(horizontal="center")
    green_red(hist, "E5:H8")
else:
    hist["A5"] = "No snapshots yet. Run run_all_scripts.py to save your first one."
    hist["A5"].font = subtitle_font

section(hist, 11, "Daily snapshots", 10)
snap_headers = ["Date", "Total Value", "Change vs. previous day ($)", "Change vs. previous day (%)",
                "Change since start ($)", "Change since start (%)", "Stock Gain/Loss ($)", "S&P 500",
                "Portfolio (start = 100)", "S&P 500 (start = 100)"]
for i, h in enumerate(snap_headers, 1):
    hist.cell(row=snap_first - 1, column=i, value=h)
style_header(hist, snap_first - 1, 10)
hist.row_dimensions[snap_first - 1].height = 45

for i, (dd, total, gl, sp) in enumerate(history):
    r = snap_first + i
    hist.cell(row=r, column=1, value=dd).number_format = "yyyy-mm-dd"
    hist.cell(row=r, column=1).font = input_font
    hist.cell(row=r, column=2, value=total).font = input_font
    hist.cell(row=r, column=7, value=gl).font = input_font
    hist.cell(row=r, column=8, value=sp).font = input_font
    if i > 0:
        hist.cell(row=r, column=3, value=f"=B{r}-B{r - 1}")
        hist.cell(row=r, column=4, value=f'=IFERROR(C{r}/B{r - 1},"")')
    hist.cell(row=r, column=5, value=f"=B{r}-$B${snap_first}")
    hist.cell(row=r, column=6, value=f'=IFERROR(B{r}/$B${snap_first}-1,"")')
    hist.cell(row=r, column=9, value=f'=IFERROR(B{r}/$B${snap_first}*100,"")')
    hist.cell(row=r, column=10, value=f'=IF(H{r}="","",IFERROR(H{r}/$H${snap_first}*100,""))')
    for c in (2, 3, 5, 7):
        hist.cell(row=r, column=c).number_format = USD
    for c in (4, 6):
        hist.cell(row=r, column=c).number_format = PCT
    hist.cell(row=r, column=8).number_format = "#,##0.00"
    for c in (9, 10):
        hist.cell(row=r, column=c).number_format = "0.0"
    for c in range(1, 11):
        hist.cell(row=r, column=c).border = thin_border
        if c in (3, 4, 5, 6, 9, 10):
            hist.cell(row=r, column=c).font = base
if m:
    green_red(hist, f"C{snap_first}:F{snap_last}")
for col, w in zip("ABCDEFGHIJ", [24, 16, 16, 16, 18, 18, 20, 14, 18, 18]):
    hist.column_dimensions[col].width = w


def line_chart(title, cols, ytitle, yfmt, anchor):
    ch = LineChart()
    ch.title = title
    ch.height, ch.width = 8.5, 19
    for c in cols:
        ch.add_data(Reference(hist, min_col=c, min_row=snap_first - 1, max_row=snap_last), titles_from_data=True)
    ch.set_categories(Reference(hist, min_col=1, min_row=snap_first, max_row=snap_last))
    for _s in ch.series:
        _s.smooth = False
    ch.x_axis = DateAxis(crossAx=100)
    ch.x_axis.number_format = "d-mmm"
    ch.x_axis.majorTimeUnit = "days"
    ch.y_axis.crossAx = 500
    ch.y_axis.title = ytitle
    ch.y_axis.number_format = yfmt
    ch.x_axis.delete = False
    ch.y_axis.delete = False
    hist.add_chart(ch, anchor)


if m >= 2:
    line_chart("Total value over time", [2], "Value ($)", '$#,##0', "L3")
    line_chart("Stock gain/loss over time", [7], "Gain/Loss ($)", '$#,##0', "L21")
    line_chart("Portfolio vs. S&P 500 (start = 100)", [9, 10], "Index", '0', "L39")
else:
    hist["L3"] = "Charts appear here once you have two or more daily snapshots (run the refresh on different days)."
    hist["L3"].font = subtitle_font

# =====================================================================
#  MY NOTES (kept between refreshes) + sheet order
# =====================================================================
if "My Notes" not in wb.sheetnames:
    mn = wb.create_sheet("My Notes")
    mn["A1"] = "My Notes"
    mn["A1"].font = title_font
    mn["A2"] = "Type anything here. This sheet, and any other sheet you add yourself, is kept when you refresh."
    mn["A2"].font = subtitle_font
    mn.column_dimensions["A"].width = 100

if fresh and "Sheet" in wb.sheetnames:
    del wb["Sheet"]
_order = {nme: i for i, nme in enumerate(GENERATED)}
wb._sheets.sort(key=lambda s: _order.get(s.title, 100))
wb.active = 0

# =====================================================================
#  SAVE
# =====================================================================
try:
    wb.save(out_path)
    print("Saved updated dashboard:", out_path)
    try:
        import portfolio_tools as pt
        _c = pt.connect()
        pt.set_meta(_c, "handled_mtime", str(os.stat(out_path).st_mtime_ns))
        pt.set_meta(_c, "exported_cash", repr(float(cash)))
        _c.close()
    except Exception:
        pass
    if hasattr(os, "startfile"):
        os.startfile(out_path)
except PermissionError:
    print("Could not save. Close portfolio_dashboard.xlsx in Excel and run this again.")
