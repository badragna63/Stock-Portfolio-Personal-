import os, sys, re, sqlite3, urllib.request, urllib.parse
import xml.etree.ElementTree as ET
from datetime import datetime, timedelta
from email.utils import parsedate_to_datetime
from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Alignment

try:
    _base = sys.base_prefix
    for _sub in (("Library", "bin"), ("Library", "mingw-w64", "bin"),
                 ("Library", "usr", "bin"), ("DLLs",), ("Scripts",)):
        _p = os.path.join(_base, *_sub)
        if os.path.isdir(_p):
            os.environ["PATH"] = _p + os.pathsep + os.environ.get("PATH", "")
            if hasattr(os, "add_dll_directory"):
                try:
                    os.add_dll_directory(_p)
                except Exception:
                    pass
except Exception:
    pass

WATCH = []
INCLUDE_MY_HOLDINGS = True
MAX_ARTICLES_PER_ITEM = 1
DAYS_BACK = 3

folder = os.path.join(os.path.expanduser("~"), "portfolio-system")
os.makedirs(folder, exist_ok=True)
db_path = os.path.join(folder, "portfolio.db")
out_path = os.path.join(folder, "morning_news.xlsx")
log_path = os.path.join(folder, "morning_news.log")

def log(msg):
    line = datetime.now().strftime("%Y-%m-%d %H:%M:%S") + "  " + str(msg)
    try:
        print(line)
    except Exception:
        pass
    try:
        with open(log_path, "a", encoding="utf-8") as f:
            f.write(line + "\n")
    except Exception:
        pass

def looks_like_ticker(s):
    return re.fullmatch(r"[A-Z]{1,5}([.-][A-Z])?", s) is not None

def google_news(query, days):
    q = urllib.parse.quote_plus(f"{query} stock when:{days}d")
    url = f"https://news.google.com/rss/search?q={q}&hl=en-US&gl=US&ceid=US:en"
    req = urllib.request.Request(url, headers={"User-Agent": "Mozilla/5.0"})
    with urllib.request.urlopen(req, timeout=20) as resp:
        root = ET.fromstring(resp.read())
    out = []
    for item in root.iter("item"):
        title = (item.findtext("title") or "").strip()
        link = (item.findtext("link") or "").strip()
        src_el = item.find("source")
        source = (src_el.text or "").strip() if src_el is not None else ""
        if source and title.endswith(" - " + source):
            title = title[: -len(" - " + source)]
        try:
            date = parsedate_to_datetime(item.findtext("pubDate")).astimezone().strftime("%Y-%m-%d")
        except Exception:
            date = datetime.now().strftime("%Y-%m-%d")
        if title:
            out.append((title, "", source, link, date))
    return out

def yahoo_news(ticker):
    try:
        import yfinance as yf
    except Exception:
        return []
    out = []
    try:
        items = yf.Ticker(ticker).news or []
    except Exception:
        return []

    for item in items:
        content = item.get("content")
        if isinstance(content, dict):
            title = content.get("title")
            summary = content.get("summary") or content.get("description") or ""
            source = (content.get("provider") or {}).get("displayName", "")
            url = ((content.get("canonicalUrl") or {}).get("url")
                   or (content.get("clickThroughUrl") or {}).get("url") or "")
            date = (content.get("pubDate") or "")[:10]
        else:
            title = item.get("title")
            summary = ""
            source = item.get("publisher", "")
            url = item.get("link", "")
            ts = item.get("providerPublishTime")
            date = datetime.fromtimestamp(ts).strftime("%Y-%m-%d") if ts else ""
        if title:
            out.append((title, summary, source, url, date or datetime.now().strftime("%Y-%m-%d")))
    return out

conn = sqlite3.connect(db_path)
cur = conn.cursor()
cur.execute("""CREATE TABLE IF NOT EXISTS news (
    id INTEGER PRIMARY KEY AUTOINCREMENT,
    date TEXT, ticker TEXT, headline TEXT, source TEXT, url TEXT, summary TEXT)""")

items = []
if INCLUDE_MY_HOLDINGS:
    try:
        items += [r[0] for r in cur.execute("SELECT ticker FROM holdings")]
    except Exception:
        log("Could not read holdings from database.")
    try:
        items += [r[0] for r in cur.execute("SELECT ticker FROM watchlist")]
    except Exception:
        pass

items = list(dict.fromkeys(items))

new_keys = set()
for item in items:
    found = []
    if looks_like_ticker(item):
        try:
            found += yahoo_news(item)
        except Exception as e:
            log(f"Yahoo news failed for {item}: {e}")

    if not found:
        try:
            found += google_news(item, DAYS_BACK)
        except Exception as e:
            log(f"Google News failed for {item}: {e}")

    added = 0
    for title, summary, source, url, date in found:
        if added >= MAX_ARTICLES_PER_ITEM:
            break
        exists = cur.execute("SELECT 1 FROM news WHERE ticker = ? AND headline = ?",
                             (item, title)).fetchone()
        if not exists:
            cur.execute("INSERT INTO news (date, ticker, headline, source, url, summary) VALUES (?, ?, ?, ?, ?, ?)",
                        (date, item, title, source, url, summary))
            new_keys.add((item, title))
            added += 1
    log(f"{item}: {added} article captured")

conn.commit()

cutoff = (datetime.now() - timedelta(days=DAYS_BACK + 7)).strftime("%Y-%m-%d")
rows = []
for item in items:
    row = cur.execute(
        "SELECT date, headline, source, url, summary FROM news "
        "WHERE ticker = ? AND date >= ? ORDER BY date DESC, id DESC LIMIT 1", (item, cutoff)).fetchone()
    if row:
        d, headline, source, url, summary = row
        rows.append(((item, headline) in new_keys, d, item, headline, source, url, summary))

conn.close()

FONT = "Calibri"
wb = Workbook()
ws = wb.active
ws.title = "Morning News"
headers = ["Status", "Date", "Stock Ticker", "Headline (click to open)", "Source", "Summary"]
for i, h in enumerate(headers, 1):
    c = ws.cell(row=1, column=i, value=h)
    c.font = Font(name=FONT, size=11, bold=True, color="FFFFFF")
    c.fill = PatternFill("solid", fgColor="1F3864")
    c.alignment = Alignment(horizontal="center", vertical="center")

wrap = Alignment(wrap_text=True, vertical="top")
for r, (is_new, d, item, headline, source, url, summary) in enumerate(rows, 2):
    vals = ["NEW" if is_new else "SAVED", d, item, headline, source, summary]
    for c, v in enumerate(vals, 1):
        cell = ws.cell(row=r, column=c, value=v)
        cell.font = Font(name=FONT, size=10, bold=(c == 1), color="C00000" if (c == 1 and is_new) else "000000")
        cell.alignment = wrap
    if url:
        h = ws.cell(row=r, column=4)
        h.hyperlink = url
        h.font = Font(name=FONT, size=10, color="0563C1", underline="single")

for col, w in zip("ABCDEF", [10, 12, 16, 75, 20, 55]):
    ws.column_dimensions[col].width = w

ws.freeze_panes = "A2"
if rows:
    ws.auto_filter.ref = f"A1:F{len(rows) + 1}"

try:
    wb.save(out_path)
    log(f"Saved {out_path} ({len(rows)} articles outputted)")
except PermissionError:
    log("Could not save morning_news.xlsx. Close it in Excel and run again.")