"""
settlement_loader.py  -  keep this file in the SAME folder as your reconciliation script.
load_settlement_df(path) reads an NPCI settlement file whatever it really is
(real .xls/.xlsx, HTML table, SpreadsheetML XML or text) and returns the same
kind of table as pd.read_excel(path, header=None).
"""
import io
import os
import xml.etree.ElementTree as ET
from html.parser import HTMLParser

import pandas as pd


def to_float(v):
    if v is None or isinstance(v, bool):
        return None
    if isinstance(v, (int, float)):
        return None if pd.isna(v) else float(v)
    s = str(v).replace(",", "").strip()
    if s == "" or s.lower() in ("nan", "none"):
        return None
    try:
        return float(s)
    except ValueError:
        return None


class _TableParser(HTMLParser):
    def __init__(self):
        super().__init__()
        self.rows, self.row, self.cell = [], None, None

    def handle_starttag(self, tag, attrs):
        if tag == "tr":
            self.row = []
        elif tag in ("td", "th") and self.row is not None:
            self.cell = []
        elif tag == "br" and self.cell is not None:
            self.cell.append(" ")

    def handle_endtag(self, tag):
        if tag in ("td", "th") and self.cell is not None and self.row is not None:
            self.row.append(" ".join("".join(self.cell).split()))
            self.cell = None
        elif tag == "tr" and self.row is not None:
            self.rows.append(self.row)
            self.row = None

    def handle_data(self, data):
        if self.cell is not None:
            self.cell.append(data)


def _rows_to_df(rows):
    """List of rows -> DataFrame. Numeric-looking cells (not column A) become numbers."""
    width = max([len(r) for r in rows] + [4])
    out = []
    for r in rows:
        r = list(r) + [None] * (width - len(r))
        new = []
        for j, v in enumerate(r):
            if isinstance(v, str) and v.strip() == "":
                v = None
            elif j > 0 and isinstance(v, str):
                f = to_float(v)
                v = f if f is not None else v
            new.append(v)
        out.append(new)
    return pd.DataFrame(out, dtype=object)


def _decode(raw):
    if raw.startswith((b"\xff\xfe", b"\xfe\xff")):
        return raw.decode("utf-16")
    for enc in ("utf-8-sig", "cp1252"):
        try:
            return raw.decode(enc)
        except UnicodeDecodeError:
            pass
    return raw.decode("latin-1")


def _spreadsheetml_rows(raw):
    ns = "{urn:schemas-microsoft-com:office:spreadsheet}"
    root = ET.fromstring(raw)
    table = root.find(f".//{ns}Table")
    rows = []
    for row in table.findall(f"{ns}Row"):
        if row.get(ns + "Index"):
            rows += [[]] * (int(row.get(ns + "Index")) - 1 - len(rows))
        cells = []
        for c in row.findall(f"{ns}Cell"):
            if c.get(ns + "Index"):
                cells += [None] * (int(c.get(ns + "Index")) - 1 - len(cells))
            d = c.find(f"{ns}Data")
            cells.append(d.text if d is not None else None)
            cells += [None] * int(c.get(ns + "MergeAcross") or 0)
        rows.append(cells)
    return rows


def load_settlement_df(path):
    """Read a settlement file whatever it really is: xls, xlsx, HTML table, SpreadsheetML XML or text."""
    if os.path.getsize(path) == 0:
        raise ValueError("file is empty (0 KB)")
    with open(path, "rb") as f:
        raw = f.read()
    if raw.startswith(b"\xd0\xcf\x11\xe0") or raw.startswith(b"PK"):
        return pd.read_excel(path, header=None)          # real .xls / .xlsx
    text = _decode(raw)
    low = text.lstrip().lower()
    if "urn:schemas-microsoft-com:office:spreadsheet" in low[:3000]:
        return _rows_to_df(_spreadsheetml_rows(raw))
    if "<table" in low:
        p = _TableParser()
        p.feed(text)
        if p.rows:
            return _rows_to_df(p.rows)
    if low[:1] != "<" and ("\t" in text or "," in text):
        sep = "\t" if text.count("\t") >= text.count(",") else ","
        df = pd.read_csv(io.StringIO(text), header=None, sep=sep, dtype=object, engine="python", on_bad_lines="skip")
        return _rows_to_df(df.where(df.notna(), None).values.tolist())
    raise ValueError(f"unrecognised file content (starts with {raw[:40]!r})")
