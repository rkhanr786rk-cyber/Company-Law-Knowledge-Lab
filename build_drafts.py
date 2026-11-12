"""Step 6: turn the Word drafts in the drafts/ folder into fill-in forms for the Drafting section.

Reads   drafts/**/*.docx        (folder name = category, file name = title)
Writes  site/data/drafts/index.json        list of drafts for the Drafting page
        site/data/drafts/<id>.json         the fields of one draft and the text for the live preview
        site/drafts/<id>.docx              the draft with every fill-in place replaced by a marker the page can fill
        data/drafts_report.md
Run it after build_site.py (which creates the site/ folder).

How the places to fill are found (the first style that exists in a draft is used for that draft):
  1. {{Company name}}  {{Date of meeting | date}}  {{Mode | choice: Physical; Video}}  {{Address | long}}  {{Chairman | optional}}
  2. Word content controls (Developer tab: text, date and drop-down boxes)
  3. Text with a highlight colour (short pieces up to 120 characters)
  4. <<Company name>>
  5. [Company name]  (also [*] and [ ] with the words around them used as the label)
  6. Nothing marked: the "auto" style finds the parts that normally change (dates, amounts, percentages,
     names after Sh./Smt./Mr., S/o D/o W/o names, R/o addresses, M/s firms, company names, (city) style
     notes, short blanks). Each one becomes a field that starts with the original wording, so the person
     only changes what is different.
Old .doc files are converted to .docx first (needs LibreOffice; the workflow installs it when .doc files exist).
"""
import json
import re
import shutil
import subprocess
import sys
import tempfile
import zipfile
from pathlib import Path
from xml.sax.saxutils import escape, unescape

SRC, OUT, DATA = Path("drafts"), Path("site"), Path("data")

P_RE = re.compile(r"<w:p(?:\s[^>]*[^/>])?>.*?</w:p>", re.S)
R_RE = re.compile(r"<w:r(?:\s[^>]*[^/>])?>.*?</w:r>", re.S)
T_RE = re.compile(r"<w:t(?:\s[^>]*[^/>])?>(.*?)</w:t>|<w:t\s*/>", re.S)
SDT_RE = re.compile(r"<w:sdt>.*?</w:sdt>", re.S)
TBL_RE = re.compile(r"<w:tbl>.*?</w:tbl>", re.S)
HIGHLIGHT_RE = re.compile(r"<w:highlight\b[^>]*/>")
PART_RE = re.compile(r"^word/(document|header\d*|footer\d*)\.xml$")
BLANK_RE = re.compile(r"_{3,}|\.{6,}|…{2,}")
NOT_FIELD = re.compile(r"^(?:\d+|[ivxlc]+|[IVXLC]+|\*+|\.+|omitted|deleted|sic)$", re.I)
LEAD_WORDS = re.compile(r"^(?:insert|enter|mention|specify|give|fill in|add|write)\s+(?:the\s+)?", re.I)

DATE_WORDS = re.compile(r"\bdate\b|\bdated\b|date of birth", re.I)
LONG_WORDS = re.compile(r"address|details|description|purpose|resolution|agenda|business|reason|objects?|terms|matter|particulars|remarks|recital|text of|nature of", re.I)
NUM_WORDS = re.compile(r"amount|number of|no\. of|percentage|quantity|\bshares?\b|capital|\bvalue\b|price|\brate\b|\bfee\b|salary|remuneration", re.I)


def slug(s):
    return re.sub(r"[^a-z0-9]+", "-", s.lower()).strip("-") or "draft"


def natural(s):
    return [int(x) if x.isdigit() else x.lower() for x in re.split(r"(\d+)", s)]


def clean_title(stem):
    t = re.sub(r"^\d+[\s._-]+", "", stem.replace("_", " ")).strip()
    t = re.sub(r"\s+", " ", t) or stem
    return t.title() if t.isupper() and len(t) > 3 else t


# ---------- reading a paragraph as runs ----------
def split_runs(pxml):
    runs = []
    pos = 0
    for m in R_RE.finditer(pxml):
        text = unescape("".join(g for g in T_RE.findall(m.group(0))))
        runs.append({"a": m.start(), "b": m.end(), "xml": m.group(0), "text": text, "s": pos, "e": pos + len(text)})
        pos += len(text)
    return runs, pos


def rebuild_run(xml, new_text):
    """Put new_text in the first w:t of a run and drop the others."""
    first = [True]

    def sub(m):
        if first[0]:
            first[0] = False
            return f'<w:t xml:space="preserve">{escape(new_text)}</w:t>' if new_text else ""
        return ""
    return T_RE.sub(sub, xml)


def apply_spans(pxml, runs, spans, drop_highlight=False):
    """spans: [(a, b, token)] in paragraph text offsets, sorted, not overlapping."""
    if not spans:
        return pxml
    out = pxml
    for r in reversed(runs):
        s, e, text = r["s"], r["e"], r["text"]
        if e == s:
            continue
        pieces, cur, touched = [], s, False
        for a, b, tok in spans:
            if b <= s or a >= e:
                continue
            touched = True
            if a > cur:
                pieces.append(text[cur - s:min(a, e) - s])
            if a >= s:
                pieces.append(tok)
            cur = max(cur, min(b, e))
        if not touched:
            continue
        if cur < e:
            pieces.append(text[cur - s:])
        xml = rebuild_run(r["xml"], "".join(pieces))
        if drop_highlight:
            xml = HIGHLIGHT_RE.sub("", xml)
        out = out[:r["a"]] + xml + out[r["b"]:]
    return out


# ---------- finding the fill-in places ----------
def context_label(text, a, b):
    junk = re.compile(r"_{2,}|\.{3,}|\u2026+|\[\s*\]")
    before = re.findall(r"[^\s]+", junk.sub(" ", text[:a]))[-4:]
    after = re.findall(r"[^\s]+", junk.sub(" ", text[b:]))[:3]
    if before:
        return "After \u201c" + " ".join(before).strip(" ,;:") + "\u201d"
    if after:
        return "Before \u201c" + " ".join(after).strip(" ,;:") + "\u201d"
    return "Details"


def parse_curly(inner):
    parts = [p.strip() for p in inner.split("|")]
    f = {"label": parts[0], "type": None, "options": [], "optional": False, "help": ""}
    for p in parts[1:]:
        low = p.lower()
        if low in ("date", "number", "long", "text"):
            f["type"] = "long" if low == "long" else low
        elif low == "optional":
            f["optional"] = True
        elif low.startswith("choice"):
            f["type"] = "choice"
            f["options"] = [o.strip() for o in p.split(":", 1)[1].split(";") if o.strip()] if ":" in p else []
        elif low.startswith("help"):
            f["help"] = p.split(":", 1)[1].strip() if ":" in p else ""
    return f


def spans_curly(text, _p, _runs):
    out = []
    for m in re.finditer(r"\{\{\s*(.+?)\s*\}\}", text):
        out.append((m.start(), m.end(), parse_curly(m.group(1)), True))
    return out


def spans_angle(text, _p, _runs):
    return [(m.start(), m.end(), {"label": m.group(1).strip()}, True) for m in re.finditer(r"<<\s*(.+?)\s*>>", text)]


def spans_bracket(text, _p, _runs):
    out = []
    for m in re.finditer(r"\[([^\[\]]{0,80})\]", text):
        inner = m.group(1).strip()
        if m.start() > 0 and text[m.start() - 1].isdigit():
            continue
        if not inner or re.fullmatch(r"[\s●•_*.…-]+", inner):
            out.append((m.start(), m.end(), {"label": context_label(text, m.start(), m.end())}, False))
        elif re.search(r"[A-Za-z]", inner) and not NOT_FIELD.match(inner):
            out.append((m.start(), m.end(), {"label": inner}, True))
    return out


def spans_blank(text, _p, _runs):
    return [(m.start(), m.end(), {"label": context_label(text, m.start(), m.end())}, False) for m in BLANK_RE.finditer(text)]


def spans_highlight(text, _p, runs):
    out, group = [], []

    def flush():
        if group:
            a, b = group[0]["s"], group[-1]["e"]
            t = text[a:b]
            if t.strip() and len(t) <= 120:
                lab = t.strip().strip("[]() ").strip()
                if not lab or re.fullmatch(r"[\s●•_*.…-]+", lab):
                    f = {"label": context_label(text, a, b)}
                    out.append((a, b, f, False))
                else:
                    out.append((a, b, {"label": lab}, True))
            group.clear()
    for r in runs:
        hl = re.search(r'<w:highlight\b[^>]*w:val="(?!none)[^"]+"', r["xml"])
        if r["e"] == r["s"]:
            continue
        if hl:
            group.append(r)
        else:
            flush()
    flush()
    return out


def spans_controls(text, pxml, runs):
    out = []
    for sm in SDT_RE.finditer(pxml):
        sx = sm.group(0)
        inside = [r for r in runs if sm.start() <= r["a"] and r["b"] <= sm.end() and r["e"] > r["s"]]
        if not inside:
            continue
        alias = re.search(r'<w:alias w:val="([^"]*)"', sx)
        tag = re.search(r'<w:tag w:val="([^"]*)"', sx)
        label = unescape((alias or tag).group(1)) if (alias or tag) else ""
        shown = text[inside[0]["s"]:inside[-1]["e"]]
        if not label:
            label = shown if len(shown) < 50 and not re.search(r"click|tap|enter text", shown, re.I) else "Field"
        f = {"label": label}
        if "<w:date" in sx:
            f["type"] = "date"
        items = re.findall(r'<w:listItem\b[^>]*w:displayText="([^"]*)"', sx)
        if items:
            f["type"] = "choice"
            f["options"] = [unescape(i) for i in items]
        out.append((inside[0]["s"], inside[-1]["e"], f, True))
    return out



MONTH = r"(?:January|February|March|April|May|June|July|August|September|October|November|December|Jan|Feb|Mar|Apr|Jun|Jul|Aug|Sept?|Oct|Nov|Dec)"
DATE_RES = [re.compile(x, re.I) for x in (
    rf"\b\d{{1,2}}(?:st|nd|rd|th)?\s+(?:day\s+of\s+)?{MONTH}\.?,?\s*\d{{4}}\b",
    rf"\b{MONTH}\.?\s+\d{{1,2}}(?:st|nd|rd|th)?,?\s*\d{{4}}\b",
    r"\b\d{1,2}[./-]\d{1,2}[./-](?:\d{4}|\d{2})\b")]
AMOUNT_RE = re.compile(r"(?:Rs\.?|INR|\u20b9)\s*\d[\d,]*(?:\.\d+)?(?:\s*/-)?(?:\s*(?:Crores?|Crs?\.?|Lacs?|Lakhs?|thousand|million))?(?:\s*\((?:Rupees|Rs\.?)[^)]*\))?", re.I)
PCT_RE = re.compile(r"\b\d+(?:\.\d+)?\s?%")
CAP = r"(?![SDWRC]/[oO]\b)[A-Z][A-Za-z.'\-]*"
HON = r"(?:Late\s+)?(?:Sh|Shri|Smt|Shrimati|Mrs|Mr|Ms|Miss|Dr|Mast)\b\.?"
NAME_RE = re.compile(rf"\b{HON}\s*{CAP}(?:\s+(?:{CAP}|Alias))*")
REL_RE = re.compile(rf"\b[SDWC]/[oO]\s+((?:{HON}\s*)?{CAP}(?:\s+{CAP})*)")
ADDR_RE = re.compile(r"\b[Rr]/[oO]\s+(.+?)(?=\s+(?:\(\d\)|hereinafter|BETWEEN|AND\b)|\s*$)")
FIRM_RE = re.compile(r"\bM/[sS]\.?\s+[A-Z][\w&.\-]*(?:\s+(?:and\s+)?[A-Z&][\w&.\-]*)*")
COMPANY_RE = re.compile(r"\b[A-Z][A-Z&.\-]*(?:\s+[A-Z&()][A-Z&.()\-]*)*\s+(?:PRIVATE\s+LIMITED|PVT\.?\s+LTD\.?|LIMITED|LTD\.?|LLP)\b")
PLACE_RE = re.compile(r"\bat\s+([A-Z][a-z]+(?:\s[A-Z][a-z]+)?)(?=,?\s+this\b)")
NOTE_RE = re.compile(r"\(\s*(city|town|place|name|date|address|amount|month|year|day|designation|state|district|pin ?code)\s*\)", re.I)
SMALL_BLANK = re.compile(r"(?<!_)_{3,14}(?!_)|(?<!\.)\.{4,14}(?!\.)|\u2026{2,}")
SIGN_RE = re.compile(r"SIGNED|presence of|signature|witness", re.I)


MONTHS_FULL = ["January", "February", "March", "April", "May", "June", "July", "August", "September", "October", "November", "December"]
_D1 = re.compile(rf"^(\d{{1,2}})(st|nd|rd|th)?(\s+(?:day\s+of\s+)?)({MONTH})(\.?,?\s*)(\d{{4}})$", re.I)
_D2 = re.compile(rf"^({MONTH})(\.?\s+)(\d{{1,2}})(st|nd|rd|th)?(,?\s*)(\d{{4}})$", re.I)
_D3 = re.compile(r"^(\d{1,2})([./-])(\d{1,2})\2(\d{4}|\d{2})$")


def month_no(word):
    w = word.lower().rstrip(".")
    for i, m in enumerate(MONTHS_FULL):
        if m.lower() == w or (len(w) >= 3 and m.lower().startswith(w)):
            return i + 1
    return 0


def ordinal(d):
    return "th" if 10 < d % 100 < 20 else {1: "st", 2: "nd", 3: "rd"}.get(d % 10, "th")


def fmt_style(y, m, d, st):
    if st["k"] == "n":
        dd, mm = (f"{d:02d}", f"{m:02d}") if st["pad"] else (str(d), str(m))
        yy = str(y) if st["y4"] else f"{y % 100:02d}"
        return dd + st["sep"] + mm + st["sep"] + yy
    mon = MONTHS_FULL[m - 1][:st["n"]]
    mon = mon.upper() if st["case"] == "u" else mon.lower() if st["case"] == "l" else mon
    day = (f"{d:02d}" if st["pad"] else str(d))
    suf = (ordinal(d).upper() if st.get("sufu") else ordinal(d)) if st["ord"] else ""
    if st["order"] == "dmy":
        return day + suf + st["sep"] + mon + st["msuf"] + st["ysep"] + str(y)
    return mon + st["msuf"] + st["sep"] + day + suf + st["ysep"] + str(y)


def parse_date(orig):
    """Returns (iso, style) when the text is a date that can be written back exactly the same way, else None."""
    o = orig.strip()
    m = _D1.match(o)
    if m:
        d, suf, sep, mon, tail, y = m.groups()
        mn = month_no(mon)
        if not mn:
            return None
        msuf = "." if tail.startswith(".") else ""
        st = {"k": "t", "order": "dmy", "ord": bool(suf), "sufu": bool(suf and suf.isupper()), "pad": len(d) == 2 and d.startswith("0"), "sep": sep,
              "n": len(mon.rstrip(".")), "case": "u" if mon.isupper() and len(mon) > 1 else "l" if mon.islower() else "t", "msuf": msuf, "ysep": tail[len(msuf):]}
        d_i = int(d)
    else:
        m = _D2.match(o)
        if m:
            mon, sepx, d, suf, ysep, y = m.groups()
            mn = month_no(mon)
            if not mn:
                return None
            msuf = "." if sepx.startswith(".") else ""
            st = {"k": "t", "order": "mdy", "ord": bool(suf), "sufu": bool(suf and suf.isupper()), "pad": len(d) == 2 and d.startswith("0"), "sep": sepx[len(msuf):],
                  "n": len(mon.rstrip(".")), "case": "u" if mon.isupper() and len(mon) > 1 else "l" if mon.islower() else "t", "msuf": msuf, "ysep": ysep}
            d_i = int(d)
        else:
            m = _D3.match(o)
            if not m:
                return None
            d, sep, mo, y = m.groups()
            mn, d_i = int(mo), int(d)
            st = {"k": "n", "sep": sep, "pad": any(len(x) == 2 and x.startswith("0") for x in (d, mo)), "y4": len(y) == 4}
    y_i = int(y) if len(y) == 4 else 2000 + int(y)
    if not (1 <= mn <= 12 and 1 <= d_i <= 31):
        return None
    if st["k"] == "n" and len(y) == 2:
        return None
    if fmt_style(y_i, mn, d_i, st) != o:
        return None
    return f"{y_i:04d}-{mn:02d}-{d_i:02d}", st


def spans_auto(text, _p, _runs):
    c = []   # (priority, a, b, base label, mergeable)

    def add(pr, a, b, base):
        while b > a and text[b - 1] in " .,;:":
            b -= 1
        if b > a:
            c.append((pr, a, b, base))
    for rx in DATE_RES:
        for m in rx.finditer(text):
            add(1, m.start(), m.end(), "Date")
    for m in AMOUNT_RE.finditer(text):
        add(1, m.start(), m.end(), "Amount")
    for m in PCT_RE.finditer(text):
        add(1, m.start(), m.end(), "Percentage")
    for m in REL_RE.finditer(text):
        add(1, m.start(1), m.end(1), "Father\u2019s or husband\u2019s name")
    for m in ADDR_RE.finditer(text):
        add(1, m.start(1), m.end(1), "Address")
    for m in FIRM_RE.finditer(text):
        add(1, m.start(), m.end(), "Firm name")
    for m in COMPANY_RE.finditer(text):
        add(1, m.start(), m.end(), "Company name")
    for m in PLACE_RE.finditer(text):
        add(1, m.start(1), m.end(1), "Place")
    for m in NOTE_RE.finditer(text):
        add(1, m.start(), m.end(), m.group(1).strip().capitalize())
    for m in NAME_RE.finditer(text):
        add(2, m.start(), m.end(), "Name")
    if re.fullmatch(r"\s*-{4,}\s*", text):
        add(1, text.index("-"), len(text.rstrip()), "Address line")
    elif not SIGN_RE.search(text):
        for m in SMALL_BLANK.finditer(text):
            c.append((1, m.start(), m.end(), None))
    c.sort(key=lambda x: (x[0], x[1]))
    taken, out = [], []
    for pr, a, b, base in c:
        if any(a < tb and b > ta for ta, tb in taken):
            continue
        taken.append((a, b))
        orig = text[a:b]
        pd = parse_date(orig) if base == "Date" else None
        if pd:
            meta = {"label": "Date", "base": "Date", "default": pd[0], "key": "date|" + pd[0], "type": "date", "style": pd[1], "orig": orig}
            out.append((a, b, meta, True))
        elif base is None:
            meta = {"label": context_label(text, a, b), "base": None, "default": "", "key": None}
            out.append((a, b, meta, False))
        else:
            meta = {"label": base, "base": base, "default": orig, "key": base + "|" + re.sub(r"\s+", " ", orig.lower())}
            out.append((a, b, meta, True))
    return out


MODES = [("curly", spans_curly, False), ("controls", spans_controls, False), ("highlight", spans_highlight, True),
         ("angle", spans_angle, False), ("bracket", spans_bracket, False), ("auto", spans_auto, False)]


def guess_type(f):
    if f.get("type"):
        return f["type"]
    lab = f["label"]
    if DATE_WORDS.search(lab):
        return "date"
    if LONG_WORDS.search(lab):
        return "long"
    if NUM_WORDS.search(lab):
        return "number"
    return "text"


_SIMPLE_CHILD = re.compile(r"<w:t(?:\s[^>]*)?>.*?</w:t>|<w:t\s*/>|<w:tab\s*/>|<w:br(?:\s[^>]*)?/>", re.S)


def split_mixed_runs(xml):
    """A run that holds text, tabs and line breaks together is split into one run per item, so a tab can never move."""
    def fix(m):
        run = m.group(0)
        head = re.match(r"<w:r(?:\s[^>]*[^/>])?>", run)
        body = run[head.end():-len("</w:r>")]
        rpr = re.match(r"\s*<w:rPr>.*?</w:rPr>", body, re.S)
        rest = body[rpr.end():] if rpr else body
        items = _SIMPLE_CHILD.findall(rest)
        if len(items) < 2 or "".join(items) != rest:
            return run
        pre = rpr.group(0) if rpr else ""
        return "".join(head.group(0) + pre + it + "</w:r>" for it in items)
    return R_RE.sub(fix, xml)


# ---------- one draft ----------
def process(path, orig=None):
    rel = (orig or path).relative_to(SRC)
    zin = zipfile.ZipFile(path)
    parts = {n: split_mixed_runs(zin.read(n).decode("utf-8")) for n in zin.namelist() if PART_RE.match(n)}
    if "word/document.xml" not in parts:
        raise ValueError("not a Word .docx file")

    paras = []   # (part, match, pxml, runs, text)
    for name in sorted(parts, key=lambda n: (n != "word/document.xml", n)):
        for m in P_RE.finditer(parts[name]):
            runs, _ = split_runs(m.group(0))
            paras.append((name, m, m.group(0), runs, "".join(r["text"] for r in runs)))

    mode = None
    for mname, fn, drop in MODES:
        if any(fn(t, px, rs) for _, _, px, rs, t in paras):
            mode, finder, drop_hl = mname, fn, drop
            break

    fields, by_key, n_blank = [], {}, 0
    new_parts = {n: [] for n in parts}   # list of (start, end, newxml)
    if mode:
        for name, m, px, runs, text in paras:
            found = sorted(finder(text, px, runs), key=lambda x: x[0])
            spans, last = [], -1
            for a, b, meta, mergeable in found:
                if a < last:
                    continue
                last = b
                label = LEAD_WORDS.sub("", meta["label"]).strip(" .:;,-") or meta["label"]
                label = label[:1].upper() + label[1:]
                key = meta.get("key") or (re.sub(r"\s+", " ", label.lower()) if mergeable else None)
                if key is None or key not in by_key:
                    f = {"id": f"f{len(fields) + 1}", "label": label, "type": meta.get("type"), "options": meta.get("options", []),
                         "optional": meta.get("optional", False), "help": meta.get("help", ""), "count": 0,
                         "default": re.sub(r"\s+", " ", meta.get("default", "")).strip(), "base": meta.get("base")}
                    shown = meta.get("orig") or f["default"]
                    if not meta.get("type"):
                        f["type"] = "long" if f["base"] in ("Address", "Address line") else guess_type(f) if not f["base"] else "text"
                    if shown:
                        f["help"] = "In the original: \u201c" + (shown[:90] + ("\u2026" if len(shown) > 90 else "")) + "\u201d"
                    fields.append(f)
                    if key is not None:
                        by_key[key] = f
                else:
                    f = by_key[key]
                f["count"] += 1
                tok = f["id"]
                if meta.get("style") is not None:
                    sl = f.setdefault("styles", [])
                    if meta["style"] not in sl:
                        sl.append(meta["style"])
                    tok += "#" + str(sl.index(meta["style"]))
                spans.append((a, b, "\u27e6" + tok + "\u27e7"))
            if spans:
                newp = apply_spans(px, runs, spans, drop_highlight=drop_hl)
                if mode == "controls":
                    newp = newp.replace("<w:showingPlcHdr/>", "")
                new_parts[name].append((m.start(), m.end(), newp))
    counts = {}
    for f in fields:
        if f.get("base"):
            counts[f["base"]] = counts.get(f["base"], 0) + 1
    seen = {}
    for f in fields:
        b = f.get("base")
        if b and counts[b] > 1:
            seen[b] = seen.get(b, 0) + 1
            f["label"] = f"{f['label']} {seen[b]}"
    for name, edits in new_parts.items():
        x = parts[name]
        for a, b, new in reversed(edits):
            x = x[:a] + new + x[b:]
        if mode == "controls":
            x = x.replace("<w:showingPlcHdr/>", "")
        parts[name] = x

    # text for the live preview (document body only)
    doc = parts["word/document.xml"]
    tables = [(m.start(), m.end()) for m in TBL_RE.finditer(doc)]
    preview, prev_blank = [], True
    for m in P_RE.finditer(doc):
        px = m.group(0)
        text = ""
        for tm in re.finditer(r"<w:t(?:\s[^>]*[^/>])?>(.*?)</w:t>|<w:tab\s*/>|<w:br\s*/>", px, re.S):
            g = tm.group(0)
            text += "\t" if g.startswith("<w:tab") else ("\n" if g.startswith("<w:br") else unescape(tm.group(1) or ""))
        if not text.strip():
            if not prev_blank:
                preview.append({"t": "", "k": "gap"})
            prev_blank = True
            continue
        prev_blank = False
        runs_x = [r for r in R_RE.findall(px) if T_RE.search(r)]
        bold = bool(runs_x) and all(re.search(r"<w:b/>|<w:b w:val=\"(?:1|true)\"", r) for r in runs_x)
        kind = []
        if bold:
            kind.append("b")
        if re.search(r'<w:jc w:val="center"', px):
            kind.append("c")
        if re.search(r'<w:jc w:val="(?:right|end)"', px):
            kind.append("r")
        if re.search(r'<w:pStyle w:val="(?:Heading\d|Title)"', px):
            kind.append("h")
        if any(a <= m.start() < b for a, b in tables):
            kind.append("t")
        preview.append({"t": text.strip("\n"), "k": " ".join(kind)})
    while preview and preview[-1]["k"] == "gap":
        preview.pop()

    return rel, mode, fields, preview, parts, zin


def convert_docs():
    """Turn old .doc files into .docx in a temporary folder. Returns {original .doc path: converted .docx path}."""
    docs = sorted(p for p in SRC.rglob("*") if p.suffix.lower() == ".doc" and not p.name.startswith("~$")) if SRC.exists() else []
    out = {}
    exe = shutil.which("soffice") or shutil.which("libreoffice")
    if not docs or not exe:
        return out, docs, bool(docs and not exe)
    tmp = Path(tempfile.mkdtemp(prefix="drafts_"))
    by_dir = {}
    for d in docs:
        by_dir.setdefault(d.parent, []).append(d)
    for i, (folder, items) in enumerate(by_dir.items()):
        sub = tmp / str(i)
        sub.mkdir()
        for k in range(0, len(items), 20):
            chunk = items[k:k + 20]
            try:
                subprocess.run([exe, "--headless", "--convert-to", "docx", "--outdir", str(sub)] + [str(x) for x in chunk],
                               stdout=subprocess.DEVNULL, stderr=subprocess.DEVNULL, timeout=600)
            except Exception:
                pass
        for d in items:
            c = sub / (d.stem + ".docx")
            if c.exists():
                out[d] = c
    return out, docs, False


def main():
    if not OUT.exists():
        sys.exit("Run build_site.py first (it creates the site folder).")
    (OUT / "data" / "drafts").mkdir(parents=True, exist_ok=True)
    (OUT / "drafts").mkdir(parents=True, exist_ok=True)
    DATA.mkdir(exist_ok=True)

    converted, old_docs, no_office = convert_docs()
    files = sorted([p for p in SRC.rglob("*") if p.suffix.lower() == ".docx" and not p.name.startswith("~$")] + list(converted), key=lambda p: natural(str(p))) if SRC.exists() else []
    index, report, used = [], ["# Drafting report", ""], set()
    rows = []
    for p in files:
        try:
            rel, mode, fields, preview, parts, zin = process(converted.get(p, p), p)
        except Exception as e:      # a damaged file must not stop the others
            rows.append(f"| {p.relative_to(SRC)} | not read | 0 | {str(e)[:80]} |")
            continue
        category = " / ".join(clean_title(x) for x in rel.parent.parts) or "General"
        title = clean_title(rel.stem)
        did = slug(f"{category} {title}")
        while did in used:
            did += "-2"
        used.add(did)
        # prepared template
        with zipfile.ZipFile(OUT / "drafts" / f"{did}.docx", "w") as zout:
            for info in zin.infolist():
                data = parts[info.filename].encode("utf-8") if info.filename in parts else zin.read(info.filename)
                zout.writestr(info, data)
        desc = next((x["t"] for x in preview if x["t"] and x["t"].strip().lower() != title.lower() and len(x["t"]) > 25), "")
        desc = re.sub("\u27e6f[\\d#]+\u27e7", "\u2026", desc)
        fields_out = [{k: v for k, v in f.items() if k not in ("count", "base")} for f in fields]
        (OUT / "data" / "drafts" / f"{did}.json").write_text(json.dumps({
            "id": did, "title": title, "category": category, "fields": fields_out, "preview": preview,
            "auto": mode == "auto", "file": f"drafts/{did}.docx", "parts": sorted(parts), "name": rel.stem}, ensure_ascii=False, separators=(",", ":")), encoding="utf-8")
        index.append({"id": did, "title": title, "category": category, "sort": [list(natural(x)) and x for x in rel.parent.parts], "n": len(fields),
                      "desc": re.sub(r"\s+", " ", desc)[:170], "labels": [f["label"] for f in fields][:40]})
        note = "" if fields else "no fill-in places found; offered as a plain download"
        rows.append(f"| {rel} | {mode or '-'} | {len(fields)} | {note} |")

    index.sort(key=lambda d: (natural("/".join(d["sort"])), natural(d["title"])))
    for d in index:
        d.pop("sort")
    (OUT / "data" / "drafts" / "index.json").write_text(json.dumps({"drafts": index}, ensure_ascii=False, separators=(",", ":")), encoding="utf-8")
    report += [f"- Drafts found: {len(files)}", f"- Drafts ready: {len(index)}",
               f"- Drafts with no fill-in places: {sum(1 for d in index if d['n'] == 0)}", "",
               "| File | Fill-in style used | Fields | Note |", "|---|---|---|---|"] + rows
    lost = [d for d in old_docs if d not in converted]
    if lost:
        report += ["", f"{len(lost)} old .doc file(s) could not be converted" + (" (LibreOffice is not installed here)." if no_office else "."), ""] + [f"- {d.relative_to(SRC)}" for d in lost]
    if not SRC.exists():
        report.append("")
        report.append("There is no drafts folder yet. Create one and upload the Word files there.")
    (DATA / "drafts_report.md").write_text("\n".join(report) + "\n", encoding="utf-8")
    print("\n".join(report[:8]))


if __name__ == "__main__":
    main()
