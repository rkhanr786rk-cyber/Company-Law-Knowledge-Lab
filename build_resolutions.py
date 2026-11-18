#!/usr/bin/env python3
"""Builds the Model Resolutions library from the book of model resolutions kept in the repository
(any .pdf or .txt whose name contains "resolution"). Writes site/data/res/ and adds, to each Companies Act
section page, a list of the model resolutions that cite that section."""
import json
import re
import sys
from pathlib import Path

SITE = Path("site/data")
FORMS = {"INC", "MGT", "PAS", "SH", "DIR", "ADT", "AOC", "CHG", "DPT", "MR", "BEN", "CAA", "STK", "FC", "GNL", "MBP", "CRA", "NCLT", "SPICE", "RUN"}
FORM_RE = re.compile(r"\b([A-Z]{2,5})\s?-\s?(\d{1,2}[A-Z]?)\b")
SEC_RE = re.compile(r"\b[Ss]ections?\s+((?:\d{1,3}[A-Z]?(?:\s*\(\w+\))*)(?:\s*(?:,|and|&|read with|to|r/w)\s*\d{1,3}[A-Z]?(?:\s*\(\w+\))*)*)")
MEETS = {("A", "1"): "Company's first Board meeting", ("A", "2"): "Board meeting that approves the annual accounts", ("A", "3"): "Periodic Board meetings",
         ("B", "1"): "Audit Committee meeting", ("B", "2"): "Nomination and Remuneration Committee meeting",
         ("C", "1"): "Annual General Meeting: usual business", ("C", "2"): "Annual or Extraordinary General Meeting: special business"}
PARTS = {"A": "Board meeting", "B": "Committee meeting", "C": "Shareholders' meeting"}


def find_source():
    for root in (Path("resolutions"), Path(".")):
        if not root.exists():
            continue
        for pat in ("*.pdf", "*.txt"):
            for f in sorted(root.rglob(pat)):
                if "resolution" in f.name.lower() and not any(x in f.parts for x in (".git", "raw", "data", "site", "drafts")):
                    return f
    return None


def read_text(f):
    if f.suffix.lower() == ".txt":
        return f.read_text(encoding="utf-8")
    import fitz
    return "\n".join(p.get_text("text") for p in fitz.open(f))


NOISE = re.compile(r"^(?:CimplyFive.{0,4}s Text of Model Resolutions under the Companies Act, 2013|Page \| \w+\s+(?:November|September) 20\d\d Edition)\s*$")
ID = re.compile(r"^(\d{1,2}\.\d{1,2}\.\d{1,2})$")
H1 = re.compile(r"^(\d{1,2})\.?\s+((?:Company.s First Board Meeting|Meeting in which Annual Accounts are approved|Periodic Board Meetings|Audit Committee Meeting|Nomination and Remuneration Committee Meeting|Annual General Meetings?(?: and Extraordinary General Meetings)?))\s*$")
H2 = re.compile(r"^(\d{1,2}\.\d{1,2})\s+([A-Z][^\n]{3,90})$")
PART = re.compile(r"^Part\s+([ABC])\s*[–-]\s*(.+)$")
LABELS = {"proposal": "Proposal", "resolution": "Proposed Resolution", "minutes": "Minutes – Gist of Discussion"}


def paras(lines):
    out, cur = [], []
    brk = re.compile(r"^(?:[“\"']?Resolved|RESOLVED|Further|Provided|\d+[.)]\s|\([a-z0-9ivx]+\)\s|[a-z]\)\s|•|Notes?\b|Explanation)")
    for l in lines:
        s = l.strip()
        if not s:
            continue
        if cur and brk.match(s):
            out.append(" ".join(cur)); cur = []
        cur.append(s)
        if re.search(r"[”\"]\s*$", s) and len(s) < 60:
            out.append(" ".join(cur)); cur = []
    if cur:
        out.append(" ".join(cur))
    return [re.sub(r"\s+", " ", p).strip() for p in out if p.strip()]


def sections_in(text):
    text = re.sub(r"[Ss]ections?\s+[^/;]{1,60}?\s+of\s+the\s+Companies\s+Act,?\s*1956", " ", text)
    got = []
    for m in SEC_RE.finditer(text):
        for n in re.findall(r"\d{1,3}[A-Z]?", re.sub(r"\(\w+\)", "", m.group(1))):
            if n not in got and 0 < int(re.match(r"\d+", n).group()) <= 470:
                got.append(n)
    return got


def forms_in(text):
    got = []
    for m in FORM_RE.finditer(text):
        if m.group(1) in FORMS:
            f = f"{m.group(1)}-{m.group(2)}"
            if f not in got:
                got.append(f)
    return got


def main():
    src = find_source()
    if not src:
        print("No book of model resolutions found (a file with 'resolution' in its name). Skipping.")
        return
    lines = [l for l in read_text(src).replace("\r", "").split("\n") if not NOISE.match(l.strip())]
    first = next((i for i, l in enumerate(lines) if l.strip().startswith("Agenda Title")), None)
    if first is None:
        print("Could not find the agenda items in the book.")
        return
    part, meet, group = "A", "", ""
    items, cur, field = [], None, None
    i = max(0, first - 40)
    # walk back to the part and heading in force before the first item
    for j in range(first, max(0, first - 60), -1):
        pass
    for k in range(0, first):
        s = lines[k].strip()
        pm = PART.match(s)
        if pm and k > 600:
            part = pm.group(1)
    start = max(0, first - 8)
    n = len(lines)
    k = start
    while k < n:
        s = lines[k].strip()
        pm = PART.match(s)
        if pm:
            part, meet, group = pm.group(1), "", ""
            cur = None
            k += 1
            continue
        m = ID.match(s)
        if m and k + 1 < n and lines[k + 1].strip().startswith("Agenda Title"):
            cur = {"id": f"{part}-{m.group(1)}", "p": part, "m": meet, "g": group, "t": "", "f": {"proposal": [], "resolution": [], "minutes": [], "explanatory": []}}
            items.append(cur)
            k += 1
            title = lines[k].strip().split(":", 1)[1].strip()
            while k + 1 < n and lines[k + 1].strip() and len(title) < 200 and not lines[k + 1].strip().startswith(("Proposal", "Proposed", "Explanatory", "Minutes")):
                k += 1
                title += " " + lines[k].strip()
            cur["t"] = re.sub(r"\s+", " ", title)
            field = None
            k += 1
            continue
        h2, h1 = H2.match(s), H1.match(s)
        if h2 and not re.search(r"[.”]$", s) and len(s) < 100:
            group = f"{h2.group(1)} {h2.group(2)}".strip(); cur = None; field = None; k += 1; continue
        if h1 and not re.search(r"[.”;,]$", s) and len(s) < 100 and not re.match(r"^\d+\.?\s+(?:Mr|Ms|Mrs|Dr)\b", s):
            meet = f"{h1.group(1)} {h1.group(2)}".strip(); group = ""; cur = None; field = None; k += 1; continue
        if cur is not None:
            if s == "Proposal":
                field = "proposal"
            elif s.startswith("Explanatory Statement") and len(s) < 30:
                field = "explanatory"
            elif s.startswith("Proposed Resolution"):
                field = "resolution"
            elif s.startswith("Minutes") and "Gist" in s:
                field = "minutes"
            elif field:
                cur["f"][field].append(s)
        k += 1

    groups = {}
    for it in items:
        pre = '.'.join(it['id'].split('-', 1)[1].split('.')[:2])
        if it['g'].startswith(pre + ' '):
            groups[(it['p'], pre)] = it['g']
    clean = []
    for it in items:
        first = it["id"].split("-", 1)[1].split(".")[0]
        d = {"id": it["id"], "t": it["t"], "p": it["p"], "m": MEETS.get((it["p"], first), it["m"]), "g": it["g"],
             "proposal": paras(it["f"]["proposal"] + it["f"]["explanatory"]), "resolution": paras(it["f"]["resolution"]), "minutes": paras(it["f"]["minutes"])}
        if not (d["proposal"] or d["resolution"] or d["minutes"]):
            continue
        pre = ".".join(it["id"].split("-", 1)[1].split(".")[:2])
        if not d["g"].startswith(pre + " "):
            d["g"] = groups.get((it["p"], pre), "")
        body = " ".join(d["proposal"] + d["resolution"])
        d["s"] = sections_in(body)
        d["f"] = forms_in(body)
        clean.append(d)
    if len(clean) < 20:
        print(f"Only {len(clean)} model resolutions were found; the layout of the book may be different from what was expected.")
        return

    out = SITE / "res"
    out.mkdir(parents=True, exist_ok=True)
    for f in out.glob("*.json"):
        f.unlink()
    for a, d in enumerate(clean):
        d["prev"] = [clean[a - 1]["id"], clean[a - 1]["t"]] if a else None
        d["next"] = [clean[a + 1]["id"], clean[a + 1]["t"]] if a + 1 < len(clean) else None
        (out / f"{d['id']}.json").write_text(json.dumps(d, ensure_ascii=False, separators=(",", ":")), encoding="utf-8")
    index = [{"id": d["id"], "t": d["t"], "p": d["p"], "m": d["m"], "g": d["g"], "s": d["s"], "f": d["f"]} for d in clean]
    (out / "index.json").write_text(json.dumps({"parts": PARTS, "items": index}, ensure_ascii=False, separators=(",", ":")), encoding="utf-8")
    (out / "text.json").write_text(json.dumps({d["id"]: " ".join(d["proposal"] + d["resolution"] + d["minutes"]).lower() for d in clean}, ensure_ascii=False, separators=(",", ":")), encoding="utf-8")

    by_sec = {}
    for d in clean:
        for n in d["s"]:
            by_sec.setdefault(n, []).append([d["id"], d["t"]])
    patched = 0
    for n, lst in by_sec.items():
        p = SITE / "s" / f"{n}.json"
        if p.exists():
            sj = json.loads(p.read_text(encoding="utf-8"))
            sj["mr"] = lst[:14]
            p.write_text(json.dumps(sj, ensure_ascii=False, separators=(",", ":")), encoding="utf-8")
            patched += 1
    idx = SITE / "index.json"
    try:
        j = json.loads(idx.read_text(encoding="utf-8"))
        j["res"] = len(clean)
        idx.write_text(json.dumps(j, ensure_ascii=False, separators=(",", ":")), encoding="utf-8")
    except (OSError, ValueError):
        pass
    print(f"{len(clean)} model resolutions written from {src}; {patched} section pages now link to them")


if __name__ == "__main__":
    try:
        main()
    except Exception as e:
        print(f"Model resolutions step ended with an error: {e}")
    sys.exit(0)
