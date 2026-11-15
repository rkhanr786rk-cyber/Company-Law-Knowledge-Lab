"""Builds the SEBI (LODR) Regulations pages for the website, from the LODR PDF kept in the lodr folder.
Also adds, to every Companies Act section page, a link to each Regulation that names that section."""
import bisect
import json
import re
import sys
from pathlib import Path

SITE = Path("site/data")
ROMAN = {"i", "ii", "iii", "iv", "v", "vi", "vii", "viii", "ix", "x", "xi", "xii", "xiii", "xiv", "xv", "xvi", "xvii", "xviii", "xix", "xx"}


def find_source():
    for root in (Path("lodr"), Path(".")):
        if not root.exists():
            continue
        for pat in ("*.pdf", "*.txt"):
            for f in sorted(root.rglob(pat)):
                if "lodr" in str(f).lower() or "listing-obligations" in str(f).lower():
                    if ".git" in f.parts or "raw" in f.parts or "data" in f.parts or "site" in f.parts:
                        continue
                    return f
    return None


def read_text(f):
    if f.suffix.lower() == ".txt":
        return f.read_text(encoding="utf-8")
    import fitz
    return "\n".join(p.get_text("text") for p in fitz.open(f))


def strip_footnotes(lines):
    """Drops the amendment footnotes printed at the foot of each page."""
    out, skip = [], False
    start = re.compile(r"^\d{1,3}\s+(?:Inserted|Substituted|Omitted|Deleted|The words?|Words?|Added|Re-?numbered|Renumbered|Sub-?regulation|Clause|Proviso|Explanation|Figure|Number|Heading|For the|Brought|Amended|Ins\b|Subs\b|Corrected|Rectified|Existing|The (?:figure|expression|word)|Regulation|Schedule|Para|Paragraph|Item|Table|Footnote|Earlier)")
    for l in lines:
        s = l.strip()
        if skip:
            if re.fullmatch(r"\d{1,3}", s):
                skip = False
            continue
        if start.match(s) and re.search(r"(?:w\.e\.f|Regulations?,?\s+20\d\d|Amendment|Gazette|ibid|notification|Circular|Board)", s, re.I):
            skip = True
            continue
        if re.fullmatch(r"\d{1,3}", s):
            continue
        out.append(l)
    return out


def clean(x):
    x = re.sub(r"\b\d{1,3}\s?\[\s*\*+\s*\]\d{0,3}", "", x)
    x = re.sub(r"\b\d{1,3}\s?\[", "", x)
    x = re.sub(r"\]\d{0,3}(?=\s|$|[,.;:])", "", x)
    x = x.replace(" ", " ")
    return re.sub(r"[ \t]+", " ", x).strip()


def head_clean(h):
    h = re.sub(r"\d{1,3}\s?[\[{]\**[\]}]?", "", h)
    h = re.sub(r"[\[\]{}]", "", h)
    return re.sub(r"\s+", " ", h).strip(" .")


def paragraphs(lines):
    paras, cur, prev_tag = [], None, ""
    for raw in lines:
        l = clean(raw)
        if not l:
            continue
        m = re.match(r"^(\(\s?[0-9A-Za-z]{1,5}\s?\))\s*(.*)", l)
        new = bool(m) or re.match(r"^(Provided|Explanation|Illustration|Note\b)", l) or re.match(r"^\d{1,3}[A-Z]{0,2}\.\s", l)
        if new or cur is None:
            lvl, tag = 0, ""
            if m:
                tag = m.group(1).strip("() ").lower()
                if tag[0].isdigit():
                    lvl = 0
                elif tag in ROMAN and not (tag in ("i", "v", "x") and prev_tag in ("h", "u", "w")):
                    lvl = 2
                else:
                    lvl = 1
                prev_tag = tag
            elif l.startswith(("Provided", "Explanation")):
                lvl = 1
            cur = {"x": l, "l": lvl}
            paras.append(cur)
        else:
            cur["x"] += " " + l
    return paras


def main():
    src = find_source()
    if not src:
        print("No LODR file found (put the LODR PDF in a folder named lodr). Skipping.")
        return
    text = read_text(src)
    raw_lines = [l.replace("\u00a0", " ").replace("\t", " ") for l in text.split("\n")]
    # some PDFs print a regulation number alone on its line: join it with the text that follows
    merged, k = [], 0
    while k < len(raw_lines):
        l = raw_lines[k]
        if re.fullmatch(r"\s*\d{1,3}[A-Z]{0,2}\s?\.\s*", l) and k + 1 < len(raw_lines):
            nxt = k + 1
            while nxt < len(raw_lines) and not raw_lines[nxt].strip():
                nxt += 1
            if nxt < len(raw_lines):
                merged.append(l.strip().replace(" ", "") + " " + raw_lines[nxt].strip())
                k = nxt + 1
                continue
        merged.append(re.sub(r"^(\s*\d{1,3}[A-Z]{0,2})\s+\.\s", r"\1. ", l))
        k += 1
    lines = strip_footnotes(merged)
    n_sched_start = next((i for i, l in enumerate(lines) if re.match(r"^SCHEDULE\s+I\b", l.strip()) and i > len(lines) * 0.5), len(lines))

    marks, chap = [], ""
    for i, l in enumerate(lines):
        s = l.strip()
        c = re.match(r"^CHAPTER\s+([IVX]+)\s*$", s)
        if c and i < n_sched_start:
            t = lines[i + 1].strip() if i + 1 < len(lines) else ""
            chap = f"Chapter {c.group(1)}: {head_clean(t).title()}" if t and len(t) < 120 else f"Chapter {c.group(1)}"
            marks.append((i, "chap", chap, ""))
            continue
        m = re.match(r"^(\d{1,3}[A-Z]{0,2})\.\s+\S", l)
        if m and 5 < i < n_sched_start and i > 0:
            p = lines[i - 1].strip()
            if p and not p.startswith("(") and len(p) < 160 and not p.endswith((";", ",", ":")):
                hh = head_clean(p)
                if (hh[:1].islower() or hh.lower().startswith(("of ", "and "))) and i > 1:
                    hh = head_clean(lines[i - 2]) + " " + hh
                if hh[:1].islower() and i > 2:
                    hh = head_clean(lines[i - 3]) + " " + hh
                marks.append((i, "reg", m.group(1), hh))
            continue
        sc = re.match(r"^SCHEDULE\s+([IVX]+)\b\s*[:\-–]?\s*(.*)", s)
        if sc and i >= n_sched_start - 1 and len(s) < 110:
            marks.append((i, "sch", sc.group(1), head_clean(sc.group(2)) or f"Schedule {sc.group(1)}"))
    # keep regulations in running order only (drops numbered list items inside schedules / notes)
    good, last = [], 0
    for mk in marks:
        if mk[1] != "reg":
            good.append(mk)
            continue
        n = int(re.match(r"\d+", mk[2]).group())
        if last <= n <= last + 4 or (last == 0 and n <= 2):
            good.append(mk)
            last = n
    seen_sched = set()
    items = []  # id, kind, title, chapter, start line
    chap_now = ""
    for mk in good:
        if mk[1] == "chap":
            chap_now = mk[2]
        elif mk[1] == "reg":
            items.append({"id": mk[2], "k": "r", "t": mk[3], "ch": chap_now, "i": mk[0]})
        elif mk[2] not in seen_sched:
            seen_sched.add(mk[2])
            items.append({"id": "S" + mk[2], "k": "s", "t": mk[3], "ch": "Schedules", "i": mk[0]})
    if len(items) < 20:
        print("LODR file read, but too few regulations found:", len(items))
        Path("data").mkdir(exist_ok=True)
        cand = [f"{i}: {l[:110]}" for i, l in enumerate(lines) if re.match(r"^\s*\d{1,3}[A-Z]{0,2}\s?\.", l)][:25]
        first = [f"{i}: {l[:110]}" for i, l in enumerate(lines[:60])]
        Path("data/lodr_report.md").write_text("# LODR report (problem)\n\n- File: %s\n- Lines read: %d\n- Regulations found: %d\n\n## First lines\n\n```\n%s\n```\n\n## Lines that look like a regulation start\n\n```\n%s\n```\n" % (src, len(lines), len(items), "\n".join(first), "\n".join(cand)), encoding="utf-8")
        return
    items.sort(key=lambda e: e["i"])
    (SITE / "lodr" / "r").mkdir(parents=True, exist_ok=True)

    offs, c = [], 0
    for l in lines:
        offs.append(c)
        c += len(l) + 1
    body = "\n".join(lines)
    starts = [e["i"] for e in items]
    cites = {}  # regulation id -> set of sections
    sec_hits = {}  # section -> [ids]
    for m in re.finditer(r"Companies\s+Act,?\s*(?:2013|1956)?", body):
        if "1956" in m.group(0):
            continue
        pre = body[max(0, m.start() - 260): m.start()]
        cand = list(re.finditer(r"(?<![A-Za-z-])sections?\s+((?:\(?\d{1,3}[A-Z]{0,2}\)?(?:\s*\([^)]*\))*\s*(?:,|and|to|or|-|–|read with|/)?\s*)+)", pre, re.I))
        if not cand or pre[cand[-1].end():].strip() not in ("of", "of the", ""):
            continue
        ln = bisect.bisect_right(offs, m.start()) - 1
        j = bisect.bisect_right(starts, ln) - 1
        if j < 0:
            continue
        for nn in dict.fromkeys(re.findall(r"(?<![\d(])(\d{1,3}[A-Z]{0,2})(?!\d)", re.sub(r"\([^)]*\)", "", cand[-1].group(1)))):
            if nn[0].isdigit() and 1 <= int(re.match(r"\d+", nn).group()) <= 470:
                cites.setdefault(items[j]["id"], set()).add(nn)
                sec_hits.setdefault(nn, set()).add(items[j]["id"])

    index, full = [], {}
    for k, e in enumerate(items):
        end = items[k + 1]["i"] if k + 1 < len(items) else len(lines)
        seg = lines[e["i"]:end]
        seg = [l for l in seg if not re.match(r"^CHAPTER\s+[IVX]+\s*$", l.strip())]
        if e["k"] == "r" and seg:
            seg[0] = re.sub(r"^\d{1,3}[A-Z]{0,2}\.\s*", "", seg[0])
        ps = paragraphs(seg)
        prev = items[k - 1] if k else None
        nxt = items[k + 1] if k + 1 < len(items) else None
        dump = {"id": e["id"], "t": e["t"], "ch": e["ch"], "k": e["k"], "paras": ps,
                "cites": sorted(cites.get(e["id"], []), key=lambda s: (int(re.match(r"\d+", s).group()), s)),
                "prev": [prev["id"], prev["t"]] if prev else None, "next": [nxt["id"], nxt["t"]] if nxt else None}
        (SITE / "lodr" / "r" / f"{e['id']}.json").write_text(json.dumps(dump, ensure_ascii=False, separators=(",", ":")), encoding="utf-8")
        index.append({"id": e["id"], "t": e["t"], "ch": e["ch"], "k": e["k"]})
        full[e["id"]] = " ".join(p["x"] for p in ps).lower()
    (SITE / "lodr" / "index.json").write_text(json.dumps(index, ensure_ascii=False, separators=(",", ":")), encoding="utf-8")
    (SITE / "lodr" / "text.json").write_text(json.dumps(full, ensure_ascii=False, separators=(",", ":")), encoding="utf-8")

    # tell the Act's section pages and the site index
    title = {e["id"]: e["t"] for e in items}
    for n, ids in sec_hits.items():
        f = SITE / "s" / f"{n}.json"
        if not f.exists():
            continue
        d = json.loads(f.read_text(encoding="utf-8"))
        gn = d.setdefault("gn", {})
        have = {x.get("id") or x.get("regulation") for x in gn.get("lodr", [])}
        lst = gn.setdefault("lodr", [])
        for rid in sorted(ids, key=lambda s: (items_pos(items, s))):
            lab = ("Reg " + rid) if not rid.startswith("S") else "Schedule " + rid[1:]
            if rid in have or lab in have:
                continue
            lst.append({"id": rid, "regulation": lab, "note": title[rid]})
        for x in lst:  # rows typed by hand in guide/lodr.csv get a link when they name a regulation
            mm = re.match(r"^Reg(?:ulation)?\s+(\d{1,3}[A-Z]{0,2})$", x.get("regulation", "").strip(), re.I)
            if mm and not x.get("id") and mm.group(1) in title:
                x["id"] = mm.group(1)
        f.write_text(json.dumps(d, ensure_ascii=False, separators=(",", ":")), encoding="utf-8")
    ix = SITE / "index.json"
    d = json.loads(ix.read_text(encoding="utf-8"))
    d["lodr"] = len(items)
    ix.write_text(json.dumps(d, ensure_ascii=False, separators=(",", ":")), encoding="utf-8")
    rep = [f"- LODR source: {src}", f"- Regulations and schedules written: {len(items)}", f"- Companies Act sections that now link to a Regulation: {len(sec_hits)}"]
    Path("data").mkdir(exist_ok=True)
    Path("data/lodr_report.md").write_text("# LODR report\n\n" + "\n".join(rep) + "\n", encoding="utf-8")
    print("\n".join(rep))


def items_pos(items, rid):
    for k, e in enumerate(items):
        if e["id"] == rid:
            return k
    return 0


if __name__ == "__main__":
    main()
