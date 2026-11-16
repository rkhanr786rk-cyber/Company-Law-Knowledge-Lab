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
    strong = re.compile(r"^\d{1,3}\s+(?:Inserted|Substituted|Omitted|Deleted|Added|Re-?numbered|Renumbered|Subs\b|Ins\b|Corrected|Rectified|Brought|Amended|Heading|For the words?|The words?|Words?|The figure|The expression|The word)\b")
    weak = re.compile(r"^\d{1,3}\s+(?:Sub-?regulation|Clause|Proviso|Explanation|Figure|Number|Regulation|Schedule|Para|Paragraph|Item|Table|Footnote|Earlier|Existing)")
    keys = re.compile(r"(?:w\.e\.f|Regulations?,?\s+20\d\d|Amendment|Gazette|ibid|notification|Circular)", re.I)
    for l in lines:
        s = l.strip()
        if skip:
            if re.fullmatch(r"\d{1,3}", s):
                skip = False
            continue
        if re.search(r"(?:Amendment\)?\s*Regulations,?\s*(?:19|20)\d\d,?\s*w\.e\.f)|^Obligations and Disclosure Requirements\)", s):
            continue
        if strong.match(s) or (weak.match(s) and keys.search(s)):
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


# Regulations and Companies Act sections that deal with the same subject, even where the text does not name the section.
# Shown to the reader as "related (topic match)", not as a citation.
RELATED = {
    "17": ["149", "152", "173"], "18": ["177"], "19": ["178"], "20": ["178"], "22": ["177"], "23": ["188", "177"],
    "24A": ["204"], "25": ["149", "150", "152"], "26": ["165"], "29": ["173"], "33": ["129", "134"], "34": ["134", "135", "136"],
    "36": ["136", "101"], "37": ["230", "232"], "40": ["56"], "42": ["91"], "43": ["123"], "44": ["108", "110"], "45": ["13"],
}
NUM = re.compile(r"(?<![\d(])(\d{1,3}[A-Z]{0,2})(?!\d)")


def para_links(x):
    """Positions of the Companies Act section numbers named in a paragraph: [[start, end, section]]."""
    out = []
    for m in re.finditer(r"Companies\s+Act,?\s*(?:2013|1956)?", x):
        if "1956" in m.group(0):
            continue
        lo = max(0, m.start() - 260)
        pre = x[lo:m.start()]
        cand = list(re.finditer(r"(?<![A-Za-z-])sections?\s+((?:\(?\d{1,3}[A-Z]{0,2}\)?(?:\s*\([^)]*\))*\s*(?:,|and|to|or|-|–|read with|/)?\s*)+)", pre, re.I))
        if not cand or pre[cand[-1].end():].strip() not in ("of", "of the", ""):
            continue
        g0 = cand[-1].start(1)
        for n in NUM.finditer(cand[-1].group(1)):
            sec = n.group(1)
            if sec[0].isdigit() and 1 <= int(re.match(r"\d+", sec).group()) <= 470:
                out.append([lo + g0 + n.start(), lo + g0 + n.end(), sec])
    return out


def main():
    src = find_source()
    if not src:
        print("No LODR file found (put the LODR PDF in a folder named lodr). Skipping.")
        return
    text = read_text(src)
    raw_lines = [re.sub(r"[\x00-\x08\x0b-\x1f\x7f\u00ad]", "", l.replace("\u00a0", " ").replace("\t", " ")) for l in text.split("\n")]
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
            j_ = i + 2
            while j_ < len(lines) and j_ < i + 4 and lines[j_].strip() and lines[j_].strip().upper() == lines[j_].strip() and len(lines[j_].strip()) < 90 and not re.match(r"^\d", lines[j_].strip()):
                t += " " + lines[j_].strip()
                j_ += 1
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
    pos = {e["id"]: k for k, e in enumerate(items)}
    sec_rows = []
    try:
        sec_rows = json.loads((SITE / "index.json").read_text(encoding="utf-8"))["sections"]
    except (OSError, ValueError, KeyError):
        pass
    act_title = {r[0]: r[1] for r in sec_rows}
    act_inforce = {r[0] for r in sec_rows if not r[4]}

    # 1. sections of the Companies Act that a Regulation names outright
    cites, why = {}, {}  # regulation id -> sections ; (reg, section) -> sentence
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
        win = re.sub(r"\s+", " ", clean(body[max(0, m.start() - 220): m.end() + 40])).strip()
        cut = re.search(r"[.;:]\s+(?=[A-Z(])", win[:150])
        snippet = win[cut.end():] if cut else win
        for nn in dict.fromkeys(re.findall(r"(?<![\d(])(\d{1,3}[A-Z]{0,2})(?!\d)", re.sub(r"\([^)]*\)", "", cand[-1].group(1)))):
            if nn[0].isdigit() and 1 <= int(re.match(r"\d+", nn).group()) <= 470:
                cites.setdefault(items[j]["id"], set()).add(nn)
                why.setdefault((items[j]["id"], nn), snippet)

    # 2. Regulations and Schedules that a Regulation names (inside the LODR itself)
    ext = re.compile(r"^\s*(?:and\s+\S+\s+)?of\s+(?:the\s+)?(?:Securities|SEBI|Depositories|Prohibition|Substantial|Issue|Delisting|Buy|Mutual|Alternative|Foreign|Prevention|Companies|Merchant|Credit|Stock|Real|Infrastructure|Self|Research|Investment|Portfolio|Custodian|Registrars|Banker|Debenture|Insolvency|Income|Contract|Indian)", re.I)
    reg_re = re.compile(r"(?<![A-Za-z])(?:sub-)?regulations?\s+(\d{1,3}[A-Z]{0,2}(?:\s*\([0-9A-Za-z]{1,4}\))*(?:\s*(?:,|and|to|or|-|–)\s*\d{1,3}[A-Z]{0,2}(?:\s*\([0-9A-Za-z]{1,4}\))*)*)", re.I)
    sch_re = re.compile(r"Schedule\s+([IVX]{1,5})\b", re.I)
    out, wo = {}, {}
    sched_ids = {e["id"] for e in items if e["k"] == "s"}
    for k, e in enumerate(items):
        end = items[k + 1]["i"] if k + 1 < len(items) else len(lines)
        seg = [l for l in lines[e["i"]:end]]
        txt = clean(" ".join(seg))
        found = {}
        for m in reg_re.finditer(txt):
            tail = txt[m.end(): m.end() + 140]
            if ext.match(tail) or re.match(r"^[^.;]{0,100}?(?:Regulations|Rules|Act),?\s*(?:19|20)\d\d", tail) and not re.match(r"^\s*of\s+(?:these|this)\b", tail):
                continue
            spec = re.sub(r"\([^)]*\)", " ", m.group(1))
            ids = re.findall(r"(?<![\d(])(\d{1,3}[A-Z]{0,2})(?![\d)])", spec)
            if re.search(r"\bto\b|-|–", spec) and len(ids) == 2 and ids[0] in pos and ids[1] in pos:
                lo, hi = sorted((pos[ids[0]], pos[ids[1]]))
                ids = [items[x]["id"] for x in range(lo, hi + 1) if items[x]["k"] == "r"]
            sn = txt[max(0, m.start() - 150): m.end() + 60].strip()
            for i_ in ids:
                if i_ in pos and i_ != e["id"]:
                    found.setdefault(i_, sn)
        for m in sch_re.finditer(txt):
            tail = txt[m.end(): m.end() + 80]
            sid = "S" + m.group(1).upper()
            if ext.match(tail) or sid not in sched_ids or sid == e["id"]:
                continue
            found.setdefault(sid, txt[max(0, m.start() - 150): m.end() + 60].strip())
        if found:
            out[e["id"]] = list(found)
            wo[e["id"]] = found
    inn, wi = {}, {}
    for a, lst in out.items():
        for b in lst:
            inn.setdefault(b, []).append(a)
            wi.setdefault(b, {})[a] = wo[a][b]

    # 3. Sections of the Act whose heading shares subject words with a Regulation heading (to be verified)
    STOPW = set("the and for of in on to by as at or an a with from under these such other regulations regulation listed entity entities certain applicable provisions provision requirements requirement obligations obligation information disclosure disclosures obligations act shall which any all".split())

    def stems(t):
        out_ = set()
        for w in re.findall(r"[A-Za-z]{3,}", t):
            w = w.lower()
            if w in STOPW:
                continue
            for suf in ("ations", "ation", "ions", "ion", "ies", "es", "s"):
                if w.endswith(suf) and len(w) - len(suf) >= 4:
                    w = w[: -len(suf)] + ("y" if suf == "ies" else "")
                    break
            out_.add(w)
        return out_
    a_st = {n: stems(t) for n, t in act_title.items() if n in act_inforce}
    df = {}
    for st in a_st.values():
        for w in st:
            df[w] = df.get(w, 0) + 1
    N = max(len(a_st), 1)
    import math
    rel = {}
    for e in items:
        rs = stems(e["t"])
        if not rs:
            continue
        got = []
        for n, st in a_st.items():
            if n in cites.get(e["id"], ()):
                continue
            sh = rs & st
            if not sh:
                continue
            score = sum(math.log(N / df[w]) for w in sh)
            whole = rs <= st and (len(rs) >= 2 or df.get(next(iter(rs)), 99) <= 3)
            if whole or (len(sh) >= 2 and score >= 5.0) or (len(sh) == 1 and score >= 4.6 and len(rs) <= 2 and len(st) <= 3):
                got.append((score, n, sorted(sh)))
        got.sort(key=lambda g: (-g[0], len(g[1]), g[1]))
        if got:
            rel[e["id"]] = [(n, sh) for _, n, sh in got[:6]]

    # 4. explanations written by the AI step, if any
    ai_dir = Path("data/ai/l")
    ai_ids = sorted(f.stem for f in ai_dir.glob("*.json")) if ai_dir.exists() else []
    if ai_ids:
        (SITE / "ai" / "l").mkdir(parents=True, exist_ok=True)
        for i_ in ai_ids:
            (SITE / "ai" / "l" / f"{i_}.json").write_bytes((ai_dir / f"{i_}.json").read_bytes())

    index, full, plain = [], {}, []
    sec_hits = {}  # section -> {regulation id: ("cite"|"rel", detail)}
    for k, e in enumerate(items):
        end = items[k + 1]["i"] if k + 1 < len(items) else len(lines)
        seg = [l for l in lines[e["i"]:end] if not re.match(r"^CHAPTER\s+[IVX]+\s*$", l.strip())]
        if e["k"] == "r" and seg:
            seg[0] = re.sub(r"^\d{1,3}[A-Z]{0,2}\.\s*", "", seg[0])
        ps = paragraphs(seg)
        prev = items[k - 1] if k else None
        nxt = items[k + 1] if k + 1 < len(items) else None
        sec_sort = lambda s_: (int(re.match(r"\d+", s_).group()), s_)
        acts = [{"n": n, "w": "c", "x": why.get((e["id"], n), "")} for n in sorted(cites.get(e["id"], []), key=sec_sort)]
        acts += [{"n": n, "w": "r", "x": ", ".join(sh)} for n, sh in rel.get(e["id"], [])]
        dump = {"id": e["id"], "t": e["t"], "ch": e["ch"], "k": e["k"], "paras": ps,
                "cites": sorted(cites.get(e["id"], []), key=sec_sort),
                "acts": [dict(a_, t=act_title.get(a_["n"], "")) for a_ in acts],
                "out": out.get(e["id"], []), "in": sorted(inn.get(e["id"], []), key=lambda x: pos.get(x, 0)),
                "wo": wo.get(e["id"], {}), "wi": wi.get(e["id"], {}),
                "prev": [prev["id"], prev["t"]] if prev else None, "next": [nxt["id"], nxt["t"]] if nxt else None}
        (SITE / "lodr" / "r" / f"{e['id']}.json").write_text(json.dumps(dump, ensure_ascii=False, separators=(",", ":")), encoding="utf-8")
        index.append({"id": e["id"], "t": e["t"], "ch": e["ch"], "k": e["k"]})
        full[e["id"]] = " ".join(p["x"] for p in ps).lower()
        plain.append({"id": e["id"], "t": e["t"], "ch": e["ch"], "k": e["k"], "text": "\n".join(p["x"] for p in ps)})
        for a_ in acts:
            sec_hits.setdefault(a_["n"], {}).setdefault(e["id"], (a_["w"], a_["x"]))
    (SITE / "lodr" / "index.json").write_text(json.dumps(index, ensure_ascii=False, separators=(",", ":")), encoding="utf-8")
    (SITE / "lodr" / "text.json").write_text(json.dumps(full, ensure_ascii=False, separators=(",", ":")), encoding="utf-8")
    Path("data").mkdir(exist_ok=True)
    Path("data/lodr.json").write_text(json.dumps(plain, ensure_ascii=False, separators=(",", ":")), encoding="utf-8")  # read by the AI explain step

    # tell the Act's section pages and the site index
    title = {e["id"]: e["t"] for e in items}
    for n, regs in sec_hits.items():
        f = SITE / "s" / f"{n}.json"
        if not f.exists():
            continue
        d = json.loads(f.read_text(encoding="utf-8"))
        gn = d.setdefault("gn", {})
        lst = gn.setdefault("lodr", [])
        have = {x.get("id") or x.get("regulation") for x in lst}
        for rid in sorted(regs, key=lambda s_: pos.get(s_, 0)):
            lab = ("Reg " + rid) if not rid.startswith("S") else "Schedule " + rid[1:]
            if rid in have or lab in have:
                continue
            w, x = regs[rid]
            lst.append({"id": rid, "regulation": lab, "note": title[rid], "rel": w == "r", "x": x})
        for x in lst:  # rows typed by hand in guide/lodr.csv get a link when they name a regulation
            mm = re.match(r"^Reg(?:ulation)?\s+(\d{1,3}[A-Z]{0,2})$", x.get("regulation", "").strip(), re.I)
            if mm and not x.get("id") and mm.group(1) in title:
                x["id"] = mm.group(1)
        lst.sort(key=lambda x: (x.get("rel", False), pos.get(x.get("id"), 999)))
        f.write_text(json.dumps(d, ensure_ascii=False, separators=(",", ":")), encoding="utf-8")
    ix = SITE / "index.json"
    d = json.loads(ix.read_text(encoding="utf-8"))
    d["lodr"] = len(items)
    d["ai_l"] = ai_ids
    ix.write_text(json.dumps(d, ensure_ascii=False, separators=(",", ":")), encoding="utf-8")
    n_cite = sum(1 for n, r in sec_hits.items() if any(v[0] == "c" for v in r.values()))
    rep = [f"- LODR source: {src}", f"- Regulations and schedules written: {len(items)}",
           f"- Companies Act sections named by a Regulation: {n_cite}",
           f"- Companies Act sections linked by subject (to verify): {len(sec_hits) - n_cite}",
           f"- Regulations that refer to other Regulations or Schedules: {len(out)}",
           f"- Regulations with a plain-language explanation: {len(ai_ids)}"]
    Path("data/lodr_report.md").write_text("# LODR report\n\n" + "\n".join(rep) + "\n", encoding="utf-8")
    print("\n".join(rep))


def items_pos(items, rid):
    for k, e in enumerate(items):
        if e["id"] == rid:
            return k
    return 0


if __name__ == "__main__":
    main()
