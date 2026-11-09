"""Step 4: turn data/*.json into the website.

Reads  data/sections.json, data/rules.json, data/links.json and web/index.html
Writes site/ (index.html + data files the page loads on demand)
  site/data/index.json      list of sections, chapters, rule sets and rule titles (small, loaded at start)
  site/data/fulltext.json   word -> sections, for keyword search (loaded on first search)
  site/data/defs.json       defined terms from section 2, for hover definitions
  site/data/s/<n>.json      one file per section: reflowed text, notes, linked rules
  site/data/r/<k>.json      one file per rule set: reflowed text of every rule
"""
import datetime
import json
import re
import shutil
import sys
from pathlib import Path

DATA, OUT, PAGE = Path("data"), Path("site"), Path("web/index.html")
ACT_NAME = "The Companies Act, 2013"

# ---------- text reflow: PDF text has a line break at every line end; rebuild real paragraphs ----------
LONE_NUM = re.compile(r"^\s*\d{1,3}\.?\s*$")
HEADER_NOISE = re.compile(r"^\s*(?:the companies act, 2013|page \d+.*)\s*$", re.I)
ACT_NOTE = re.compile(r"^\s*\d{1,3}\.\s+(?:Ins\.|Subs\.|Omitted\b|Rep\.|The words\b|Added\b|Re-?numbered\b|Substituted\b|Inserted\b|Amended\b)")
RULE_NOTE = re.compile(r"^\s*(?:Substituted|Inserted|Omitted|Amended|Added|Renumbered)\s+by\b", re.I)
LABEL = re.compile(r"^(?:\[|\d{1,2}\[)*\(\s*(\d{1,3}[A-Z]{0,2}|[a-z]{1,3}|[A-Z]{1,2})\s*\)\s*")
PROV = re.compile(r"^(?:\[|\d{1,2}\[)*\s*(Provided|Explanation|Illustration|Exception)\b")
CONTINUES = re.compile(r"\b(?:section|sections|sub-section|sub-sections|clause|clauses|sub-clause|sub-clauses|paragraph|rule|rules|"
                       r"regulation|item|items|under|of|in|to|by|with|from|the|a|an)$", re.I)
AND_OR = re.compile(r"\b(?:and|or)$", re.I)
ROMAN = re.compile(r"^(?:i{1,3}|iv|v|vi{0,3}|ix|x{1,3})$")
MARK = re.compile(r"(?<![\w)\]])\d{1,2}\[")          # footnote markers such as 3[ ... ]


def sort_key(num):
    m = re.match(r"(\d+)([A-Z]*)", num)
    return int(m.group(1)), m.group(2)


def strip_trailing_headings(text):
    """The body of a section ends with the heading of the next chapter or part; remove it."""
    lines = text.rstrip().split("\n")

    def heading(line):
        t = line.strip()
        return (not t or LONE_NUM.match(t) or re.match(r"^(?:CHAPTER|PART)\s+[IVXLC]+[A-Z]?\b", t)
                or (t.isupper() and len(t) <= 120))

    while lines and heading(lines[-1]):
        lines.pop()
    return "\n".join(lines)


def is_start(line, prev):
    if PROV.match(line):
        return True
    m = LABEL.match(line)
    if not m:
        return False
    p = prev.strip()
    if p and CONTINUES.search(p):
        return False
    if p and m.group(1)[0].isdigit() and AND_OR.search(p):
        return False
    return True


def kind_of(label, prev_kind, prev_label):
    if label[0].isdigit():
        return "sub"
    if label == "i":
        return "subclause" if prev_kind in ("clause", "subclause") and prev_label != "h" else "clause"
    if len(label) > 1 and ROMAN.match(label):
        return "subclause"
    return "clause"


def to_paragraphs(text, rule=False):
    note_re = RULE_NOTE if rule else ACT_NOTE
    paras, notes, cur, note, prev = [], [], None, None, ""
    for raw in text.replace("\r", "").split("\n"):
        ln = raw.strip()
        if not ln or LONE_NUM.match(ln) or HEADER_NOISE.match(ln):
            continue
        if note is not None:
            note.append(ln)
            if ln.endswith((".", '."', ".”")) or len(note) >= 8:
                notes.append(" ".join(note))
                note = None
            continue
        if note_re.match(ln):
            note = [ln]
            if ln.endswith("."):
                notes.append(ln)
                note = None
            continue
        if cur is None or is_start(ln, prev):
            pm = PROV.match(ln)
            lm = LABEL.match(ln)
            if pm:
                cur = {"k": "proviso" if pm.group(1) == "Provided" else "expl", "l": "", "x": ln}
            elif lm:
                last = paras[-1] if paras else {"k": "", "l": ""}
                lab = lm.group(1)
                cur = {"k": kind_of(lab, last["k"], last["l"].strip("()")), "l": f"({lab})", "x": ln[lm.end():]}
            else:
                cur = {"k": "p", "l": "", "x": ln}
            paras.append(cur)
        else:
            cur["x"] = cur["x"] + ("" if cur["x"].endswith("-") else " ") + ln
        prev = ln
    for p in paras:
        p["x"] = MARK.sub("[", p["x"]).strip()
    return [p for p in paras if p["x"] or p["l"]], notes


# ---------- definitions from section 2 ----------
LEAD_TERMS = re.compile(r"^(?:[“\"][^”\"]{2,60}[”\"](?:\s*(?:,|or|and)\s*)?)+")


def extract_definitions(paras):
    defs, cur = {}, None
    for p in paras:
        if p["k"] == "sub":
            cur = None
            m = LEAD_TERMS.match(p["x"])
            if m:
                terms = re.findall(r"[“\"]([^”\"]{2,60})[”\"]", m.group(0))
                cur = {"d": f'{p["l"]} {p["x"]}', "terms": terms}
                for t in terms:
                    defs[t.lower()] = cur
        elif cur is not None and p["k"] in ("clause", "subclause") and len(cur["d"]) < 700:
            cur["d"] += f' {p["l"]} {p["x"]}'
    out = {}
    for key, v in defs.items():
        if re.match(r"^[A-Za-z][A-Za-z\- ']{2,60}$", key):
            out[key] = {"t": key, "d": v["d"][:700]}
    return out



# ---------- which paragraph of a rule refers to which sub-section of a section ----------
SEC_RE = re.compile(
    r"\b(?i:sections?)\s+(\d{1,3}[A-Z]{0,2}(?:\s*\([^)]{1,8}\))*"
    r"(?:\s*(?:,|and|or|to|&)\s*\d{1,3}[A-Z]{0,2}(?:\s*\([^)]{1,8}\))*)*)")
SUBOF_RE = re.compile(
    r"\b(?i:sub-sections?)\s*\(\s*(\d{1,3}[A-Z]?)\s*\)(?:\s*\([a-zA-Z0-9]{1,4}\))*"
    r"(?:\s*(?:,|and|or|to)\s*\(\s*\d{1,3}[A-Z]?\s*\))*"
    r"(?:[^()]{0,45}?)?\s*(?i:of|to|in)\s+(?i:the\s+)?(?i:section)\s+(\d{1,3}[A-Z]{0,2})")
OF_RE = re.compile(
    r"\bof\s+(?:the\s+)?((?:Companies\s+)?Act\b|[A-Z][A-Za-z’'()\- ,]{2,80}?\b(?:Act|Code|Regulations|Rules)\b)")


def cites_companies_act(tail):
    m = OF_RE.search(tail[:110])
    if not m:
        return True
    if not re.match(r"(?:Companies\s+)?Act\b", m.group(1)):
        return False
    return not re.match(r"\s*,?\s*19\d\d", tail[m.end(1):m.end(1) + 10])   # the old Companies Act, 1956


def snippet(text, start, width=200):
    """A short quotation around a citation, cut at word boundaries."""
    a, b = max(0, start - 70), min(len(text), start + width - 70)
    if a:
        k = text.find(" ", a)
        a = k + 1 if 0 <= k < a + 25 else a
    if b < len(text):
        k = text.rfind(" ", a, b)
        b = k if k > a else b
    return ("\u2026" if a else "") + text[a:b].strip() + ("\u2026" if b < len(text) else "")


def para_refs(text, pos=None):
    """{section number: set of first-level sub-section labels ('' = the section as a whole)}; pos gets where each first appears"""
    out, spans = {}, []
    for m in SUBOF_RE.finditer(text):
        if cites_companies_act(text[m.end():m.end() + 110]):
            out.setdefault(m.group(2), set()).add(m.group(1))
            if pos is not None:
                pos.setdefault(m.group(2), m.start())
        spans.append((m.start(), m.end()))
    for m in SEC_RE.finditer(text):
        if any(a <= m.start() < b for a, b in spans) or not cites_companies_act(text[m.end():m.end() + 110]):
            continue
        for it in re.finditer(r"(\d{1,3}[A-Z]{0,2})((?:\s*\([^)]{1,8}\))*)", m.group(1)):
            subs = re.findall(r"\(([^)]{1,8})\)", it.group(2))
            out.setdefault(it.group(1), set()).add(subs[0] if subs else "")
            if pos is not None:
                pos.setdefault(it.group(1), m.start())
    return out


STOP = set("the and for are was were that this with from have has had not but any all may shall such under which their its "
           "his her they them than then thereof therein thereto into onto upon also been being per who whom whose where when "
           "what other more most some each every either neither within without".split())


def main():
    for f in ("sections.json", "rules.json", "links.json"):
        if not (DATA / f).exists():
            sys.exit(f"Missing data/{f}. Run the earlier steps first.")
    if not PAGE.exists():
        sys.exit("Missing web/index.html. Create it in the repo (Add file, Create new file, name it web/index.html).")

    sections = json.loads((DATA / "sections.json").read_text(encoding="utf-8"))
    rules_doc = json.loads((DATA / "rules.json").read_text(encoding="utf-8"))
    links = json.loads((DATA / "links.json").read_text(encoding="utf-8"))["by_section"]
    sections.sort(key=lambda s: sort_key(s["section"]))

    if OUT.exists():
        shutil.rmtree(OUT)
    (OUT / "data" / "s").mkdir(parents=True)
    (OUT / "data" / "r").mkdir(parents=True)
    (OUT / ".nojekyll").write_text("", encoding="utf-8")
    shutil.copy(PAGE, OUT / "index.html")

    def dump(path, obj):
        path.write_text(json.dumps(obj, ensure_ascii=False, separators=(",", ":")), encoding="utf-8")

    # ----- rule sets: one file per set with the reflowed text of every rule -----
    set_index, set_pos, rule_paras = [], {}, {}
    by_id = {}
    for r in rules_doc["rules"]:
        by_id[r["id"]] = r
    for k, rs in enumerate(rules_doc["rule_sets"]):
        set_pos[rs["id"]] = k
        body = {}
        for r in rules_doc["rules"]:
            if r["rule_set_id"] != rs["id"]:
                continue
            title, text = r["title"], r["text"]
            cut = re.search(r"\.\s+(?=\(1\)\s)", title)          # a heading that swallowed the first paragraph
            if cut and len(title) > 90:
                title, text = title[:cut.start()], title[cut.end():] + "\n" + text
            r["title"] = title
            paras, notes = to_paragraphs(text, rule=True) if r["status"] == "in force" else ([], [])
            rule_paras[r["id"]] = paras
            body[r["rule"]] = {"t": title, "st": r["status"], "paras": paras, "notes": notes}
        dump(OUT / "data" / "r" / f"{k}.json", {"name": rs["name"], "rules": body})
        set_index.append({"k": k, "name": rs["name"], "n": rs["rule_count"]})

    # ----- sections -----
    order = [s["section"] for s in sections]
    title_of = {s["section"]: s["title"] for s in sections}
    rank = {"high": 0, "medium": 1, "low": 2}
    defs, rows, ft = {}, [], {}
    best_for_rule = {}
    for n, lst in links.items():
        for item in lst:
            cur = best_for_rule.get(item["rule_id"])
            if cur is None or rank[item["confidence"]] < rank[cur[1]]:
                best_for_rule[item["rule_id"]] = (n, item["confidence"])

    # which sections of the Act does each section refer to (and so which refer back to it)
    parsed, outs, why_out, valid = {}, {}, {}, set(order)
    for s in sections:
        paras, notes = to_paragraphs(strip_trailing_headings(s["text"]))
        parsed[s["section"]] = (paras, notes)
        why, cur = {}, ""
        if s["status"] == "in force":
            for p in paras:
                if p["k"] == "sub":
                    cur = p["l"].strip("()")
                pos = {}
                for num in para_refs(p["x"], pos):
                    if num in valid and num != s["section"] and num not in why:
                        why[num] = {"l": cur, "x": snippet(p["x"], pos[num])}
        why_out[s["section"]] = why
        outs[s["section"]] = sorted(why, key=sort_key)
    ins = {}
    for a, lst in outs.items():
        for b in lst:
            ins.setdefault(b, []).append(a)

    # sections that share a rule with this one (the rule cites both), apart from the direct rules themselves
    rule_secs = {}
    for n, lst in links.items():
        for item in lst:
            if item["confidence"] != "low":
                rule_secs.setdefault(item["rule_id"], set()).add(n)

    for i, s in enumerate(sections):
        n = s["section"]
        paras, notes = parsed[n]
        if n == "2":
            defs = extract_definitions(paras)
        linked = []
        for item in links.get(n, []):
            r = by_id.get(item["rule_id"])
            if not r or r["status"] != "in force":
                continue
            ps, w = [], ""
            if item["confidence"] != "low":
                for pi, p in enumerate(rule_paras.get(r["id"], [])):
                    pos = {}
                    refs = para_refs(p["x"], pos)
                    for sb in sorted(refs.get(n, ())):
                        ps.append([pi, sb])
                    if n in refs and not w:
                        w = snippet(p["x"], pos[n])
            linked.append({"set": set_pos[r["rule_set_id"]], "rule": r["rule"], "t": r["title"],
                           "conf": item["confidence"], "m": item["mentions"],
                           "sb": item.get("subs", []), "ps": ps, "w": w})
        direct = sum(1 for x in linked if x["conf"] != "low")
        co = {}
        for item in links.get(n, []):
            if item["confidence"] == "low" or item["rule_id"] not in by_id:
                continue
            for m_ in rule_secs.get(item["rule_id"], ()):
                if m_ not in (n, "2") and m_ not in outs[n] and n not in outs.get(m_, ()):
                    co.setdefault(m_, []).append(item["rule_id"])
        co_list = []
        for m_, rids in sorted(co.items(), key=lambda kv: (-len(kv[1]), sort_key(kv[0])))[:40]:
            co_list.append({"m": m_, "c": len(rids),
                            "r": [[set_pos[by_id[x]["rule_set_id"]], by_id[x]["rule"]] for x in rids[:3]]})
        prev_n = order[i - 1] if i else None
        next_n = order[i + 1] if i + 1 < len(order) else None
        dump(OUT / "data" / "s" / f"{n}.json", {
            "n": n, "t": s["title"], "c": s["chapter"], "ct": s["chapter_title"], "s": s["status"],
            "paras": paras, "notes": notes, "rules": linked,
            "out": outs[n], "in": sorted(ins.get(n, []), key=sort_key),
            "wo": why_out[n], "wi": {b: why_out[b][n] for b in ins.get(n, [])}, "co": co_list,
            "prev": [prev_n, title_of[prev_n]] if prev_n else None,
            "next": [next_n, title_of[next_n]] if next_n else None,
        })
        rows.append([n, s["title"], s["chapter"], s["chapter_title"], 0 if s["status"] == "in force" else 1, direct])
        words = set(re.findall(r"[a-z]{3,}", (s["title"] + " " + " ".join(p["x"] for p in paras)).lower())) - STOP
        for w in words:
            ft.setdefault(w, []).append(i)

    total = len(sections)
    ft = {w: v for w, v in ft.items() if len(v) < total * 0.6}

    rule_rows = []
    for r in rules_doc["rules"]:
        if r["status"] != "in force" or not r["title"] or r["id"] not in best_for_rule:
            continue
        rule_rows.append([set_pos[r["rule_set_id"]], r["rule"], r["title"], best_for_rule[r["id"]][0]])

    chapters = []
    for s in sections:
        if s["chapter"] and (not chapters or chapters[-1]["c"] != s["chapter"]):
            chapters.append({"c": s["chapter"], "t": s["chapter_title"], "first": s["section"], "last": s["section"]})
        if chapters and s["chapter"] == chapters[-1]["c"]:
            chapters[-1]["last"] = s["section"]

    dump(OUT / "data" / "index.json", {
        "act": ACT_NAME, "built": datetime.date.today().isoformat(),
        "sections": rows, "chapters": chapters, "sets": set_index, "rules": rule_rows,
    })
    dump(OUT / "data" / "fulltext.json", ft)
    dump(OUT / "data" / "defs.json", defs)

    size = sum(f.stat().st_size for f in OUT.rglob("*") if f.is_file())
    report = [
        "# Site report", "",
        f"- Sections written: {total}",
        f"- Rule sets written: {len(set_index)}",
        f"- Rules searchable by title: {len(rule_rows)}",
        f"- Defined terms found in section 2: {len(defs)}",
        f"- Search words indexed: {len(ft)}",
        f"- Total site size: {size / 1e6:.1f} MB",
    ]
    (DATA / "site_report.md").write_text("\n".join(report) + "\n", encoding="utf-8")
    print("\n".join(report))


if __name__ == "__main__":
    main()
