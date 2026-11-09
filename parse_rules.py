"""Step 3: split every Rules document into rules, find which sections of the Act each rule cites,
and build the section-to-rules map.

Reads  data/raw/rules/*.txt and data/sections.json
Writes data/rules.json, data/links.json, data/rules_report.md, data/rules_debug.md
This step never fails the build: problems are listed in the report instead.
"""
import bisect
import json
import re
from pathlib import Path

RULE_MAX = 400
D = r"[—–─‒−\-]"                                  # every kind of dash seen in the PDFs
PREFIX = r"(?:\[\d{1,3}\]|\d{1,3}\[|\[)*"            # footnote markers such as [13][ or 3[
START = r"^(?=[ \t]*" + PREFIX + r"[ \t]*(?P<num>\d{1,3}[A-Z]{0,2})\.[ \t]*"
NOTNUM = r"(?![ \t]*" + PREFIX + r"\d{1,3}[A-Z]{0,2}\.)"

# Heading with a dash: "6. Title.-", "6. Title.─", "6. Title:-", "13. Title- text", title may wrap over lines
DASH_RE = re.compile(
    START + r"(?P<nl>\n?)[ \t]*(?:\(\d{1,2}\)[ \t]*)?"
    r"(?P<title>[A-Z\[“‘][^\n]{1,200}?(?:\n(?:[ \t]*\n)*" + NOTNUM + r"[^\n]{1,120}?){0,10}?)"
    r"(?:(?P<t1>[ \t]*\.\]?[ \t]*" + D + r"+|[ \t]*" + D + r"+\.\]?(?=[ \t]*\n))"
    r"|(?P<t2>[ \t]*:[ \t]*" + D + r"+)"
    r"|(?P<t3>[ \t]*" + D + r"(?=[ \t]+\S)))"
    r"(?P<end>))",
    re.MULTILINE,
)
# Heading with no dash: "2. Definitions." alone on a line
PLAIN_RE = re.compile(
    START + r"(?P<title>[A-Z\[“‘][^\n]{1,100}?(?:\n" + NOTNUM + r"[^\n]{1,100}?){0,2}?)\.\]?[ \t]*\n(?P<end>))",
    re.MULTILINE,
)
# a rule that has no heading, or whose heading is followed by its text on the same line
BARE_RE = re.compile(r"^[ \t]*" + PREFIX + r"[ \t]*(?P<num>\d{1,3}[A-Z]{0,2})\.[ \t]+(?P<body>[A-Z(\[“‘][^\n]*)", re.MULTILINE)
OLD_TEXT = re.compile(r"stood as under|Prior to (?:omission|substitution)|for the (?:rule|rules|sub-rule|words|proviso|clause)", re.IGNORECASE)
# vocabulary of footnotes such as "Substituted by ... vide Notification No. G.S.R. 42(E) ... w.e.f."
FOOT_RE = re.compile(r"\bvide\b|\bNotification No\b|w\.e\.f|G\.S\.R|(?:Substituted|Inserted|Omitted|Amended) by\b")
STOP = re.compile(r"\b(?:shall|may|must|should|will|is|are|was|were|has|have|can|if|unless|where|whose|that|being)\b")
OMIT1 = re.compile(r"^[ \t]*" + PREFIX + r"[ \t]*(?P<num>\d{1,3}[A-Z]{0,2})\.[^\n]{0,80}\n?[ \t]*(?:\[\d{1,3}\])?\[\*{3}\]", re.MULTILINE)
OMIT_AFTER = re.compile(r"[ \t]*\n?[ \t]*(?:\[\d{1,3}\])?\[\*{3}\]")
OMIT2 = re.compile(r"\[Rule[ \t]+(?P<num>\d{1,3}[A-Z]{0,2})[ \t]+omitted\]", re.IGNORECASE)

FURNITURE_RE = re.compile(
    r"^[ \t]*(?:IBC Laws|®|www\.ibclaw\.in|Printed on [^\n]{0,60}|Page \d+ of \d+|Download Mobile App|"
    r"Subscribe Daily Email Newsletter|Last amendment on[^\n]{0,30}|No amendment|Amended upto date)[ \t]*$\n?",
    re.MULTILINE,
)
NAMELY_RE = re.compile(r"(namely\s*[:\-—–─]*)[ \t]*(?=(?:\d{0,2}\[)?1\.[ \t]*[A-Z\[(])")
NOISE_RE = re.compile(r"^[ \t]*\d{1,3}[ \t]*$", re.MULTILINE)
# after the last rule, forms and annexures begin
ENDMARK_RE = re.compile(r"^[ \t]*(?:FORM[ \t]+NO\b|ANNEXURE\b|APPENDIX\b|SCHEDULE\b|THE[ \t]+(?:FIRST|SECOND)[ \t]+SCHEDULE)", re.MULTILINE)
CALLED_RE = re.compile(
    r"may\s+be\s+(?:called|cited)\s+(?:as\s+)?(?:the\s+)?(.{5,200}?(?:Rules|Regulations|Scheme|Order|Directions)[\s,]*(?:19|20)\d{2})",
    re.DOTALL,
)
# "section 164", "sections 149, 150 and 152", "section 164(1)(a)"
SEC_RE = re.compile(
    r"\b(?i:sections?)\s+(\d{1,3}[A-Z]{0,2}(?:\s*\([^)]{1,8}\))*"
    r"(?:\s*(?:,|and|or|to|&)\s*\d{1,3}[A-Z]{0,2}(?:\s*\([^)]{1,8}\))*)*)"
)
OF_RE = re.compile(
    r"\bof\s+(?:the\s+)?((?:Companies\s+)?Act\b|[A-Z][A-Za-z’'()\- ,]{2,80}?\b(?:Act|Code|Regulations|Rules)\b)"
)


def sort_key(num: str):
    m = re.match(r"(\d+)([A-Z]*)", num)
    return int(m.group(1)), m.group(2)


def longest_increasing(cands):
    tails, tail_keys, prev = [], [], [None] * len(cands)
    for i, (key, _) in enumerate(cands):
        j = bisect.bisect_left(tail_keys, key)
        if j < len(tail_keys) and tail_keys[j] == key:
            continue
        prev[i] = tails[j - 1] if j else None
        if j == len(tails):
            tails.append(i)
            tail_keys.append(key)
        else:
            tails[j] = i
            tail_keys[j] = key
    chain, i = [], (tails[-1] if tails else None)
    while i is not None:
        chain.append(i)
        i = prev[i]
    return [cands[i][1] for i in reversed(chain)]


def trim_chain(chain):
    """Drop a short tail that sits far beyond the real rules (for example item 92 of a form)."""
    nums = [sort_key(m.group("num"))[0] for m in chain]
    for i in range(1, len(nums)):
        if nums[i] - nums[i - 1] > 20 and len(nums) - i < 6:
            return chain[:i]
    return chain


def title_ok(title, kind):
    t = title.strip()
    if FOOT_RE.search(t):
        return False
    if kind in ("t1", "t2"):
        return True
    if kind == "plain":
        return len(t) <= 170 and ". " not in t and not STOP.search(t)
    return len(t) <= 100 and "\n" not in t and ". " not in t and not STOP.search(t)


def clean(text):
    text = FURNITURE_RE.sub("", text.replace("\r", ""))
    return NAMELY_RE.sub(r"\1\n", text)


def derive_title(body_line):
    """For '12. Application for incorporation. An application for ...' return (title, characters used)."""
    sentence = body_line.split(". ")[0] if ". " in body_line else ""
    if sentence and len(sentence) <= 100 and not STOP.search(sentence) and not FOOT_RE.search(sentence):
        return sentence.strip(" ."), len(sentence) + 2
    return "", 0


def fill_gaps(text, items, omitted_nums):
    """Rules with no recognisable heading: look for a bare '10. Where the articles ...' line between neighbours."""
    out = []
    for idx, a in enumerate(items):
        out.append(a)
        if idx + 1 >= len(items):
            break
        b = items[idx + 1]
        na, nb = sort_key(a["num"])[0], sort_key(b["num"])[0]
        pos = a["body"]
        for n in range(na + 1, nb):
            ns = str(n)
            if ns in omitted_nums:
                continue
            for m in BARE_RE.finditer(text, pos, b["start"]):
                if m.group("num") != ns:
                    continue
                if OLD_TEXT.search(text[max(0, m.start() - 350):m.start()]) or FOOT_RE.search(m.group("body")):
                    continue
                title, used = derive_title(m.group("body"))
                out.append({"num": ns, "title": title, "start": m.start(), "body": m.start("body") + used, "untitled": not title})
                pos = m.end()
                break
    return out


def split_rules(text):
    text = clean(text)
    cands = {}
    for rx, plain in ((DASH_RE, False), (PLAIN_RE, True)):
        for m in rx.finditer(text):
            key = sort_key(m.group("num"))
            if not 1 <= key[0] <= RULE_MAX:
                continue
            kind = "plain" if plain else ("t1" if m.group("t1") else "t2" if m.group("t2") else "t3")
            if kind == "t3" and m.group("nl"):
                continue                      # "31.\nOmitted by ..." is a footnote, not a heading
            if OMIT_AFTER.match(text, m.end("end")):
                continue                      # "5. Penalty. [13][***]" is an omitted rule, not a live one
            if title_ok(m.group("title"), kind) and m.start() not in cands:
                cands[m.start()] = (key, m)
    kept = trim_chain(longest_increasing([cands[p] for p in sorted(cands)]))
    items = [{"num": m.group("num"), "title": re.sub(r"\s+", " ", m.group("title")).strip(),
              "start": m.start(), "body": m.end("end"), "untitled": False} for m in kept]

    top = max((sort_key(i["num"])[0] for i in items), default=0)
    found = {i["num"] for i in items}
    omitted = {}
    for rx in (OMIT1, OMIT2):
        for m in rx.finditer(text):
            n = m.group("num")
            if n not in found and n not in omitted and 1 <= sort_key(n)[0] <= top:
                omitted[n] = {"rule": n, "title": "Omitted", "status": "omitted", "text": "This rule has been omitted."}
    items = fill_gaps(text, items, set(omitted))

    rules = []
    for i, it in enumerate(items):
        end = items[i + 1]["start"] if i + 1 < len(items) else len(text)
        if i + 1 == len(items):
            cm = ENDMARK_RE.search(text, it["body"])
            if cm:
                end = cm.start()
        body = re.sub(r"\n{3,}", "\n\n", NOISE_RE.sub("", text[it["body"]:end])).strip()
        rules.append({"rule": it["num"], "title": it["title"], "untitled": it["untitled"], "status": "in force", "text": body})
    preamble = text[:items[0]["start"]] if items else ""
    return rules, list(omitted.values()), preamble, text


def rule_set_name(text, stem):
    m = CALLED_RE.search(text[:40000])
    if m:
        name = re.sub(r"\s+", " ", m.group(1))
        name = re.sub(r"\d{1,2}\[|\]", "", name)                       # footnote markers such as 1[...]
        name = re.sub(r"^(?:(?:as|of)\s+)?(?:the\s+)?", "", name, flags=re.IGNORECASE)
        return name.strip(" ,")
    return stem.replace("-", " ").title()


def classify(tail):
    """'high' if the citation is to the Companies Act, None if to another law, else 'medium'."""
    m = OF_RE.search(tail[:110])
    if not m:
        return "medium"
    return "high" if re.match(r"(?:Companies\s+)?Act\b", m.group(1)) else None


TOKEN_RE = re.compile(r"(\d{1,3}[A-Z]{0,2})((?:\s*\([^)]{1,8}\))*)")
SUBBEFORE_RE = re.compile(r"sub-?\s?sections?\s*\(\s*([0-9]{1,2}[A-Z]{0,2})\s*\)\s*(?:\([^)]{1,8}\)\s*)*(?:and|or|,|\s)*(?:\(\s*[0-9]{1,2}[A-Z]{0,2}\s*\)\s*)*(?:of|in|under)?\s*$", re.I)


def cited_sections(text, valid):
    """{section: {mentions, confidence, subs}} where subs lists the sub-sections cited, e.g. ["1", "2"]."""
    found = {}
    for m in SEC_RE.finditer(text):
        conf = classify(text[m.end():m.end() + 110])
        if conf is None:
            continue
        toks = TOKEN_RE.findall(m.group(1))
        before = SUBBEFORE_RE.search(text[max(0, m.start() - 40):m.start()])
        for num, parens in toks:
            if num not in valid:
                continue
            d = found.setdefault(num, {"mentions": 0, "confidence": "medium", "subs": []})
            d["mentions"] += 1
            if conf == "high":
                d["confidence"] = "high"
            p = re.match(r"\s*\(\s*([0-9]{1,2}[A-Z]{0,2})\s*\)", parens)
            sub = p.group(1) if p else (before.group(1) if before and len(toks) == 1 else None)
            if sub and sub not in d["subs"]:
                d["subs"].append(sub)
    return found


RANK = {"high": 0, "medium": 1, "low": 2}


def missing_numbers(rules):
    if not rules:
        return []
    found = {r["rule"] for r in rules}
    last = max(sort_key(r["rule"])[0] for r in rules)
    return [str(n) for n in range(1, last + 1) if str(n) not in found]


def context_for(text, n):
    pat = re.compile(r"(?m)^[ \t]*(?:\d{0,2}\[)?%s[A-Z]?[ \t]*[.\-]|(?<!\d)%s\.[ \t]*[A-Z]" % (n, n))
    for m in list(pat.finditer(text))[:2]:
        a, b = max(0, m.start() - 60), min(len(text), m.end() + 200)
        yield repr(text[a:b])


def main():
    Path("data").mkdir(exist_ok=True)
    sp = Path("data/sections.json")
    valid = {s["section"] for s in json.loads(sp.read_text(encoding="utf-8"))} if sp.exists() else set()
    files = sorted(Path("data/raw/rules").rglob("*.txt"))

    docs = []
    for f in files:
        text = f.read_text(encoding="utf-8").replace("\r", "")
        rules, omitted, preamble, text = split_rules(text)
        docs.append({"stem": f.stem, "file": f.name, "text": text, "rules": rules, "omitted": omitted,
                     "preamble": preamble, "name": rule_set_name(text, f.stem)})

    # The same Rules uploaded twice: keep the copy with most rules (latest file name on a tie)
    best = {}
    for d in docs:
        k = re.sub(r"\W+", " ", d["name"]).lower().strip()
        if k not in best or (len(d["rules"]), d["stem"]) > (len(best[k]["rules"]), best[k]["stem"]):
            best[k] = d
    kept_ids = {id(d) for d in best.values()}
    skipped = [d for d in docs if id(d) not in kept_ids]
    docs = [d for d in docs if id(d) in kept_ids]

    rule_sets, all_rules, rows, probs = [], [], [], []
    for d in docs:
        stem, name, rules = d["stem"], d["name"], d["rules"]
        enabling = sorted(cited_sections(d["preamble"][-3000:], valid), key=sort_key)
        fallback = [s for s in enabling if s != "469"]
        for r in rules:
            cites = cited_sections(r["title"] + "\n" + r["text"], valid)
            cl = [{"section": s, **cites[s]} for s in sorted(cites, key=sort_key)]
            if not cl:  # no explicit citation: use the section(s) the Rules were made under
                cl = [{"section": s, "mentions": 0, "confidence": "low"} for s in fallback]
            all_rules.append({"id": f"{stem}#{r['rule']}", "rule_set_id": stem, "rule_set": name,
                              "rule": r["rule"], "title": r["title"], "untitled": r["untitled"], "status": "in force", "text": r["text"], "cites": cl})
        for r in d["omitted"]:
            all_rules.append({"id": f"{stem}#{r['rule']}", "rule_set_id": stem, "rule_set": name, "rule": r["rule"],
                              "title": r["title"], "status": "omitted", "text": r["text"], "cites": []})
        mine = sorted((r for r in all_rules if r["rule_set_id"] == stem), key=lambda r: sort_key(r["rule"]))
        all_rules[:] = [r for r in all_rules if r["rule_set_id"] != stem] + mine
        rule_sets.append({"id": stem, "name": name, "file": d["file"], "rule_count": len(rules),
                          "omitted_count": len(d["omitted"]), "enabling_sections": enabling})
        miss = missing_numbers(rules + d["omitted"])
        first_last = f"{rules[0]['rule']} to {rules[-1]['rule']}" if rules else "-"
        explicit = sum(1 for r in all_rules if r["rule_set_id"] == stem and any(c["confidence"] != "low" for c in r["cites"]))
        om = len(d["omitted"])
        shown = ", ".join(miss[:6]) + (" ..." if len(miss) > 6 else "")
        rows.append(f"| {stem[:45]} | {name[:60]} | {len(rules)} (+{om} omitted) | {first_last} | {len(miss)}: {shown} | {explicit} |")
        if not rules or len(miss) >= 3:
            probs.append((len(miss) if rules else 999, d, miss))

    debug = ["# Rules debug", ""]
    for _, d, miss in sorted(probs, key=lambda x: -x[0])[:10]:
        if not d["rules"]:
            debug += [f"## {d['stem']}: NO rules found", "```", repr(d["text"][:900]), "```", ""]
            continue
        debug.append(f"## {d['stem']}: {len(d['rules'])} found, {len(miss)} numbers missing")
        for n in miss[:3]:
            debug.append(f"Rule {n}:")
            hits = list(context_for(d["text"], n))
            debug += (["```", *hits, "```"] if hits else ["(nothing that looks like a heading)"])
        debug.append("")

    by_section = {}
    for r in all_rules:
        for c in r["cites"]:
            by_section.setdefault(c["section"], []).append({
                "rule_id": r["id"], "rule_set": r["rule_set"], "rule": r["rule"], "title": r["title"],
                "mentions": c["mentions"], "confidence": c["confidence"], "subs": c.get("subs", []),
            })
    for lst in by_section.values():
        lst.sort(key=lambda x: (RANK[x["confidence"]], x["rule_set"], sort_key(x["rule"])))
    ordered = {k: by_section[k] for k in sorted(by_section, key=sort_key)}

    Path("data/rules.json").write_text(json.dumps({"rule_sets": rule_sets, "rules": all_rules}, ensure_ascii=False, indent=1), encoding="utf-8")
    Path("data/links.json").write_text(json.dumps({"by_section": ordered}, ensure_ascii=False, indent=1), encoding="utf-8")

    live = [r for r in all_rules if r["status"] == "in force"]
    no_link = sum(1 for r in live if not r["cites"])
    low_only = sum(1 for r in live if r["cites"] and all(c["confidence"] == "low" for c in r["cites"]))
    report = [
        "# Rules report", "",
        f"- Rules documents read: {len(files)} (duplicates skipped: {len(skipped)})",
        f"- Individual rules found: {len(live)} (plus {len(all_rules) - len(live)} omitted rules recorded)",
        f"- Sections of the Act that have at least one linked rule: {len(ordered)}",
        f"- Rules linked only through their rule set's enabling section (low confidence): {low_only}",
        f"- Rules with no link at all: {no_link}",
    ]
    if skipped:
        report += ["- Skipped duplicates: " + "; ".join(d["stem"][:50] for d in skipped)]
    report += ["", "| File | Rule set | Rules | Numbers | Missing numbers | Rules citing a section |",
               "|---|---|---|---|---|---|"] + rows
    Path("data/rules_report.md").write_text("\n".join(report) + "\n", encoding="utf-8")
    Path("data/rules_debug.md").write_text("\n".join(debug) + "\n", encoding="utf-8")
    print("\n".join(report))


if __name__ == "__main__":
    main()
