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

# A rule heading: optional footnote marker, number, full stop, title (may wrap over lines), then ".—"
RULE_RE = re.compile(
    r"^(?=[ \t]*(?:\d{0,2}\[)?(?P<num>\d{1,3}[A-Z]{0,2})\.[ \t]*\n?[ \t]*"
    r"(?P<title>[A-Z\[“‘][^\n]{1,200}?(?:\n(?![ \t]*\d{1,3}[A-Z]{0,2}\.)[^\n]{1,120}?){0,10})"
    r"\.\]?[ \t]*[—–-](?P<end>))",
    re.MULTILINE,
)
NOISE_RE = re.compile(r"^[ \t]*\d{1,3}[ \t]*$", re.MULTILINE)
# after the last rule, forms and annexures begin
ENDMARK_RE = re.compile(r"^[ \t]*(?:FORM[ \t]+NO\b|ANNEXURE\b|APPENDIX\b|SCHEDULE\b|THE[ \t]+(?:FIRST|SECOND)[ \t]+SCHEDULE)", re.MULTILINE)
CALLED_RE = re.compile(
    r"may\s+be\s+called\s+(?:the\s+)?(.{5,200}?(?:Rules|Regulations|Scheme|Order|Directions)[\s,]*(?:19|20)\d{2})",
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


def split_rules(text):
    cands = []
    for m in RULE_RE.finditer(text):
        key = sort_key(m.group("num"))
        if 1 <= key[0] <= RULE_MAX:
            cands.append((key, m))
    kept = longest_increasing(cands)
    rules = []
    for i, m in enumerate(kept):
        end = kept[i + 1].start() if i + 1 < len(kept) else len(text)
        if i + 1 == len(kept):
            cm = ENDMARK_RE.search(text, m.end("end"))
            if cm:
                end = cm.start()
        body = re.sub(r"\n{3,}", "\n\n", NOISE_RE.sub("", text[m.end("end"):end])).strip()
        rules.append({"rule": m.group("num"), "title": re.sub(r"\s+", " ", m.group("title")).strip(), "text": body})
    preamble = text[:kept[0].start()] if kept else ""
    return rules, preamble


def rule_set_name(text, stem):
    m = CALLED_RE.search(text[:40000])
    if m:
        return re.sub(r"\s+", " ", m.group(1)).strip(" ,")
    return stem.replace("-", " ").title()


def classify(tail):
    """'high' if the citation is to the Companies Act, None if to another law, else 'medium'."""
    m = OF_RE.search(tail[:110])
    if not m:
        return "medium"
    return "high" if re.match(r"(?:Companies\s+)?Act\b", m.group(1)) else None


def cited_sections(text, valid):
    found = {}
    for m in SEC_RE.finditer(text):
        conf = classify(text[m.end():m.end() + 110])
        if conf is None:
            continue
        plain = re.sub(r"\([^)]{1,8}\)", "", m.group(1))
        for num in re.findall(r"\d{1,3}[A-Z]{0,2}", plain):
            if num in valid:
                d = found.setdefault(num, {"mentions": 0, "confidence": "medium"})
                d["mentions"] += 1
                if conf == "high":
                    d["confidence"] = "high"
    return found


def main():
    Path("data").mkdir(exist_ok=True)
    sp = Path("data/sections.json")
    valid = {s["section"] for s in json.loads(sp.read_text(encoding="utf-8"))} if sp.exists() else set()
    files = sorted(Path("data/raw/rules").rglob("*.txt"))

    rule_sets, all_rules, debug, rows = [], [], ["# Rules debug", ""], []
    for f in files:
        text = f.read_text(encoding="utf-8").replace("\r", "")
        stem = f.stem
        rules, preamble = split_rules(text)
        name = rule_set_name(text, stem)
        enabling = sorted(cited_sections(preamble[-3000:], valid), key=sort_key)
        for r in rules:
            cites = cited_sections(r["title"] + "\n" + r["text"], valid)
            all_rules.append({
                "id": f"{stem}#{r['rule']}", "rule_set_id": stem, "rule_set": name,
                "rule": r["rule"], "title": r["title"], "text": r["text"],
                "cites": [{"section": s, **cites[s]} for s in sorted(cites, key=sort_key)],
            })
        rule_sets.append({"id": stem, "name": name, "file": f.name, "rule_count": len(rules), "enabling_sections": enabling})
        first_last = f"{rules[0]['rule']} to {rules[-1]['rule']}" if rules else "-"
        linked = sum(1 for r in all_rules if r["rule_set_id"] == stem and r["cites"])
        rows.append(f"| {stem[:45]} | {name[:60]} | {len(rules)} | {first_last} | {linked} |")
        if len(rules) < 3:
            debug += [f"## {stem}: only {len(rules)} rule(s) found", "```", repr(text[:900]), "```", ""]

    by_section = {}
    for r in all_rules:
        for c in r["cites"]:
            by_section.setdefault(c["section"], []).append({
                "rule_id": r["id"], "rule_set": r["rule_set"], "rule": r["rule"], "title": r["title"],
                "mentions": c["mentions"], "confidence": c["confidence"],
            })
    for lst in by_section.values():
        lst.sort(key=lambda x: (x["rule_set"], sort_key(x["rule"])))
    ordered = {k: by_section[k] for k in sorted(by_section, key=sort_key)}

    Path("data/rules.json").write_text(json.dumps({"rule_sets": rule_sets, "rules": all_rules}, ensure_ascii=False, indent=1), encoding="utf-8")
    Path("data/links.json").write_text(json.dumps({"by_section": ordered}, ensure_ascii=False, indent=1), encoding="utf-8")

    unlinked = sum(1 for r in all_rules if not r["cites"])
    report = [
        "# Rules report", "",
        f"- Rules documents read: {len(files)}",
        f"- Individual rules found: {len(all_rules)}",
        f"- Sections of the Act that have at least one linked rule: {len(ordered)}",
        f"- Rules that cite no section of the Act: {unlinked}",
        "",
        "| File | Rule set | Rules | Numbers | Rules citing a section |",
        "|---|---|---|---|---|",
    ] + rows
    Path("data/rules_report.md").write_text("\n".join(report) + "\n", encoding="utf-8")
    Path("data/rules_debug.md").write_text("\n".join(debug) + "\n", encoding="utf-8")
    print("\n".join(report))


if __name__ == "__main__":
    main()
