"""Step 2: split the extracted Act text into sections and validate the result.

Reads  data/raw/act/*.txt
Writes data/sections.json, data/parse_report.md and data/parse_debug.md
Omitted sections (for example 253 to 269) are kept as entries marked "omitted".
The build fails only if sections are missing that are neither found nor listed as omitted.
"""
import bisect
import json
import re
import sys
from pathlib import Path

EXPECTED_MAX = 470       # the Act has section numbers 1 to 470 (plus inserted ones like 10A)
MAX_UNEXPLAINED = 15     # fail the build if more sections than this are unaccounted for

# A heading: optional footnote marker like "3[", number, full stop, optional space, title, then ".—"
SECTION_RE = re.compile(
    r"^[ \t]*(?:\d{0,2}\[)?(?P<num>\d{1,3}[A-Z]{0,2})\.[ \t]*"
    r"(?P<title>[A-Z\[“‘][^\n]{2,200}?(?:\n[^\n]{1,120}?)?)\.\]?[ \t]*[—–-]",
    re.MULTILINE,
)
CHAPTER_RE = re.compile(r"^[ \t]*CHAPTER[ \t]+([IVXLC]+[A-Z]?)\b[ \t]*\n?[ \t]*([^\n]*)", re.MULTILINE)
NOISE_RE = re.compile(r"^[ \t]*(?:\d{1,3}|THE COMPANIES ACT, 2013|Page \d+.*)[ \t]*$", re.MULTILINE)
SCHEDULE_RE = re.compile(
    r"^[ \t]*(?:THE[ \t]+)?(?:FIRST|SECOND|THIRD|FOURTH|FIFTH|SIXTH|SEVENTH|EIGHTH|NINTH|TENTH)[ \t]+SCHEDULE\b"
    r"|^[ \t]*SCHEDULE[ \t]+[IVX]+\b",
    re.MULTILINE,
)
# Omitted sections appear twice: in the contents list ("253. [Omitted.].") and in a footnote
# ("253. [Determination of sickness.] Omitted by the Insolvency and Bankruptcy Code ...").
OMIT_TOC_RE = re.compile(r"^[ \t]*(?P<num>\d{1,3}[A-Z]{0,2})\.[ \t]*\[(?:Omitted|Repealed|Deleted)\b", re.MULTILINE)
OMIT_NOTE_RE = re.compile(
    r"^[ \t]*(?P<num>\d{1,3}[A-Z]{0,2})\.[ \t]*\[(?P<title>[^\]\n]{3,200})\][ \t]*"
    r"(?P<note>(?:Omitted|Repealed|Deleted)\b[^\n]*(?:\n(?![ \t]*(?:\d{1,3}[A-Z]{0,2}\.|\[|\n))[^\n]*)?)",
    re.MULTILINE,
)


def sort_key(num: str):
    m = re.match(r"(\d+)([A-Z]*)", num)
    return int(m.group(1)), m.group(2)


def longest_increasing(cands):
    """From (key, match) pairs in document order, keep the longest run whose keys keep increasing.
    This drops false hits (list items, footnotes) however far the real numbers jump."""
    tails, tail_keys, prev = [], [], [None] * len(cands)
    for i, (key, _) in enumerate(cands):
        j = bisect.bisect_left(tail_keys, key)
        if j < len(tail_keys) and tail_keys[j] == key:
            continue  # same number seen again later: keep the earlier one
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


def chapter_at(chapters, pos):
    return next(((c, t) for p, c, t in reversed(chapters) if p < pos), (None, None))


def parse(text: str):
    text = text.replace("\r", "")
    chapters = [(m.start(), m.group(1), re.sub(r"\s+", " ", m.group(2)).strip()) for m in CHAPTER_RE.finditer(text)]

    cands = []
    for m in SECTION_RE.finditer(text):
        key = sort_key(m.group("num"))
        if 1 <= key[0] <= EXPECTED_MAX:
            cands.append((key, m))
    kept = longest_increasing(cands)

    sections = []
    for i, m in enumerate(kept):
        end = kept[i + 1].start() if i + 1 < len(kept) else len(text)
        if i + 1 == len(kept):  # last section: stop where the Schedules begin
            sm = SCHEDULE_RE.search(text, m.end())
            if sm:
                end = sm.start()
        body = NOISE_RE.sub("", text[m.end():end])
        body = re.sub(r"\n{3,}", "\n\n", body).strip()
        chapter = chapter_at(chapters, m.start())
        sections.append({
            "section": m.group("num"),
            "title": re.sub(r"\s+", " ", m.group("title")).strip(),
            "status": "in force",
            "chapter": chapter[0],
            "chapter_title": chapter[1],
            "text": body,
        })

    found = {s["section"] for s in sections}
    omitted = {}
    for m in OMIT_TOC_RE.finditer(text):
        num = m.group("num")
        if num not in found and sort_key(num)[0] <= EXPECTED_MAX and num not in omitted:
            chapter = chapter_at(chapters, m.start())
            omitted[num] = {"section": num, "title": "Omitted", "status": "omitted",
                            "chapter": chapter[0], "chapter_title": chapter[1], "text": "This section has been omitted."}
    for m in OMIT_NOTE_RE.finditer(text):
        num = m.group("num")
        if num in omitted:
            omitted[num]["title"] = re.sub(r"\s+", " ", m.group("title")).strip()
            omitted[num]["text"] = re.sub(r"\s+", " ", m.group("note")).strip()[:400]

    everything = sections + list(omitted.values())
    everything.sort(key=lambda s: sort_key(s["section"]))
    return everything


def write_debug(text, sections, unexplained):
    """data/parse_debug.md: raw text around sections that could not be accounted for."""
    out = ["# Parse debug", "", f"- Act text length: {len(text)} characters"]
    in_force = [s for s in sections if s["status"] == "in force"]
    if in_force:
        out.append(f"- Last section found: {in_force[-1]['section']} ({in_force[-1]['title']})")
    if not unexplained:
        out.append("- Every section number from 1 to %d is either found or listed as omitted." % EXPECTED_MAX)
    for n in [x for x in unexplained if x.isdigit()][:8]:
        pat = re.compile(r"(?m)^[ \t]*(?:\d{0,2}\[)?%s[A-Z]?[ \t]*[.\-]|(?<!\d)%s\.[ \t]*[A-Z]" % (n, n))
        hits = list(pat.finditer(text))[:3]
        out += ["", f"## Section {n}: {len(hits)} possible heading(s) shown"]
        for m in hits:
            a, b = max(0, m.start() - 60), min(len(text), m.end() + 200)
            out += ["```", repr(text[a:b]), "```"]
        if not hits:
            out.append("(nothing that looks like a heading for this number)")
    Path("data/parse_debug.md").write_text("\n".join(out) + "\n", encoding="utf-8")


def main():
    files = sorted(Path("data/raw/act").rglob("*.txt"))
    if not files:
        sys.exit("No extracted Act text found in data/raw/act")
    text = "\n".join(f.read_text(encoding="utf-8") for f in files)
    sections = parse(text)

    present = {s["section"] for s in sections}
    in_force = [s for s in sections if s["status"] == "in force"]
    omitted = [s["section"] for s in sections if s["status"] == "omitted"]
    unexplained = [str(n) for n in range(1, EXPECTED_MAX + 1) if str(n) not in present]
    empty = [s["section"] for s in in_force if len(s["text"]) < 20]
    lettered = [s["section"] for s in in_force if not s["section"].isdigit()]

    Path("data").mkdir(exist_ok=True)
    Path("data/sections.json").write_text(json.dumps(sections, ensure_ascii=False, indent=1), encoding="utf-8")
    write_debug(text, sections, unexplained)

    def short(items, n=40):
        return ", ".join(items[:n]) + (" ..." if len(items) > n else "") if items else "none"

    report = [
        "# Act parsing report", "",
        f"- Sections in force found: {len(in_force)} (including {len(lettered)} inserted ones such as 10A)",
        f"- Omitted or repealed sections recorded: {len(omitted)}",
        f"- Section numbers unaccounted for ({len(unexplained)}): {short(unexplained)}",
        f"- Sections with almost no text: {short(empty)}",
        f"- Omitted section numbers: {short(omitted, 60)}",
    ]
    Path("data/parse_report.md").write_text("\n".join(report) + "\n", encoding="utf-8")
    print("\n".join(report))

    if len(unexplained) > MAX_UNEXPLAINED or len(in_force) < 300:
        sys.exit(f"FAILED: {len(unexplained)} sections unaccounted for, {len(in_force)} found. See data/parse_report.md and data/parse_debug.md")


if __name__ == "__main__":
    main()
