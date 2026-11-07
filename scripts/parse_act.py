"""Step 2: split the extracted Act text into sections and validate the result.

Reads  data/raw/act/*.txt
Writes data/sections.json and data/parse_report.md
Exits with an error if too many sections are missing, so a bad parse never gets published.
"""
import json
import re
import sys
from pathlib import Path

EXPECTED_MAX = 470      # Companies Act, 2013 has 470 sections (numbers 1 to 470, plus inserted ones like 10A)
MIN_FOUND = 430         # fail the build below this

# A section starts at the beginning of a line: number, full stop, title, then ".—" (or ".-")
SECTION_RE = re.compile(
    r"^[ \t]*(?P<num>\d{1,3}[A-Z]{0,2})\.[ \t]+(?P<title>[^\n]{3,200}?(?:\n[^\n]{1,120}?)?)\.[ \t]*[—–-]",
    re.MULTILINE,
)
CHAPTER_RE = re.compile(r"^[ \t]*CHAPTER[ \t]+([IVXLC]+[A-Z]?)\b[ \t]*\n?[ \t]*([^\n]*)", re.MULTILINE)
# page furniture such as "THE COMPANIES ACT, 2013" headers and lone page numbers
NOISE_RE = re.compile(r"^[ \t]*(?:\d{1,3}|THE COMPANIES ACT, 2013|Page \d+.*)[ \t]*$", re.MULTILINE)


def sort_key(num: str):
    m = re.match(r"(\d+)([A-Z]*)", num)
    return int(m.group(1)), m.group(2)


def parse(text: str):
    text = text.replace("\r", "")
    chapters = [(m.start(), m.group(1), m.group(2).strip()) for m in CHAPTER_RE.finditer(text)]
    hits = list(SECTION_RE.finditer(text))

    # Keep matches whose numbers increase through the document. This drops false hits
    # such as numbered list items inside a section ("2. the company shall ...").
    kept, last = [], (0, "")
    for m in hits:
        key = sort_key(m.group("num"))
        if key > last and key[0] <= last[0] + 12:
            kept.append(m)
            last = key

    sections = []
    for i, m in enumerate(kept):
        end = kept[i + 1].start() if i + 1 < len(kept) else len(text)
        body = text[m.end():end]
        body = NOISE_RE.sub("", body)
        body = re.sub(r"\n{3,}", "\n\n", body).strip()
        chapter = next(((c, t) for pos, c, t in reversed(chapters) if pos < m.start()), (None, None))
        sections.append({
            "section": m.group("num"),
            "title": re.sub(r"\s+", " ", m.group("title")).strip(),
            "chapter": chapter[0],
            "chapter_title": chapter[1],
            "text": body,
        })
    return sections


def main():
    files = sorted(Path("data/raw/act").rglob("*.txt"))
    if not files:
        sys.exit("No extracted Act text found in data/raw/act")
    text = "\n".join(f.read_text(encoding="utf-8") for f in files)
    sections = parse(text)

    found = {s["section"] for s in sections}
    missing = [str(n) for n in range(1, EXPECTED_MAX + 1) if str(n) not in found]
    dupes = len(sections) - len(found)
    empty = [s["section"] for s in sections if len(s["text"]) < 20]

    Path("data").mkdir(exist_ok=True)
    Path("data/sections.json").write_text(json.dumps(sections, ensure_ascii=False, indent=1), encoding="utf-8")

    report = [
        "# Act parsing report", "",
        f"- Sections found: {len(sections)} (expected about {EXPECTED_MAX})",
        f"- Duplicate section numbers: {dupes}",
        f"- Sections with almost no text: {', '.join(empty) or 'none'}",
        f"- Section numbers not found ({len(missing)}): {', '.join(missing[:60]) or 'none'}{' ...' if len(missing) > 60 else ''}",
        "",
        "Omitted sections (for example 'Omitted.' or 'Repealed.') may show up as missing or empty. That is expected.",
    ]
    Path("data/parse_report.md").write_text("\n".join(report) + "\n", encoding="utf-8")
    print("\n".join(report))

    if len(sections) < MIN_FOUND:
        sys.exit(f"FAILED: only {len(sections)} sections found, need at least {MIN_FOUND}. See data/parse_report.md")


if __name__ == "__main__":
    main()
