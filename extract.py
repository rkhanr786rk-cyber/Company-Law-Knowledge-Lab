"""Step 1: turn every PDF in the repo into plain text under data/raw/act and data/raw/rules.

A PDF whose name starts with "Companies Act" is treated as the Act; every other PDF is a Rules document.
"""
from pathlib import Path

import fitz  # PyMuPDF

OUT = Path("data/raw")
SKIP = {"data", ".git", ".github", "node_modules"}


def category(pdf: Path) -> str:
    folders = [p.lower() for p in pdf.parts[:-1]]
    if "act" in folders:
        return "act"
    if "rules" in folders:
        return "rules"
    return "act" if pdf.name.lower().startswith("companies act") else "rules"


pdfs = sorted(
    p for p in Path(".").rglob("*")
    if p.is_file() and p.suffix.lower() == ".pdf" and not (set(p.parts) & SKIP)
)
if not pdfs:
    raise SystemExit("No PDF files found in the repo")

lines = ["# PDF extraction report", ""]
acts = 0
for pdf in pdfs:
    cat = category(pdf)
    acts += cat == "act"
    doc = fitz.open(pdf)
    pages = [page.get_text("text") for page in doc]
    chars = sum(len(t.strip()) for t in pages)
    out = OUT / cat / (pdf.stem + ".txt")
    out.parent.mkdir(parents=True, exist_ok=True)
    out.write_text("\n".join(pages), encoding="utf-8")
    per_page = chars / max(len(pages), 1)
    flag = "OK" if per_page > 200 else "WARNING: very little text, probably a scanned PDF (needs OCR)"
    lines.append(f"- [{cat}] {pdf}: {len(pages)} pages, {chars} characters. {flag}")

Path("data").mkdir(exist_ok=True)
Path("data/extract_report.md").write_text("\n".join(lines) + "\n", encoding="utf-8")
print("\n".join(lines))

if acts == 0:
    raise SystemExit("No Act PDF found. The Act file name must start with 'Companies Act'.")
