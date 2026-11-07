"""Step 1: turn every PDF in /act and /rules into plain text files under data/raw/."""
from pathlib import Path

import fitz  # PyMuPDF

OUT = Path("data/raw")
lines = ["# PDF extraction report", ""]

for folder in ("act", "rules"):
    for pdf in sorted(Path(folder).rglob("*.pdf")):
        doc = fitz.open(pdf)
        pages = [page.get_text("text") for page in doc]
        chars = sum(len(t.strip()) for t in pages)
        out = OUT / pdf.with_suffix(".txt")
        out.parent.mkdir(parents=True, exist_ok=True)
        out.write_text("\n".join(pages), encoding="utf-8")
        per_page = chars / max(len(pages), 1)
        flag = "OK" if per_page > 200 else "WARNING: very little text, probably a scanned PDF (needs OCR)"
        lines.append(f"- {pdf}: {len(pages)} pages, {chars} characters. {flag}")

if len(lines) == 2:
    raise SystemExit("No PDFs found in /act or /rules")

Path("data").mkdir(exist_ok=True)
Path("data/extract_report.md").write_text("\n".join(lines) + "\n", encoding="utf-8")
print("\n".join(lines))
