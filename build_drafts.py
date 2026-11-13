"""Step 6: build the Drafts library, a separate page that shows every draft in its original wording.

Reads   drafts/**   (.doc .docx .rtf .odt; folder name = category, file name = title)
Writes  site/drafts/index.html          the library page (open it at  <site>/drafts/)
        site/drafts/data/index.json     the list of drafts
        site/drafts/data/text.json      the words of every draft, for searching inside drafts
        site/drafts/data/d/<id>.json    one draft: original text, formatted for reading
        site/drafts/files/<id>.<ext>    a copy of each original file, for download
        site/drafts/drafts_index.csv    the index for Google Sheets (also saved as data/drafts_index.csv)
        data/drafts_report.md
Run it after build_site.py (which creates the site/ folder).

Word files are turned into clean HTML with LibreOffice (the workflow installs it). If LibreOffice is not
available, .docx files are still read directly, with plainer formatting.
"""
import csv
import html
import json
import os
import re
import shutil
import subprocess
import sys
import tempfile
import zipfile
from html.parser import HTMLParser
from pathlib import Path

SRC, OUT, DATA = Path("drafts"), Path("site"), Path("data")
EXTS = {".doc", ".docx", ".rtf", ".odt"}


def slug(s):
    return re.sub(r"[^a-z0-9]+", "-", s.lower()).strip("-") or "draft"


def natural(s):
    return [int(x) if x.isdigit() else x.lower() for x in re.split(r"(\d+)", s)]


def clean_title(stem):
    t = re.sub(r"^\d+[\s._-]+", "", stem.replace("_", " ")).strip()
    t = re.sub(r"\s+", " ", t) or stem
    return t.title() if t.isupper() and len(t) > 3 else t


# ---------- cleaning the HTML that LibreOffice writes ----------
BLOCK = {"p", "div", "h1", "h2", "h3", "h4", "h5", "h6", "li", "tr", "blockquote", "table", "ul", "ol"}
KEEP = {"p", "br", "b", "i", "u", "sub", "sup", "table", "tbody", "thead", "tr", "td", "th", "ul", "ol", "li", "h1", "h2", "h3", "h4", "h5", "h6", "blockquote", "center"}
RENAME = {"strong": "b", "em": "i", "div": "p"}
VOID = {"br"}
UNIT = re.compile(r"^-?\d+(?:\.\d+)?(?:in|cm|mm|pt|px|em)$")


def css_len(v):
    v = v.strip().lower()
    if not UNIT.match(v):
        return None
    n = float(re.match(r"-?\d+(?:\.\d+)?", v).group(0))
    u = v[len(re.match(r"-?\d+(?:\.\d+)?", v).group(0)):]
    px = {"in": 96, "cm": 37.8, "mm": 3.78, "pt": 1.333, "px": 1, "em": 16}[u] * n
    return max(-40, min(400, round(px)))


class Clean(HTMLParser):
    def __init__(self):
        super().__init__(convert_charrefs=True)
        self.out, self.lines, self.cur = [], [], []
        self.body = False
        self.skip = 0
        self.stack = []

    def _flush(self):
        t = re.sub(r"[ \t]+", " ", "".join(self.cur).replace("\xa0", " ")).strip()
        self.cur = []
        if t:
            self.lines.append(t)
        elif self.lines and self.lines[-1] != "":
            self.lines.append("")

    def handle_starttag(self, tag, attrs):
        if tag == "body":
            self.body = True
            return
        if tag in ("style", "script", "head", "title"):
            self.skip += 1
            return
        if not self.body or self.skip:
            return
        tag = RENAME.get(tag, tag)
        if tag not in KEEP:
            return
        a = dict(attrs)
        extra = ""
        style = []
        al = (a.get("align") or "").lower()
        if al in ("left", "center", "right", "justify") and tag in ("p", "td", "th", "h1", "h2", "h3", "h4", "h5", "h6", "tr", "table"):
            style.append("text-align:" + al)
        st = a.get("style") or ""
        if tag in ("p", "li", "h1", "h2", "h3", "h4", "h5", "h6", "blockquote"):
            for prop in ("margin-left", "text-indent"):
                m = re.search(r"(?<![\w-])" + prop + r"\s*:\s*([^;]+)", st)
                if m:
                    px = css_len(m.group(1))
                    if px:
                        style.append(f"{prop}:{px}px")
        if tag in ("td", "th"):
            for k in ("colspan", "rowspan"):
                if (a.get(k) or "").isdigit():
                    extra += f' {k}="{int(a[k])}"'
            m = re.search(r"border(?:-\w+)?\s*:\s*([^;]+)", st)
            if m and "none" not in m.group(1) and not re.search(r"\b0(?:px|in|pt)?\b", m.group(1).split()[0]):
                extra += ' class="bd"'
        if tag == "ol" and (a.get("start") or "").isdigit():
            extra += f' start="{int(a["start"])}"'
        if tag in BLOCK:
            self._flush()
        if tag == "br":
            self.out.append("<br>")
            self._flush()
            return
        self.out.append(f"<{tag}{extra}{' style=' + chr(34) + ';'.join(style) + chr(34) if style else ''}>")
        self.stack.append(tag)

    def handle_endtag(self, tag):
        if tag in ("style", "script", "head", "title"):
            self.skip = max(0, self.skip - 1)
            return
        if not self.body or self.skip:
            return
        tag = RENAME.get(tag, tag)
        if tag not in KEEP or tag in VOID:
            return
        if tag in self.stack:
            while self.stack:
                t = self.stack.pop()
                self.out.append(f"</{t}>")
                if t in BLOCK:
                    self._flush()
                if t == tag:
                    break
        if tag in ("td", "th"):
            self.cur.append(" \t ")

    def handle_data(self, data):
        if not self.body or self.skip:
            return
        d = re.sub(r"[ \t\r\n]+", " ", data)
        if d == " " and (not self.out or self.out[-1].endswith(">")):
            if not self.stack:
                return
        self.out.append(html.escape(d, quote=False))
        self.cur.append(d)

    def result(self):
        while self.stack:
            self.out.append(f"</{self.stack.pop()}>")
        self._flush()
        body = "".join(self.out)
        body = re.sub(r"(?:<p[^>]*>\s*(?:<br>)?\s*</p>\s*){2,}", "<p><br></p>", body)
        body = re.sub(r"<(b|i|u)>\s*</\1>", "", body)
        lines = []
        for ln in self.lines:
            if ln == "" and (not lines or lines[-1] == ""):
                continue
            lines.append(ln)
        while lines and lines[-1] == "":
            lines.pop()
        return body.strip(), "\n".join(lines)


def clean_html(raw):
    c = Clean()
    c.feed(raw)
    c.close()
    return c.result()


# ---------- fallback: read a .docx directly ----------
def docx_fallback(path):
    xml = zipfile.ZipFile(path).read("word/document.xml").decode("utf-8")
    parts, lines = [], []

    def para(px):
        runs = []
        for rm in re.finditer(r"<w:r(?:\s[^>]*[^/>])?>.*?</w:r>", px, re.S):
            r = rm.group(0)
            t = ""
            for tm in re.finditer(r"<w:t(?:\s[^>]*[^/>])?>(.*?)</w:t>|<w:tab\s*/>|<w:br\s*/>", r, re.S):
                g = tm.group(0)
                t += "\t" if g.startswith("<w:tab") else ("\n" if g.startswith("<w:br") else html.unescape(tm.group(1) or ""))
            if not t:
                continue
            e = html.escape(t, quote=False).replace("\t", "&emsp;").replace("\n", "<br>")
            if re.search(r"<w:b/>|<w:b w:val=\"(?:1|true)\"", r):
                e = f"<b>{e}</b>"
            if re.search(r"<w:i/>", r):
                e = f"<i>{e}</i>"
            if re.search(r"<w:u [^>]*w:val=\"(?!none)", r):
                e = f"<u>{e}</u>"
            runs.append((t, e))
        text = "".join(t for t, _ in runs)
        al = re.search(r'<w:jc w:val="(center|right|both|left|end)"', px)
        st = {"center": "center", "right": "right", "end": "right", "both": "justify"}.get(al.group(1)) if al else None
        ind = re.search(r'<w:ind [^>]*w:left="(\d+)"', px)
        css = ([f"text-align:{st}"] if st else []) + ([f"margin-left:{int(int(ind.group(1)) / 15)}px"] if ind and int(ind.group(1)) else [])
        return text, f"<p{' style=' + chr(34) + ';'.join(css) + chr(34) if css else ''}>{''.join(e for _, e in runs) or '<br>'}</p>"

    for m in re.finditer(r"<w:tbl>.*?</w:tbl>|<w:p(?:\s[^>]*)?>.*?</w:p>", xml, re.S):
        blk = m.group(0)
        if blk.startswith("<w:tbl>"):
            rows = []
            for rm in re.finditer(r"<w:tr(?:\s[^>]*)?>.*?</w:tr>", blk, re.S):
                cells = []
                for cm in re.finditer(r"<w:tc>.*?</w:tc>", rm.group(0), re.S):
                    ps = [para(p.group(0)) for p in re.finditer(r"<w:p(?:\s[^>]*)?>.*?</w:p>", cm.group(0), re.S)]
                    cells.append((" ".join(t for t, _ in ps).strip(), "".join(h for _, h in ps)))
                rows.append(cells)
                lines.append("\t".join(t for t, _ in cells))
            parts.append("<table>" + "".join("<tr>" + "".join(f'<td class="bd">{h}</td>' for _, h in r) + "</tr>" for r in rows) + "</table>")
        else:
            t, h = para(blk)
            parts.append(h)
            lines.append(t.strip())
    out = []
    for ln in lines:
        if ln == "" and (not out or out[-1] == ""):
            continue
        out.append(ln)
    return "".join(parts), "\n".join(out).strip()


# ---------- converting with LibreOffice ----------
def convert_all(files):
    """Returns {original path: html text} for the files LibreOffice could convert."""
    exe = shutil.which("soffice") or shutil.which("libreoffice")
    res = {}
    if not exe or not files:
        return res, bool(files and not exe)
    tmp = Path(tempfile.mkdtemp(prefix="drafts_"))
    (tmp / "in").mkdir()
    (tmp / "out").mkdir()
    names = {}
    for i, f in enumerate(files):
        n = f"d{i:05d}{f.suffix.lower()}"
        shutil.copyfile(f, tmp / "in" / n)
        names[f] = n
    items = list(names.values())
    for k in range(0, len(items), 20):
        chunk = items[k:k + 20]
        try:
            subprocess.run([exe, "--headless", "--convert-to", "html", "--outdir", str(tmp / "out")] + [str(tmp / "in" / x) for x in chunk],
                           stdout=subprocess.DEVNULL, stderr=subprocess.DEVNULL, timeout=900)
        except Exception:
            pass
    for f, n in names.items():
        h = tmp / "out" / (Path(n).stem + ".html")
        if h.exists():
            try:
                res[f] = h.read_text(encoding="utf-8", errors="replace")
            except Exception:
                pass
    return res, False


def main():
    OUT.mkdir(exist_ok=True)
    DATA.mkdir(exist_ok=True)
    files = sorted([p for p in SRC.rglob("*") if p.is_file() and p.suffix.lower() in EXTS and not p.name.startswith("~$")], key=lambda p: natural(str(p))) if SRC.exists() else []
    converted, no_office = convert_all(files)
    lib = OUT / "drafts"
    if lib.exists():
        shutil.rmtree(lib)
    (lib / "data" / "d").mkdir(parents=True)
    (lib / "files").mkdir()

    repo = os.environ.get("GITHUB_REPOSITORY", "")
    base = f"https://{repo.split('/')[0]}.github.io/{repo.split('/')[1]}/drafts/" if "/" in repo else ""

    index, texts, rows, report, used = [], {}, [], [], set()
    for p in files:
        rel = p.relative_to(SRC)
        how = "LibreOffice"
        try:
            if p in converted:
                body, text = clean_html(converted[p])
                if not text.strip():
                    raise ValueError("empty")
            elif p.suffix.lower() == ".docx":
                body, text = docx_fallback(p)
                how = "direct"
            else:
                raise ValueError("needs LibreOffice")
        except Exception as e:
            report.append(f"| {rel} | not read | 0 | {str(e)[:60]} |")
            continue
        category = " / ".join(clean_title(x) for x in rel.parent.parts) or "General"
        title = clean_title(rel.stem)
        did = slug(f"{category} {title}")
        while did in used:
            did += "-2"
        used.add(did)
        fname = f"{did}{p.suffix.lower()}"
        shutil.copyfile(p, lib / "files" / fname)
        nwords = len(text.split())
        desc = next((ln for ln in text.split("\n") if len(ln) >= 30 and ln.strip().lower() != title.lower()), "")
        desc = re.sub(r"\s+", " ", desc)[:170]
        (lib / "data" / "d" / f"{did}.json").write_text(json.dumps({"id": did, "title": title, "category": category, "html": body, "text": text,
            "words": nwords, "file": f"files/{fname}", "name": p.name}, ensure_ascii=False, separators=(",", ":")), encoding="utf-8")
        texts[did] = re.sub(r"\s+", " ", text)[:60000]
        index.append({"id": did, "title": title, "category": category, "desc": desc, "words": nwords})
        rows.append([category, title, nwords, p.name, (base + "#" + did) if base else f"#{did}", re.sub(r"\s+", " ", text)[:300], text[:45000]])
        report.append(f"| {rel} | {category} | {nwords} | {how} |")

    index.sort(key=lambda d: (natural(d["category"]), natural(d["title"])))
    rows.sort(key=lambda r: (natural(r[0]), natural(r[1])))
    (lib / "data" / "index.json").write_text(json.dumps({"drafts": index}, ensure_ascii=False, separators=(",", ":")), encoding="utf-8")
    (lib / "data" / "text.json").write_text(json.dumps(texts, ensure_ascii=False, separators=(",", ":")), encoding="utf-8")
    (lib / "index.html").write_text(PAGE, encoding="utf-8")
    for target in (lib / "drafts_index.csv", DATA / "drafts_index.csv"):
        with open(target, "w", encoding="utf-8-sig", newline="") as fh:
            w = csv.writer(fh)
            w.writerow(["Category", "Title", "Words", "Original file", "Open in library", "Preview", "Full text"])
            w.writerows(rows)

    out = ["# Drafts library report", "", f"- Drafts found: {len(files)}", f"- Drafts published: {len(index)}",
           f"- Categories: {len({d['category'] for d in index})}", ""]
    if no_office:
        out += ["LibreOffice is not installed here, so only .docx files could be read.", ""]
    out += ["| File | Category | Words | Read with |", "|---|---|---|---|"] + report
    if not SRC.exists():
        out += ["", "There is no drafts folder yet. Create one and upload the Word files there."]
    (DATA / "drafts_report.md").write_text("\n".join(out) + "\n", encoding="utf-8")
    print("\n".join(out[:6]))


PAGE = r'''<!doctype html>
<html lang="en" data-theme="auto">
<head>
<meta charset="utf-8">
<meta name="viewport" content="width=device-width,initial-scale=1">
<title>Drafts library</title>
<meta name="description" content="All drafts with their original wording. Search, read, copy or download the original Word file.">
<link rel="preconnect" href="https://fonts.googleapis.com">
<link rel="preconnect" href="https://fonts.gstatic.com" crossorigin>
<link href="https://fonts.googleapis.com/css2?family=Newsreader:ital,opsz,wght@0,6..72,400;0,6..72,500;0,6..72,600;1,6..72,400&family=Schibsted+Grotesk:wght@400;500;600&display=swap" rel="stylesheet">
<style>
:root{
  --paper:#eceff3;--page:#fafbfc;--ink:#18222f;--muted:#566377;--rule:#d3d9e2;
  --tape:#a51f1d;--link:#1d4f91;--mark:#fff0b3;--fs:1;
  --serif:"Newsreader",Georgia,"Times New Roman",serif;
  --sans:"Schibsted Grotesk",system-ui,-apple-system,"Segoe UI",Roboto,sans-serif;
}
:root[data-theme=light]{color-scheme:light}
:root[data-theme=sepia]{--paper:#e6dcc4;--page:#f4ecd8;--ink:#2d2316;--muted:#675740;--rule:#d2c5a5;--tape:#9b2218;--link:#7a3b10;--mark:#f7dc8a}
:root[data-theme=dark]{--paper:#0d1218;--page:#131a22;--ink:#e4e9f0;--muted:#93a0b3;--rule:#26313f;--tape:#ee6a60;--link:#8db7ee;--mark:#4a3f12;color-scheme:dark}
@media (prefers-color-scheme:dark){
  :root[data-theme=auto]{--paper:#0d1218;--page:#131a22;--ink:#e4e9f0;--muted:#93a0b3;--rule:#26313f;--tape:#ee6a60;--link:#8db7ee;--mark:#4a3f12;color-scheme:dark}
}
*{box-sizing:border-box}
html{-webkit-text-size-adjust:100%}
body{margin:0;background:var(--paper);color:var(--ink);font:400 1rem/1.5 var(--sans)}
a{color:var(--link)}
button{font:inherit;color:inherit}
:focus-visible{outline:2px solid var(--link);outline-offset:2px}
[hidden]{display:none!important}
mark{background:var(--mark);color:inherit;border-radius:2px;padding:0 .08em}

.bar{position:sticky;top:0;z-index:20;background:var(--paper);border-bottom:1px solid var(--rule)}
.bar-in{max-width:1320px;margin:auto;padding:.65rem 1.5rem;display:flex;align-items:center;gap:1.1rem}
.brand{font:600 1.2rem/1 var(--serif);color:var(--ink);text-decoration:none;white-space:nowrap;letter-spacing:-.01em}
.back{font:500 .9rem var(--sans);color:var(--muted);text-decoration:none;white-space:nowrap}
.back:hover{color:var(--ink)}
.find{flex:1;max-width:30rem;min-width:0}
.field{display:flex;align-items:center;background:var(--page);border:1px solid var(--rule);border-radius:3px}
.field:focus-within{border-color:var(--link);box-shadow:0 0 0 3px color-mix(in srgb,var(--link) 22%,transparent)}
.field .sym{padding:0 .1rem 0 .85rem;color:var(--tape);font:500 1.1rem/1 var(--serif)}
.field input{flex:1;min-width:0;border:0;background:transparent;color:inherit;font:inherit;padding:.6rem .8rem;outline:0}
.field input::placeholder{color:var(--muted)}
.tools{display:flex;align-items:center;gap:.35rem;margin-left:auto}
.tools button{background:none;border:1px solid transparent;border-radius:3px;cursor:pointer;min-width:2rem;height:2rem;padding:0 .3rem;font:500 .9rem var(--sans);color:var(--muted)}
.tools button:hover{border-color:var(--rule);color:var(--ink)}
.sw{position:relative}
.sw::after{content:"";position:absolute;inset:6px;border-radius:50%;border:1px solid var(--muted);background:var(--c)}
.sw[aria-pressed=true]{border-color:var(--ink)!important}
.sw[data-t=light]{--c:#fafbfc}.sw[data-t=sepia]{--c:#f1e5c8}.sw[data-t=dark]{--c:#131a22}

.wrap{max-width:1320px;margin:0 auto;padding:1.2rem 1.5rem 3rem;display:grid;grid-template-columns:minmax(17rem,24rem) minmax(0,1fr);gap:1.6rem;align-items:start}
.lp{position:sticky;top:4.3rem;max-height:calc(100vh - 5.2rem);overflow:auto;padding-right:.3rem}
.chips{display:flex;flex-wrap:wrap;gap:.4rem;margin:0 0 .9rem}
.chips button{border:1px solid var(--rule);border-radius:99px;background:var(--page);padding:.18rem .7rem;font:500 .84rem var(--sans);cursor:pointer;color:var(--ink)}
.chips button span{color:var(--muted);margin-left:.3rem;font-weight:400}
.chips button[aria-pressed=true]{background:var(--ink);color:var(--page);border-color:var(--ink)}
.chips button[aria-pressed=true] span{color:inherit;opacity:.8}
.cnt{margin:0 0 .6rem;font:400 .88rem var(--sans);color:var(--muted)}
.grp{margin:1.1rem 0 .3rem;font:600 1.05rem var(--serif);display:flex;gap:.5rem;align-items:baseline}
.grp small{font:400 .8rem var(--sans);color:var(--muted)}
.it{display:block;width:100%;text-align:left;background:none;border:0;border-left:3px solid transparent;padding:.42rem .6rem;cursor:pointer;color:var(--ink);text-decoration:none}
.it:hover{background:color-mix(in srgb,var(--link) 8%,transparent)}
.it[aria-current=true]{background:var(--page);border-left-color:var(--tape)}
.it b{display:block;font:600 .98rem/1.3 var(--serif)}
.it span{display:block;margin-top:.1rem;font:400 .8rem/1.35 var(--sans);color:var(--muted)}
.it .sn{margin-top:.2rem;font-size:.82rem}
.none{padding:1rem 0;color:var(--muted)}

.intro h1{margin:0;font:600 clamp(2.2rem,5vw,3.4rem)/1 var(--serif);letter-spacing:-.025em}
.intro p{margin:1rem 0 0;max-width:36rem;font:400 1.15rem/1.5 var(--serif);color:var(--muted)}
.cats{display:grid;grid-template-columns:repeat(auto-fill,minmax(14rem,1fr));gap:.7rem;margin-top:1.6rem}
.cat{display:block;text-align:left;background:var(--page);border:1px solid var(--rule);border-left:3px solid var(--tape);border-radius:3px;padding:.75rem .9rem;cursor:pointer}
.cat:hover{border-color:var(--ink);border-left-color:var(--tape)}
.cat b{display:block;font:600 1.05rem/1.25 var(--serif)}
.cat span{font:400 .85rem var(--sans);color:var(--muted)}
.sheet-tools{display:flex;flex-wrap:wrap;gap:.5rem;align-items:center;margin:0 0 .9rem}
.sheet-tools h1{flex:1 1 100%;margin:.2rem 0 0;font:600 clamp(1.5rem,3vw,2.1rem)/1.15 var(--serif);letter-spacing:-.015em}
.sheet-tools .meta{flex:1 1 100%;margin:0 0 .4rem;color:var(--muted);font:400 .9rem var(--sans)}
.btn{border:1px solid var(--rule);background:var(--page);border-radius:3px;padding:.45rem .85rem;font:500 .9rem var(--sans);cursor:pointer;color:var(--ink);text-decoration:none;display:inline-block}
.btn:hover{border-color:var(--ink)}
.btn.main{background:var(--ink);color:var(--page);border-color:var(--ink)}
.btn.main:hover{background:var(--tape);border-color:var(--tape)}
.mback{display:none}
.sheet{background:var(--page);border:1px solid var(--rule);box-shadow:0 14px 32px -22px rgba(8,16,30,.5);padding:clamp(1.1rem,4vw,2.6rem);font:400 calc(1.05rem*var(--fs))/1.7 var(--serif);overflow-wrap:break-word}
.sheet p{margin:0 0 .6rem}
.sheet table{border-collapse:collapse;max-width:100%;margin:.5rem 0}
.sheet td,.sheet th{padding:.25rem .5rem;vertical-align:top}
.sheet td.bd,.sheet th.bd{border:1px solid var(--muted)}
.sheet ul,.sheet ol{margin:.3rem 0 .6rem;padding-left:1.6rem}
.sheet h1,.sheet h2,.sheet h3,.sheet h4{margin:1rem 0 .5rem;font-family:var(--serif)}
.toast{position:fixed;left:50%;bottom:1.4rem;transform:translateX(-50%);background:var(--ink);color:var(--page);padding:.5rem .9rem;border-radius:3px;font:500 .9rem var(--sans);z-index:50}
@media (max-width:860px){
  .wrap{grid-template-columns:minmax(0,1fr);padding:1rem}
  .lp{position:static;max-height:none;padding:0}
  .bar-in{padding:.55rem 1rem;gap:.6rem;flex-wrap:wrap}
  .find{order:5;flex:1 1 100%;max-width:none}
  .back{display:none}
  body.reading .lp{display:none}
  body:not(.reading) .rp{display:none}
  .mback{display:inline-block}
}
@media print{.bar,.lp,.sheet-tools .btn,.mback,.toast{display:none!important}.wrap{display:block;padding:0}.sheet{border:0;box-shadow:none;padding:0}body{background:#fff}}
</style>
</head>
<body>
<header class="bar"><div class="bar-in">
  <a class="brand" href="./">Drafts library</a>
  <a class="back" href="../">&larr; Companies Act reader</a>
  <div class="find"><div class="field"><span class="sym" aria-hidden="true">&sect;</span>
    <input id="q" type="search" placeholder="Search drafts by name or any words inside" autocomplete="off" spellcheck="false" aria-label="Search drafts"></div></div>
  <div class="tools" role="group" aria-label="Reading settings">
    <button data-fs="-1" aria-label="Smaller text" title="Smaller text">A&minus;</button>
    <button data-fs="1" aria-label="Larger text" title="Larger text">A+</button>
    <button class="sw" data-t="light" aria-label="Light page" title="Light page" aria-pressed="false"></button>
    <button class="sw" data-t="sepia" aria-label="Sepia page" title="Sepia page" aria-pressed="false"></button>
    <button class="sw" data-t="dark" aria-label="Dark page" title="Dark page" aria-pressed="false"></button>
  </div>
</div></header>
<main class="wrap">
  <nav class="lp" aria-label="Drafts"><div class="chips" id="chips"></div><p class="cnt" id="cnt"></p><div id="list"></div></nav>
  <section class="rp" id="rp"></section>
</main>
<script>
(() => {
"use strict";
const $ = (s, r = document) => r.querySelector(s);
const $$ = (s, r = document) => [...r.querySelectorAll(s)];
const esc = s => String(s).replace(/[&<>"]/g, c => ({ "&": "&amp;", "<": "&lt;", ">": "&gt;", '"': "&quot;" }[c]));
const store = { get(k) { try { return localStorage.getItem(k); } catch (e) { return null; } }, set(k, v) { try { localStorage.setItem(k, v); } catch (e) {} } };
const S = { idx: [], text: null, low: null, cat: "", q: "", cur: null, loading: null };

function applyPrefs() {
  const t = store.get("theme") || "auto";
  document.documentElement.dataset.theme = t;
  $$(".sw").forEach(b => b.setAttribute("aria-pressed", String(b.dataset.t === t)));
  document.documentElement.style.setProperty("--fs", parseFloat(store.get("fs") || "1"));
}
document.addEventListener("click", e => {
  const sw = e.target.closest(".sw"), fb = e.target.closest("[data-fs]");
  if (sw) { store.set("theme", sw.dataset.t); applyPrefs(); }
  if (fb) { const cur = parseFloat(store.get("fs") || "1"); store.set("fs", String(Math.min(1.45, Math.max(.85, +(cur + fb.dataset.fs * .08).toFixed(2))))); applyPrefs(); }
});
const toast = m => { const t = document.createElement("div"); t.className = "toast"; t.textContent = m; document.body.appendChild(t); setTimeout(() => t.remove(), 2200); };
const getJSON = u => fetch(u).then(r => { if (!r.ok) throw new Error(u); return r.json(); });
const words = q => q.toLowerCase().split(/\s+/).filter(Boolean);

function loadText() {
  if (S.text) return Promise.resolve();
  return S.loading || (S.loading = getJSON("data/text.json").then(t => { S.text = t; S.low = {}; Object.keys(t).forEach(k => { S.low[k] = t[k].toLowerCase(); }); }).catch(() => { S.text = {}; S.low = {}; }));
}
function cats() { const m = new Map(); S.idx.forEach(d => m.set(d.category, (m.get(d.category) || 0) + 1)); return [...m]; }

function matches() {
  const w = words(S.q);
  let list = S.idx.filter(d => !S.cat || d.category === S.cat);
  if (!w.length) return list.map(d => ({ d }));
  const out = [];
  for (const d of list) {
    const head = (d.title + " " + d.category).toLowerCase(), body = S.low ? (S.low[d.id] || "") : "";
    if (!w.every(x => head.includes(x) || body.includes(x))) continue;
    let score = 0; w.forEach(x => { if (d.title.toLowerCase().includes(x)) score += 5; else if (head.includes(x)) score += 3; else score += 1; });
    let sn = "";
    const x = w.find(t => !head.includes(t)) || w[0], i = body.indexOf(x);
    if (i >= 0 && S.text) { const t = S.text[d.id]; sn = (i > 40 ? "…" : "") + t.slice(Math.max(0, i - 40), i + 90).replace(/\s+/g, " ") + "…"; }
    out.push({ d, score, sn });
  }
  return out.sort((a, b) => b.score - a.score || a.d.title.localeCompare(b.d.title));
}
function hl(text, w) {
  if (!w.length) return esc(text);
  const re = new RegExp("(" + w.map(x => x.replace(/[.*+?^${}()|[\]\\]/g, "\\$&")).join("|") + ")", "ig");
  return esc(text).replace(re, "<mark>$1</mark>");
}
function drawChips() {
  const all = S.idx.length;
  $("#chips").innerHTML = `<button data-cat="" aria-pressed="${!S.cat}">All<span>${all}</span></button>` + cats().map(([c, n]) => `<button data-cat="${esc(c)}" aria-pressed="${S.cat === c}">${esc(c)}<span>${n}</span></button>`).join("");
}
function drawList() {
  const res = matches(), w = words(S.q);
  $("#cnt").textContent = S.q ? `${res.length} ${res.length === 1 ? "draft matches" : "drafts match"} “${S.q}”` : `${res.length} ${res.length === 1 ? "draft" : "drafts"}${S.cat ? " in " + S.cat : ""}`;
  if (!res.length) { $("#list").innerHTML = `<p class="none">Nothing matches. Try fewer or shorter words.</p>`; return; }
  const item = r => `<a class="it" href="#${esc(r.d.id)}"${S.cur && S.cur.id === r.d.id ? ' aria-current="true"' : ""}><b>${hl(r.d.title, w)}</b><span>${esc(r.d.category)} &middot; ${r.d.words} words</span>${r.sn ? `<span class="sn">${hl(r.sn, w)}</span>` : ""}</a>`;
  if (S.q) { $("#list").innerHTML = res.map(item).join(""); return; }
  const g = new Map(); res.forEach(r => { if (!g.has(r.d.category)) g.set(r.d.category, []); g.get(r.d.category).push(r); });
  $("#list").innerHTML = [...g].map(([c, a]) => `${g.size > 1 ? `<h3 class="grp">${esc(c)}<small>${a.length}</small></h3>` : ""}${a.map(item).join("")}`).join("");
}
function intro() {
  S.cur = null; document.body.classList.remove("reading"); document.title = "Drafts library";
  const c = cats();
  $("#rp").innerHTML = `<div class="intro"><h1>Drafts library</h1><p>${S.idx.length} drafts in their original wording. Search by name or by any words inside a draft, open one to read it, then copy the text or download the original Word file.</p>
    ${c.length > 1 ? `<div class="cats">${c.map(([n, k]) => `<button class="cat" data-cat="${esc(n)}"><b>${esc(n)}</b><span>${k} ${k === 1 ? "draft" : "drafts"}</span></button>`).join("")}</div>` : ""}</div>`;
}
async function openDraft(id) {
  const meta = S.idx.find(d => d.id === id); if (!meta) return intro();
  S.cur = meta; document.body.classList.add("reading"); document.title = meta.title + " | Drafts library";
  $("#rp").innerHTML = `<p class="none">Opening…</p>`; drawList();
  let d; try { d = await getJSON(`data/d/${encodeURIComponent(id)}.json`); } catch (e) { $("#rp").innerHTML = `<p class="none">This draft could not be loaded.</p>`; return; }
  if (!S.cur || S.cur.id !== id) return;
  S.doc = d;
  $("#rp").innerHTML = `<div class="sheet-tools"><button class="btn mback" id="mb">&larr; All drafts</button><h1>${esc(d.title)}</h1><p class="meta">${esc(d.category)} &middot; ${d.words} words &middot; original file: ${esc(d.name)}</p>
    <button class="btn main" id="cp">Copy text</button><a class="btn" href="${esc(d.file)}" download="${esc(d.name)}">Download original</a><button class="btn" id="pr">Print</button><button class="btn" id="ln">Copy link</button></div>
    <article class="sheet" id="sheet">${d.html}</article>`;
  const w = words(S.q); if (w.length) markIn($("#sheet"), w);
  window.scrollTo(0, 0);
}
function markIn(root, w) {
  const re = new RegExp("(" + w.map(x => x.replace(/[.*+?^${}()|[\]\\]/g, "\\$&")).join("|") + ")", "ig");
  const tw = document.createTreeWalker(root, NodeFilter.SHOW_TEXT), nodes = []; let n;
  while ((n = tw.nextNode())) if (re.test(n.nodeValue)) nodes.push(n);
  nodes.slice(0, 400).forEach(t => { const f = document.createDocumentFragment(); t.nodeValue.split(re).forEach((p, i) => { if (i % 2) { const m = document.createElement("mark"); m.textContent = p; f.appendChild(m); } else f.appendChild(document.createTextNode(p)); }); t.parentNode.replaceChild(f, t); });
  const first = $("#sheet mark"); if (first) first.scrollIntoView({ block: "center" });
}
async function copyText(t) {
  try { await navigator.clipboard.writeText(t); return true; }
  catch (e) { const a = document.createElement("textarea"); a.value = t; document.body.appendChild(a); a.select(); let ok = false; try { ok = document.execCommand("copy"); } catch (x) {} a.remove(); return ok; }
}
document.addEventListener("click", async e => {
  const c = e.target.closest("[data-cat]");
  if (c && !c.closest(".sheet")) { S.cat = c.dataset.cat; drawChips(); drawList(); if (location.hash) history.replaceState(null, "", location.pathname); intro(); drawList(); window.scrollTo(0, 0); return; }
  if (e.target.closest("#mb")) { history.replaceState(null, "", location.pathname + location.search); intro(); drawList(); return; }
  if (e.target.closest("#cp")) { const t = (S.doc && S.doc.text) || ""; toast(await copyText(t) ? "Text copied" : "Could not copy. Select the text and copy it."); }
  if (e.target.closest("#pr")) window.print();
  if (e.target.closest("#ln")) toast(await copyText(location.href) ? "Link copied" : "Could not copy the link");
});
let timer;
$("#q").addEventListener("input", e => {
  S.q = e.target.value.trim(); loadText().then(() => { if (S.q === e.target.value.trim()) drawList(); });
  clearTimeout(timer); timer = setTimeout(drawList, 120);
});
document.addEventListener("keydown", e => {
  if (e.metaKey || e.ctrlKey || e.altKey) return;
  if (e.key === "/" && !/input|textarea/i.test(e.target.tagName)) { e.preventDefault(); $("#q").focus(); }
  else if (e.key === "Escape" && document.activeElement === $("#q")) { $("#q").value = ""; S.q = ""; drawList(); $("#q").blur(); }
});
function route() {
  const id = decodeURIComponent(location.hash.replace(/^#/, ""));
  if (id) openDraft(id); else { intro(); drawList(); }
}
window.addEventListener("hashchange", route);
(async function init() {
  applyPrefs();
  try { S.idx = (await getJSON("data/index.json")).drafts; }
  catch (e) { $("#rp").innerHTML = `<div class="intro"><h1>Drafts library</h1><p>No drafts are published yet. Upload Word files to the drafts folder of the repository and the library appears after the next build.</p></div>`; return; }
  drawChips(); drawList(); route();
})();
})();
</script>
</body>
</html>
'''

if __name__ == "__main__":
    main()
