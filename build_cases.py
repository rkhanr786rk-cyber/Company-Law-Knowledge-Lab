#!/usr/bin/env python3
"""Builds the Case laws pages from the company-law case book kept in the repository
(any .pdf or .txt whose name contains "company_law", "company law" or "lb-303" or "case").
The list of cases, their topics and the related sections of the Companies Act, 2013 are written below
from the course syllabus; the judgments themselves are read from the book.
Writes site/data/cl/ and adds to each section page a list of the cases that relate to it."""
import json
import re
import sys
from pathlib import Path

SITE = Path("site/data")

TOPICS = [
    (1, "Nature and kinds of companies", "Corporate personality, lifting the veil, and the kinds of companies.", ["2", "9"]),
    (2, "Promotion and formation of a company", "Promoters and their duties; formation and the effect of registration.", ["3", "7", "9"]),
    (3, "Constitutional documents", "Memorandum and Articles, ultra vires, constructive notice and indoor management.", ["4", "5", "10", "13", "14"]),
    (4, "Capital market instruments", "Prospectus, share capital and debentures.", ["23", "26", "43", "44"]),
    (5, "Board of Directors", "Appointment, duties and liabilities of directors; corporate governance.", ["149", "152", "166", "153"]),
    (6, "General meetings", "Kinds of meetings, conditions of a valid meeting and how they are called.", ["96", "100", "101"]),
    (7, "Prevention of oppression and mismanagement", "Protection of minority shareholders; the rule in Foss v. Harbottle and its exceptions.", ["241", "242", "244", "245"]),
    (8, "Winding up of companies", "Winding up by the Tribunal.", ["270", "271"]),
    (9, "Adjudicatory bodies", "The NCLT and NCLAT: constitution, powers and judicial review.", ["407", "408", "410"]),
    (10, "Contemporary developments", "CSR, compromises and amalgamations, insolvency, ESG.", ["135", "230", "232"]),
]

# num, text key, name, citation, court, topic, related sections of the 2013 Act
CASES = [
    (1, "Salomon v. Salomon", "Salomon v. Salomon & Co. Ltd.", "(1897) AC 22", "House of Lords", 1, ["2", "9"]),
    (2, "Bennett Coleman", "Bennett Coleman & Co. v. Union of India", "(1972) 2 SCC 788; AIR 1973 SC 106", "Supreme Court of India", 1, ["2"]),
    (3, "Daimler Co", "Daimler Co. Ltd. v. Continental Tyre and Rubber Co. (Great Britain) Ltd.", "[1916] 2 AC 307", "House of Lords", 1, ["2"]),
    (4, "Lee v. Lee", "Lee v. Lee's Air Farming Ltd.", "[1960] 3 All ER 420", "Privy Council", 1, ["2", "9"]),
    (5, "In re Sir Dinshaw", "In re Sir Dinshaw Maneckjee Petit", "AIR 1927 Bom 371", "Bombay High Court", 1, ["9"]),
    (6, "Workmen v. Associated Rubber", "Workmen v. Associated Rubber Industry Ltd.", "(1985) 4 SCC 114", "Supreme Court of India", 1, ["2", "9"]),
    (7, "Gilford Motor", "Gilford Motor Co. Ltd. v. Horne", "[1933] Ch 935", "Court of Appeal", 1, ["2", "9"]),
    (8, "Subhra Mukherjee", "Subhra Mukherjee v. Bharat Coking Coal Ltd.", "(2000) 3 SCC 312", "Supreme Court of India", 1, ["9"]),
    (9, "Erlanger v. New Sombrero", "Erlanger v. New Sombrero Phosphate Co.", "(1878) 3 App Cas 1218", "House of Lords", 2, ["2", "26", "34", "35"]),
    (10, "Ashbury Railway", "Ashbury Railway Carriage and Iron Co. Ltd. v. Riche", "(1875) LR 7 HL 653", "House of Lords", 3, ["4", "13"]),
    (11, "Cotman v. Brougham", "Cotman v. Brougham", "[1918] AC 514", "House of Lords", 3, ["4", "13"]),
    (12, "In re Jon Beauforte", "In re (Jon) Beauforte (London) Ltd.", "[1953] Ch 131", "Chancery Division", 3, ["4", "13"]),
    (13, "Bell Houses", "Bell Houses Ltd. v. City Wall Properties Ltd.", "[1966] 2 All ER 674", "Court of Appeal", 3, ["4", "13"]),
    (14, "Introductions, Ltd. v. National", "Re Introductions Ltd.; Introductions Ltd. v. National Provincial Bank Ltd.", "[1969] 1 All ER 887", "Court of Appeal", 3, ["4", "13"]),
    (15, "The Royal British Bank", "Royal British Bank v. Turquand", "(1856) 119 ER 886", "Exchequer Chamber", 3, ["5", "10"]),
    (16, "Freeman and Lockyer", "Freeman & Lockyer v. Buckhurst Park Properties (Mangal) Ltd.", "[1964] 1 All ER 630", "Court of Appeal", 3, ["5", "10"]),
    (17, "Kotla Venkataswamy", "Kotla Venkataswamy v. Chinta Ramamurthy", "AIR 1934 Mad 579", "Madras High Court", 3, ["5", "10"]),
    (18, "Dehradun-Mussoorie", "Dehradun-Mussoorie Electric Tramway Co. v. Jagmandar Das", "AIR 1932 All 141", "Allahabad High Court", 3, ["5", "10"]),
    (19, "Percival v. Wright", "Percival v. Wright", "[1902] 2 Ch 421", "Chancery Division", 5, ["166"]),
    (20, "Burland v. Earle", "Burland v. Earle", "[1902] AC 83", "Privy Council", 5, ["166", "245"]),
    (22, "Regal (Hastings)", "Regal (Hastings) Ltd. v. Gulliver", "[1967] 2 AC 134", "House of Lords", 5, ["166"]),
    (23, "Industrial Development Consultants", "Industrial Development Consultants Ltd. v. Cooley", "[1972] 1 WLR 443", "Chancery Division", 5, ["166"]),
    (24, "Foss v. Harbottle", "Foss v. Harbottle", "(1843) 67 ER 189; 2 Hare 461", "Chancery Court", 7, ["241", "242", "245"]),
    (25, "Shanti Prasad Jain", "Shanti Prasad Jain v. Kalinga Tubes Ltd.", "AIR 1965 SC 1535", "Supreme Court of India", 7, ["241", "242"]),
    (26, "Rajahmundry Electric", "Rajahmundry Electric Supply Corporation Ltd. v. A. Nageshwara Rao", "AIR 1956 SC 213", "Supreme Court of India", 7, ["241", "242"]),
    (27, "Bharat Insurance", "Bharat Insurance Co. Ltd. v. Kanhaya Lal Gauba", "AIR 1935 Lah 792", "Lahore High Court", 7, ["241", "242"]),
    (28, "Radharamanan", "M.S.D.C. Radharamanan v. M.S.D. Chandrasekara Raja", "(2008) 6 SCC 750; AIR 2008 SC 1738", "Supreme Court of India", 7, ["241", "242", "244"]),
    (29, "Tata Consultancy Services Limited v. Cyrus", "Tata Consultancy Services Ltd. v. Cyrus Investments Pvt. Ltd.", "(2021) 9 SCC 449", "Supreme Court of India", 7, ["241", "242", "244"]),
    (30, "In re German Date Coffee", "In re German Date Coffee Co.", "(1882) 20 Ch D 169", "Court of Appeal", 8, ["271"]),
    (31, "Seth Mohan Lal", "Seth Mohan Lal v. Grain Chambers Ltd.", "AIR 1968 SC 772", "Supreme Court of India", 8, ["271"]),
    (32, "Madras Bar Association", "Madras Bar Association v. Union of India", "Writ Petition (C) No. 1072 of 2013, decided 14 May 2015", "Supreme Court of India", 9, ["408", "409", "410", "411", "412"]),
    (33, "Embassy Property", "Embassy Property Developments Pvt. Ltd. v. State of Karnataka", "MANU/SC/1661/2019", "Supreme Court of India", 9, ["408"]),
]

# named in the syllabus as suggested reading; the book does not carry their text
SUGGESTED = [
    (1, "State Trading Corporation v. CTO", "AIR 1963 SC 811"), (1, "TELCO v. State of Bihar", "AIR 1965 SC 40"),
    (1, "R.C. Cooper v. Union of India", "(1970) 3 SCR 530"), (1, "CIT v. Meenakshi Mills Ltd.", "AIR 1967 SC 819"),
    (1, "Kapila Hingorani v. State of Bihar", "(2003) 6 SCC 1"),
    (3, "Dr. A. Lakshmanaswami Mudaliar v. Life Insurance Corporation of India", "AIR 1963 SC 1185"),
    (3, "Shuttleworth v. Cox Brothers and Company (Maidenhead) Ltd.", "[1927] 2 K.B. 9"), (3, "Rayfield v. Hands", "[1957 R. No. 603]"),
    (3, "Jehangir Rastamji Modi v. Shamji Ladha", "(1867) BHCR 185"),
    (5, "City Equitable Fire Insurance Co., Re", "(1925) Ch. 407"), (5, "Standard Chartered Bank v. Pakistan National Shipping Corp.", "(2003) 1 All ER 173 (HL)"),
    (5, "Sanjay Dutt v. State of Haryana", "Supreme Court of India, 2 January 2025, Criminal Appeal No. 11 of 2025"),
    (7, "H.R. Harmer Ltd., Re", "(1958) 3 All E.R. 689"), (7, "Scottish Co-operative Wholesale Society Ltd. v. Meyer", "1959 AC 324"),
    (7, "Needle Industries (India) Ltd. v. Needle Industries Newey (India) Holdings Ltd.", "AIR 1981 SC 1298"),
    (8, "Aluminum Corporation of India Ltd. v. Lakshmi Rattan Cotton Mills Co. Ltd.", "AIR 1970 All 452"), (8, "Yenidje Tobacco Co. Ltd., Re", "(1916) 2 Ch. D. 169"),
    (8, "Devas Employees Mauritius Private Limited v. Union of India", "Karnataka High Court, W.P. No. 6191 of 2021, 28 April 2021"),
    (10, "Arun Kumar Jagatramka v. Jindal Steel and Power Ltd.", "AIR 2021 SC 136"),
]
STUDIES = [(5, "Satyam scandal", "The syllabus names a case study on the Satyam scandal under Board of Directors. The course book carries no text for it. Add your material to the guide folder and it can be built here.")]


def find_source():
    for root in (Path("cases"), Path(".")):
        if not root.exists():
            continue
        for pat in ("*.pdf", "*.txt"):
            for f in sorted(root.rglob(pat)):
                n = f.name.lower().replace("_", " ")
                if ("company law" in n or "lb-303" in n or "lb 303" in n or "case book" in n) and not any(x in f.parts for x in (".git", "raw", "data", "site", "drafts", "lodr")):
                    return f
    return None


def read_text(f):
    if f.suffix.lower() == ".txt":
        return f.read_text(encoding="utf-8")
    import fitz
    return "\n".join(p.get_text("text") for p in fitz.open(f))


def para_build(lines):
    lines = [l.rstrip() for l in lines]
    lens = sorted(len(l) for l in lines if len(l) > 40)
    typical = lens[int(len(lens) * 0.8)] if lens else 90
    out, cur = [], ""
    judge = re.compile(r"^(?:[A-Z][A-Z .,'’\-]{3,50}(?:,? (?:J|JJ|L\.?J|C\.?J|L\.?C|M\.?R|V\.?C)\.?)?\s*[–\-—:]|(?:LORD|LORDS|MR\.?|MRS|JUSTICE|THE LORD)\b[^a-z]{0,60}[–\-—:]|\d{1,3}\.\s+[A-Z“\"(]|\([a-z0-9ivx]+\)\s|[A-Z]\.\s)")
    for i, l in enumerate(lines):
        s = l.strip()
        if not s:
            continue
        if re.fullmatch(r"\d{1,3}", s):
            continue
        if cur:
            prev_end = cur.rstrip()[-1:]
            short = len(lines[i - 1].strip()) < typical * 0.86 if i else False
            if (prev_end in ".?:;”\")" and (short or judge.match(s))) or (judge.match(s) and prev_end in ".?:;”\")"):
                out.append(cur); cur = s
                continue
            if cur.endswith("-") and s[:1].islower():
                cur = cur[:-1] + s
            else:
                cur += " " + s
        else:
            cur = s
    if cur:
        out.append(cur)
    return [re.sub(r"\s+", " ", p).strip() for p in out if len(p.strip()) > 1]


def main():
    src = find_source()
    if not src:
        print("No case book found (a file with 'company law' or 'LB-303' in its name). Skipping.")
        return
    lines = read_text(src).replace("\r", "").replace("\x0c", "\n").split("\n")
    # skip the syllabus: the first case heading begins the body
    body0 = next((i for i, l in enumerate(lines) if l.strip().startswith("Salomon v. Salomon") and i > 100), None)
    if body0 is None:
        print("Could not find the first case in the book; the layout may be different from what was expected.")
        return
    starts, at = [], body0
    for num, key, *_ in CASES:
        j = None
        for i in range(at, len(lines)):
            s = lines[i].strip()
            if not (5 < len(s) < 150 and s.lower().startswith(key.lower())):
                continue
            nxt = lines[i + 1].strip() if i + 1 < len(lines) else ""
            if re.search(r"[,;]\s*$", s) or (re.match(r"^[a-z]", nxt) and not re.match(r"^(?:ors|and|ors\.)\b", nxt)):
                continue
            if re.match(r"^(\(|\[|AIR|Writ|MANU|Civil|Ors|With|\d{4}\b)", nxt) or re.search(r"(?:Ltd|Limited|Ltd\.|Company|Ors\.?|Co\.?|and|&)\s*$", s):
                j = i
                break
        if j is not None and lines[j - 1].strip().startswith("Civil Appeal"):
            j -= 1
        if j is None:
            print(f"  case {num} ({key}) not found in the book")
            starts.append(None)
        else:
            starts.append(j)
            at = j + 1
    found = [(n, s) for n, s in zip(CASES, starts) if s is not None]
    if len(found) < 20:
        print(f"Only {len(found)} cases were found in the book; the layout may be different.")
        return
    out = SITE / "cl"
    out.mkdir(parents=True, exist_ok=True)
    for f in out.glob("*.json"):
        f.unlink()
    recs = []
    for k, (c, s) in enumerate(found):
        end = found[k + 1][1] if k + 1 < len(found) else len(lines)
        seg = lines[s:end]
        # drop trailing page number and chapter-header lines belonging to the next case
        while seg and (not seg[-1].strip() or re.fullmatch(r"\d{1,3}", seg[-1].strip()) or (seg[-1].strip().isupper() and len(seg[-1].strip()) < 70)):
            seg.pop()
        # the title and citation lines at the top are shown in the header
        i = 0
        while i < min(6, len(seg)) and (i == 0 or re.match(r"^(\(|\[|AIR|Writ|MANU|Ors|Ors\.|\d{4})", seg[i].strip()) or len(seg[i].strip()) < 25):
            i += 1
        if seg and seg[0].strip().startswith("Civil Appeal"):
            i = 0
        paras = para_build(seg[i:])
        num, key, name, cit, court, topic, secs = c
        recs.append({"id": f"c{num}", "n": num, "t": name, "cit": cit, "court": court, "topic": topic, "s": secs, "paras": paras})
    for k, r in enumerate(recs):
        r["prev"] = [recs[k - 1]["id"], recs[k - 1]["t"]] if k else None
        r["next"] = [recs[k + 1]["id"], recs[k + 1]["t"]] if k + 1 < len(recs) else None
        (out / f"{r['id']}.json").write_text(json.dumps(r, ensure_ascii=False, separators=(",", ":")), encoding="utf-8")
    topics = []
    for tn, tt, tsyl, tsecs in TOPICS:
        topics.append({"n": tn, "t": tt, "syl": tsyl, "s": tsecs,
                       "cases": [{"id": r["id"], "t": r["t"], "cit": r["cit"], "court": r["court"]} for r in recs if r["topic"] == tn],
                       "ref": [{"t": a, "cit": b} for (x, a, b) in SUGGESTED if x == tn],
                       "study": [{"t": a, "note": b} for (x, a, b) in STUDIES if x == tn]})
    (out / "index.json").write_text(json.dumps({"topics": topics}, ensure_ascii=False, separators=(",", ":")), encoding="utf-8")
    (out / "text.json").write_text(json.dumps({r["id"]: (r["t"] + " " + r["cit"] + " " + " ".join(r["paras"])).lower() for r in recs}, ensure_ascii=False, separators=(",", ":")), encoding="utf-8")

    # data for the explanation step
    Path("data").mkdir(exist_ok=True)
    ai = [{"id": r["id"], "t": r["t"], "cit": r["cit"], "court": r["court"], "topic": TOPICS[r["topic"] - 1][1], "text": "\n".join(r["paras"])} for r in recs]
    Path("data/caselaw.json").write_text(json.dumps(ai, ensure_ascii=False, separators=(",", ":")), encoding="utf-8")
    ai_dir = Path("data/ai/c")
    ai_ids = sorted(f.stem for f in ai_dir.glob("*.json")) if ai_dir.exists() else []
    if ai_ids:
        (SITE / "ai" / "c").mkdir(parents=True, exist_ok=True)
        for i_ in ai_ids:
            (SITE / "ai" / "c" / f"{i_}.json").write_bytes((ai_dir / f"{i_}.json").read_bytes())

    by_sec = {}
    for r in recs:
        for n in r["s"]:
            by_sec.setdefault(n, []).append([r["id"], r["t"], r["cit"]])
    patched = 0
    for n, lst in by_sec.items():
        p = SITE / "s" / f"{n}.json"
        if p.exists():
            sj = json.loads(p.read_text(encoding="utf-8"))
            sj["cl"] = lst
            p.write_text(json.dumps(sj, ensure_ascii=False, separators=(",", ":")), encoding="utf-8")
            patched += 1
    idx = SITE / "index.json"
    try:
        j = json.loads(idx.read_text(encoding="utf-8"))
        j["cl"] = len(recs)
        j["ai_c"] = ai_ids
        idx.write_text(json.dumps(j, ensure_ascii=False, separators=(",", ":")), encoding="utf-8")
    except (OSError, ValueError):
        pass
    print(f"{len(recs)} cases written from {src}; {patched} section pages now link to them; {sum(len(r['paras']) for r in recs)} paragraphs")


if __name__ == "__main__":
    try:
        main()
    except Exception as e:
        print(f"Case laws step ended with an error: {e}")
    sys.exit(0)
