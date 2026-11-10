"""Step 5 (runs on its own schedule): write a plain-language explanation and an example for every section and rule.

Uses the free tier of the Google Gemini API. The key is kept in GitHub (Settings > Secrets > GEMINI_API_KEY),
never in the website. Explanations are written once, saved in data/ai/, and shown by the site. Each run does
as many as the free limits allow, then stops. The next run carries on. It never fails the build.

Reads   data/sections.json, data/rules.json
Writes  data/ai/s/<section>.json          one explanation per section
        data/ai/r/<rule set id>.json      explanations of all rules in one set, keyed by rule number
        data/ai_report.md
Settings (all optional, set as environment variables):
        GEMINI_API_KEY    required, otherwise nothing is done
        GEMINI_MODELS     comma-separated model names, tried in order (default below)
        AI_RPM            requests per minute to stay under (default 8)
        AI_MAX_PER_RUN    most explanations in one run (default 450)
        AI_MAX_MINUTES    stop after this many minutes (default 110)
        GEMINI_ENDPOINT   only for testing against a fake server
"""
import datetime
import hashlib
import json
import os
import re
import sys
import time
import urllib.error
import urllib.request
from pathlib import Path

DATA = Path("data")
AI = DATA / "ai"
PROMPT_VERSION = "1"
DEFAULT_MODELS = "gemini-3.5-flash-lite,gemini-3.1-flash-lite,gemini-3.8-flash"
ENDPOINT = (os.environ.get("GEMINI_ENDPOINT") or "https://generativelanguage.googleapis.com/v1beta").rstrip("/")
KEY = (os.environ.get("GEMINI_API_KEY") or "").strip()
MODELS = [m.strip() for m in (os.environ.get("GEMINI_MODELS") or DEFAULT_MODELS).split(",") if m.strip()]
RPM = max(1.0, float(os.environ.get("AI_RPM") or 8))
MAX_PER_RUN = int(os.environ.get("AI_MAX_PER_RUN") or 450)
MAX_MINUTES = float(os.environ.get("AI_MAX_MINUTES") or 110)

SCHEMA = {
    "type": "object",
    "properties": {
        "summary": {"type": "string"},
        "points": {"type": "array", "items": {"type": "string"}},
        "example": {"type": "string"},
        "terms": {"type": "array", "items": {"type": "object", "properties": {
            "term": {"type": "string"}, "meaning": {"type": "string"}}, "required": ["term", "meaning"]}},
        "watch": {"type": "string"},
    },
    "required": ["summary", "points", "example"],
}

RULES_OF_THE_ROAD = (
    "Use only the text given below. Do not add requirements, limits, time periods, amounts or penalties that are not in it. "
    "If the text is too short or unclear to explain, say that plainly in the summary. "
    "Write for a student or an ordinary reader with no legal training: short sentences, everyday words, and explain any legal term you must use. "
    "The example must be an invented, realistic situation with fictional names (for instance 'Meera Textiles Private Limited'), must follow the text exactly, "
    "and must not state figures or deadlines that the text does not contain. Do not give legal advice. "
    'Return JSON with: "summary" (2 to 4 sentences), "points" (list of short bullets), "example" (a short scenario), '
    '"terms" (list of {"term","meaning"} for hard words in the text), "watch" (one short sentence on a common confusion, or an empty string).'
)


def clean(text, limit):
    t = re.sub(r"[ \t]+", " ", text.replace("\r", ""))
    t = re.sub(r"\n{2,}", "\n", t).strip()
    return t[:limit] + ("\n[text continues]" if len(t) > limit else "")


def section_prompt(s):
    return (f"You are explaining one section of the Companies Act, 2013 (India).\n{RULES_OF_THE_ROAD}\n"
            'For "points" give 3 to 6 bullets on what the section requires, allows or forbids. For "example" write 70 to 130 words. '
            'Give up to 4 "terms".\n\n'
            f"Section {s['section']}. {s['title']}\nChapter {s.get('chapter', '')}: {s.get('chapter_title', '')}\n\nTEXT:\n{clean(s['text'], 9000)}")


def rule_prompt(r, section_titles):
    rel = "; ".join(section_titles) if section_titles else "not stated"
    return (f"You are explaining one rule made under the Companies Act, 2013 (India).\n{RULES_OF_THE_ROAD}\n"
            'For "points" give 2 to 5 bullets. For "example" write 50 to 100 words. Give up to 3 "terms".\n\n'
            f"Rules: {r['rule_set']}\nRule {r['rule']}. {r['title']}\nRelated sections of the Act: {rel}\n\nTEXT:\n{clean(r['text'], 7000)}")


class Stop(Exception):
    pass


state = {"model_i": 0, "last_call": 0.0, "calls": 0}


def pace():
    wait = 60.0 / RPM - (time.time() - state["last_call"])
    if wait > 0:
        time.sleep(wait)
    state["last_call"] = time.time()


def call(prompt):
    """One request. Returns the parsed JSON object, or None if this item failed. Raises Stop when the run must end."""
    for attempt in range(4):
        if state["model_i"] >= len(MODELS):
            raise Stop("every model has reached its limit for today")
        model = MODELS[state["model_i"]]
        body = json.dumps({
            "contents": [{"parts": [{"text": prompt}]}],
            "generationConfig": {"temperature": 0.3, "maxOutputTokens": 4096,
                                 "responseMimeType": "application/json", "responseSchema": SCHEMA},
        }).encode("utf-8")
        req = urllib.request.Request(f"{ENDPOINT}/models/{model}:generateContent", data=body, method="POST",
                                     headers={"Content-Type": "application/json", "x-goog-api-key": KEY})
        pace()
        try:
            with urllib.request.urlopen(req, timeout=90) as resp:
                payload = json.loads(resp.read().decode("utf-8"))
            state["calls"] += 1
            parts = (((payload.get("candidates") or [{}])[0].get("content") or {}).get("parts") or [])
            text = "".join(p.get("text", "") for p in parts if not p.get("thought"))
            return json.loads(text), model
        except urllib.error.HTTPError as e:
            msg = e.read().decode("utf-8", "replace")
            if e.code == 429:
                if re.search(r"PerDay|per day|daily", msg, re.I):
                    print(f"  {model}: daily limit reached")
                    state["model_i"] += 1
                    continue
                m = re.search(r'"retryDelay"\s*:\s*"(\d+)', msg)
                delay = min(90, int(m.group(1)) + 2 if m else 30)
                print(f"  {model}: too many requests this minute, waiting {delay}s")
                time.sleep(delay)
                continue
            if e.code in (404, 400) and "model" in msg.lower() and ("not found" in msg.lower() or "not supported" in msg.lower()):
                print(f"  {model}: not available, trying the next model")
                state["model_i"] += 1
                continue
            if e.code in (401, 403) or (e.code == 400 and "API key" in msg):
                raise Stop(f"the key was refused (HTTP {e.code}). Check the GEMINI_API_KEY secret")
            if e.code >= 500:
                time.sleep(10 * (attempt + 1))
                continue
            print(f"  HTTP {e.code}: {msg[:200]}")
            return None, model
        except (urllib.error.URLError, TimeoutError, json.JSONDecodeError, ValueError) as e:
            time.sleep(5 * (attempt + 1))
            last = str(e)
    return None, MODELS[min(state["model_i"], len(MODELS) - 1)]


def tidy(obj):
    """Check the answer has the right shape and trim it. Returns a clean dict or None."""
    if not isinstance(obj, dict):
        return None
    summary = str(obj.get("summary") or "").strip()
    example = str(obj.get("example") or "").strip()
    points = [str(x).strip() for x in (obj.get("points") or []) if str(x).strip()][:7]
    terms = []
    for t in (obj.get("terms") or [])[:5]:
        if isinstance(t, dict) and str(t.get("term") or "").strip() and str(t.get("meaning") or "").strip():
            terms.append({"term": str(t["term"]).strip()[:80], "meaning": str(t["meaning"]).strip()[:300]})
    if len(summary) < 20 or len(example) < 20:
        return None
    return {"summary": summary[:900], "points": [p[:300] for p in points], "example": example[:1100],
            "terms": terms, "watch": str(obj.get("watch") or "").strip()[:300]}


def digest(*parts):
    return hashlib.sha1(("\x1f".join(parts) + PROMPT_VERSION).encode("utf-8")).hexdigest()[:12]


def load(path):
    try:
        return json.loads(path.read_text(encoding="utf-8"))
    except (OSError, ValueError):
        return {}


def save(path, obj):
    path.parent.mkdir(parents=True, exist_ok=True)
    path.write_text(json.dumps(obj, ensure_ascii=False, indent=1), encoding="utf-8")


def main():
    started = time.time()
    out_new = 0
    report = ["# AI explanations report", ""]

    def finish(note):
        report.append(f"- Result: {note}")
        report.append(f"- Written in this run: {out_new}")
        sec_done = len(list((AI / "s").glob("*.json"))) if (AI / "s").exists() else 0
        rule_done = sum(len(load(p)) for p in (AI / "r").glob("*.json")) if (AI / "r").exists() else 0
        report.append(f"- Sections explained so far: {sec_done}")
        report.append(f"- Rules explained so far: {rule_done}")
        DATA.mkdir(exist_ok=True)
        (DATA / "ai_report.md").write_text("\n".join(report) + "\n", encoding="utf-8")
        print("\n".join(report))
        gh = os.environ.get("GITHUB_OUTPUT")
        if gh:
            with open(gh, "a", encoding="utf-8") as f:
                f.write(f"new={out_new}\n")

    if not KEY:
        return finish("GEMINI_API_KEY is not set, so no explanations were written. Add the key as a repository secret.")
    if not (DATA / "sections.json").exists() or not (DATA / "rules.json").exists():
        return finish("data/sections.json or data/rules.json is missing. Run the main build first.")

    sections = json.loads((DATA / "sections.json").read_text(encoding="utf-8"))
    rules_doc = json.loads((DATA / "rules.json").read_text(encoding="utf-8"))
    sec_title = {s["section"]: s["title"] for s in sections}

    work = []   # (kind, key, prompt, hash)
    for s in sections:
        if s["status"] != "in force" or len(s["text"].strip()) < 120:
            continue
        h = digest(s["title"], clean(s["text"], 9000))
        if load(AI / "s" / f"{s['section']}.json").get("h") != h:
            work.append(("s", s["section"], section_prompt(s), h, s))
    cites = {}
    for r in rules_doc["rules"]:
        cites[r["id"]] = [c["section"] for c in r.get("cites", []) if c.get("confidence") != "low"][:3]
    cache = {}
    # rules that cite a section directly first, then the rest
    for r in sorted(rules_doc["rules"], key=lambda r: (0 if cites[r["id"]] else 1)):
        if r["status"] != "in force" or len(r["text"].strip()) < 60:
            continue
        stem = r["rule_set_id"]
        cache.setdefault(stem, load(AI / "r" / f"{stem}.json"))
        h = digest(r["title"], clean(r["text"], 7000))
        if cache[stem].get(r["rule"], {}).get("h") != h:
            titles = [f"section {n} ({sec_title.get(n, '')})" for n in cites[r["id"]]]
            work.append(("r", (stem, r["rule"]), rule_prompt(r, titles), h, r))

    report.append(f"- Waiting to be explained at the start of this run: {len(work)}")
    report.append(f"- Models in order: {', '.join(MODELS)}")
    print(f"{len(work)} items to explain; at most {MAX_PER_RUN} in this run")
    stopped = "all done" if not work else "stopped at the limit for this run"
    try:
        for kind, key, prompt, h, src in work:
            if out_new >= MAX_PER_RUN:
                stopped = f"stopped after {MAX_PER_RUN} items (the per-run limit); the next run continues"
                break
            if (time.time() - started) / 60 > MAX_MINUTES:
                stopped = "stopped at the time limit; the next run continues"
                break
            result = call(prompt)
            obj, model = result if isinstance(result, tuple) else (None, "")
            entry = tidy(obj)
            if not entry:
                print(f"  skipped {kind} {key}: no usable answer")
                continue
            entry.update({"h": h, "m": model, "d": datetime.date.today().isoformat()})
            if kind == "s":
                save(AI / "s" / f"{key}.json", entry)
            else:
                stem, rule = key
                cache[stem][rule] = entry
                save(AI / "r" / f"{stem}.json", cache[stem])
            out_new += 1
            if out_new % 25 == 0:
                print(f"  {out_new} done")
        else:
            if work:
                stopped = "everything waiting was explained"
    except Stop as e:
        stopped = f"stopped: {e}"
    finish(stopped)


if __name__ == "__main__":
    try:
        main()
    except Exception as e:     # this step must never break the website build
        print(f"AI step ended with an error: {e}")
    sys.exit(0)
