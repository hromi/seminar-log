#!/usr/bin/env python3
"""Attendance / activity-point overview for the AI-literacy courses.

All AI calls go ONLY to the in-house gateway (api.udk.digital, OpenAI-compatible).

Two paper forms are recognised automatically:
  * Participant_List_Landscape_15_Rows      -> 1 point per person per date ("Liste")
  * Participant_Attendance_Portrait_Signatures (Activity Log)
                                            -> 0.5 check-in + 0.5 check-out signature,
                                               Activity column (0/1) recorded, weight in sheet
Weights are editable in the sheet "Einstellungen".

  seminar_log.py models            list models available to your key
  seminar_log.py extract [--force] read NEW photos/PDF scans in DATA -> DATA/extracted/*.json
  seminar_log.py build             Moodle list + extracted/*.json -> DATA/Teilnahme.xlsx
  seminar_log.py run               extract + build  (use this each week after adding scans)

DATA = folder given by --data, $ATTENDANCE_DIR, or ./data (holds `moodle`, photos, scans).
Env: UDK_API_KEY (required for models/extract), UDK_MODEL (vision model),
     UDK_BASE_URL (default https://api.udk.digital/v1), UDK_MAX_EDGE (default 2400 px)
Corrections: edit participant_id / date in AI/extracted/<file>.json, then run `build`.
"""
import argparse, base64, difflib, hashlib, io, json, os, re, subprocess, sys, tempfile
from datetime import date, datetime
from pathlib import Path

import requests
from PIL import Image, ImageOps
from openpyxl import Workbook, load_workbook
from openpyxl.styles import Alignment, Font, PatternFill
from openpyxl.utils import get_column_letter as L

BASE_URL = os.environ.get("UDK_BASE_URL", "https://api.udk.digital/v1").rstrip("/")
def set_data_dir(path):
    global DATA_DIR, MOODLE, EXTRACT_DIR, OUT_XLSX
    DATA_DIR = Path(path).expanduser().resolve()
    MOODLE = DATA_DIR / "moodle"
    EXTRACT_DIR = DATA_DIR / "extracted"
    OUT_XLSX = DATA_DIR / "Teilnahme.xlsx"


set_data_dir(os.environ.get("ATTENDANCE_DIR", "data"))
IMG_EXT = {".jpg", ".jpeg", ".png", ".webp", ".tif", ".tiff"}
MAX_EDGE = int(os.environ.get("UDK_MAX_EDGE", "2400"))
SCHEMA = 2

# kind -> default weight (editable in sheet "Einstellungen")
WEIGHTS = {"Liste": 1, "Check-in": 0.5, "Check-out": 0.5, "Activity": 1}
RANK = {"high": 2, "medium": 1, "low": 0}

YELLOW = PatternFill("solid", fgColor="FFF2CC")
HEAD = PatternFill("solid", fgColor="D9E1F2")


# ---------- Moodle participants ----------
def strip_avatar_initials(name):
    """Moodle's copy&paste prepends the avatar initials: 'ABAnna Berg' -> 'Anna Berg'."""
    for k in range(1, 5):
        pre, rest = name[:k], name[k:]
        if pre.isalpha() and pre.isupper() and len(rest) > 1 and rest[0] == pre[0] and rest[1].islower():
            return rest
    return name


def load_participants(path=None):
    """Parse the copy&paste of Moodle's participant page (tab-separated rows with an e-mail)."""
    people, seen = [], set()
    for line in (path or MOODLE).read_text(encoding="utf-8").splitlines():
        if "@" not in line or "\t" not in line:
            continue
        fields = [f.strip() for f in line.split("\t")]
        email_i = next(i for i, f in enumerate(fields) if "@" in f)
        name = strip_avatar_initials(next((f for f in fields[:email_i] if f), ""))
        email = fields[email_i].lower()
        if email in seen:
            continue
        seen.add(email)
        people.append({"id": len(people) + 1, "name": name, "email": email})
    return people


# ---------- Input files ----------
def sha1(path):
    return hashlib.sha1(path.read_bytes()).hexdigest()


def input_files():
    return sorted(p for p in DATA_DIR.rglob("*")
                  if p.is_file() and EXTRACT_DIR not in p.parents
                  and (p.suffix.lower() in IMG_EXT or p.suffix.lower() == ".pdf"))


def pages(path):
    """Yield (page_no, PIL.Image) – PDFs are rasterised page by page."""
    if path.suffix.lower() != ".pdf":
        yield 1, ImageOps.exif_transpose(Image.open(path)).convert("RGB")
        return
    with tempfile.TemporaryDirectory() as td:
        subprocess.run(["pdftoppm", "-r", "200", "-jpeg", str(path), f"{td}/p"], check=True)
        for i, f in enumerate(sorted(Path(td).glob("p*.jpg")), 1):
            yield i, Image.open(f).convert("RGB")


def to_data_url(img):
    img = img.copy()
    img.thumbnail((MAX_EDGE, MAX_EDGE))
    buf = io.BytesIO()
    img.save(buf, "JPEG", quality=85)
    return "data:image/jpeg;base64," + base64.b64encode(buf.getvalue()).decode()


def capture_date(path):
    try:
        raw = Image.open(path).getexif().get(306)
        return datetime.strptime(raw, "%Y:%m:%d %H:%M:%S").date().isoformat()
    except Exception:
        return date.fromtimestamp(path.stat().st_mtime).isoformat()


# ---------- Gateway ----------
def _headers():
    key = os.environ.get("UDK_API_KEY")
    if not key:
        sys.exit("UDK_API_KEY is not set (create a token at https://api.udk.digital -> Console -> Tokens).")
    return {"Authorization": f"Bearer {key}"}


def list_models():
    r = requests.get(f"{BASE_URL}/models", headers=_headers(), timeout=30)
    r.raise_for_status()
    for m in sorted(x["id"] for x in r.json()["data"]):
        print(m)


PROMPT = """You are reading a photo/scan of a hand-filled university course form (it may be rotated).
It is one of two printed forms:

A) "PARTICIPANT LIST" (A4 landscape). Header: "Course / seminar:" and "Date:".
   Columns: Pronoun | Given name | Family name | Study type (Ba/MA/PhD/Absolvent/MS) |
   Faculty / study program | Mail | Modality (A/V/L/I) | Teacher? (checkbox) | Signature
B) "PARTICIPANT ACTIVITY LOG" (A4 portrait).
   Columns: Datum | Given name | Family name | Check-in (signature) | Check-out (signature) | Activity (0 or 1)
   Datum may be abbreviated (e.g. "3.11.") or a ditto mark (", ", 〃, -) meaning "same as row above".

The photo was taken on {photo_date}; use that year when a written date has no year.
Dates are European (day.month).

Official course participants from Moodle (id. name <mail>):
{roster}

For every filled row (skip empty rows) return the values exactly as written and the id of the matching
Moodle participant (spelling variants, nicknames, swapped names and mail addresses help; null if
nobody matches; never invent ids). confidence: "high" only if handwriting AND match are clear.

Answer with JSON only:
{{"doc_type": "participant_list" | "activity_log" | "other",
  "course": string|null,
  "date": "YYYY-MM-DD"|null,
  "rows": [{{
    "date": "YYYY-MM-DD"|null,
    "given_name": string|null, "family_name": string|null,
    "pronoun": string|null, "study_type": string|null, "faculty": string|null,
    "mail": string|null, "modality": string|null, "teacher": true|false|null,
    "signature": true|false|null,
    "check_in": true|false|null, "check_out": true|false|null, "activity": 0|1|null,
    "participant_id": int|null, "confidence": "high"|"medium"|"low"}}]}}
"date" on top level = header date of form A; "date" per row = Datum column of form B.
check_in / check_out = true only if a signature/mark is present in that cell."""


def parse_json(text):
    text = re.sub(r"^```(?:json)?|```$", "", text.strip(), flags=re.M).strip()
    m = re.search(r"\{.*\}", text, re.S)
    return json.loads(m.group(0) if m else text)


def ask_vision(model, img, people, photo_date):
    roster = "\n".join(f"{p['id']}. {p['name']} <{p['email']}>" for p in people)
    body = {
        "model": model,
        "temperature": 0,
        "messages": [{"role": "user", "content": [
            {"type": "text", "text": PROMPT.format(roster=roster, photo_date=photo_date)},
            {"type": "image_url", "image_url": {"url": to_data_url(img)}},
        ]}],
        "response_format": {"type": "json_object"},
    }
    r = requests.post(f"{BASE_URL}/chat/completions", headers=_headers(), json=body, timeout=900)
    if r.status_code == 400 and "response_format" in r.text:  # backend without JSON mode
        body.pop("response_format")
        r = requests.post(f"{BASE_URL}/chat/completions", headers=_headers(), json=body, timeout=900)
    if not r.ok:
        raise RuntimeError(f"Gateway error {r.status_code}: {r.text[:500]}")
    return parse_json(r.json()["choices"][0]["message"]["content"])


def cached_keys():
    keys = set()
    for f in EXTRACT_DIR.glob("*.json"):
        try:
            d = json.loads(f.read_text(encoding="utf-8"))
        except Exception:
            continue
        if d.get("schema") == SCHEMA:
            keys.add((d.get("sha1"), d.get("page")))
    return keys


def extract(force=False):
    model = os.environ.get("UDK_MODEL")
    if not model:
        sys.exit("UDK_MODEL is not set (pick a vision model from `seminar_log.py models`).")
    people = load_participants()
    EXTRACT_DIR.mkdir(exist_ok=True)
    done = set() if force else cached_keys()
    new = failed = 0
    for path in input_files():
        h = sha1(path)
        rel = path.relative_to(DATA_DIR)
        for page, img in pages(path):
            if (h, page) in done:
                continue
            print(f"reading {rel} p{page} with {model} …", flush=True)
            try:
                res = ask_vision(model, img, people, capture_date(path))
            except Exception as e:  # keep going, retry next run
                print(f"  FAILED: {e}")
                failed += 1
                continue
            res.update(schema=SCHEMA, source=str(rel), page=page, sha1=h,
                       photo_date=capture_date(path), model=model,
                       extracted_at=datetime.now().isoformat(timespec="seconds"))
            out = EXTRACT_DIR / f"{re.sub(r'[^A-Za-z0-9._-]', '_', str(rel))}.p{page}.json"
            out.write_text(json.dumps(res, ensure_ascii=False, indent=2), encoding="utf-8")
            rows = res.get("rows", [])
            print(f"  {res.get('doc_type')}: {len(rows)} rows, "
                  f"{sum(1 for r in rows if not r.get('participant_id'))} unmatched")
            new += 1
    print(f"{new} new page(s) read" + (f", {failed} failed (re-run to retry)" if failed else ""))


# ---------- Consolidation ----------
def resolve(row, people, by_id, by_mail):
    """Validate the model's match; fall back to (fuzzy) mail match."""
    pid, conf = row.get("participant_id"), row.get("confidence") or "low"
    mail = (row.get("mail") or "").strip().lower().replace(" ", "")
    if mail:
        if mail in by_mail:
            if pid != by_mail[mail]:
                return by_mail[mail], "medium" if pid else conf
            return pid, conf
        close = difflib.get_close_matches(mail, list(by_mail), 1, 0.85)
        if close and not pid:
            return by_mail[close[0]], "medium"
    return (pid, conf) if pid in by_id else (None, conf)


def consolidate(people):
    by_id = {p["id"]: p for p in people}
    by_mail = {p["email"]: p["id"] for p in people}
    entries, review, profile = {}, [], {}
    docs = [json.loads(f.read_text(encoding="utf-8")) for f in sorted(EXTRACT_DIR.glob("*.json"))]
    docs = [d for d in docs if d.get("schema") == SCHEMA]

    def add(day, kind, pid, value, conf, src, written):
        key = (day, kind, pid)
        old = entries.get(key)
        if old is None or (value, RANK.get(conf, 0)) > (old["value"], RANK.get(old["conf"], 0)):
            entries[key] = dict(date=day, kind=kind, pid=pid, value=value, conf=conf, src=src, written=written)

    for d in sorted(docs, key=lambda d: d.get("extracted_at", "")):
        src = f"{d['source']} p{d['page']}"
        kind = d.get("doc_type")
        if kind not in ("participant_list", "activity_log"):
            review.append([d.get("date"), src, None, None, None, f"Dokumenttyp: {kind}"])
            continue
        prev_day = None
        for row in d.get("rows", []):
            written = " ".join(x for x in (row.get("given_name"), row.get("family_name")) if x) or row.get("mail") or "?"
            day = (d.get("date") if kind == "participant_list" else row.get("date") or prev_day) \
                or d.get("date") or d.get("photo_date")
            prev_day = day
            pid, conf = resolve(row, people, by_id, by_mail)
            if pid is None:
                guess = difflib.get_close_matches(written, [p["name"] for p in people], 1, 0.5)
                review.append([day, src, written, None, guess[0] if guess else None, "nicht zugeordnet"])
                continue
            if conf != "high":
                review.append([day, src, written, by_id[pid]["name"], None, f"Konfidenz {conf}"])
            if kind == "participant_list":
                if row.get("teacher"):
                    review.append([day, src, written, by_id[pid]["name"], None, "als Teacher markiert"])
                add(day, "Liste", pid, 1, conf, src, written)
                prof = profile.setdefault(pid, {})
                for k in ("pronoun", "study_type", "faculty", "modality", "mail"):
                    if row.get(k):
                        prof[k] = row[k]
            else:
                if row.get("check_in"):
                    add(day, "Check-in", pid, 1, conf, src, written)
                if row.get("check_out"):
                    add(day, "Check-out", pid, 1, conf, src, written)
                if row.get("activity") is not None:
                    add(day, "Activity", pid, int(row["activity"] or 0), conf, src, written)
    return sorted(entries.values(), key=lambda e: (e["date"], e["pid"], e["kind"])), review, profile


# ---------- Excel ----------
def _old_bonus():
    """Keep manually entered 'Zusatzpunkte' across rebuilds (keyed by e-mail)."""
    if not OUT_XLSX.exists():
        return {}
    try:
        ws = load_workbook(OUT_XLSX)["Übersicht"]
    except KeyError:
        return {}
    hdr = [c.value for c in ws[1]]
    if "Zusatzpunkte" not in hdr or "E-Mail" not in hdr:
        return {}
    ei, bi = hdr.index("E-Mail"), hdr.index("Zusatzpunkte")
    return {r[ei]: r[bi] for r in ws.iter_rows(min_row=2, values_only=True) if r[ei] and r[bi] not in (None, "")}


def style_header(ws, widths):
    for c in ws[1]:
        c.font, c.fill = Font(bold=True), HEAD
        c.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
    for i, w in enumerate(widths, 1):
        ws.column_dimensions[L(i)].width = w
    ws.freeze_panes = "A2"


def build():
    people = load_participants()
    entries, review, profile = consolidate(people)
    bonus = _old_bonus()
    dates = sorted({e["date"] for e in entries})
    n = len(people)
    wb = Workbook()

    # Overview
    ov = wb.active
    ov.title = "Übersicht"
    kinds = list(WEIGHTS)
    ov.append(["Nr", "Name", "E-Mail", "Liste", "Check-in", "Check-out", "Activity",
               "Termine", "Zusatzpunkte", "Punkte"])
    E = "Einträge!"
    for i, p in enumerate(people, 2):
        row = [p["id"], p["name"], p["email"]]
        row += [f'=SUMIFS({E}$G:$G,{E}$C:$C,$A{i},{E}$B:$B,"{k}")' for k in kinds]
        row += [f"=SUMPRODUCT((Termine!$D{i}:${L(3 + max(len(dates), 1))}{i}>0)*1)",
                bonus.get(p["email"]),
                f"=SUMIFS({E}$H:$H,{E}$C:$C,$A{i})+N(I{i})"]
        ov.append(row)
    style_header(ov, [5, 30, 34, 8, 9, 10, 9, 9, 12, 9])
    ov.freeze_panes = "D2"
    ov.auto_filter.ref = f"A1:J{n + 1}"

    # Per-date matrix (points per day)
    tm = wb.create_sheet("Termine")
    tm.append(["Nr", "Name", "E-Mail", *dates])
    for i, p in enumerate(people, 2):
        tm.append([p["id"], p["name"], p["email"]] +
                  [f'=SUMIFS({E}$H:$H,{E}$C:$C,$A{i},{E}$A:$A,{L(c)}$1)' for c in range(4, 4 + len(dates))])
    tm.append([None, "Anwesende"] + [None] +
              [f'=COUNTIF({L(c)}2:{L(c)}{n + 1},">0")' for c in range(4, 4 + len(dates))])
    tm.cell(n + 2, 2).font = Font(bold=True)
    style_header(tm, [5, 30, 34] + [11] * len(dates))
    tm.freeze_panes = "D2"

    # Entries (source of truth for all formulas)
    en = wb.create_sheet("Einträge")
    en.append(["Datum", "Art", "Nr", "Name", "Gelesen", "Konfidenz", "Wert", "Punkte", "Quelle"])
    by_id = {p["id"]: p for p in people}
    for r, e in enumerate(entries, 2):
        en.append([e["date"], e["kind"], e["pid"], by_id[e["pid"]]["name"], e["written"], e["conf"],
                   e["value"], f"=G{r}*IFERROR(VLOOKUP(B{r},Einstellungen!$A$2:$B$10,2,0),0)", e["src"]])
        if e["conf"] != "high":
            for c in en[r]:
                c.fill = YELLOW
    style_header(en, [12, 11, 5, 30, 30, 10, 6, 7, 40])
    en.auto_filter.ref = f"A1:I{len(entries) + 1}"

    # Profile data from participant lists
    sd = wb.create_sheet("Stammdaten")
    sd.append(["Nr", "Name", "E-Mail (Moodle)", "Pronoun", "Study type", "Faculty / program", "Modality", "Mail (Liste)"])
    for p in people:
        pr = profile.get(p["id"], {})
        sd.append([p["id"], p["name"], p["email"], pr.get("pronoun"), pr.get("study_type"),
                   pr.get("faculty"), pr.get("modality"), pr.get("mail")])
    style_header(sd, [5, 30, 34, 10, 12, 30, 10, 34])

    rv = wb.create_sheet("Prüfen")
    rv.append(["Datum", "Quelle", "Gelesen", "Zugeordnet", "Vorschlag", "Hinweis"])
    for row in sorted(review, key=lambda r: (str(r[0]), r[1])):
        rv.append(row)
    style_header(rv, [12, 40, 30, 30, 30, 22])

    cfg = wb.create_sheet("Einstellungen")
    cfg.append(["Art", "Punkte"])
    for k, w in WEIGHTS.items():
        cfg.append([k, w])
    cfg.append([])
    cfg.append(["Liste = Eintrag in Participant List (Landscape); Check-in/-out = Unterschrift im Activity Log; "
                "Activity = Wert 0/1 aus dem Activity Log"])
    style_header(cfg, [14, 10])

    wb.save(OUT_XLSX)
    print(f"{n} participants, {len(dates)} dates, {len(entries)} entries, {len(review)} rows to check -> {OUT_XLSX}")


if __name__ == "__main__":
    ap = argparse.ArgumentParser(description=__doc__, formatter_class=argparse.RawDescriptionHelpFormatter)
    ap.add_argument("cmd", choices=["models", "extract", "build", "run"])
    ap.add_argument("--force", action="store_true", help="re-read all pages even if cached")
    ap.add_argument("--data", help="data folder (default: $ATTENDANCE_DIR or ./data)")
    a = ap.parse_args()
    if a.data:
        set_data_dir(a.data)
    if a.cmd != "models" and not MOODLE.exists():
        sys.exit(f"Moodle participant list not found: {MOODLE}")
    if a.cmd == "models":
        list_models()
    if a.cmd in ("extract", "run"):
        extract(a.force)
    if a.cmd in ("build", "run"):
        build()
