# seminar-log

Turns photos or scans of hand-filled seminar attendance forms, together with the Moodle participant list, into an Excel overview of attendance and activity points.

The handwriting is read by a vision-language model through an **in-house, OpenAI-compatible AI gateway** (default: `https://api.udk.digital/v1`, GenKI Lab, UdK Berlin). Student data is sent **only** to that gateway and never to a commercial cloud AI. Everything else (Moodle parsing, matching checks, scoring, the Excel file) runs locally.

## Paper forms

Print the blank templates in [`templates/`](templates/):

| Template | Filled in | Points (default) |
|---|---|---|
| [`Participant_List_Landscape_15_Rows.pdf`](templates/Participant_List_Landscape_15_Rows.pdf) | once per session: course, date, pronoun, name, study type, faculty, mail, modality, teacher?, signature | **1** for appearing in the list |
| [`Participant_Attendance_Portrait_Signatures.pdf`](templates/Participant_Attendance_Portrait_Signatures.pdf) (activity log) | per row: date, name, check-in signature, check-out signature, activity 0/1 | **0.5** check-in, **0.5** check-out, **1** for activity = 1 |

Each kind of point counts at most once per person per date. The model detects which form is on each page by itself.

## Setup

```bash
git clone git@github.com:hromi/seminar-log.git && cd seminar-log
python3 -m venv .venv && .venv/bin/pip install -r requirements.txt
# PDF scans additionally need `pdftoppm` (poppler-utils)

export UDK_API_KEY=sk-...          # token from the gateway console (Console → Tokens)
.venv/bin/python seminar_log.py models   # list models available to your key
export UDK_MODEL=<vision-model>     # a model that accepts images
# optional: export UDK_BASE_URL=https://your-gateway/v1
```

## Data folder

By default this is `./data`. You can change it with `--data DIR` or `$ATTENDANCE_DIR`. The folder is git-ignored and **must never be committed**.

```
data/
├── moodle              # copy & paste of Moodle → Participants page (select all, paste as text)
├── week1/IMG_….jpg     # photos or PDF scans of filled forms, any sub-folder structure
├── week2/scan.pdf      # multi-page PDFs are split into pages
├── extracted/          # created: one JSON per page with what the model read (cache)
└── Teilnahme.xlsx      # created: the overview
```

## Usage

```bash
.venv/bin/python seminar_log.py run      # = extract + build
```

- `extract` sends only **new** pages to the gateway. Each page is identified by a content hash, so you can add new scans every week and simply run `run` again. Renamed or duplicated files are not read twice, and failed pages are retried on the next run. Use `--force` to re-read everything.
- `build` recreates `Teilnahme.xlsx` from the Moodle list and `extracted/*.json`. Values typed into the **Zusatzpunkte** (bonus points) column are kept.

### Workbook

| Sheet | Content |
|---|---|
| Übersicht | per participant: counts per kind, sessions attended, bonus points, **total points** |
| Termine | points per participant per date, attendees per date |
| Einträge | one row per person × date × kind, with source page and confidence. All formulas are based on this sheet. Uncertain rows are highlighted yellow |
| Stammdaten | pronoun, study type, faculty, modality, as written on the participant lists |
| Prüfen | rows to check: unmatched names (with a suggestion), low-confidence readings, people marked as teacher, unknown documents |
| Einstellungen | point weight per kind. Edits last until the next `build`; change `WEIGHTS` in the script to make them permanent |

### Corrections

Fix the `participant_id` (Moodle row number, see `Nr`) or a `date` in the matching `data/extracted/<file>.pN.json`, then run `seminar_log.py build`.

### Matching

The model receives the Moodle roster (name + mail) and proposes a participant id for each row. On the participant list, the handwritten mail is also checked locally: an exact or close mail match overrides the model's choice, and the row is then marked for review.

## Tips

- Photograph forms flat, filling the frame and in good light. If handwriting is misread, try `UDK_MAX_EDGE=3200` or a stronger vision model.
- Make sure the date is filled in on every form. Otherwise the date the photo was taken (EXIF data) is used.
