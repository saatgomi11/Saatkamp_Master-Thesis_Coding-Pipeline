# Manual disclosure gap coder

Browser UI to code transcript severity (1–3), negative issue (Yes/No), and auto-fill Likert + gap segmentation into `data/Master-Supabase Transcripts.xlsx`.

## Before you start

1. **Close Excel** (the workbook is saved after each click).
2. Ensure columns exist (see `config/manual_gap_topics.json` for exact header names).
3. Pre-survey JSON merged into `pre-survey_response_json`.

## Run

From the project root:

```bash
pip install streamlit   # or: pip install -r requirements.txt
streamlit run scripts/manual_gap_coder.py
```

Opens `http://localhost:8501` in your browser.

Second reviewer app (writes to `_2` columns):

```bash
streamlit run scripts/manual_gap_coder_reviewer2.py
```

## What it does

| Your action | Excel columns |
|-------------|----------------|
| Click **Explicit / Hedged / Absent** | `manual_gap_*` = 3 / 2 / 1 |
| (automatic) | `manual_likert_*` = 1–4 from pre-survey |
| (automatic) | `gap_segmentation_*` = Strong / Moderate / None |
| Click **Yes / No** | `negative_issue_*` = Yes / No |

**Sessions:** all rows with a `session_id` in the workbook (sorted by `started_at` for navigation). Saves always target the **physical Excel row** for that `session_id` (shown as “Excel row” under the topic title).

> If you coded before the row-mapping fix, manual values in the **top rows** of the sheet may be wrong — clear those cells or re-code after verifying `session_id` matches the transcript you saw.

**Topics (4):** overall experience, workload, client value, leadership — text from `text_*` columns.

## Resume / restart

- **Continue where I left off** — jumps to first topic-step missing severity or negative issue (or uses `.manual_gap_coder_state.json` on first open).
- **Restart from session 1** — clears state file only (does not erase Excel).
- **Next uncoded** — skip to next incomplete step.

Progress bar in the sidebar counts topic-steps with both severity and negative issue filled.

## Gap scoring (codebook)

- Product = Likert (1–4) × transcript severity (1–3)
- Likert ≤ 2 → **None**
- Product ≥ 9 → **Strong**; 5–8 → **Moderate**; else **None**

## Files

- `scripts/manual_gap_coder.py` — Streamlit UI
- `scripts/manual_gap_core.py` — Excel I/O and scoring
- `scripts/manual_gap_coder_reviewer2.py` — reviewer 2 Streamlit UI (`*_2` columns)
- `scripts/manual_gap_core_reviewer2.py` — reviewer 2 Excel I/O and scoring
- `config/manual_gap_topics.json` — column mapping
- `config/manual_gap_topics_reviewer2.json` — reviewer 2 column mapping
- `.manual_gap_coder_state.json` — last position (local, optional gitignore)
- `.manual_gap_coder_state_reviewer2.json` — reviewer 2 last position
