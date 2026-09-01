# Gioia qualitative coding — Stage 1

Stage 1 Gioia analysis for disclosure quotes in AI check-in transcripts. Stage 2 (second-order themes) and Stage 3 (aggregate dimensions) are not part of this repo’s runnable scripts, as they are assessed and identified manually.

## Excel layout

- **Transcripts:** `data/Master-Supabase Transcripts.xlsx` (sheet resolved by `workbook_utils`) — **unchanged** by this pipeline.
- **Output:** sheet **`Qualitative Analysis`** (header row 1).

| Column | Source |
|--------|--------|
| Session_ID | transcript `session_id` |
| Interview_Number | `session_number_per_employee` (1 = first interview for that employee) |
| Response_JSON | full Stage 1 API JSON (same on each quote row for that session) |
| Quote | `quote_text` |
| Topic | topic label from model |
| Flaggin_Reason | `flagging_reason` |
| First_Code_Label | `first_order_code` |
| Confidence_Score | `confidence` |
| Second_Order | left for human / prior runs — Stage 1 does not write this |
| My_Manual_Code | **human only** — never overwritten by scripts |
| Include_Flag | **human only** — never overwritten by scripts |

## Do other scripts break?

**No.** Existing scripts only read/write the **transcripts** sheet by name. The second tab is ignored unless you run Stage 1.

## Prerequisites

- Close Excel before running.
- `ANTHROPIC_API_KEY` in `env/.env` (or project-root `.env`)
- Optional: `ANTHROPIC_MODEL` (e.g. Opus), `ANTHROPIC_MAX_TOKENS`, `ANTHROPIC_SLEEP_SEC`
- Topic segments populated (`text_overall_experience`, etc.)
- `session_number_per_employee` from `scripts/process_transcripts.py`

## Stage 1 — one call per session

From the project root:

```bash
python3 scripts/qualitative_gioia_stage1.py              # only new sessions
python3 scripts/qualitative_gioia_stage1.py --all-rows   # re-run all eligible sessions
python3 scripts/qualitative_gioia_stage1.py --limit 3    # test
python3 scripts/qualitative_gioia_stage1.py --dry-run
```

- Prompt: `prompts/Prompt_Gioia_Quotes`
- Skips rows without `session_id` or `transcription`
- One **row per quote**; sessions with zero quotes add no rows
- Re-run with `--all-rows` keeps `My_Manual_Code` / `Include_Flag` when quote text matches
