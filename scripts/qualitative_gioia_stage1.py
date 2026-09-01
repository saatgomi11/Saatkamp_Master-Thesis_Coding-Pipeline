#!/usr/bin/env python3
"""
Gioia Stage 1 — extraction + first-order coding (one Anthropic call per session).

Reads topic-segmented speech from the transcripts sheet; writes one row per quote to
``Qualitative Analysis``. Preserves ``My_Manual_Code`` and ``Include_Flag`` when
re-processing a session (--all-rows) if the quote text matches.

Close Excel before running.

Usage:
  python3 qualitative_gioia_stage1.py
  python3 qualitative_gioia_stage1.py --limit 3
  python3 qualitative_gioia_stage1.py --all-rows
  python3 qualitative_gioia_stage1.py --dry-run
"""

from __future__ import annotations

import argparse
import json
import sys
import time
from pathlib import Path

import pandas as pd

from qualitative_gioia_core import (
    DEFAULT_WORKBOOK,
    build_stage1_user_prompt,
    build_topic_speech_payload,
    call_anthropic_json,
    anthropic_client,
    load_transcript_sessions,
    manual_override_index,
    read_qualitative_sheet,
    sessions_already_in_qa,
    stage1_rows_from_response,
    write_qualitative_sheet,
    _load_prompt,
    PROMPT_STAGE1,
)


def main() -> int:
    parser = argparse.ArgumentParser(description="Gioia Stage 1: quote extraction + 1st-order codes.")
    parser.add_argument("--input", type=Path, default=DEFAULT_WORKBOOK)
    parser.add_argument("--all-rows", action="store_true", help="Re-run all eligible sessions (overwrite their QA rows).")
    parser.add_argument("--limit", type=int, default=0)
    parser.add_argument("--dry-run", action="store_true")
    args = parser.parse_args()

    path = args.input.expanduser().resolve()
    if not path.is_file():
        print(f"File not found: {path}", file=sys.stderr)
        return 1

    template = _load_prompt(PROMPT_STAGE1)
    sessions, transcript_sheet = load_transcript_sessions(path)
    df_qa = read_qualitative_sheet(path)
    overrides = manual_override_index(df_qa)
    done = sessions_already_in_qa(df_qa)

    indices: list[int] = []
    for i in range(len(sessions)):
        sid = str(sessions.iloc[i]["session_id"]).strip()
        if not args.all_rows and sid in done:
            continue
        indices.append(i)
    if args.limit > 0:
        indices = indices[: args.limit]

    if not indices:
        print(f"No sessions to process on {transcript_sheet!r} (all already in Qualitative Analysis?).")
        return 0

    if args.dry_run:
        print(f"Dry run: would call Anthropic for {len(indices)} session(s).")
        for i in indices[:10]:
            r = sessions.iloc[i]
            speech = build_topic_speech_payload(r)
            n_topics = sum(1 for v in speech.values() if v)
            print(f"  {r['session_id']} interview#{r.get('session_number_per_employee')} topics_with_text={n_topics}")
        if len(indices) > 10:
            print(f"  … +{len(indices) - 10} more")
        return 0

    client, model, max_tokens, sleep_s = anthropic_client()
    new_rows: list[dict] = []
    processed = 0

    # Keep rows for sessions we are NOT re-processing.
    if not df_qa.empty:
        rerun_ids = {str(sessions.iloc[i]["session_id"]).strip() for i in indices}
        keep = df_qa[~df_qa["Session_ID"].astype(str).str.strip().isin(rerun_ids)].copy()
    else:
        keep = df_qa.copy()

    for i in indices:
        row = sessions.iloc[i]
        sid = str(row["session_id"]).strip()
        interview_num = row.get("session_number_per_employee")
        speech = build_topic_speech_payload(row)
        if not any(speech.values()):
            print(f"Skip {sid}: no topic segment text.", file=sys.stderr)
            continue

        user_prompt = build_stage1_user_prompt(template, sid, interview_num, speech)
        try:
            data = call_anthropic_json(client, model, max_tokens, user_prompt)
        except Exception as e:
            print(f"API/parse error for {sid}: {e}", file=sys.stderr)
            continue

        response_json_text = json.dumps(data, ensure_ascii=False)
        quote_rows = stage1_rows_from_response(
            sid, interview_num, data, response_json_text, overrides
        )
        new_rows.extend(quote_rows)
        processed += 1
        print(f"  {sid}: {len(quote_rows)} quote(s)")
        if sleep_s > 0:
            time.sleep(sleep_s)

    out = pd.concat([keep, pd.DataFrame(new_rows)], ignore_index=True)
    if not out.empty:
        out = out.sort_values(["Session_ID", "Topic", "Quote"], kind="mergesort", na_position="last")
    write_qualitative_sheet(out, path)
    print(f"Done. Stage 1 processed {processed} session(s); {len(out)} total row(s) in Qualitative Analysis.")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
