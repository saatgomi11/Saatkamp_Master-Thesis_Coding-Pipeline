#!/usr/bin/env bash
# Run the full transcript + survey + NLP pipeline in order.
# Close Microsoft Excel (and any app locking the workbooks) before running.
#
# Usage (from project root):
#   ./run_full_pipeline.sh
#   bash run_full_pipeline.sh
#
# Optional: pass extra args to the first segmentation step only, e.g.:
#   SEGMENT_EXTRA="--limit 2" ./run_full_pipeline.sh

set -euo pipefail
ROOT="$(cd "$(dirname "${BASH_SOURCE[0]}")" && pwd)"
cd "$ROOT"

SEGMENT_EXTRA="${SEGMENT_EXTRA:-}"
PY=(python3)
export PYTHONPATH="${ROOT}/scripts${PYTHONPATH:+:$PYTHONPATH}"

run() {
  echo "==> $*"
  "${PY[@]}" "$@"
}

run scripts/merge_survey_to_transcripts.py
run scripts/extract_post_feedback_topic.py
run scripts/process_transcripts.py
run scripts/compute_week_numbers.py
run scripts/compute_turn_metrics.py
run scripts/segment_transcripts_anthropic.py ${SEGMENT_EXTRA}
run scripts/compute_disclosure_gap_anthropic.py
run scripts/compute_word_counts.py
run scripts/segment_transcripts_anthropic.py --refresh-negative-issue

echo "==> Pipeline finished."
