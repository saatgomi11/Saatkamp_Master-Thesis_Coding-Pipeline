#!/usr/bin/env python3
"""
Manual disclosure-gap coder (reviewer 2) - browser UI (Streamlit).

Close Master-Supabase Transcripts.xlsx before saving.

Usage:
  pip install streamlit   # if not installed
  streamlit run manual_gap_coder_reviewer2.py
"""

from __future__ import annotations

import streamlit as st

from manual_gap_core_reviewer2 import (
    DEFAULT_WORKBOOK,
    SEVERITY_LABELS,
    clear_state,
    compute_gap_segmentation,
    find_resume_position,
    load_sessions,
    load_state,
    load_topics_config,
    parse_pre_likert,
    read_topic_values,
    save_negative_issue,
    save_severity,
    save_state,
    topic_text_to_chat_html,
)

st.set_page_config(page_title="Manual Gap Coder (Reviewer 2)", layout="wide", initial_sidebar_state="expanded")

CSS = """
<style>
.help-small { font-size: 0.78rem; color: #555; line-height: 1.35; margin: 0.2rem 0 0.6rem 0; }
.help-small ul { margin: 0.2rem 0 0.4rem 1rem; padding: 0; }
.transcript-box {
  background: #f8f9fa; border: 1px solid #dee2e6; border-radius: 8px;
  padding: 0.75rem; min-height: 420px; max-height: 70vh; overflow-y: auto;
}
.chat-wrap { display: flex; flex-direction: column; gap: 0.65rem; }
.chat-topic-title {
  font-size: 0.85rem; font-weight: 600; color: #495057; text-align: center;
  padding: 0.35rem 0.5rem; border-bottom: 1px solid #dee2e6; margin-bottom: 0.25rem;
}
.chat-row { display: flex; flex-direction: column; max-width: 92%; }
.chat-bubble-agent { align-self: flex-start; }
.chat-bubble-user { align-self: flex-end; }
.chat-bubble-other { align-self: center; }
.chat-role {
  font-size: 0.7rem; font-weight: 600; text-transform: uppercase; letter-spacing: 0.03em;
  margin-bottom: 0.15rem; color: #6c757d;
}
.chat-bubble-user .chat-role { text-align: right; color: #0a58ca; }
.chat-bubble-agent .chat-role { color: #5c636a; }
.chat-text {
  padding: 0.55rem 0.75rem; border-radius: 12px; font-size: 0.95rem; line-height: 1.45;
  word-wrap: break-word;
}
.chat-bubble-agent .chat-text {
  background: #fff; border: 1px solid #dee2e6; border-bottom-left-radius: 4px;
}
.chat-bubble-user .chat-text {
  background: #cfe2ff; border: 1px solid #9ec5fe; border-bottom-right-radius: 4px;
}
.chat-bubble-other .chat-text { background: #f1f3f5; border: 1px solid #dee2e6; }
.chat-plain { white-space: pre-wrap; font-size: 0.95rem; line-height: 1.5; padding: 0.5rem; }
.chat-empty { padding: 0.5rem; color: #6c757d; }
.progress-pill { display: inline-block; background: #0d6efd; color: white; padding: 0.35rem 0.75rem;
  border-radius: 999px; font-size: 0.9rem; margin-right: 0.5rem; }
</style>
"""

SEVERITY_HELP = {
    3: {
        "title": "Explicit (3)",
        "def": "Explicit problem, named cause/consequence, or direct criticism. Little hedging.",
        "examples": [
            '"The workload is unsustainable. We are consistently understaffed."',
            '"My manager is not engaged and it is affecting the team."',
            '"We spent three days on a deliverable the client did not ask for."',
        ],
    },
    2: {
        "title": "Hedged (2)",
        "def": "Clear concern with hedging, qualification, or softening. Concern identifiable but attenuated.",
        "examples": [
            '"It has not been the most sustainable few weeks, to be honest."',
            '"Leadership could maybe be a bit more visible on the case."',
            '"Some alignment issues with the client, not terrible, but there."',
        ],
    },
    1: {
        "title": "Absent (1)",
        "def": "No negative content on this construct, OR Likert already negative (no gap possible).",
        "examples": [
            '"The workload has been very manageable."',
            '"Leadership has been very present and clear."',
            "Likert = disagree -> no gap by definition.",
        ],
    },
}

NEGATIVE_ISSUE_HELP = """
**Negative issue = Yes** if speech in this topic segment contains any substantive work-related
problem, concern, strain, friction, or criticism (regardless of Likert). Session/topic-level for this segment.

**Keywords (anchors):** overload, workload (strain), stress, pressure, overtime, late, unsustainable,
unpredictable; problem, issue, challenge, difficult, unclear, inefficient, rework, bottleneck, blocker,
delay, missing, lack; frustration, conflict, concern, misaligned, ambiguity, unclear priorities.

Do **not** code Yes for clearly negated positives (e.g. "no overload at all").
"""


@st.cache_data
def _load_data(_reload_token: int = 0):
    sessions, sheet = load_sessions(DEFAULT_WORKBOOK)
    topics = load_topics_config()
    return sessions, sheet, topics


def _init_session_state(sessions, topics) -> None:
    n_sess = len(sessions)
    n_top = len(topics)
    if "si" not in st.session_state:
        saved = load_state()
        if saved and "session_index" in saved:
            st.session_state.si = min(int(saved["session_index"]), n_sess - 1)
            st.session_state.ti = min(int(saved["topic_index"]), n_top - 1)
        else:
            si, ti = find_resume_position(sessions, topics)
            st.session_state.si = si
            st.session_state.ti = ti
    if "reload" not in st.session_state:
        st.session_state.reload = 0


def _reload_sessions():
    st.cache_data.clear()
    st.session_state.reload = st.session_state.get("reload", 0) + 1


def _go(si: int, ti: int, n_sess: int, n_top: int) -> None:
    st.session_state.si = max(0, min(si, n_sess - 1))
    st.session_state.ti = max(0, min(ti, n_top - 1))
    save_state(st.session_state.si, st.session_state.ti)
    _reload_sessions()


def main() -> None:
    st.markdown(CSS, unsafe_allow_html=True)
    sessions, sheet_name, topics = _load_data(st.session_state.get("reload", 0))
    _init_session_state(sessions, topics)

    n_sess = len(sessions)
    n_top = len(topics)
    si = st.session_state.si
    ti = st.session_state.ti
    row = sessions.iloc[si]
    topic = topics[ti]
    vals = read_topic_values(row, topic)

    total_steps = n_sess * n_top
    step = si * n_top + ti + 1
    likert_preview = parse_pre_likert(row.get("pre-survey_response_json"), topic["pre_survey_question_id"])

    with st.sidebar:
        st.header("Game controls")
        if st.button("Continue where I left off", use_container_width=True):
            rsi, rti = find_resume_position(sessions, topics)
            _go(rsi, rti, n_sess, n_top)
            st.rerun()
        if st.button("Restart from session 1 / topic 1", use_container_width=True):
            clear_state()
            _go(0, 0, n_sess, n_top)
            st.rerun()
        st.divider()
        st.caption("Jump to session")
        jump_s = st.number_input("Session #", min_value=1, max_value=n_sess, value=si + 1, step=1)
        jump_t = st.selectbox(
            "Topic",
            range(n_top),
            format_func=lambda i: topics[i]["label"],
            index=ti,
        )
        if st.button("Jump", use_container_width=True):
            _go(int(jump_s) - 1, int(jump_t), n_sess, n_top)
            st.rerun()
        st.divider()
        st.markdown(f"**Workbook:** `{DEFAULT_WORKBOOK.name}`")
        st.markdown(f"**Sheet:** `{sheet_name}`")
        st.warning("Close Excel before each save.", icon="⚠️")
        coded = 0
        for i in range(n_sess):
            r = sessions.iloc[i]
            for t in topics:
                v = read_topic_values(r, t)
                if v["manual_gap"] is not None and v["negative_issue"] is not None:
                    coded += 1
        st.progress(coded / total_steps if total_steps else 0.0)
        st.caption(f"Coded topic-steps: **{coded}** / **{total_steps}**")

    st.title("Manual disclosure gap coder (Reviewer 2)")
    st.markdown(
        f'<span class="progress-pill">Session {si + 1} / {n_sess}</span>'
        f'<span class="progress-pill">Topic {ti + 1} / {n_top}</span>'
        f'<span class="progress-pill">Step {step} / {total_steps}</span>',
        unsafe_allow_html=True,
    )
    st.subheader(topic["label"])
    meta = (
        f"`session_id`: {row.get('session_id', '')} · "
        f"Excel row: **{int(row['_excel_row'])}** (writes go to this row)"
    )
    emp = row.get("employee_id")
    if emp is not None and str(emp).strip() and str(emp).lower() != "nan":
        meta += f" · `employee_id`: {emp}"
    st.caption(meta)

    if likert_preview is not None:
        st.info(f"Pre-survey Likert (auto): **{likert_preview}** -> saved to `{topic['manual_likert']}` when you pick severity.")
    else:
        st.error("No pre-survey Likert for this topic - merge/fill `pre-survey_response_json` before coding.")

    col_l, col_r = st.columns([1.05, 1], gap="large")

    with col_l:
        st.markdown("#### Transcript segment")
        text = row.get(topic["text_column"])
        text_s = str(text).strip() if text is not None and str(text).lower() != "nan" else ""
        if not text_s:
            st.warning("No text in this column for this session.")
        chat_html = topic_text_to_chat_html(text if text_s else None)
        st.markdown(f'<div class="transcript-box">{chat_html}</div>', unsafe_allow_html=True)

    with col_r:
        st.markdown("#### Transcript severity")
        st.caption("Stores **1-3** in `" + topic["manual_gap"] + "` · gap label in `" + topic["gap_segmentation"] + "`")

        current_sev = vals["manual_gap"]
        for sev in (3, 2, 1):
            info = SEVERITY_HELP[sev]
            selected = current_sev == sev
            ex_html = "<ul>" + "".join(f"<li>{_escape_html(x)}</li>" for x in info["examples"]) + "</ul>"
            st.markdown(
                f'<div class="help-small"><strong>{info["title"]}</strong> - {info["def"]}{ex_html}</div>',
                unsafe_allow_html=True,
            )
            if st.button(
                info["title"].split(" (")[0],
                key=f"sev_{si}_{ti}_{sev}",
                use_container_width=True,
                type="primary" if selected else "secondary",
            ):
                try:
                    result = save_severity(row, topic, sev, sheet_name, DEFAULT_WORKBOOK)
                    st.success(
                        f"Saved severity **{sev}** · Likert **{result['likert']}** · Gap **{result['segmentation']}**"
                    )
                    _reload_sessions()
                    st.rerun()
                except Exception as e:
                    st.error(str(e))

        if current_sev is not None:
            seg = vals["gap_segmentation"]
            if seg is None and likert_preview is not None:
                seg = compute_gap_segmentation(likert_preview, current_sev)
            st.markdown(f"**Current:** {SEVERITY_LABELS.get(current_sev, current_sev)} ({current_sev}) -> **{seg}**")

        st.divider()
        st.markdown("#### Negative issue (this topic segment)")
        st.markdown(f'<div class="help-small">{NEGATIVE_ISSUE_HELP}</div>', unsafe_allow_html=True)
        st.caption("Stores **Yes** / **No** in `" + topic["negative_issue"] + "`")

        neg = vals["negative_issue"]
        c1, c2 = st.columns(2)
        with c1:
            if st.button(
                "Yes",
                key=f"neg_y_{si}_{ti}",
                use_container_width=True,
                type="primary" if neg == "Yes" else "secondary",
            ):
                try:
                    save_negative_issue(row, topic, "Yes", sheet_name, DEFAULT_WORKBOOK)
                    st.success("Saved **Yes**")
                    _reload_sessions()
                    st.rerun()
                except Exception as e:
                    st.error(str(e))
        with c2:
            if st.button(
                "No",
                key=f"neg_n_{si}_{ti}",
                use_container_width=True,
                type="primary" if neg == "No" else "secondary",
            ):
                try:
                    save_negative_issue(row, topic, "No", sheet_name, DEFAULT_WORKBOOK)
                    st.success("Saved **No**")
                    _reload_sessions()
                    st.rerun()
                except Exception as e:
                    st.error(str(e))

    st.divider()
    nav1, nav2, nav3, nav4 = st.columns([1, 1, 1, 2])
    with nav1:
        if st.button("<- Back", use_container_width=True):
            if ti > 0:
                _go(si, ti - 1, n_sess, n_top)
            elif si > 0:
                _go(si - 1, n_top - 1, n_sess, n_top)
            st.rerun()
    with nav2:
        if st.button("Next ->", use_container_width=True):
            if ti < n_top - 1:
                _go(si, ti + 1, n_sess, n_top)
            elif si < n_sess - 1:
                _go(si + 1, 0, n_sess, n_top)
            st.rerun()
    with nav3:
        if st.button("Next uncoded", use_container_width=True):
            rsi, rti = find_resume_position(sessions, topics)
            _go(rsi, rti, n_sess, n_top)
            st.rerun()
    with nav4:
        st.caption(
            f"Likert: {vals['manual_likert'] or '-'} · "
            f"Severity: {vals['manual_gap'] or '-'} · "
            f"Gap: {vals['gap_segmentation'] or '-'} · "
            f"Negative issue: {vals['negative_issue'] or '-'}"
        )


def _escape_html(s: str) -> str:
    return (
        s.replace("&", "&amp;")
        .replace("<", "&lt;")
        .replace(">", "&gt;")
        .replace('"', "&quot;")
    )


if __name__ == "__main__":
    main()
