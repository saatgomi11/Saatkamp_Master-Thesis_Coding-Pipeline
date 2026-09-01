"""Shared helpers for transcript Excel: project paths, sheet resolution, safer saving."""

from __future__ import annotations

import os
from pathlib import Path

from openpyxl.workbook.workbook import Workbook

# scripts/ -> project root
SCRIPTS_DIR = Path(__file__).resolve().parent
PROJECT_ROOT = SCRIPTS_DIR.parent

DATA_DIR = PROJECT_ROOT / "data"
PROMPTS_DIR = PROJECT_ROOT / "prompts"
CONFIG_DIR = PROJECT_ROOT / "config"
ENV_DIR = PROJECT_ROOT / "env"

DEFAULT_TRANSCRIPTS = DATA_DIR / "Master-Supabase Transcripts.xlsx"
DEFAULT_SURVEY = DATA_DIR / "Master-Supabase Survey Results.xlsx"


def load_project_dotenv() -> None:
    """Load env/.env first, then project-root .env."""
    try:
        from dotenv import load_dotenv
    except ImportError:
        return
    for p in (ENV_DIR / ".env", PROJECT_ROOT / ".env"):
        if p.is_file():
            load_dotenv(p)
            return


def transcripts_sheet_name(sheet_names: list[str]) -> str:
    """
    Pick the data sheet for Master-Supabase Transcripts.xlsx.

    Order: env TRANSCRIPTS_SHEET_NAME if set and present, then 'Database Sessions',
    then first sheet whose name contains 'transcript' (case-insensitive), else first sheet.
    """
    env = os.environ.get("TRANSCRIPTS_SHEET_NAME", "").strip()
    if env and env in sheet_names:
        return env
    if "Database Sessions" in sheet_names:
        return "Database Sessions"
    for n in sheet_names:
        if "transcript" in n.lower():
            return n
    return sheet_names[0]


def save_workbook_atomic(wb: Workbook, path: Path) -> None:
    """
    Write to a temp file next to path, then os.replace into place.

    If save fails, the original path is left unchanged (temp may remain; safe to delete *.tmp~).
    """
    path = path.resolve()
    tmp = path.with_name(path.stem + ".tmp~" + path.suffix)
    try:
        wb.save(tmp)
        os.replace(tmp, path)
    except BaseException:
        if tmp.exists():
            try:
                tmp.unlink()
            except OSError:
                pass
        raise
