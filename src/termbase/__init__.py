# Termbase Module - 術語庫與 placeholder 機制
from .termbase import (
    Termbase,
    TermEntry,
    TermProtection,
    load_termbase_from_json,
    load_termbase_from_csv,
    create_combined_termbase
)

__all__ = [
    "Termbase",
    "TermEntry",
    "TermProtection",
    "load_termbase_from_json",
    "load_termbase_from_csv",
    "create_combined_termbase"
]
