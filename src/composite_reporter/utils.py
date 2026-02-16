from __future__ import annotations

import re
import unicodedata
from difflib import SequenceMatcher
from typing import Iterable

SYNONYM_REPLACEMENTS = {
    "sales": "revenue",
    "income": "revenue",
    "cogs": "cost of goods sold",
    "costs": "cost",
    "wages": "labor",
    "technician": "tech",
    "technicians": "tech",
    "a/r": "accounts receivable",
    "a p": "accounts payable",
    "ar": "accounts receivable",
    "ap": "accounts payable",
    "ro": "repair order",
    "shop supplies": "shop supply",
    "hazmat": "hazardous waste",
    "tire disposal": "tire fee",
    "sublet": "outsourced repair",
}


def normalize_text(value: object) -> str:
    if value is None:
        return ""
    text = str(value)
    # Normalize dash variants to spaces so account words do not collapse.
    text = text.replace("\u2010", " ").replace("\u2011", " ").replace("\u2012", " ").replace("\u2013", " ").replace("\u2014", " ").replace("\u2212", " ")
    text = unicodedata.normalize("NFKD", text)
    text = text.encode("ascii", "ignore").decode("ascii")
    text = text.casefold().strip()
    text = text.replace("&", " and ")
    text = re.sub(r"[^a-z0-9\s]", " ", text)
    text = re.sub(r"\s+", " ", text).strip()
    for source, target in SYNONYM_REPLACEMENTS.items():
        text = re.sub(rf"\b{re.escape(source)}\b", target, text)
    text = re.sub(r"\s+", " ", text).strip()
    return text


def normalize_account_name(value: object) -> str:
    text = normalize_text(value)
    # QBO account names often begin with account numbers that can change over time.
    text = re.sub(r"^\d{2,}\s+", "", text).strip()
    return text


def parse_amount(value: object) -> float | None:
    if value is None:
        return None
    if isinstance(value, (int, float)):
        return float(value)

    text = str(value).strip()
    if not text:
        return None

    negative = False
    if text.startswith("(") and text.endswith(")"):
        negative = True
        text = text[1:-1].strip()

    text = text.replace("$", "").replace(",", "")
    if text in {"-", "--", "N/A"}:
        return None

    try:
        number = float(text)
    except ValueError:
        return None
    return -number if negative else number


def similarity_score(left: str, right: str) -> float:
    left_norm = normalize_text(left)
    right_norm = normalize_text(right)
    if not left_norm or not right_norm:
        return 0.0

    ratio = SequenceMatcher(None, left_norm, right_norm).ratio()
    left_tokens = set(left_norm.split())
    right_tokens = set(right_norm.split())
    overlap = len(left_tokens & right_tokens) / max(len(left_tokens | right_tokens), 1)
    return max(0.0, min(1.0, (ratio * 0.7) + (overlap * 0.3)))


def top_matches(query: str, candidates: Iterable[str], limit: int = 5) -> list[tuple[str, float]]:
    scored = [(label, similarity_score(query, label)) for label in candidates]
    scored.sort(key=lambda item: item[1], reverse=True)
    return scored[:limit]
