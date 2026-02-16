from __future__ import annotations

import json
from datetime import datetime, timezone
from pathlib import Path

from openpyxl import Workbook

from .mapping import MappingDecision


def write_unmapped_workbook(path: Path, rows: list[MappingDecision]) -> None:
    wb = Workbook()
    ws = wb.active
    ws.title = "Unmapped"
    ws.append([
        "qbo_account_name",
        "qbo_amount",
        "top_5_suggested_template_labels",
        "confidence_per_suggestion",
        "reason",
    ])

    for row in rows:
        labels = [label for label, _ in row.suggestions]
        confidences = [round(score, 4) for _, score in row.suggestions]
        ws.append([
            row.qbo_account_name,
            row.qbo_amount,
            "; ".join(labels),
            "; ".join(str(score) for score in confidences),
            row.reason,
        ])

    wb.save(path)


def write_tieout(path: Path, payload: dict) -> None:
    payload["generated_at_utc"] = datetime.now(timezone.utc).isoformat()
    with path.open("w", encoding="utf-8") as handle:
        json.dump(payload, handle, indent=2)
