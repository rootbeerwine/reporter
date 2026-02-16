from __future__ import annotations

import csv
from dataclasses import dataclass
from pathlib import Path

from .utils import normalize_account_name, normalize_text, top_matches


@dataclass
class MappingDecision:
    qbo_account_name: str
    qbo_amount: float
    mapped_template_label: str | None
    confidence: float
    reason: str
    suggestions: list[tuple[str, float]]
    mapping_source: str = "none"


DOMAIN_RULES_PL: list[tuple[str, tuple[str, ...]]] = [
    ("parts", ("parts",)),
    ("labor", ("labor",)),
    ("sublet", ("sublet", "outsourced repair")),
    ("outsourced repair", ("sublet", "outsourced repair")),
    ("tire", ("tire",)),
    ("shop supply", ("shop supply",)),
    ("hazardous waste", ("hazardous waste", "hazmat", "disposal", "tire fee")),
    ("discount", ("discount",)),
    ("credit card fee", ("merchant fee", "credit card fee", "bank fee")),
    ("bank charge", ("bank fee", "merchant fee")),
    ("rent", ("rent",)),
    ("payroll tax", ("payroll tax",)),
    ("training", ("employee training", "training")),
    ("insurance", ("insurance",)),
    ("depreciation", ("depreciation",)),
    ("interest expense", ("interest",)),
    ("salaries technicians", ("labor", "technician", "wages")),
    ("salaries service advisors", ("labor", "service advisor", "wages")),
    ("salaries management", ("labor", "management", "wages")),
    ("salaries owners officers", ("labor", "owner", "officer", "wages")),
    ("salary", ("labor", "wages")),
]

DOMAIN_RULES_BS: list[tuple[str, tuple[str, ...]]] = [
    ("cash", ("cash", "checking", "savings", "money market")),
    ("accounts receivable", ("accounts receivable", "trade receivable", "a r")),
    ("inventory", ("inventory",)),
    ("prepaid", ("prepaid",)),
    ("fixed asset", ("fixed asset", "equipment", "vehicle", "building", "land")),
    ("accumulated depreciation", ("accumulated depreciation",)),
    ("accounts payable", ("accounts payable", "a p")),
    ("credit card payable", ("credit card", "card payable")),
    ("payroll liabilities", ("payroll liabilit",)),
    ("sales tax payable", ("sales tax payable",)),
    ("loan payable", ("loan", "note payable")),
    ("line of credit", ("line of credit",)),
    ("retained earnings", ("retained earnings",)),
    ("owner equity", ("owner", "equity", "capital", "draw")),
    ("total equity", ("total equity",)),
]


def load_mapping_csv(path: Path) -> dict[str, str]:
    mapping: dict[str, str] = {}
    with path.open("r", encoding="utf-8-sig", newline="") as handle:
        lines = [line for line in handle if line.strip()]
        reader = csv.DictReader(lines)
        required = {"qbo_account_name", "template_label"}
        if not required.issubset(set(reader.fieldnames or [])):
            raise ValueError(f"Invalid mapping file: {path}. Expected columns: qbo_account_name, template_label")
        for row in reader:
            qbo_name_raw = row.get("qbo_account_name")
            qbo_name = normalize_text(qbo_name_raw)
            qbo_name_canonical = normalize_account_name(qbo_name_raw)
            template_label = (row.get("template_label") or "").strip()
            if qbo_name and template_label:
                mapping[qbo_name] = template_label
            if qbo_name_canonical and template_label:
                mapping[qbo_name_canonical] = template_label
    return mapping


def load_mapping_csvs(paths: list[Path]) -> dict[str, str]:
    merged: dict[str, str] = {}
    for path in paths:
        if not path.exists():
            continue
        merged.update(load_mapping_csv(path))
    return merged


def upsert_mapping_csv(path: Path, mapping_rows: dict[str, str]) -> int:
    if not mapping_rows:
        return 0

    existing: dict[str, str] = {}
    if path.exists():
        existing = load_mapping_csv(path)

    new_rows: list[tuple[str, str]] = []
    for qbo_account_name, template_label in mapping_rows.items():
        key = normalize_account_name(qbo_account_name)
        if not key:
            continue
        if key in existing:
            continue
        existing[key] = template_label
        new_rows.append((key, template_label))

    if not new_rows:
        return 0

    path.parent.mkdir(parents=True, exist_ok=True)
    write_header = not path.exists()
    with path.open("a", encoding="utf-8", newline="") as handle:
        writer = csv.writer(handle)
        if write_header:
            writer.writerow(["qbo_account_name", "template_label"])
        writer.writerows(new_rows)
    return len(new_rows)


def _score_domain_label(label: str, keywords: tuple[str, ...]) -> float:
    label_norm = normalize_text(label)
    score = 0.0
    for keyword in keywords:
        if keyword in label_norm:
            score += 1.0
    if score == 0:
        return 0.0
    return score / float(len(keywords))


def _apply_domain_rules(
    qbo_account_name: str,
    template_labels: list[str],
    statement_type: str | None,
) -> tuple[str, float, str] | None:
    name_norm = normalize_account_name(qbo_account_name)
    rule_set = DOMAIN_RULES_BS if statement_type == "bs" else DOMAIN_RULES_PL

    for trigger, keywords in rule_set:
        if trigger not in name_norm:
            continue
        candidates = sorted(
            ((label, _score_domain_label(label, keywords)) for label in template_labels),
            key=lambda item: item[1],
            reverse=True,
        )
        if not candidates:
            continue
        best_label, best_score = candidates[0]
        if best_score < 0.6:
            continue
        second_best = candidates[1][1] if len(candidates) > 1 else 0.0
        if best_score - second_best < 0.2:
            continue
        return best_label, min(0.95, 0.75 + best_score * 0.2), f"Mapped via auto-repair domain rule: {trigger}"
    return None


def evaluate_mapping(
    qbo_account_name: str,
    amount: float,
    static_mapping: dict[str, str],
    template_labels: list[str],
    min_confidence: float,
    statement_type: str | None = None,
) -> MappingDecision:
    normalized = normalize_text(qbo_account_name)
    canonical = normalize_account_name(qbo_account_name)

    mapped = static_mapping.get(normalized)
    if mapped:
        # Mapping table matches are explicit operator-approved mappings.
        confidence = 1.0
        if confidence >= min_confidence:
            return MappingDecision(
                qbo_account_name=qbo_account_name,
                qbo_amount=amount,
                mapped_template_label=mapped,
                confidence=confidence,
                reason="Matched via mapping file and passed confidence threshold.",
                suggestions=[(mapped, confidence)],
                mapping_source="mapping_exact",
            )

    mapped = static_mapping.get(canonical)
    if mapped:
        confidence = 1.0
        if confidence >= min_confidence:
            return MappingDecision(
                qbo_account_name=qbo_account_name,
                qbo_amount=amount,
                mapped_template_label=mapped,
                confidence=confidence,
                reason="Matched via canonicalized mapping key and passed confidence threshold.",
                suggestions=[(mapped, confidence)],
                mapping_source="mapping_canonical",
            )

    domain_rule = _apply_domain_rules(qbo_account_name, template_labels, statement_type)
    if domain_rule is not None:
        label, confidence, reason = domain_rule
        if confidence >= min_confidence:
            return MappingDecision(
                qbo_account_name=qbo_account_name,
                qbo_amount=amount,
                mapped_template_label=label,
                confidence=confidence,
                reason=reason,
                suggestions=[(label, confidence)],
                mapping_source="domain_rule",
            )

    suggestions = top_matches(qbo_account_name, template_labels, limit=5)
    if suggestions and suggestions[0][1] >= min_confidence:
        best_label, best_score = suggestions[0]
        return MappingDecision(
            qbo_account_name=qbo_account_name,
            qbo_amount=amount,
            mapped_template_label=best_label,
            confidence=best_score,
            reason="Mapped via fuzzy/semantic suggestion above threshold.",
            suggestions=suggestions,
            mapping_source="fuzzy",
        )

    return MappingDecision(
        qbo_account_name=qbo_account_name,
        qbo_amount=amount,
        mapped_template_label=None,
        confidence=suggestions[0][1] if suggestions else 0.0,
        reason="No suggestion met confidence threshold.",
        suggestions=suggestions,
        mapping_source="none",
    )
