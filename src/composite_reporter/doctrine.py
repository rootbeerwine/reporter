from __future__ import annotations

import csv
from dataclasses import dataclass
from pathlib import Path

from .utils import normalize_account_name


@dataclass
class DoctrineRule:
    template_label: str
    rule_type: str
    statement: str
    source_key: str
    source_section: str
    critical: bool


def load_doctrine(path: Path) -> list[DoctrineRule]:
    if not path.exists():
        return []
    rules: list[DoctrineRule] = []
    with path.open("r", encoding="utf-8-sig", newline="") as handle:
        reader = csv.DictReader(handle)
        for row in reader:
            template_label = (row.get("template_label") or "").strip()
            rule_type = (row.get("rule_type") or "").strip().lower()
            statement = (row.get("statement") or "").strip().lower()
            source_key = (row.get("source_key") or "").strip()
            source_section = (row.get("source_section") or "").strip().lower()
            critical = (row.get("critical") or "").strip().lower() in {"1", "true", "yes", "y"}
            if not template_label or not rule_type:
                continue
            rules.append(
                DoctrineRule(
                    template_label=template_label,
                    rule_type=rule_type,
                    statement=statement,
                    source_key=source_key,
                    source_section=source_section,
                    critical=critical,
                )
            )
    return rules


def apply_doctrine_rules(
    rules: list[DoctrineRule],
    *,
    pl_statement,
    bs_statement,
    section_amounts_by_label: dict[str, dict[str, float]],
) -> tuple[dict[str, float], list[str], list[str]]:
    overrides: dict[str, float] = {}
    warnings: list[str] = []
    critical_misses: list[str] = []

    for rule in rules:
        value = None
        if rule.rule_type == "direct_from_qbo_total":
            src = pl_statement.totals_reported if rule.statement == "pl" else bs_statement.totals_reported
            value = src.get(normalize_account_name(rule.source_key))

        elif rule.rule_type == "direct_from_account":
            src_lines = pl_statement.lines if rule.statement == "pl" else bs_statement.lines
            key = normalize_account_name(rule.source_key)
            section_key = normalize_account_name(rule.source_section)
            for line in src_lines:
                name = normalize_account_name(line.account_name)
                section = normalize_account_name(line.section or "")
                if key and key not in name:
                    continue
                if section_key and section_key not in section:
                    continue
                value = float(line.amount)
                break

        elif rule.rule_type == "section_sum":
            by_section = section_amounts_by_label.get(rule.template_label, {})
            value = by_section.get(rule.source_section, 0.0)

        elif rule.rule_type == "zero_if_missing":
            src_lines = pl_statement.lines if rule.statement == "pl" else bs_statement.lines
            key = normalize_account_name(rule.source_key)
            section_key = normalize_account_name(rule.source_section)
            found = None
            for line in src_lines:
                name = normalize_account_name(line.account_name)
                section = normalize_account_name(line.section or "")
                if key and key not in name:
                    continue
                if section_key and section_key not in section:
                    continue
                found = float(line.amount)
                break
            value = 0.0 if found is None else found

        if value is None:
            msg = f"Doctrine unresolved: {rule.template_label} ({rule.rule_type})"
            if rule.critical:
                critical_misses.append(msg)
            else:
                warnings.append(msg)
            continue

        overrides[rule.template_label] = round(float(value), 2)

    return overrides, warnings, critical_misses

