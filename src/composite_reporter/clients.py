from __future__ import annotations

import json
from dataclasses import dataclass
from pathlib import Path


@dataclass
class ClientProfile:
    client_id: str
    display_name: str
    base_dir: Path
    template_path: Path
    mapping_pl_path: Path
    mapping_bs_path: Path
    doctrine_path: Path | None = None
    calibration_path: Path | None = None
    confidence_threshold: float = 0.87
    tolerance: float = 1.0
    learned_confidence_threshold: float = 0.96


def _resolve_profile_path(base_dir: Path, raw_path: str) -> Path:
    path = Path(raw_path)
    if path.is_absolute():
        return path
    return (base_dir / path).resolve()


def load_client_profile(profile_path: Path) -> ClientProfile:
    payload = json.loads(profile_path.read_text(encoding="utf-8-sig"))
    base_dir = profile_path.parent

    client_id = str(payload.get("client_id") or base_dir.name).strip()
    display_name = str(payload.get("display_name") or client_id).strip()
    template_path = _resolve_profile_path(base_dir, str(payload["template_path"]))
    mapping_pl_path = _resolve_profile_path(base_dir, str(payload["mapping_pl_path"]))
    mapping_bs_path = _resolve_profile_path(base_dir, str(payload["mapping_bs_path"]))
    doctrine_raw = payload.get("doctrine_path")
    doctrine_path = _resolve_profile_path(base_dir, str(doctrine_raw)) if doctrine_raw else None
    calibration_raw = payload.get("calibration_path")
    calibration_path = _resolve_profile_path(base_dir, str(calibration_raw)) if calibration_raw else None

    return ClientProfile(
        client_id=client_id,
        display_name=display_name,
        base_dir=base_dir,
        template_path=template_path,
        mapping_pl_path=mapping_pl_path,
        mapping_bs_path=mapping_bs_path,
        doctrine_path=doctrine_path,
        calibration_path=calibration_path,
        confidence_threshold=float(payload.get("confidence_threshold", 0.87)),
        tolerance=float(payload.get("tolerance", 1.0)),
        learned_confidence_threshold=float(payload.get("learned_confidence_threshold", 0.96)),
    )


def list_client_profiles(clients_root: Path) -> list[ClientProfile]:
    profiles: list[ClientProfile] = []
    project_root = clients_root.parent

    default_template = project_root / "Composite Report Spreadsheet.xlsx"
    default_mapping_pl = project_root / "mapping_pl.csv"
    default_mapping_bs = project_root / "mapping_bs.csv"
    if default_template.exists() and default_mapping_pl.exists() and default_mapping_bs.exists():
        profiles.append(
            ClientProfile(
                client_id="default-root",
                display_name="Default Root Client",
                base_dir=project_root,
                template_path=default_template,
                mapping_pl_path=default_mapping_pl,
                mapping_bs_path=default_mapping_bs,
            )
        )

    if not clients_root.exists():
        return profiles

    for profile_path in sorted(clients_root.glob("*/profile.json")):
        profile = load_client_profile(profile_path)
        if profile.template_path.exists() and profile.mapping_pl_path.exists() and profile.mapping_bs_path.exists():
            profiles.append(profile)
    return profiles


def get_client_profile(clients_root: Path, client_id: str) -> ClientProfile:
    for profile in list_client_profiles(clients_root):
        if profile.client_id == client_id:
            return profile
    raise ValueError(f"Client profile not found: {client_id}")
