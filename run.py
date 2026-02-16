from __future__ import annotations

import argparse
import fnmatch
import os
import subprocess
import sys
from pathlib import Path

ROOT = Path(__file__).resolve().parent
SRC = ROOT / "src"
if str(SRC) not in sys.path:
    sys.path.insert(0, str(SRC))

from composite_reporter.clients import get_client_profile, list_client_profiles
from composite_reporter.pipeline import RunConfig, run_pipeline


def _find_one(base: Path, candidates: list[str]) -> Path | None:
    for candidate in candidates:
        direct = base / candidate
        if direct.exists():
            return direct

    excluded_dirs = {".tmp_tests", "tests", "out", "out_real", "out_synth", ".venv", "__pycache__"}
    for root, dirs, files in os.walk(base):
        dirs[:] = [d for d in dirs if d not in excluded_dirs and not d.startswith(".git")]
        for pattern in candidates:
            for filename in files:
                if fnmatch.fnmatch(filename, pattern):
                    return Path(root) / filename
    return None


def main() -> int:
    parser = argparse.ArgumentParser(description="Run self-tests then execute composite reporter pipeline.")
    parser.add_argument("--skip-tests", action="store_true", help="Skip running python -m pytest -q before pipeline.")
    parser.add_argument("--client", help="Client ID from clients/<client_id>/profile.json")
    parser.add_argument("--clients-root", default="clients", help="Client profiles root directory")
    args = parser.parse_args()

    base = Path.cwd()

    if not args.skip_tests:
        result = subprocess.run([sys.executable, "-m", "pytest", "-q"], cwd=base)
        if result.returncode != 0:
            print("ERROR: Tests failed. Fix tests before running pipeline.")
            return result.returncode

    pl = _find_one(base, ["*Profit+and+Loss*.xlsx", "*Profit*Loss*.xlsx", "*P&L*.xlsx"])
    bs = _find_one(base, ["*Balance+Sheet*.xlsx", "*Balance*Sheet*.xlsx"])

    clients_root = base / args.clients_root
    if args.client:
        profile = get_client_profile(clients_root, args.client)
    else:
        profiles = list_client_profiles(clients_root)
        profile = profiles[0] if len(profiles) == 1 else None

    template = profile.template_path if profile else _find_one(base, ["Composite Report Spreadsheet.xlsx"])
    mapping_pl = profile.mapping_pl_path if profile else _find_one(base, ["mapping_pl.csv"])
    mapping_bs = profile.mapping_bs_path if profile else _find_one(base, ["mapping_bs.csv"])
    learned_pl = (profile.base_dir / "mapping_pl.learned.csv") if profile else None
    learned_bs = (profile.base_dir / "mapping_bs.learned.csv") if profile else None

    missing = []
    if not template:
        missing.append("Composite Report Spreadsheet.xlsx")
    if not mapping_pl:
        missing.append("mapping_pl.csv")
    if not mapping_bs:
        missing.append("mapping_bs.csv")
    if not pl:
        missing.append("one QBO Profit & Loss export .xlsx")
    if not bs:
        missing.append("one QBO Balance Sheet export .xlsx")

    if missing:
        print("ERROR: Missing required files for hands-off run.")
        print("Expected filenames:")
        print("- Composite Report Spreadsheet.xlsx")
        print("- mapping_pl.csv")
        print("- mapping_bs.csv")
        print("- one QBO Profit & Loss export .xlsx")
        print("- one QBO Balance Sheet export .xlsx")
        if profile:
            print(f"- using client profile: {profile.client_id}")
        elif not args.client and len(list_client_profiles(clients_root)) > 1:
            print("- multiple client profiles found; pass --client <client_id>")
        print("Missing now:")
        for item in missing:
            print(f"- {item}")
        return 2

    outdir = base / "out"
    tieout = run_pipeline(
        RunConfig(
            pl_path=pl,
            bs_path=bs,
            template_path=template,
            mapping_pl_path=mapping_pl,
            mapping_bs_path=mapping_bs,
            outdir=outdir,
            confidence_threshold=profile.confidence_threshold if profile else 0.85,
            tolerance=profile.tolerance if profile else 1.0,
            client_id=profile.client_id if profile else None,
            mapping_pl_extra_paths=[learned_pl] if learned_pl else [],
            mapping_bs_extra_paths=[learned_bs] if learned_bs else [],
            learned_mapping_pl_path=learned_pl,
            learned_mapping_bs_path=learned_bs,
            learned_confidence_threshold=profile.learned_confidence_threshold if profile else 0.96,
            calibration_path=profile.calibration_path if profile else None,
        )
    )

    print(f"Outputs written to: {outdir}")
    print(
        f"status={tieout['status']} filled_labels={tieout['filled_labels_count']} "
        f"unmapped_pl={tieout['unmapped_counts']['pl']} unmapped_bs={tieout['unmapped_counts']['bs']}"
    )
    return 0 if tieout["status"] == "PASSED" else 1


if __name__ == "__main__":
    raise SystemExit(main())
