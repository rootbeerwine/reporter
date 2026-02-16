from __future__ import annotations

import argparse
from pathlib import Path

from .clients import get_client_profile
from .pipeline import RunConfig, run_pipeline


def build_parser() -> argparse.ArgumentParser:
    parser = argparse.ArgumentParser(description="Fill coaching composite report from QBO exports.")
    parser.add_argument("--pl", help="Path to QBO Profit & Loss export (.xlsx)")
    parser.add_argument("--bs", help="Path to QBO Balance Sheet export (.xlsx)")
    parser.add_argument("--template", help="Path to Composite Report Spreadsheet.xlsx")
    parser.add_argument("--mapping-pl", default="mapping_pl.csv", help="Path to mapping_pl.csv")
    parser.add_argument("--mapping-bs", default="mapping_bs.csv", help="Path to mapping_bs.csv")
    parser.add_argument("--client", help="Client ID from clients/<client_id>/profile.json")
    parser.add_argument("--clients-root", default="clients", help="Client profiles root directory")
    parser.add_argument("--outdir", default="out", help="Output directory")
    parser.add_argument("--confidence-threshold", type=float, default=0.85)
    parser.add_argument("--tolerance", type=float, default=1.0)
    return parser


def main(argv: list[str] | None = None) -> int:
    args = build_parser().parse_args(argv)
    if not args.pl or not args.bs:
        raise ValueError("Both --pl and --bs are required.")

    client_profile = None
    template_path = Path(args.template) if args.template else None
    mapping_pl_path = Path(args.mapping_pl)
    mapping_bs_path = Path(args.mapping_bs)
    confidence_threshold = float(args.confidence_threshold)
    tolerance = float(args.tolerance)
    mapping_pl_extra_paths: list[Path] = []
    mapping_bs_extra_paths: list[Path] = []
    learned_mapping_pl_path = None
    learned_mapping_bs_path = None
    learned_confidence_threshold = 0.96

    if args.client:
        client_profile = get_client_profile(Path(args.clients_root), args.client)
        template_path = client_profile.template_path
        mapping_pl_path = client_profile.mapping_pl_path
        mapping_bs_path = client_profile.mapping_bs_path
        confidence_threshold = client_profile.confidence_threshold
        tolerance = client_profile.tolerance
        learned_mapping_pl_path = client_profile.base_dir / "mapping_pl.learned.csv"
        learned_mapping_bs_path = client_profile.base_dir / "mapping_bs.learned.csv"
        mapping_pl_extra_paths = [learned_mapping_pl_path]
        mapping_bs_extra_paths = [learned_mapping_bs_path]
        learned_confidence_threshold = client_profile.learned_confidence_threshold

    if template_path is None:
        raise ValueError("Template path is required when --client is not provided.")

    config = RunConfig(
        pl_path=Path(args.pl),
        bs_path=Path(args.bs),
        template_path=template_path,
        mapping_pl_path=mapping_pl_path,
        mapping_bs_path=mapping_bs_path,
        outdir=Path(args.outdir),
        confidence_threshold=confidence_threshold,
        tolerance=tolerance,
        client_id=client_profile.client_id if client_profile else None,
        mapping_pl_extra_paths=mapping_pl_extra_paths,
        mapping_bs_extra_paths=mapping_bs_extra_paths,
        learned_mapping_pl_path=learned_mapping_pl_path,
        learned_mapping_bs_path=learned_mapping_bs_path,
        learned_confidence_threshold=learned_confidence_threshold,
        doctrine_path=client_profile.doctrine_path,
        calibration_path=client_profile.calibration_path if client_profile else None,
    )
    tieout = run_pipeline(config)
    print(
        f"status={tieout['status']} filled_labels={tieout['filled_labels_count']} "
        f"unmapped_pl={tieout['unmapped_counts']['pl']} unmapped_bs={tieout['unmapped_counts']['bs']}"
    )
    return 0 if tieout["status"] == "PASSED" else 1


if __name__ == "__main__":
    raise SystemExit(main())
