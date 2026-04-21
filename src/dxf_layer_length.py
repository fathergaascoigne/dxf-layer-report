
import argparse
import json
from pathlib import Path

from dxf_layer_report import (
    ScanOptions,
    create_temp_output_path,
    normalize_layers,
    open_in_default_spreadsheet,
    scan_dxf_files,
    write_results,
)


def load_config(config_path: Path) -> ScanOptions:
    with config_path.open("r", encoding="utf-8") as file:
        raw = json.load(file)

    return ScanOptions(
        source_path=Path(raw["source_path"]),
        target_layers=normalize_layers(raw.get("target_layers", ["14"])),
        unit_to_mm_factor=float(raw.get("unit_to_mm_factor", 1.0)),
        recursive=bool(raw.get("recursive", True)),
        workers=max(1, int(raw.get("workers", 4))),
        include_missing_layers=bool(raw.get("include_missing_layers", True)),
        include_read_errors=bool(raw.get("include_read_errors", True)),
        spline_segments=max(8, int(raw.get("spline_segments", 96))),
        ellipse_segments=max(8, int(raw.get("ellipse_segments", 96))),
    )


def parse_args() -> argparse.Namespace:
    parser = argparse.ArgumentParser(
        description="Measure one or more layers in a single DXF file or across all DXF files in a folder."
    )
    parser.add_argument("--config", required=True, type=Path, help="Path to the JSON config file.")
    parser.add_argument("--open", action="store_true", help="Open the generated workbook automatically.")
    return parser.parse_args()


def main() -> None:
    args = parse_args()
    options = load_config(args.config)

    results, total_files, stopped_early = scan_dxf_files(options)
    output_path = create_temp_output_path()
    write_results(output_path, options.target_layers, results)

    print(f"Files processed: {total_files}")
    print(f"Stopped early: {stopped_early}")
    print(f"Workbook: {output_path}")

    if args.open:
        open_in_default_spreadsheet(output_path)


if __name__ == "__main__":
    main()
