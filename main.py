#!/usr/bin/env python3
"""
main.py

Terminal-based converter for SharePoint Excel to PDF.
Interactive CLI supports selecting subfolders, scanning all,
exporting checklist, and converting.
"""
import argparse
import logging
import sys
import tempfile
from pathlib import Path

import pythoncom
import yaml

from py_files.config import LOG_PATH, CONFIG_PATH, CHECKLIST_CSV
from py_files.sharepoint_gateway import SharePointGateway
from py_files.mock_sharepoint_gateway import MockSharePointGateway
from py_files.checklist import load_checklist, save_checklist
from py_files.excel_converter import ExcelConverter


def parse_args():
    parser = argparse.ArgumentParser(
        description="Terminal converter for SharePoint Excel to PDF"
    )
    parser.add_argument(
        "--mock-local",
        action="store_true",
        default=False,
        help="Use local mock instead of SharePoint"
    )
    parser.add_argument(
        "--export-checklist",
        metavar="CSV_PATH",
        help="Export the checklist to a CSV file and exit"
    )
    parser.add_argument(
        "folders",
        nargs="*",
        help="Server-relative folder URLs to process"
    )
    return parser.parse_args()


def setup_logging():
    logging.basicConfig(
        level=logging.INFO,
        format="%(asctime)s [%(levelname)s] %(message)s",
        handlers=[
            logging.FileHandler(LOG_PATH, encoding="utf-8"),
            logging.StreamHandler()
        ]
    )


def load_cfg() -> dict:
    """
    Read CONFIG_PATH (set by config.py). Exit if missing.
    """
    if not CONFIG_PATH.exists():
        sys.exit(
            f"[ERROR] Config file '{CONFIG_PATH}' was not found. "
            "Place config.yaml next to the executable or in your project root."
        )
    return yaml.safe_load(CONFIG_PATH.read_text())


def list_subfolders_with_stats(gateway):
    subs = gateway.list_immediate_subfolders()
    stats = []
    for idx, fld in enumerate(subs, start=1):
        rel = fld.serverRelativeUrl
        file_count = len(list(Path(rel).rglob("*_lease_leadpaint_xrf.pdf")))
        stats.append((idx, fld.name, rel, file_count))
    return stats


def scan_all_folders(gateway):
    stats = list_subfolders_with_stats(gateway)
    done_map = load_checklist()
    for _, _, rel, _ in stats:
        for pdf in Path(rel).rglob("*_lease_leadpaint_xrf.pdf"):
            parts = pdf.stem.replace("_lease_leadpaint_xrf", "").split("_")
            if len(parts) >= 2:
                done_map[f"{parts[0]}_{parts[1]}"] = True
    save_checklist(done_map)
    print("Scan complete. Checklist updated.")
    for _, name, rel, count in list_subfolders_with_stats(gateway):
        print(f"{name}: {count} completed")
    return done_map


def select_subfolders(gateway):
    done_map = load_checklist()
    while True:
        stats = list_subfolders_with_stats(gateway)
        if not stats:
            print("No subfolders found.")
            return [], done_map
        print("Folders:")
        for idx, name, rel, done in stats:
            print(f"  {idx}. {name} — {done} completed")
        print(
            "Please choose an option:\n"
            "  • Enter folder number(s) (e.g. 1 or 1,3,5) to convert those folders\n"
            "  • Type 'all'   to convert every folder\n"
            "  • Type 's'     to scan all folders and update progress\n"
            "  • Type 'e'     to export the checklist to a CSV file\n"
            "  • Type 'q'     to quit the program"
        )
        choice = input("Enter Choice: ").strip().lower()
        if choice == 'e':
            print("Enter full path, e.g. C:\\path\\to\\XRF_checklist.csv or ./my_checklist.csv")
            path = Path(input("Export checklist to (path or directory): ").strip())
            if path.is_dir():
                path = path / CHECKLIST_CSV.name
            path.parent.mkdir(parents=True, exist_ok=True)
            save_checklist(done_map, path)
            print(f"Checklist exported to {path}")
            continue
        if choice == 's':
            done_map = scan_all_folders(gateway)
            continue
        if choice == 'all':
            return [rel for _, _, rel, _ in stats], done_map
        if choice == 'q':
            return [], done_map
        rels = []
        for part in choice.split(','):
            try:
                idx = int(part)
                entry = next((rel for i, n, rel, d in stats if i == idx), None)
                if entry:
                    rels.append(entry)
                else:
                    print(f"{idx} out of range, ignored.")
            except ValueError:
                print(f"'{part}' invalid, ignored.")
        if rels:
            return rels, done_map
        print("No valid selection, try again.")


def convert_folder(rel_path: str, gateway, done_map):
    folder = Path(rel_path)
    out_dir = folder / "automation_output"
    out_dir.mkdir(parents=True, exist_ok=True)
    tmpdir = tempfile.TemporaryDirectory(prefix="lp_src_")
    sources = list(gateway.download_sources(rel_path, Path(tmpdir.name)))
    jobs = []
    for src in sources:
        prop, unit = ExcelConverter._extract_ids(src.stem)
        if prop and unit and not done_map.get(f"{prop}_{unit}"):
            jobs.append(src)
    if not jobs:
        print(f"No new files in {rel_path}")
        tmpdir.cleanup()
        return
    pythoncom.CoInitialize()
    conv = ExcelConverter(out_dir)
    conv.__enter__()
    try:
        print(f"Converting {len(jobs)} files in {rel_path}...")
        for i, src in enumerate(jobs, 1):
            print(f"[{i}/{len(jobs)}] {src.name} ... ", end="", flush=True)
            pdf = conv.convert(src)
            if pdf:
                key = Path(pdf).stem.replace("_lease_leadpaint_xrf", "")
                done_map[key] = True
                print("Done")
            else:
                print("Skipped")
    finally:
        conv.__exit__(None, None, None)
        pythoncom.CoUninitialize()
        tmpdir.cleanup()


def main():
    args = parse_args()
    setup_logging()

    # Load external or bundled config.yaml
    cfg = load_cfg()

    gateway = MockSharePointGateway(cfg) if args.mock_local else SharePointGateway(cfg)

    if args.export_checklist:
        dm = load_checklist()
        p = Path(args.export_checklist)
        if p.is_dir():
            p = p / CHECKLIST_CSV.name
        p.parent.mkdir(parents=True, exist_ok=True)
        save_checklist(dm, p)
        print(f"Checklist exported to {p}")
        return

    if args.folders:
        done_map = load_checklist()
        rels = args.folders
    else:
        rels, done_map = select_subfolders(gateway)
        if not rels:
            print("No folders selected, exiting.")
            return

    for rel in rels:
        convert_folder(rel, gateway, done_map)

    save_checklist(done_map)
    print("All done.")


if __name__ == "__main__":
    main()
