#!/usr/bin/env python3
"""
converter_cli.py

Terminal-based converter for SharePoint Excel to PDF.
Interactive CLI supports selecting subfolders, resetting/exporting checklist, scanning, and converting.
"""
import argparse
import logging
import shutil
import tempfile
from pathlib import Path
import pythoncom
import yaml

from py_files.config import CONFIG_PATH, CHECKLIST_CSV
from py_files.sharepoint_gateway import SharePointGateway
from py_files.mock_sharepoint_gateway import MockSharePointGateway
from py_files.checklist import load_checklist, save_checklist, reset_checklist
from py_files.excel_converter import ExcelConverter


def parse_args():
    parser = argparse.ArgumentParser(
        description="Terminal converter for SharePoint Excel to PDF"
    )
    parser.add_argument(
        "--config",
        default=str(CONFIG_PATH),
        help="YAML config file"
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
        handlers=[logging.StreamHandler()]
    )


def list_subfolders_with_stats(gateway, done_map):
    subs = gateway.list_immediate_subfolders()
    stats = []
    for idx, fld in enumerate(subs, start=1):
        name = fld.name
        rel = fld.serverRelativeUrl
        prefix = f"{Path(rel).name}_"
        done_count = sum(1 for k,v in done_map.items() if k.startswith(prefix) and v)
        stats.append((idx, name, rel, done_count))
    return stats


def scan_all_folders(gateway, done_map):
    # Update done_map by scanning each folder for existing PDFs
    stats = list_subfolders_with_stats(gateway, done_map)
    for _, _, rel, _ in stats:
        for pdf in Path(rel).rglob("*_lease_leadpaint_xrf.pdf"):
            parts = pdf.stem.replace("_lease_leadpaint_xrf", "").split("_")
            if len(parts) >= 2:
                key = f"{parts[0]}_{parts[1]}"
                done_map[key] = True
    save_checklist(done_map)
    print("Scan complete. Checklist updated.")
    # show updated stats
    stats = list_subfolders_with_stats(gateway, done_map)
    for idx, name, rel, done in stats:
        print(f"  {name}: {done} completed")


def select_subfolders(gateway, done_map):
    """
    Interactive menu: select subfolders, reset, scan, export, convert, quit.
    """
    while True:
        stats = list_subfolders_with_stats(gateway, done_map)
        if not stats:
            print("No subfolders found.")
            return []
        print("Folders:")
        for idx, name, rel, done in stats:
            print(f"  {idx}. {name} — {done} completed")
        print("Options: [numbers] convert, 'all' convert all, 's' scan all, 'r' reset checklist, 'e' export checklist, 'q' quit")
        choice = input("Choice: ").strip().lower()
        if choice == 'r':
            reset_checklist()
            done_map.clear()
            print("Checklist reset.")
            continue
        if choice == 'e':
            path_str = input("Export checklist to (path or directory): ").strip()
            path = Path(path_str)
            if path.is_dir():
                path = path / CHECKLIST_CSV.name
            save_checklist(done_map, path)
            print(f"Checklist exported to {path}")
            continue
        if choice == 's':
            scan_all_folders(gateway, done_map)
            continue
        if choice == 'all':
            return [rel for _,_,rel,_ in stats]
        if choice == 'q':
            return []
        # parse comma-separated numbers
        rels = []
        for part in choice.split(','):
            try:
                i = int(part)
                entry = next((rel for idx,name,rel,done in stats if idx == i), None)
                if entry:
                    rels.append(entry)
                else:
                    print(f"{i} is out of range, ignored.")
            except ValueError:
                print(f"'{part}' not a valid option, ignored.")
        if rels:
            return rels
        print("No valid selection, try again.")

...
# rest of file unchanged
