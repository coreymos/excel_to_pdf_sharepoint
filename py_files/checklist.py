# checklist.py
import csv
from pathlib import Path
import logging
from typing import Dict

from py_files.config import CHECKLIST_CSV


def load_checklist() -> Dict[str, bool]:
    """
    Load the checklist into a mapping keyed by "Property_Unit" with completion booleans.
    Supports both new (Property,Unit,Complete) CSV format and legacy (Folder,Complete) where
    we derive Property and Unit from the PDF filenames in each folder.
    """
    state: Dict[str, bool] = {}
    try:
        if not CHECKLIST_CSV.exists():
            return state
        with CHECKLIST_CSV.open(newline="", encoding="utf-8") as f:
            reader = csv.DictReader(f)
            headers = reader.fieldnames or []
            if set(["Property","Unit","Complete"]).issubset(headers):
                # New format: direct columns
                for row in reader:
                    prop = row.get("Property", "").strip()
                    unit = row.get("Unit", "").strip()
                    complete = row.get("Complete", "").strip().lower()
                    key = f"{prop}_{unit}"
                    state[key] = complete in ("x","yes","true","1","✓")
            elif set(["Folder","Complete"]).issubset(headers):
                # Legacy format: derive from PDF filenames within folder
                for row in reader:
                    folder = row.get("Folder", "").strip()
                    complete = row.get("Complete", "").strip().lower()
                    folder_path = Path(folder)
                    # Look for any matching PDF in the folder
                    for pdf in folder_path.glob("*_lease_leadpaint_xrf.pdf"):
                        stem = pdf.stem
                        # Remove suffix
                        prefix = stem.replace("_lease_leadpaint_xrf", "")
                        parts = prefix.split("_")
                        if len(parts) >= 2:
                            prop, unit = parts[0], parts[1]
                            key = f"{prop}_{unit}"
                            state[key] = complete in ("x","yes","true","1","✓")
            else:
                logging.warning("Checklist CSV has unexpected headers, ignoring file")
    except Exception as e:
        logging.warning(f"Error reading checklist CSV: {e}")
    return state


def save_checklist(state: Dict[str, bool], path: Path = None) -> None:
    """
    Save the checklist as CSV with columns Property,Unit,Complete.

    :param state: Mapping of "Property_Unit" keys to completion status.
    :param path: Optional Path to save the CSV; if None, use CHECKLIST_CSV.
    """
    out_path = path or CHECKLIST_CSV
    with out_path.open("w", newline="", encoding="utf-8") as f:
        writer = csv.DictWriter(f, fieldnames=["Property","Unit","Complete"])
        writer.writeheader()
        for key, done in sorted(state.items()):
            if "_" in key:
                prop, unit = key.split("_", 1)
            else:
                prop, unit = key, ""
            writer.writerow({"Property": prop, "Unit": unit, "Complete": "X" if done else ""})