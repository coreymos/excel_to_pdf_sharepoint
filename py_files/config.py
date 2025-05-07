from pathlib import Path
import csv
import sys

# --------------------------------------------------------------------------- #
# CONFIG_PATH: external override next to the exe, else bundled, else dev
# --------------------------------------------------------------------------- #
if getattr(sys, "frozen", False):
    # 1) external config alongside the exe?
    exe_dir = Path(sys.executable).parent
    ext = exe_dir / "config.yaml"
    if ext.is_file():
        CONFIG_PATH = ext
    else:
        # 2) fallback to the one embedded by PyInstaller
        CONFIG_PATH = Path(sys._MEIPASS) / "config.yaml"
else:
    # dev mode: look in your project root
    CONFIG_PATH = Path(__file__).parent.parent / "config.yaml"


# --------------------------------------------------------------------------- #
# Other shared paths / constants
# --------------------------------------------------------------------------- #
LOG_PATH      = Path("excel_converter.log")
CHECKLIST_CSV = Path("XRF_checklist.csv")

PUBLIC_GRAPH_CLIENT_ID = "04f0c124-f2bc-4f7a-ac24-a29dd5d43626"

MARGIN_LEFT_RIGHT = 0.25
MARGIN_TOP        = 0.50
MARGIN_BOTTOM     = 0.55
HEADER_MARGIN     = 0.30
FOOTER_MARGIN     = 0.30
STRIPE_RGB        = (242, 242, 242)


# --------------------------------------------------------------------------- #
# Helper: read valid unit codes once at import time
# --------------------------------------------------------------------------- #
def _load_valid_units() -> set[str]:
    units: set[str] = set()
    try:
        if CHECKLIST_CSV.exists():
            with CHECKLIST_CSV.open(newline="", encoding="utf-8") as f:
                reader = csv.DictReader(f)
                if reader.fieldnames and "Unit" in reader.fieldnames:
                    for row in reader:
                        unit = row.get("Unit", "").strip()
                        if unit:
                            units.add(unit.upper())
    except Exception:
        pass
    return units

VALID_UNIT_CODES: set[str] = _load_valid_units()
