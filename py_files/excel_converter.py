#!/usr/bin/env python3
"""
excel_converter.py

Context-managed COM Excel converter that formats spreadsheets
and exports them as one-page-wide landscape PDFs, using DispatchEx
for unique Excel instances per process.

Optimized: disables UI, events, and switches to manual calculation to speed up.
Handles export errors gracefully to avoid crashing the worker pool.
"""
import logging
import re
import shutil
from pathlib import Path
from typing import Optional, Tuple
import pythoncom
import win32com.client as win32
from win32com.client import constants as xl

from py_files.config import (
    MARGIN_LEFT_RIGHT,
    MARGIN_TOP,
    MARGIN_BOTTOM,
    HEADER_MARGIN,
    FOOTER_MARGIN,
    STRIPE_RGB,
    VALID_UNIT_CODES,
)

logger = logging.getLogger(__name__)

class ExcelConverter:
    """
    Context manager to convert Excel/CSV files to formatted PDFs.
    Uses DispatchEx to spawn a new Excel COM instance in STA, then disables
    Excel overhead and sets manual calculation for speed.
    Handles export failures gracefully.
    """
    SUPPORTED_EXTENSIONS = (".xlsx", ".xls", ".csv")

    def __init__(self, output_dir: Path):
        self.output_dir = output_dir
        self.output_dir.mkdir(parents=True, exist_ok=True)
        self._excel = None

    def __enter__(self):
        pythoncom.CoInitialize()
        self._excel = win32.DispatchEx("Excel.Application")
        for attr in ("Visible", "ScreenUpdating", "DisplayAlerts", "EnableEvents", "AskToUpdateLinks"):  
            try:
                setattr(self._excel, attr, False)
            except Exception:
                pass
        try:
            self._excel.Calculation = xl.xlCalculationManual
        except Exception:
            pass
        return self

    def __exit__(self, exc_type, exc_val, exc_tb):
        if self._excel:
            try:
                self._excel.Calculation = xl.xlCalculationAutomatic
            except Exception:
                pass
            try:
                self._excel.Quit()
            except Exception:
                pass
        pythoncom.CoUninitialize()
        self._excel = None

    def convert(self, src: Path) -> Optional[Path]:
        """Convert a single spreadsheet to PDF."""
        try:
            wb = self._excel.Workbooks.Open(str(src))
        except Exception as e:
            logger.error("Failed to open %s: %s", src.name, e)
            return None
        try:
            ws = wb.Worksheets(1)
            if self._is_empty(ws):
                logger.info("Skipping empty workbook %s", src.name)
                return None

            prop, unit = self._extract_ids(src.stem)
            if not prop or not unit:
                logger.warning("Pattern not recognised for %s, skipping", src.name)
                return None

            self._format_sheet(ws)
            temp_pdf = src.with_suffix(".pdf")
            try:
                wb.ExportAsFixedFormat(0, str(temp_pdf))
            except Exception as ee:
                logger.error("Error exporting %s: %s", src.name, ee)
                return None
        finally:
            wb.Close(False)

        final_name = f"{prop}_{unit}_lease_leadpaint_xrf.pdf"
        final_path = self.output_dir / final_name
        try:
            temp_pdf.replace(final_path)
        except Exception:
            try:
                shutil.move(str(temp_pdf), str(final_path))
            except Exception as mv:
                logger.error("Failed to move PDF for %s: %s", src.name, mv)
                return None
        logger.info("Created %s", final_name)
        return final_path

    @staticmethod
    def _is_empty(ws) -> bool:
        ur = ws.UsedRange
        if ur.Rows.Count == 1 and ur.Columns.Count == 1:
            val = ur.Cells(1,1).Value
            return val is None or str(val).strip() == ""
        return False

    @staticmethod
    def _extract_ids(stem: str) -> Tuple[Optional[str], Optional[str]]:
        m = re.match(r'^([^-]+)', stem)
        prop = m.group(1) if m else None
        m2 = re.search(r'-([^-]+)-XRF', stem)
        unit = m2.group(1) if m2 else None
        if unit and unit.upper() not in VALID_UNIT_CODES:
            unit = None
        return prop, unit

    def _format_sheet(self, ws):
        ps = ws.PageSetup
        inch = ps.Application.InchesToPoints
        ps.LeftMargin = inch(MARGIN_LEFT_RIGHT)
        ps.RightMargin = inch(MARGIN_LEFT_RIGHT)
        ps.TopMargin = inch(MARGIN_TOP)
        ps.BottomMargin = inch(MARGIN_BOTTOM)
        ps.HeaderMargin = inch(HEADER_MARGIN)
        ps.FooterMargin = inch(FOOTER_MARGIN)
        ps.Orientation = xl.xlLandscape
        ps.CenterHorizontally = False
        ps.CenterVertically = False
        ps.FitToPagesWide = 1
        ps.FitToPagesTall = False
        ps.Zoom = False
        ps.PrintGridlines = False
        ps.PrintHeadings = False
        used = ws.UsedRange
        rows = used.Rows.Count
        cols = used.Columns.Count
        header_row = 1
        maxpop = 0
        for r in range(1, rows+1):
            cnt = sum(1 for c in range(1, cols+1)
                      if used.Cells(r,c).Value not in (None, ""))
            if cnt > maxpop:
                maxpop = cnt
                header_row = r
        ps.PrintTitleRows = f"${header_row}:${header_row}"
        ps.CenterFooter = "Page &P of &N"
        def col_letter(n: int) -> str:
            s = ''
            while n > 0:
                n, r = divmod(n-1, 26)
                s = chr(65+r) + s
            return s
        sheet_name = ws.Name
        if ' ' in sheet_name:
            sheet_name = f"'{sheet_name}'"
        start_col = col_letter(1)
        end_col = col_letter(cols)
        ps.PrintArea = f"{sheet_name}!${start_col}$1:${end_col}${rows}"
        ws.Rows(header_row).Font.Bold = True
        ws.Range(ws.Cells(header_row,1), ws.Cells(rows,cols)).Columns.AutoFit()
        bgr = STRIPE_RGB[0] | (STRIPE_RGB[1] << 8) | (STRIPE_RGB[2] << 16)
        for r in range(header_row+1, rows+1, 2):
            ws.Rows(r).Interior.Color = bgr
