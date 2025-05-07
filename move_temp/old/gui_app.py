#!/usr/bin/env python3
# gui_app.py

import threading
import tempfile
import shutil
from pathlib import Path
import tkinter as tk
from tkinter import ttk, messagebox
import pythoncom
from concurrent.futures import ProcessPoolExecutor, as_completed

from py_files.sharepoint_gateway import SharePointGateway
from py_files.mock_sharepoint_gateway import MockSharePointGateway
from py_files.excel_converter import ExcelConverter
from py_files.checklist import load_checklist, save_checklist

# Number of parallel Excel worker processes
WORKERS = 4

# Shared ExcelConverter instance per worker process
_converter_instance = None

def _worker_initializer():
    """
    Called once per worker process. Initializes COM and a single ExcelConverter.
    """
    pythoncom.CoInitialize()
    global _converter_instance
    _converter_instance = ExcelConverter(Path("."))
    _converter_instance.__enter__()


def _convert_job(src_path: str, out_dir: str):
    """
    Worker job: convert one file using the shared _converter_instance.
    Returns (src_path, pdf_path_or_None).
    """
    src = Path(src_path)
    outd = Path(out_dir)
    _converter_instance.output_dir = outd
    pdf = _converter_instance.convert(src)
    return src_path, str(pdf) if pdf else None


class FolderPickerApp(tk.Tk):
    """
    GUI for selecting SharePoint folders and converting Excel files to PDF in parallel.
    """
    def __init__(self, gateway):
        super().__init__()
        self.gateway = gateway
        self.done = load_checklist()
        self.title("Lead-Paint Report Converter")
        self.geometry("800x600")
        self.status_var = tk.StringVar()
        self._build_tree()
        self._build_buttons()
        ttk.Label(self, textvariable=self.status_var).pack(
            fill=tk.X, padx=10, pady=(0,10)
        )

    def _build_tree(self):
        self.tree = ttk.Treeview(
            self, columns=("Complete",), show=("tree","headings"), selectmode="extended"
        )
        self.tree.heading("#0", text="Folder")
        self.tree.column("#0", width=400)
        self.tree.heading("Complete", text="Done")
        self.tree.column("Complete", width=100, anchor="center")
        for fld in self.gateway.list_immediate_subfolders():
            rel = fld.serverRelativeUrl
            done_flag = any(
                key.startswith(f"{Path(rel).name}_") and val
                for key,val in self.done.items()
            )
            self.tree.insert("", "end", iid=rel, text=fld.name,
                             values=("✓" if done_flag else "",))
        self.tree.pack(fill=tk.BOTH, expand=True, padx=10, pady=10)

    def _build_buttons(self):
        frm = ttk.Frame(self)
        frm.pack(pady=5)
        self.convert_btn = ttk.Button(
            frm, text="Convert Selected", command=self._convert
        )
        self.convert_btn.pack(side=tk.LEFT, padx=5)
        ttk.Button(frm, text="Scan Checklist", command=self._scan).pack(side=tk.LEFT, padx=5)
        ttk.Button(frm, text="Export Checklist", command=self._export).pack(side=tk.LEFT, padx=5)
        ttk.Button(frm, text="Quit", command=self.destroy).pack(side=tk.LEFT, padx=5)

    def _scan(self):
        """
        Manual scan: pick up any already‐converted PDFs and update checklist.
        """
        initial = len(self.done)
        items = self.tree.selection() or self.tree.get_children()
        for rel in items:
            for pdf in Path(rel).rglob("*_lease_leadpaint_xrf.pdf"):
                parts = pdf.stem.replace("_lease_leadpaint_xrf","").split("_")
                if len(parts) >= 2:
                    key = f"{parts[0]}_{parts[1]}"
                    self.done[key] = True
        save_checklist(self.done)
        added = len(self.done) - initial
        messagebox.showinfo(
            "Scanned",
            f"Checklist updated: {added} new entr{'y' if added==1 else 'ies'}."
        )

    def _export(self):
        """
        Export the current checklist to CSV.
        """
        from tkinter import filedialog
        path = filedialog.asksaveasfilename(
            defaultextension=".csv",
            filetypes=[("CSV Files","*.csv")],
            initialfile="XRF_checklist.csv",
            title="Save Checklist As"
        )
        if path:
            save_checklist(self.done, Path(path))
            messagebox.showinfo("Saved", f"Checklist exported to {path}")

    def _convert(self):
        """
        Called when the user hits "Convert Selected".
        1) Quick implicit scan of existing PDFs.
        2) Build job list skipping already-done items.
        3) Fire off the worker pool.
        """
        sel = self.tree.selection()
        if not sel:
            messagebox.showinfo("Nothing to do","No folders selected.")
            return

        # 1) implicit scan
        for rel in sel:
            for pdf in Path(rel).rglob("*_lease_leadpaint_xrf.pdf"):
                parts = pdf.stem.replace("_lease_leadpaint_xrf","").split("_")
                if len(parts) >= 2:
                    key = f"{parts[0]}_{parts[1]}"
                    self.done[key] = True
        save_checklist(self.done)

        # 2) refresh tree checkmarks
        for rel in sel:
            done_flag = any(
                key.startswith(f"{Path(rel).name}_") and val
                for key,val in self.done.items()
            )
            self.tree.set(rel, "Complete", "✓" if done_flag else "")

        # disable re-entry
        self.convert_btn.config(state=tk.DISABLED)

        # popup & start thread
        self._make_progress_window()
        threading.Thread(
            target=self._do_convert,
            args=(list(sel),),
            daemon=True
        ).start()

    def _make_progress_window(self):
        self.prog_win = tk.Toplevel(self)
        self.prog_win.title("Converting…")
        self.prog_win.geometry("350x120")
        self.prog_win.resizable(False, False)
        self.prog_label = ttk.Label(self.prog_win, text="Starting conversion...")
        self.prog_label.pack(padx=20, pady=(20,5))
        # start in indeterminate while we figure out total
        self.progbar = ttk.Progressbar(self.prog_win, mode="indeterminate")
        self.progbar.pack(fill=tk.X, padx=20, pady=10)
        self.progbar.start(50)

    def _init_progress(self, total):
        """
        Switch the bar to determinate mode once we know how many jobs there are.
        """
        self.progbar.stop()
        self.progbar.config(mode="determinate", maximum=total, value=0)

    def _do_convert(self, sel):
        # build the list of (src,outdir) jobs
        jobs = []
        temp_dirs = []
        for rel in sel:
            folder = Path(rel)
            out_dir = folder/"automation_output"
            out_dir.mkdir(parents=True, exist_ok=True)
            tmp = Path(tempfile.mkdtemp(prefix="lp_src_"))
            temp_dirs.append(tmp)
            for src in self.gateway.download_sources(rel, tmp):
                prop,unit = ExcelConverter._extract_ids(src.stem)
                key = f"{prop}_{unit}" if prop and unit else None
                if key and not self.done.get(key):
                    jobs.append((str(src), str(out_dir)))

        total = len(jobs)
        if total == 0:
            # nothing new: tear down and re-enable
            self.after(0, self._cancel_convert, temp_dirs)
            return

        # switch to determinate bar
        self.after(0, self._init_progress, total)

        done_count = 0
        with ProcessPoolExecutor(
            max_workers=WORKERS,
            initializer=_worker_initializer
        ) as executor:
            futures = {
                executor.submit(_convert_job, src, out): (src,out)
                for src,out in jobs
            }
            for future in as_completed(futures):
                src_str, pdf_str = future.result()
                done_count += 1
                # UI update must happen on mainloop
                self.after(
                    0,
                    self._update_progress,
                    done_count, total, src_str
                )
                if pdf_str:
                    stem = Path(pdf_str).stem.replace("_lease_leadpaint_xrf","")
                    self.done[stem] = True

        # all done
        self.after(0, self._finish_convert, sel, temp_dirs)

    def _update_progress(self, done_count, total, filename):
        """
        Update label + determinate progressbar.
        """
        self.prog_label.config(
            text=f"Converted {done_count}/{total}: {Path(filename).name}"
        )
        self.progbar['value'] = done_count

    def _cancel_convert(self, temp_dirs):
        """
        Called if there was nothing to do.
        """
        for d in temp_dirs:
            shutil.rmtree(d, ignore_errors=True)
        self.progbar.stop()
        self.prog_win.destroy()
        self.convert_btn.config(state=tk.NORMAL)
        messagebox.showinfo("Nothing to do","No new files to convert.")

    def _finish_convert(self, sel, temp_dirs):
        """
        Clean up after a full run.
        """
        for d in temp_dirs:
            shutil.rmtree(d, ignore_errors=True)

        # re‐mark the tree
        for rel in sel:
            done_flag = any(
                key.startswith(f"{Path(rel).name}_") and val
                for key,val in self.done.items()
            )
            self.tree.set(rel,"Complete","✓" if done_flag else "")

        save_checklist(self.done)
        self.progbar.stop()
        self.prog_win.destroy()
        self.convert_btn.config(state=tk.NORMAL)
        messagebox.showinfo("Done","All selected folders processed.")