# This is a sample Python script.

# Press Shift+F10 to execute it or replace it with your code.
# Press Double Shift to search everywhere for classes, files, tool windows, actions, and settings.
"""
build_master_ud_tables_threaded.py

- GUI tool to combine UDxx_*.xlsx files (one-level subfolders) into prefix + UDxx.csv
- File-based progress bar, cancel button, threaded worker.
- Requires: pandas, openpyxl
    python -m pip install pandas openpyxl
"""

import tkinter as tk
from tkinter import ttk, filedialog, messagebox
from pathlib import Path
import pandas as pd
import datetime
import threading
import queue
import traceback

# ---------- Configuration ----------
UD_TABLES = ["UD01", "UD02", "UD03", "UD05", "UD06", "UD07", "UD08"]
EXCEL_EXT = ".xlsx"
LOG_FILENAME = "MasterUD_Log.txt"

# ---------- Helpers ----------
def now_str():
    return datetime.datetime.now().strftime("%Y-%m-%d %H:%M:%S")

def log_write(log_path: Path, msg: str):
    try:
        log_path.parent.mkdir(parents=True, exist_ok=True)
        with open(log_path, "a", encoding="utf-8") as f:
            f.write(f"{now_str()} - {msg}\n")
    except Exception:
        # can't do much if logging fails
        pass

# ---------- Worker logic (runs in background thread) ----------
def count_matching_files(input_folder: Path) -> int:
    total = 0
    for sub in input_folder.iterdir():
        if not sub.is_dir():
            continue
        for f in sub.iterdir():
            if not f.is_file():
                continue
            if f.suffix.lower() != EXCEL_EXT:
                continue
            if "MASTER" in f.name.upper():
                continue
            # Only process files containing a UD table anywhere in the filename
            name = f.name.upper()
            for ud in UD_TABLES:
                if ud in name and "MASTER" not in name and (name.endswith(".XLSX") or name.endswith(".CSV")):
                    # Process the file here
                    total += 1
                    # Optionally, break if you only want to count once per file
                    break
    return total

def worker_process(input_folder: Path, output_folder: Path, prefix: str,
                   status_q: queue.Queue, cancel_event: threading.Event):
    """
    Worker that does the scanning + combining. Communicates via status_q.
    status_q receives tuples: ("status", text) / ("progress", processed, total) / ("done", msg) / ("error", msg)
    """
    log_path = output_folder / LOG_FILENAME
    log_write(log_path, f"START - Input: {input_folder}  Output: {output_folder}  Prefix: '{prefix}'")
    status_q.put(("status", "Counting files..."))
    try:
        total_files = count_matching_files(input_folder)
    except Exception as e:
        status_q.put(("error", f"Failed counting files: {e}"))
        log_write(log_path, f"ERROR counting files: {e}")
        status_q.put(("done", "Failed"))
        return

    if total_files == 0:
        log_write(log_path, "No matching files found.")
        status_q.put(("status", "No files found."))
        status_q.put(("done", "No files"))
        return

    status_q.put(("status", f"Found {total_files} files. Starting..."))
    processed = 0

    # Ensure output folder exists
    output_folder.mkdir(parents=True, exist_ok=True)

    # Process per UD table, streaming output
    for ud in UD_TABLES:
        if cancel_event.is_set():
            status_q.put(("status", "Cancelled"))
            log_write(log_path, "Cancelled by user.")
            status_q.put(("done", "Cancelled"))
            return

        status_q.put(("status", f"Preparing {ud}..."))
        header_written = False
        header_cols = None
        out_filename = f"{prefix}{ud}.csv"
        out_path = output_folder / out_filename

        # If out_path exists from previous run, remove it to start fresh
        try:
            if out_path.exists():
                out_path.unlink()
        except Exception as e:
            log_write(log_path, f"Warning: couldn't remove existing {out_path}: {e}")

        # Iterate subfolders
        for sub in sorted(input_folder.iterdir()):
            if cancel_event.is_set():
                status_q.put(("status", "Cancelled"))
                log_write(log_path, "Cancelled by user.")
                status_q.put(("done", "Cancelled"))
                return

            if not sub.is_dir():
                continue

            for f in sorted(sub.iterdir()):
                if cancel_event.i s_set():
                    status_q.put(("status", "Cancelling..."))
                    log_write(log_path, "Cancelled by user.")
                    status_q.put(("done", "Cancelled"))
                    return

                if not f.is_file():
                    continue

                # Allow .xlsx OR .csv
                if f.suffix.lower() not in (".xlsx", ".csv"):
                    continue

                if "MASTER" in f.name.upper():
                    continue

                name_upper = f.name.upper()

                # ✔ Match UDxx anywhere in filename (not only at start)
                if ud not in name_upper:
                    continue

                # Update UI: current file
                status_q.put(("status", f"Processing {ud}: {f.name}"))

                # Read file (unchanged)
                try:
                    df = pd.read_excel(f, engine="openpyxl", sheet_name=0, header=0, dtype=object)
                except Exception as e:
                    log_write(log_path, f"ERROR reading {f}: {e}")
                    status_q.put(("log", f"ERROR reading {f}: {e}"))
                    processed += 1
                    status_q.put(("progress", processed, total_files))
                    continue

                # If file has no data rows under header
                if df.shape[0] < 1:
                    log_write(log_path, f"Skipped (no data rows): {f}")
                    status_q.put(("log", f"Skipped (no data rows): {f}"))
                    processed += 1
                    status_q.put(("progress", processed, total_files))
                    continue

                # Normalize columns to header from first file
                if header_cols is None:
                    header_cols = list(df.columns)

                # Align columns: ensure df has header_cols; add missing as NaN; reorder
                if list(df.columns) != header_cols:
                    # Create a DataFrame with header_cols and fill from df where possible
                    new_df = pd.DataFrame(columns=header_cols)
                    for c in df.columns:
                        if c in header_cols:
                            new_df[c] = df[c]
                    # ensure all header cols exist
                    for c in header_cols:
                        if c not in new_df.columns:
                            new_df[c] = pd.NA
                    new_df = new_df[header_cols]
                    df_to_write = new_df
                else:
                    df_to_write = df[header_cols]

                # Write header if not yet written; then append rows
                try:
                    if not header_written:
                        # write with header
                        df_to_write.to_csv(out_path, index=False, header=True, encoding="utf-8")
                        header_written = True
                    else:
                        # append without header
                        df_to_write.to_csv(out_path, index=False, header=False, mode="a", encoding="utf-8")
                except Exception as e:
                    log_write(log_path, f"ERROR writing to {out_path}: {e}")
                    status_q.put(("log", f"ERROR writing to {out_path}: {e}"))

                processed += 1
                status_q.put(("progress", processed, total_files))

        # After finishing this UD, if no file produced, log it
        if not out_path.exists():
            log_write(log_path, f"No valid files found for {ud} (no output produced)")

    # Done all UD tables
    status_q.put(("status", "All done"))
    log_write(log_path, f"COMPLETE - processed {processed} files.")
    status_q.put(("done", "Complete"))

# ---------- GUI App ----------
class App(tk.Tk):
    def __init__(self):
        super().__init__()
        self.title("Build Master UD Tables (CSV) - Threaded")
        self.geometry("700x280")
        self.resizable(False, False)

        # Threading / queue
        self.status_q = queue.Queue()
        self.cancel_event = threading.Event()
        self.worker_thread = None

        # UI
        frm = ttk.Frame(self, padding=12)
        frm.pack(fill="both", expand=True)

        # Input folder
        ttk.Label(frm, text="Input (parent) folder:").grid(column=0, row=0, sticky="w")
        self.input_var = tk.StringVar()
        self.input_entry = ttk.Entry(frm, textvariable=self.input_var, width=72)
        self.input_entry.grid(column=0, row=1, columnspan=3, sticky="w")
        ttk.Button(frm, text="Browse...", command=self.browse_input).grid(column=3, row=1, sticky="e")

        # Output folder
        ttk.Label(frm, text="Output folder (Master UD Tables):").grid(column=0, row=2, sticky="w", pady=(8,0))
        self.output_var = tk.StringVar()
        self.output_entry = ttk.Entry(frm, textvariable=self.output_var, width=72)
        self.output_entry.grid(column=0, row=3, columnspan=3, sticky="w")
        ttk.Button(frm, text="Browse...", command=self.browse_output).grid(column=3, row=3, sticky="e")

        # Prefix
        ttk.Label(frm, text="Output filename prefix:").grid(column=0, row=4, sticky="w", pady=(8,0))
        self.prefix_var = tk.StringVar(value="AUGUST - ")
        self.prefix_entry = ttk.Entry(frm, textvariable=self.prefix_var, width=40)
        self.prefix_entry.grid(column=0, row=5, sticky="w")

        # Buttons: Run / Cancel / Close
        self.run_button = ttk.Button(frm, text="Run", command=self.on_run)
        self.run_button.grid(column=0, row=6, pady=(12,0), sticky="w")
        self.cancel_button = ttk.Button(frm, text="Cancel", command=self.on_cancel, state="disabled")
        self.cancel_button.grid(column=1, row=6, pady=(12,0), sticky="w")
        self.close_button = ttk.Button(frm, text="Close", command=self.on_close)
        self.close_button.grid(column=2, row=6, pady=(12,0), sticky="w")

        # Progress bar and percent
        self.prg = ttk.Progressbar(frm, orient="horizontal", length=480, mode="determinate")
        self.prg.grid(column=0, row=7, columnspan=3, pady=(12,0), sticky="w")
        self.percent_label = ttk.Label(frm, text="0%")
        self.percent_label.grid(column=3, row=7, sticky="w")

        # Status labels
        self.status_label = ttk.Label(frm, text="Ready", relief="sunken", anchor="w", width=100)
        self.status_label.grid(column=0, row=8, columnspan=4, pady=(12,0), sticky="we")
        self.log_short = tk.StringVar(value="")
        self.log_label = ttk.Label(frm, textvariable=self.log_short, anchor="w", width=100)
        self.log_label.grid(column=0, row=9, columnspan=4, pady=(6,0), sticky="we")

        for c in range(4):
            frm.columnconfigure(c, weight=1)

        # Poll queue
        self.after(150, self.process_queue)

    def browse_input(self):
        folder = filedialog.askdirectory(title="Select parent folder containing product subfolders")
        if folder:
            self.input_var.set(folder)
            if not self.output_var.get():
                self.output_var.set(str(Path(folder) / "Master UD Tables"))

    def browse_output(self):
        folder = filedialog.askdirectory(title="Select output folder for master tables")
        if folder:
            self.output_var.set(folder)

    def on_run(self):
        input_path = self.input_var.get().strip()
        output_path = self.output_var.get().strip()
        prefix = self.prefix_var.get()

        if not input_path:
            messagebox.showerror("Error", "Please select an input folder.")
            return

        if not output_path:
            output_path = str(Path(input_path) / "Master UD Tables")
            self.output_var.set(output_path)

        input_folder = Path(input_path)
        output_folder = Path(output_path)

        # Quick sanity checks
        if not input_folder.exists() or not input_folder.is_dir():
            messagebox.showerror("Error", "Input folder does not exist or is not a folder.")
            return

        # Disable UI controls
        self.run_button.config(state="disabled")
        self.cancel_button.config(state="normal")
        self.cancel_event.clear()
        self.prg['value'] = 0
        self.percent_label.config(text="0%")
        self.status_label.config(text="Counting files...")
        self.log_short.set("")

        # Start worker thread
        self.worker_thread = threading.Thread(
            target=worker_process,
            args=(input_folder, output_folder, prefix, self.status_q, self.cancel_event),
            daemon=True
        )
        self.worker_thread.start()

    def on_cancel(self):
        if messagebox.askyesno("Cancel", "Are you sure you want to cancel the run?"):
            self.cancel_event.set()
            self.status_label.config(text="Cancel requested...")

    def on_close(self):
        if self.worker_thread and self.worker_thread.is_alive():
            if not messagebox.askyesno("Quit", "A process is running. Quit anyway?"):
                return
            # set cancel and give thread a moment
            self.cancel_event.set()
        self.destroy()

    def process_queue(self):
        """
        Poll status queue and update UI.
        """
        try:
            while True:
                item = self.status_q.get_nowait()
                if not item:
                    continue
                tag = item[0]
                if tag == "status":
                    self.status_label.config(text=item[1])
                elif tag == "progress":
                    processed, total = item[1], item[2]
                    # init prg max if needed
                    if self.prg['maximum'] != total:
                        self.prg.config(maximum=total)
                    self.prg['value'] = processed
                    pct = int((processed / total) * 100) if total else 0
                    self.percent_label.config(text=f"{pct}%")
                    self.log_short.set(f"{processed} / {total}")
                elif tag == "log":
                    # show last log-ish message in small area
                    self.log_short.set(item[1])
                elif tag == "done":
                    # finished or cancelled
                    self.status_label.config(text=item[1])
                    self.run_button.config(state="normal")
                    self.cancel_button.config(state="disabled")
                elif tag == "error":
                    messagebox.showerror("Error", item[1])
                    self.run_button.config(state="normal")
                    self.cancel_button.config(state="disabled")
                else:
                    # unknown tag - ignore
                    pass
                self.status_q.task_done()
        except queue.Empty:
            pass
        # re-schedule
        self.after(150, self.process_queue)

def main():
    app = App()
    app.mainloop()

if __name__ == "__main__":
    main()
