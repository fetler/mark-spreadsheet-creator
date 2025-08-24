# helpers/dnd_gui.py
from __future__ import annotations

import csv
import sys
from pathlib import Path

try:
    import tkinter as tk
    from tkinter import filedialog, messagebox
except Exception:
    raise

# Drag & drop support
try:
    from tkinterdnd2 import DND_FILES, TkinterDnD  # pip install tkinterdnd2
except ImportError:
    print(
        "tkinterdnd2 is not installed. Install it with:\n"
        "  pip install tkinterdnd2"
    )
    sys.exit(1)

# --- Fields map retained from your existing script ---
CSV_FIELD_MAP = {
    0: "LineType",
    1: "StudentStudyItemAssessmentCurriculumItemCode",
    2: "StudentStudyItemAssessmentCurriculumItemVersionNumber",
    3: "StudentStudyItemAssessmentCurriculumItemFullTitle",
    4: "StudentStudyItemAssessmentDeliveryYear",
    5: "StudentStudyItemAssessmentDeliveryStudyPeriodCode",
    6: "StudentStudyItemAssessmentDeliveryStudyPeriodDescription",
    7: "StudentStudyItemAssessmentDeliveryLocationCode",
    8: "StudentStudyItemAssessmentDeliveryLocationDescription",
    9: "StudentStudyItemAssessmentDeliveryNumber",
    10: "StudentStudyItemAssessmentStudentID",
    11: "StudentStudyItemAssessmentStudentStudyItemAttemptNumber",
    12: "StudentStudyItemAssessmentID",
    13: "StudentStudyItemAssessmentTypeDescription",
    14: "StudentStudyItemAssessmentDescription",
    15: "StudentStudyItemAssessmentBarcode",
}

SSPASSESS = "SSPASSESS"


class DnDApp(TkinterDnD.Tk):
    """
    A drag-and-drop CSV loader that filters to 'SSPASSESS' rows.
    When the user clicks 'Send to Main' (or drops a file if auto_send is True),
    we call callback(rows) where rows is a list of dicts mapped by CSV_FIELD_MAP.
    """

    def __init__(self, callback=None, auto_send=False):
        super().__init__()
        self.callback = callback
        self.auto_send = auto_send
        self.title("Drop a CSV file (prints or sends ONLY SSPASSESS rows)")
        self.geometry("560x320")

        self._rows_accumulator = []  # store parsed rows across multiple drops

        self._build_ui()

    # ---------------------- UI ----------------------
    def _build_ui(self):
        instruction = tk.Label(
            self,
            text=(
                "Drag and drop a CSV file here.\n"
                "The app filters ONLY rows whose first column is 'SSPASSESS'\n"
                "(SSPASSESSHIST and header/format rows are ignored)."
            ),
            justify="center",
            font=("Segoe UI", 11),
            wraplength=520,
        )
        instruction.pack(padx=16, pady=(16, 8))

        self.drop_area = tk.Label(
            self,
            text="Drop CSV here",
            relief="ridge",
            width=50,
            height=8,
            bg="#f5f5f5",
            anchor="center",
            justify="center",
            font=("Segoe UI", 12),
        )
        self.drop_area.pack(padx=16, pady=8, fill="both", expand=True)

        # DnD bindings
        self.drop_area.drop_target_register(DND_FILES)
        self.drop_area.dnd_bind("<<Drop>>", self._on_drop)

        # Buttons row
        btns = tk.Frame(self)
        btns.pack(pady=(0, 12))

        tk.Button(btns, text="…or click to choose a CSV", command=self._browse_file).pack(
            side="left", padx=4
        )
        tk.Button(btns, text="Clear Loaded Rows", command=self._clear_rows).pack(
            side="left", padx=4
        )
        tk.Button(btns, text="Send to Main", command=self._send_to_main).pack(
            side="left", padx=4
        )

        # Status
        self.status_var = tk.StringVar(value="Ready")
        tk.Label(self, textvariable=self.status_var, anchor="w").pack(fill="x", padx=8, pady=(0, 8))

    # ---------------------- Events ----------------------
    def _on_drop(self, event):
        files = self.tk.splitlist(event.data)
        loaded = 0
        for file_path in files:
            p = Path(file_path)
            if p.is_file() and p.suffix.lower() == ".csv":
                try:
                    rows = self._extract_sspassess_rows(p)
                    self._rows_accumulator.extend(rows)
                    loaded += 1
                except Exception as e:
                    messagebox.showerror("Error", f"Error reading {p}:\n{e}")
            else:
                messagebox.showinfo("Skipped", f"Not a CSV file: {p}")

        self.drop_area.configure(text="Drop another CSV here…")
        self._set_status(f"Loaded {loaded} file(s); {len(self._rows_accumulator)} SSPASSESS rows in memory.")

        if self.auto_send and self.callback and self._rows_accumulator:
            self._send_to_main()

    def _browse_file(self):
        path = filedialog.askopenfilename(
            title="Select CSV file",
            filetypes=[("CSV files", "*.csv"), ("All files", "*.*")],
        )
        if path:
            # Simulate drop
            self._on_drop(type("Evt", (), {"data": path}))

    # ---------------------- Logic ----------------------
    def _extract_sspassess_rows(self, csv_path: Path):
        """
        Return a list[dict] of rows whose first column is exactly 'SSPASSESS'.
        Dictionary keys are mapped using CSV_FIELD_MAP.
        """
        rows = []
        with open(csv_path, "r", encoding="utf-8-sig", newline="") as f:
            reader = csv.reader(f)
            for raw in reader:
                if not raw:
                    continue
                if raw[0].strip().upper() != SSPASSESS:
                    continue
                # map fields using CSV_FIELD_MAP; ignore missing indices safely
                mapped = {
                    CSV_FIELD_MAP[i]: (raw[i].strip() if i < len(raw) else "")
                    for i in CSV_FIELD_MAP
                }
                rows.append(mapped)
        return rows

    def _clear_rows(self):
        self._rows_accumulator.clear()
        self._set_status("Cleared loaded rows.")

    def _send_to_main(self):
        if not self._rows_accumulator:
            messagebox.showinfo("Nothing to send", "No SSPASSESS rows loaded yet.")
            return

        if self.callback:
            # Send a copy to avoid accidental mutation by receivers
            payload = list(self._rows_accumulator)
            try:
                self.callback(payload)
                self._set_status(f"Sent {len(payload)} row(s) to main app.")
            except Exception as e:
                messagebox.showerror("Callback failed", str(e))
                return
        else:
            # Fallback: print to console if no callback given
            for row in self._rows_accumulator:
                print(row)
            self._set_status("Printed rows to console (no callback provided).")

    def _set_status(self, msg: str):
        self.status_var.set(msg)


# Backwards-compatible function (now with optional callback)
def run_dnd_gui(callback=None, auto_send=False):
    """
    Start the DnD GUI. If callback is provided, the user can click 'Send to Main'
    (or, if auto_send=True, results are sent immediately after a drop).
    """
    app = DnDApp(callback=callback, auto_send=auto_send)
    app.mainloop()
