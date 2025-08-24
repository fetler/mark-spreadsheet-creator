import json
import tkinter as tk
from tkinter import ttk, messagebox
from tkinter.scrolledtext import ScrolledText
from pathlib import Path

APP_TITLE = "Student IDs & Marks - Spreadsheet Style"
DATA_FILE = "student_marks.json"


def data_file_path() -> Path:
    try:
        base = Path(__file__).resolve().parent
    except NameError:
        base = Path.cwd()
    return base / DATA_FILE


class GridApp(tk.Tk):
    def __init__(self, callback=None):
        super().__init__()
        self.callback = callback
        self.title(APP_TITLE)
        self.geometry("500x650")
        self.minsize(400, 600)

        # In-memory store (list[tuple[str, str]])
        self.paired_rows = []

        # UI state for in-place edit
        self._edit_entry = None
        self._edit_var = None
        self._edit_item = None
        self._edit_col = None

        self._build_ui()
        self._style_treeview()
        self._insert_initial_rows(20)  # some empty rows to start
        self._try_load_on_start()

    def process_ids_and_marks(self, pairs):
        if self.callback:
            self.callback(pairs)
        else:
            for sid, mark in pairs:
                print(f"{sid} {mark}")

    # ---------------- UI ----------------
    def _build_ui(self):
        root = ttk.Frame(self, padding=12)
        root.pack(fill="both", expand=True)

        ttk.Label(
            root,
            text="Spreadsheet-style input: double-click a cell to edit; paste via buttons on the right.",
            font=("Segoe UI", 10)
        ).pack(anchor="w", pady=(0, 8))

        # === Main 2-column area: left (treeview) and right (button panel) ===
        main = ttk.Frame(root)
        main.pack(fill="both", expand=True)

        # Left: Treeview + scrollbars
        left = ttk.Frame(main)
        left.grid(row=0, column=0, sticky="nsew")
        main.columnconfigure(0, weight=1)   # left expands
        main.rowconfigure(0, weight=1)

        self.tree = ttk.Treeview(
            left, columns=("id", "mark"), show="headings",
            selectmode="extended", height=16
        )
        self.tree.heading("id", text="Student ID")
        self.tree.heading("mark", text="Mark")
        self.tree.column("id", width=150, anchor="w")
        self.tree.column("mark", width=150, anchor="center")

        vsb = ttk.Scrollbar(left, orient="vertical", command=self.tree.yview)
        hsb = ttk.Scrollbar(left, orient="horizontal", command=self.tree.xview)
        self.tree.configure(yscrollcommand=vsb.set, xscrollcommand=hsb.set)

        self.tree.grid(row=0, column=0, sticky="nsew")
        vsb.grid(row=0, column=1, sticky="ns")
        hsb.grid(row=1, column=0, sticky="ew")
        left.rowconfigure(0, weight=1)
        left.columnconfigure(0, weight=1)

        self.tree.bind("<Double-1>", self._on_cell_double_click)

        # Right: vertical button panel
        right = ttk.Frame(main)
        right.grid(row=0, column=1, sticky="ns", padx=(10, 0))
        # Optional: fix width so it doesn’t stretch
        right.grid_propagate(True)

        # Group 1: Paste
        paste_grp = ttk.Labelframe(right, text="Paste")
        paste_grp.pack(fill="x", anchor="n", pady=(0, 8))
        ttk.Button(paste_grp, text="Paste IDs", command=self.on_paste_ids).pack(fill="x", pady=2)
        ttk.Button(paste_grp, text="Paste Marks", command=self.on_paste_marks).pack(fill="x", pady=2)
        ttk.Button(paste_grp, text="Paste 2-Column", command=self.on_paste_two_columns).pack(fill="x", pady=2)

        # Group 2: Rows & Grid
        rows_grp = ttk.Labelframe(right, text="Rows & Grid")
        rows_grp.pack(fill="x", anchor="n", pady=(0, 8))
        ttk.Button(rows_grp, text="Add 10 Rows", command=lambda: self._insert_initial_rows(10)).pack(fill="x", pady=2)
        ttk.Button(rows_grp, text="Delete Selected", command=self.on_delete_selected).pack(fill="x", pady=2)
        ttk.Button(rows_grp, text="Clear Grid", command=self.on_clear_grid).pack(fill="x", pady=2)

        # Spacer to push persistence group to bottom if desired
        ttk.Frame(right).pack(expand=True, fill="both")

        # Group 3: Persistence
        persist_grp = ttk.Labelframe(right, text="Persistence")
        persist_grp.pack(fill="x", anchor="s")
        ttk.Button(persist_grp, text="Store", command=self.on_store).pack(fill="x", pady=2)
        ttk.Button(persist_grp, text="Retrieve", command=self.on_retrieve).pack(fill="x", pady=2)
        ttk.Button(persist_grp, text="Load from File", command=self.on_load_from_file).pack(fill="x", pady=2)
        ttk.Button(persist_grp, text="Save to File", command=self.on_save_to_file).pack(fill="x", pady=2)
        ttk.Button(persist_grp, text="Save", command=self.on_process).pack(fill="x", pady=2)

        # === Output area (below the 2-column area) ===
        ttk.Label(root, text="Output (ID\\tMark)", font=("Segoe UI", 10, "bold")).pack(anchor="w", pady=(10, 0))
        self.output = ScrolledText(root, height=10, wrap="none", font=("Consolas", 10), state="disabled")
        self.output.pack(fill="both", expand=False)

        out_btns = ttk.Frame(root)
        out_btns.pack(fill="x", pady=(6, 0))
        ttk.Button(out_btns, text="Copy Output", command=self.on_copy_output).pack(side="left", padx=4)
        ttk.Button(out_btns, text="Clear Output", command=self.on_clear_output).pack(side="left", padx=4)

        # Status bar
        self.status_var = tk.StringVar(value="Ready")
        ttk.Label(root, textvariable=self.status_var, anchor="w", relief="sunken").pack(fill="x", pady=(8, 0))

    def _style_treeview(self):
        style = ttk.Style(self)
        try:
            # Use a modern theme if available
            style.theme_use("clam")
        except Exception:
            pass
        style.configure("Treeview", rowheight=24, font=("Segoe UI", 10))
        style.configure("Treeview.Heading", font=("Segoe UI", 10, "bold"))
        # Zebra striping tags
        self.tree.tag_configure("odd", background="#ffffff")
        self.tree.tag_configure("even", background="#f6f6f6")

    def _insert_initial_rows(self, count: int):
        for _ in range(count):
            self.tree.insert("", "end", values=("", ""))
        self._retag_rows()
        self._set_status(f"Added {count} empty row(s).")

    # ------------- Cell Editing -------------
    def _on_cell_double_click(self, event):
        # Determine the clicked cell
        region = self.tree.identify_region(event.x, event.y)
        if region != "cell":
            return
        item = self.tree.identify_row(event.y)
        col = self.tree.identify_column(event.x)  # '#1' or '#2'
        if not item or not col:
            return
        bbox = self.tree.bbox(item, col)
        if not bbox:
            return
        x, y, w, h = bbox
        current_value = self.tree.set(item, col)

        # Destroy any existing editor
        self._destroy_editor()

        self._edit_var = tk.StringVar(value=current_value)
        self._edit_entry = tk.Entry(self.tree, textvariable=self._edit_var, font=("Segoe UI", 10))
        self._edit_entry.place(x=x, y=y, width=w, height=h)
        self._edit_item = item
        self._edit_col = col
        self._edit_entry.focus()
        self._edit_entry.icursor("end")

        def commit(event=None):
            new_val = self._edit_var.get()
            try:
                self.tree.set(self._edit_item, self._edit_col, new_val)
            finally:
                self._destroy_editor()

        def cancel(event=None):
            self._destroy_editor()

        self._edit_entry.bind("<Return>", commit)
        self._edit_entry.bind("<Escape>", cancel)
        self._edit_entry.bind("<FocusOut>", commit, add="+")

    def _destroy_editor(self):
        if self._edit_entry is not None:
            try:
                self._edit_entry.destroy()
            except Exception:
                pass
        self._edit_entry = None
        self._edit_var = None
        self._edit_item = None
        self._edit_col = None

    # ------------- Paste Helpers -------------

    def _selected_start_index(self) -> int:
        """Start at the first selected row; fallback to first empty-row index; else append at end."""
        items = self.tree.get_children()
        sel = self.tree.selection()
        if sel:
            try:
                return items.index(sel[0])
            except ValueError:
                pass
        # First empty row (both cells empty)
        for idx, it in enumerate(items):
            if not self.tree.set(it, "id").strip() and not self.tree.set(it, "mark").strip():
                return idx
        return len(items)

    def _ensure_rows(self, n: int):
        """Ensure at least n rows exist."""
        current = len(self.tree.get_children())
        if n > current:
            self._insert_initial_rows(n - current)

    def on_paste_ids(self):
        try:
            text = self.clipboard_get()
        except Exception:
            self._set_status("Clipboard is empty or not text.")
            return
        ids = [ln.strip() for ln in text.splitlines() if ln.strip()]  # ignore blank lines
        if not ids:
            self._set_status("No IDs found to paste.")
            return

        start = self._selected_start_index()
        self._ensure_rows(start + len(ids))
        items = self.tree.get_children()
        for i, sid in enumerate(ids):
            self.tree.set(items[start + i], "id", sid)
        self._retag_rows()
        self._set_status(f"Pasted {len(ids)} ID(s) starting at row {start + 1}.")

    def on_paste_marks(self):
        try:
            text = self.clipboard_get()
        except Exception:
            self._set_status("Clipboard is empty or not text.")
            return
        # Preserve blank lines: each becomes an empty mark cell
        marks = [ln.strip() for ln in text.splitlines()]
        if not marks:
            self._set_status("No marks found to paste.")
            return

        start = self._selected_start_index()
        self._ensure_rows(start + len(marks))
        items = self.tree.get_children()
        for i, mk in enumerate(marks):
            self.tree.set(items[start + i], "mark", mk)
        self._retag_rows()
        self._set_status(f"Pasted {len(marks)} mark value(s) starting at row {start + 1}.")

    def on_paste_two_columns(self):
        try:
            text = self.clipboard_get()
        except Exception:
            self._set_status("Clipboard is empty or not text.")
            return
        rows = text.splitlines()
        parsed = []
        for r in rows:
            # Prefer TSV (what Excel copies), fallback to CSV
            if "\t" in r:
                parts = r.split("\t")
            else:
                parts = r.split(",")
            id_part = parts[0].strip() if len(parts) >= 1 else ""
            mark_part = parts[1].strip() if len(parts) >= 2 else ""
            parsed.append((id_part, mark_part))

        if not parsed:
            self._set_status("Nothing to paste.")
            return

        start = self._selected_start_index()
        self._ensure_rows(start + len(parsed))
        items = self.tree.get_children()
        for i, (sid, mk) in enumerate(parsed):
            self.tree.set(items[start + i], "id", sid)
            self.tree.set(items[start + i], "mark", mk)
        self._retag_rows()
        self._set_status(f"Pasted {len(parsed)} row(s) (2-column) starting at row {start + 1}.")

    # ------------- Grid Ops -------------

    def _retag_rows(self):
        for i, item in enumerate(self.tree.get_children()):
            self.tree.item(item, tags=("even" if i % 2 == 0 else "odd",))

    def on_delete_selected(self):
        sel = self.tree.selection()
        if not sel:
            self._set_status("No rows selected to delete.")
            return
        for it in sel:
            self.tree.delete(it)
        self._retag_rows()
        self._set_status(f"Deleted {len(sel)} row(s).")

    def on_clear_grid(self):
        for it in self.tree.get_children():
            self.tree.delete(it)
        self._insert_initial_rows(20)
        self._set_status("Grid cleared.")

    # ------------- Store / Retrieve / Persist -------------

    def _collect_pairs_from_grid(self):
        pairs = []
        for it in self.tree.get_children():
            sid = self.tree.set(it, "id").strip()
            mk = self.tree.set(it, "mark").strip()
            if sid:  # ignore rows without an ID
                pairs.append((sid, mk))
        return pairs

    def on_store(self):
        pairs = self._collect_pairs_from_grid()
        if not pairs:
            messagebox.showinfo("Store", "No student IDs found in the grid.")
            self._set_status("Store: no IDs found.")
            return
        self.paired_rows = pairs
        saved_ok = self._save_data_silent()
        msg = f"Stored {len(self.paired_rows)} row(s)."
        if saved_ok:
            msg += f" Saved to '{data_file_path().name}'."
        self._set_status(msg)
        messagebox.showinfo("Stored", msg)

    def on_retrieve(self):
        if not self.paired_rows:
            loaded = self._load_data_silent()
            if not loaded:
                messagebox.showinfo("Retrieve", "No data stored yet (in memory or file).")
                self._set_status("Retrieve: no data found.")
                return
        # Output ID \t mark (mark may be blank)
        lines = [f"{sid}\t{mk}" for sid, mk in self.paired_rows]
        self._set_output("\n".join(lines))
        self._set_status(f"Retrieved {len(self.paired_rows)} row(s).")

    def on_load_from_file(self):
        if self._load_data_silent():
            # Reflect loaded data into the grid
            for it in self.tree.get_children():
                self.tree.delete(it)
            self._ensure_rows(len(self.paired_rows))
            items = self.tree.get_children()
            for i, (sid, mk) in enumerate(self.paired_rows):
                self.tree.set(items[i], "id", sid)
                self.tree.set(items[i], "mark", mk)
            self._retag_rows()
            messagebox.showinfo("Loaded", f"Loaded {len(self.paired_rows)} row(s) from '{data_file_path().name}'.")
            self._set_status(f"Loaded {len(self.paired_rows)} row(s) from file.")
        else:
            messagebox.showinfo("Load", "No file found or file is empty.")
            self._set_status("Load: no data file or empty file.")

    def on_save_to_file(self):
        # If memory empty, try collecting from grid before saving
        if not self.paired_rows:
            pairs = self._collect_pairs_from_grid()
            if not pairs:
                messagebox.showinfo("Save", "Nothing to save yet. Paste/edit grid or retrieve first.")
                self._set_status("Save: nothing to save.")
                return
            self.paired_rows = pairs
        if self._save_data_silent():
            messagebox.showinfo("Saved", f"Saved {len(self.paired_rows)} row(s) to '{data_file_path().name}'.")
            self._set_status("Saved to file.")
        else:
            messagebox.showerror("Save failed", "Could not save to file.")
            self._set_status("Save failed.")

    def _save_data_silent(self) -> bool:
        try:
            payload = [{"student_id": sid, "mark": mk} for sid, mk in self.paired_rows]
            with open(data_file_path(), "w", encoding="utf-8") as f:
                json.dump(payload, f, indent=2, ensure_ascii=False)
            return True
        except Exception:
            return False

    def _load_data_silent(self) -> bool:
        try:
            path = data_file_path()
            if not path.exists():
                return False
            with open(path, "r", encoding="utf-8") as f:
                payload = json.load(f)
            self.paired_rows = [
                (str(it.get("student_id", "")).strip(), str(it.get("mark", "")).strip())
                for it in payload if str(it.get("student_id", "")).strip()
            ]
            return len(self.paired_rows) > 0
        except Exception:
            return False

    def _try_load_on_start(self):
        if self._load_data_silent():
            self._set_status(f"Loaded {len(self.paired_rows)} row(s) from previous session.")

    def on_process(self):
        pairs = self._collect_pairs_from_grid()
        if not pairs:
            messagebox.showinfo("Process", "No data to process yet.")
            self._set_status("Process: no data found.")
            return
        self.process_ids_and_marks(pairs)
        self._set_status(f"Processed {len(pairs)} row(s).")

    # ------------- Output helpers -------------

    def _set_output(self, text: str):
        self.output.config(state="normal")
        self.output.delete("1.0", "end")
        if text:
            self.output.insert("1.0", text)
        self.output.config(state="disabled")

    def on_copy_output(self):
        txt = self.output.get("1.0", "end").strip("\n")
        if not txt.strip():
            self._set_status("Output is empty; nothing to copy.")
            return
        self.clipboard_clear()
        self.clipboard_append(txt)
        self._set_status("Output copied to clipboard.")

    def on_clear_output(self):
        self._set_output("")
        self._set_status("Output cleared.")

    def _set_status(self, msg: str):
        self.status_var.set(msg)


def run_gui_ids_marks():
    # Optional HiDPI tweak (Windows)
    try:
        from ctypes import windll
        windll.shcore.SetProcessDpiAwareness(1)
    except Exception:
        pass

    app = GridApp()
    app.mainloop()

if __name__ == "__main__":
    run_gui_ids_marks()
