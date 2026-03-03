"""
Excel Column Merger
===================
Copies column(s) from a source Excel sheet to a destination Excel sheet,
matching rows by a common identifier column (e.g. animal ID).

Requirements:  pip install pandas openpyxl
"""

import os
import tkinter as tk
from tkinter import ttk, filedialog, messagebox

import pandas as pd


# ──────────────────────────────────────────────────────────────────────────────
# Helper
# ──────────────────────────────────────────────────────────────────────────────

def load_sheet(file_path: str, sheet_name: str) -> pd.DataFrame:
    return pd.read_excel(file_path, sheet_name=sheet_name, dtype=str)


# ──────────────────────────────────────────────────────────────────────────────
# GUI
# ──────────────────────────────────────────────────────────────────────────────

class Panel(ttk.LabelFrame):
    """One side (source OR destination) of the interface."""

    def __init__(self, parent, label: str, **kwargs):
        super().__init__(parent, text=label, padding=10, **kwargs)

        self.file_path: str = ""
        self.df: pd.DataFrame | None = None

        # ── File row ──────────────────────────────────────────────────────────
        file_row = ttk.Frame(self)
        file_row.pack(fill=tk.X, pady=(0, 4))

        self._path_lbl = ttk.Label(file_row, text="No file selected",
                                   anchor=tk.W, foreground="gray",
                                   width=38, relief=tk.SUNKEN)
        self._path_lbl.pack(side=tk.LEFT, fill=tk.X, expand=True)

        self._browse_btn = ttk.Button(file_row, text="Browse…",
                                      command=self._browse)
        self._browse_btn.pack(side=tk.LEFT, padx=(4, 0))

        ttk.Separator(self, orient=tk.HORIZONTAL).pack(fill=tk.X, pady=6)

        # ── Sheet ─────────────────────────────────────────────────────────────
        ttk.Label(self, text="Sheet:").pack(anchor=tk.W)
        self.sheet_var = tk.StringVar()
        self.sheet_combo = ttk.Combobox(self, textvariable=self.sheet_var,
                                        state="disabled", width=35)
        self.sheet_combo.pack(fill=tk.X, pady=(0, 6))
        self.sheet_combo.bind("<<ComboboxSelected>>", self._on_sheet_selected)

        ttk.Separator(self, orient=tk.HORIZONTAL).pack(fill=tk.X, pady=6)

        # ── Common identifier ─────────────────────────────────────────────────
        ttk.Label(self, text="Common Identifier Column:").pack(anchor=tk.W)
        self.id_var = tk.StringVar()
        self.id_combo = ttk.Combobox(self, textvariable=self.id_var,
                                     state="disabled", width=35)
        self.id_combo.pack(fill=tk.X, pady=(0, 6))

        ttk.Separator(self, orient=tk.HORIZONTAL).pack(fill=tk.X, pady=6)

        # ── Column selector (overridden by subclasses) ────────────────────────
        self._build_column_section()

    # ── Overridable ───────────────────────────────────────────────────────────
    def _build_column_section(self):
        pass

    # ── File browsing ─────────────────────────────────────────────────────────
    def _browse(self):
        path = filedialog.askopenfilename(
            title="Select Excel File",
            filetypes=[("Excel files", "*.xlsx *.xls"), ("All files", "*.*")],
        )
        if not path:
            return
        self.file_path = path
        self._path_lbl.config(text=os.path.basename(path), foreground="black")

        try:
            xl = pd.ExcelFile(path)
            sheets = xl.sheet_names
        except Exception as exc:
            messagebox.showerror("Error", f"Cannot open file:\n{exc}")
            return

        self.sheet_combo.config(state="readonly")
        self.sheet_combo["values"] = sheets
        self.sheet_combo.current(0)
        self._load_sheet()

    def _on_sheet_selected(self, _event=None):
        self._load_sheet()

    def _load_sheet(self):
        if not self.file_path or not self.sheet_var.get():
            return
        try:
            self.df = load_sheet(self.file_path, self.sheet_var.get())
        except Exception as exc:
            messagebox.showerror("Error", f"Cannot load sheet:\n{exc}")
            return

        cols = list(self.df.columns)
        self.id_combo.config(state="readonly")
        self.id_combo["values"] = cols
        if cols:
            self.id_combo.current(0)

        self._on_columns_loaded(cols)

    def _on_columns_loaded(self, columns: list[str]):
        pass


# ──────────────────────────────────────────────────────────────────────────────

class SourcePanel(Panel):
    def __init__(self, parent, **kwargs):
        self._import_all_var = tk.BooleanVar(value=False)
        super().__init__(parent, "Source", **kwargs)

    def _build_column_section(self):
        ttk.Label(self, text="Column(s) to Copy:").pack(anchor=tk.W)

        list_frame = ttk.Frame(self)
        list_frame.pack(fill=tk.BOTH, expand=True, pady=(2, 4))

        sb = ttk.Scrollbar(list_frame, orient=tk.VERTICAL)
        sb.pack(side=tk.RIGHT, fill=tk.Y)

        self.col_listbox = tk.Listbox(
            list_frame, selectmode=tk.MULTIPLE,
            yscrollcommand=sb.set, height=7,
            exportselection=False,
        )
        self.col_listbox.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        sb.config(command=self.col_listbox.yview)

        ttk.Checkbutton(
            self,
            text="Import ALL columns from source",
            variable=self._import_all_var,
            command=self._toggle_import_all,
        ).pack(anchor=tk.W)

    def _on_columns_loaded(self, columns: list[str]):
        self.col_listbox.delete(0, tk.END)
        for col in columns:
            self.col_listbox.insert(tk.END, col)

    def _toggle_import_all(self):
        state = tk.DISABLED if self._import_all_var.get() else tk.NORMAL
        self.col_listbox.config(state=state)

    # ── Public API ────────────────────────────────────────────────────────────
    @property
    def import_all(self) -> bool:
        return self._import_all_var.get()

    def selected_columns(self) -> list[str]:
        """Return the list of columns chosen by the user (respects import_all)."""
        if self.import_all:
            id_col = self.id_var.get()
            return [c for c in self.df.columns if c != id_col]
        return [self.col_listbox.get(i) for i in self.col_listbox.curselection()]


# ──────────────────────────────────────────────────────────────────────────────

class DestinationPanel(Panel):
    def __init__(self, parent, **kwargs):
        super().__init__(parent, "Destination", **kwargs)

    def _build_column_section(self):
        ttk.Label(self, text="Paste into Column:").pack(anchor=tk.W)
        ttk.Label(
            self,
            text="Select an existing column to overwrite, or type a new name.\n"
                 "(When copying multiple columns, their source names are used.)",
            foreground="gray", font=("TkDefaultFont", 8), wraplength=280,
        ).pack(anchor=tk.W, pady=(0, 4))

        self.dst_col_var = tk.StringVar()
        self.dst_col_combo = ttk.Combobox(self, textvariable=self.dst_col_var,
                                          state="disabled", width=35)
        self.dst_col_combo.pack(fill=tk.X)

    def _on_columns_loaded(self, columns: list[str]):
        self.dst_col_combo.config(state="normal")
        self.dst_col_combo["values"] = columns
        self.dst_col_var.set("")


# ──────────────────────────────────────────────────────────────────────────────
# Main application window
# ──────────────────────────────────────────────────────────────────────────────

class ExcelMergerApp(tk.Tk):
    def __init__(self):
        super().__init__()
        self.title("Excel Column Merger")
        self.minsize(780, 500)
        self._build_ui()

    def _build_ui(self):
        # ── Main layout ───────────────────────────────────────────────────────
        root_frame = ttk.Frame(self, padding=10)
        root_frame.pack(fill=tk.BOTH, expand=True)
        root_frame.columnconfigure(0, weight=1)
        root_frame.columnconfigure(1, weight=1)
        root_frame.rowconfigure(0, weight=1)

        # ── Left / Right panels ───────────────────────────────────────────────
        self.src_panel = SourcePanel(root_frame)
        self.src_panel.grid(row=0, column=0, sticky="nsew", padx=(0, 5))

        self.dst_panel = DestinationPanel(root_frame)
        self.dst_panel.grid(row=0, column=1, sticky="nsew", padx=(5, 0))

        # ── Merge button ──────────────────────────────────────────────────────
        btn_frame = ttk.Frame(root_frame)
        btn_frame.grid(row=1, column=0, columnspan=2, pady=10)

        ttk.Button(btn_frame, text="⬅  Merge  ➡", command=self._do_merge,
                   width=20).pack()

        # ── Status bar ────────────────────────────────────────────────────────
        self._status = tk.StringVar(value="Ready.")
        ttk.Label(root_frame, textvariable=self._status,
                  relief=tk.SUNKEN, anchor=tk.W).grid(
            row=2, column=0, columnspan=2, sticky="ew")

    # ── Merge logic ───────────────────────────────────────────────────────────
    def _do_merge(self):
        src = self.src_panel
        dst = self.dst_panel

        # ── Validation ────────────────────────────────────────────────────────
        if src.df is None:
            messagebox.showwarning("Missing Input", "Please select a Source file and sheet.")
            return
        if dst.df is None:
            messagebox.showwarning("Missing Input", "Please select a Destination file and sheet.")
            return

        src_id = src.id_var.get().strip()
        dst_id = dst.id_var.get().strip()
        if not src_id or not dst_id:
            messagebox.showwarning("Missing Input",
                                   "Please select a Common Identifier column on both sides.")
            return

        cols_to_copy = src.selected_columns()
        if not cols_to_copy:
            messagebox.showwarning("Missing Input",
                                   "Please select at least one column to copy "
                                   "(or check \"Import ALL columns\").")
            return

        # ── Build lookup (source ID → values) ────────────────────────────────
        src_df = src.df.copy()
        dst_df = dst.df.copy()

        # Drop-duplicates on ID so set_index doesn't raise
        if src_df[src_id].duplicated().any():
            keep = messagebox.askyesno(
                "Duplicate IDs in source",
                f"The source column '{src_id}' has duplicate values. "
                "Only the FIRST occurrence will be used for each ID. Continue?",
            )
            if not keep:
                return
            src_df = src_df.drop_duplicates(subset=[src_id], keep="first")

        lookup = src_df.set_index(src_id)

        # ── Copy columns ──────────────────────────────────────────────────────
        matched = 0
        not_found = 0

        dst_col_override = dst.dst_col_var.get().strip()

        for src_col in cols_to_copy:
            # Column name in destination
            if len(cols_to_copy) == 1 and dst_col_override:
                dst_col = dst_col_override
            else:
                dst_col = src_col          # use source name when multiple cols

            values = []
            for key in dst_df[dst_id]:
                if key in lookup.index:
                    values.append(lookup.at[key, src_col])
                    matched += 1
                else:
                    values.append(None)
                    not_found += 1

            dst_df[dst_col] = values

        # ── Save output ───────────────────────────────────────────────────────
        default_name = (
            os.path.splitext(os.path.basename(dst.file_path))[0]
            + "_merged.xlsx"
        )
        out_path = filedialog.asksaveasfilename(
            title="Save Merged File As",
            defaultextension=".xlsx",
            initialfile=default_name,
            filetypes=[("Excel files", "*.xlsx")],
        )
        if not out_path:
            return

        try:
            # Preserve all sheets from the destination file; replace the edited one
            with pd.ExcelFile(dst.file_path) as xl:
                all_sheets = {s: xl.parse(s, dtype=str) for s in xl.sheet_names}

            all_sheets[dst.sheet_var.get()] = dst_df

            with pd.ExcelWriter(out_path, engine="openpyxl") as writer:
                for sheet_name, df in all_sheets.items():
                    df.to_excel(writer, sheet_name=sheet_name, index=False)

        except Exception as exc:
            messagebox.showerror("Error", f"Failed to save file:\n{exc}")
            return

        summary = (
            f"Done! — Matched: {matched} value(s) | "
            f"Not found: {not_found} | "
            f"Saved: {os.path.basename(out_path)}"
        )
        self._status.set(summary)
        messagebox.showinfo(
            "Merge complete",
            f"Merge finished successfully!\n\n"
            f"  Matched rows : {matched}\n"
            f"  Unmatched IDs: {not_found}\n\n"
            f"Output saved to:\n{out_path}",
        )


# ──────────────────────────────────────────────────────────────────────────────

if __name__ == "__main__":
    app = ExcelMergerApp()
    app.mainloop()
