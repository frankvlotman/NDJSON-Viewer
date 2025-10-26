import json
import os
import re
import datetime as dt
import tkinter as tk
from tkinter import ttk, filedialog, messagebox

# ---- Dependencies ----
try:
    import pandas as pd
    PANDAS_AVAILABLE = True
except Exception:
    pd = None
    PANDAS_AVAILABLE = False


class NDJSONViewerApp(ttk.Frame):
    DISPLAY_ROW_CAP = 10000  # cap rows shown for responsiveness

    def __init__(self, master):
        super().__init__(master, padding=8)
        self.master.title("NDJSON Viewer • Filter • Export to XLSX")
        self.pack(fill="both", expand=True)

        # State
        self.df_full = None    # full dataset (DataFrame)
        self.df_view = None    # filtered view (DataFrame)
        self.current_path = None

        # UI
        self._build_style()
        self._build_toolbar()
        self._build_table()
        self._build_status()

    # ---------- UI ----------
    def _build_style(self):
        style = ttk.Style()
        try:
            style.theme_use("clam")
        except Exception:
            pass
        style.configure("TButton", padding=6)
        style.configure("TEntry", padding=4)

    def _build_toolbar(self):
        bar = ttk.Frame(self)
        bar.pack(side="top", fill="x", pady=(0, 6))

        ttk.Button(bar, text="Open NDJSON…", command=self.open_ndjson).pack(side="left")
        ttk.Separator(bar, orient="vertical").pack(side="left", fill="y", padx=6)

        # Filter controls
        ttk.Label(bar, text="Filter (any column):").pack(side="left")
        self.filter_var = tk.StringVar()
        ent = ttk.Entry(bar, textvariable=self.filter_var, width=32)
        ent.pack(side="left", padx=(6, 6))
        ent.bind("<Return>", lambda e: self.apply_filter())

        ttk.Button(bar, text="Apply", command=self.apply_filter).pack(side="left")
        ttk.Button(bar, text="Clear", command=self.clear_filter).pack(side="left", padx=(6, 10))

        # Linkage controls
        self.include_linked_var = tk.BooleanVar(value=True)
        self.link_col_var = tk.StringVar(value="")  # will be populated after load

        ttk.Checkbutton(bar, text="Include linked rows", variable=self.include_linked_var)\
            .pack(side="left", padx=(0, 6))

        ttk.Label(bar, text="Link column:").pack(side="left")
        self.link_col_combo = ttk.Combobox(bar, textvariable=self.link_col_var, width=20, state="disabled")
        self.link_col_combo.pack(side="left", padx=(6, 10))

        # Spacer
        ttk.Label(bar, text="").pack(side="left", expand=True, fill="x")

        # Export
        ttk.Button(bar, text="Export to .xlsx", command=self.export_xlsx).pack(side="right")

    def _build_table(self):
        frame = ttk.Frame(self)
        frame.pack(side="top", fill="both", expand=True)

        self.tree = ttk.Treeview(frame, columns=(), show="headings")
        self.tree.pack(side="left", fill="both", expand=True)

        vsb = ttk.Scrollbar(frame, orient="vertical", command=self.tree.yview)
        hsb = ttk.Scrollbar(self, orient="horizontal", command=self.tree.xview)
        self.tree.configure(yscrollcommand=vsb.set, xscrollcommand=hsb.set)

        vsb.pack(side="right", fill="y")
        hsb.pack(side="bottom", fill="x")

    def _build_status(self):
        self.status = tk.StringVar(value="Open an NDJSON file to begin.")
        ttk.Label(self, textvariable=self.status, anchor="w").pack(side="bottom", fill="x", pady=(6, 0))

    # ---------- Data ----------
    def open_ndjson(self):
        path = filedialog.askopenfilename(
            title="Select NDJSON file",
            filetypes=[("NDJSON files", "*.ndjson"), ("JSON Lines", "*.jsonl"), ("All files", "*.*")]
        )
        if not path:
            return
        if not PANDAS_AVAILABLE:
            messagebox.showerror("Dependency missing",
                                 "This app requires 'pandas' to load NDJSON files.\n\n"
                                 "Install with: pip install pandas")
            return

        self.current_path = path
        self._set_status(f"Loading: {os.path.basename(path)} …")
        self.master.update_idletasks()

        try:
            df = pd.read_json(path, lines=True, dtype=False)
            if not isinstance(df, pd.DataFrame):
                df = pd.DataFrame(df)

            # Normalize timestamps & pretty-print complex objects
            df = self._normalize_df(df)

            self.df_full = df
            self.df_view = df

            self._prepare_link_col_controls(df)

            self._populate_tree(self.df_view)
            self._set_status(self._mk_status_line())
        except Exception as e:
            messagebox.showerror("Load Error", f"Failed to load NDJSON.\n\n{e}")
            self._set_status("Failed to load file.")

    def _prepare_link_col_controls(self, df):
        cols = [str(c) for c in df.columns.tolist()]
        self.link_col_combo["values"] = cols
        self.link_col_combo["state"] = "readonly" if cols else "disabled"

        # Heuristics to default to a bin-like column
        default = None
        lower_map = {c.lower(): c for c in cols}
        for candidate in ("bin", "bin_id", "bin_no", "bin_number", "location_bin", "rack_bin", "to_bin", "from_bin"):
            if candidate in lower_map:
                default = lower_map[candidate]
                break
        if default is None and cols:
            default = cols[0]

        if default:
            self.link_col_var.set(default)

    # ---------- Timestamp normalization ----------
    @staticmethod
    def _format_epoch(val):
        """
        Convert epoch (seconds / milliseconds / microseconds, int/float) to 'YYYY-MM-DD HH:MM:SS' (UTC).
        Return original value if it doesn't look like an epoch.
        """
        try:
            if isinstance(val, (int, float)) and not (isinstance(val, bool)):
                x = float(val)

                # Decide unit by magnitude
                if x > 1e14:           # microseconds
                    x = x / 1e6
                elif x > 1e12:         # milliseconds
                    x = x / 1e3
                # else: treat as seconds (can include fractional)

                # Reasonable epoch window (1970 .. 2100)
                if 0 <= x <= 4102444800:
                    dt_utc = dt.datetime.utcfromtimestamp(x)
                    return dt_utc.strftime("%Y-%m-%d %H:%M:%S")
        except Exception:
            pass
        return val

    @classmethod
    def _convert_mapping_ts(cls, obj):
        """
        Recursively walk dict/list and convert any key that looks like '*_ts' (or contains 'ts')
        from epoch to readable string. Also stringify the result for display.
        """
        try:
            if isinstance(obj, dict):
                new_d = {}
                for k, v in obj.items():
                    if isinstance(v, (dict, list)):
                        new_d[k] = cls._convert_mapping_ts(v)
                    else:
                        if isinstance(k, str) and ('ts' in k.lower()):
                            new_d[k] = cls._format_epoch(v)
                        else:
                            new_d[k] = v
                return new_d
            elif isinstance(obj, list):
                return [cls._convert_mapping_ts(v) for v in obj]
        except Exception:
            return obj
        return obj

    @classmethod
    def _normalize_df(cls, df: "pd.DataFrame") -> "pd.DataFrame":
        df = df.copy()

        # 1) Convert obvious timestamp columns (column name contains 'ts')
        for col in df.columns:
            if 'ts' in str(col).lower():
                df[col] = df[col].apply(cls._format_epoch)

        # 2) For object columns with dict/list values, convert nested '*_ts' & pretty-print to JSON strings
        for col in df.columns:
            if df[col].dtype == 'object':
                # only map if we actually see dict/list values (keeps it fast-ish)
                sample = df[col].head(50)
                if sample.map(lambda x: isinstance(x, (dict, list))).any():
                    def _transform_cell(x):
                        if isinstance(x, (dict, list)):
                            converted = cls._convert_mapping_ts(x)
                            try:
                                return json.dumps(converted, ensure_ascii=False, sort_keys=True)
                            except Exception:
                                return str(converted)
                        return x
                    df[col] = df[col].map(_transform_cell)

        return df

    # ---------- Linking helper ----------
    @staticmethod
    def _extract_link_keys(val):
        """
        Extract individual bin/location keys from a cell.
        Handles patterns like:
          'W2->Z3', 'W2 to Z3', 'W2/Z3', 'W2 , Z3'
        Returns a set of uppercase tokens (e.g., {'W2','Z3'}).
        """
        if val is None or (isinstance(val, float) and pd is not None and pd.isna(val)):
            return set()
        s = str(val).upper()
        tokens = re.findall(r"[A-Z]+[A-Z0-9]*", s)
        return set(tokens)

    # ---------- Filtering ----------
    def apply_filter(self):
        if self.df_full is None:
            return

        q = self.filter_var.get().strip()
        df = self.df_full

        if not q:
            self.df_view = df
            self._populate_tree(self.df_view)
            self._set_status(self._mk_status_line())
            return

        # Step 1: direct text match across any column (case-insensitive)
        query = q.lower()
        mask = None
        for col in df.columns:
            s = df[col].astype(str).str.lower()
            m = s.str.contains(query, na=False)
            mask = m if mask is None else (mask | m)

        direct_matches = df[mask] if mask is not None else df.iloc[0:0]

        # Step 2: optionally include linked rows sharing any key in the link column
        include_linked = self.include_linked_var.get()
        link_col = self.link_col_var.get()

        if include_linked and link_col and link_col in df.columns and not direct_matches.empty:
            matched_keys = set()
            for v in direct_matches[link_col].dropna().tolist():
                matched_keys |= self._extract_link_keys(v)

            if matched_keys:
                def has_any_key(v):
                    return len(self._extract_link_keys(v) & matched_keys) > 0

                linked_mask = df[link_col].apply(has_any_key)
                df_final = df[linked_mask]
            else:
                df_final = direct_matches
        else:
            df_final = direct_matches

        self.df_view = df_final
        self._populate_tree(self.df_view)

        extra = ""
        if include_linked and link_col and link_col in df.columns and not direct_matches.empty:
            extra = f" • linked by {link_col}"
        self._set_status(self._mk_status_line(filtering=True) + extra)

    def clear_filter(self):
        self.filter_var.set("")
        self.apply_filter()

    # ---------- Table ----------
    def _populate_tree(self, df):
        # Clear headers & rows
        for c in self.tree["columns"]:
            self.tree.heading(c, text="")
        self.tree.delete(*self.tree.get_children())

        if df is None or len(df) == 0:
            self.tree["columns"] = ()
            return

        cols = [str(c) for c in df.columns.tolist()]
        self.tree["columns"] = cols
        for c in cols:
            self.tree.heading(c, text=c)
            self.tree.column(c, anchor="w", width=120, stretch=True)

        # Autosize with a small sample
        self._autosize_columns(df.head(min(200, len(df))))

        # Insert (capped) rows
        to_show = min(len(df), self.DISPLAY_ROW_CAP)
        for _, row in df.iloc[:to_show].iterrows():
            values = ["" if (pd.isna(v) if PANDAS_AVAILABLE else (v is None)) else str(v) for v in row.tolist()]
            self.tree.insert("", "end", values=values)

        if len(df) > to_show:
            self._set_status(
                f"Showing first {to_show:,} of {len(df):,} rows (display capped). Use filter to narrow results."
            )

    def _autosize_columns(self, df_sample):
        # Simple width guess: header vs sample content
        px_per_char = 7
        for c in df_sample.columns:
            header_w = max(10, len(str(c)) * px_per_char + 16)
            sample_vals = df_sample[c].astype(str).tolist()
            max_cell = max((len(s) for s in sample_vals), default=0)
            cell_w = min(900, max_cell * px_per_char + 24)
            self.tree.column(c, width=max(header_w, cell_w))

    # ---------- Export ----------
    def export_xlsx(self):
        if self.df_view is None or len(self.df_view) == 0:
            messagebox.showinfo("Nothing to Export", "No data available to export.")
            return
        if not PANDAS_AVAILABLE:
            messagebox.showerror("Dependency missing",
                                 "Excel export requires 'pandas' and 'openpyxl'.\n\n"
                                 "Install with: pip install pandas openpyxl")
            return

        default_name = "ndjson_export.xlsx"
        if self.current_path:
            base = os.path.splitext(os.path.basename(self.current_path))[0]
            default_name = f"{base}.xlsx"

        path = filedialog.asksaveasfilename(
            title="Save as Excel",
            defaultextension=".xlsx",
            initialfile=default_name,
            filetypes=[("Excel Workbook", "*.xlsx")]
        )
        if not path:
            return

        try:
            self.df_view.to_excel(path, index=False, engine="openpyxl")
            messagebox.showinfo("Export Complete", f"Saved:\n{path}")
        except Exception as e:
            messagebox.showerror("Export Error", f"Failed to export.\n\n{e}")

    # ---------- Helpers ----------
    def _mk_status_line(self, filtering=False):
        if self.df_view is None:
            return "No data loaded."
        total = len(self.df_full) if self.df_full is not None else 0
        shown = len(self.df_view) if self.df_view is not None else 0
        name = os.path.basename(self.current_path) if self.current_path else "(unsaved)"
        return f"{name} — Filtered rows: {shown:,} / {total:,}" if filtering else f"{name} — Rows: {shown:,}"

    def _set_status(self, text):
        self.status.set(text)


def main():
    root = tk.Tk()
    root.geometry("1200x650")
    app = NDJSONViewerApp(root)
    root.mainloop()


if __name__ == "__main__":
    main()
