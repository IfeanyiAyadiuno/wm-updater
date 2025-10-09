import tkinter as tk
from tkinter import ttk, messagebox, filedialog
import pandas as pd
import pyodbc
from pathlib import Path

# ============================================================
# CONFIG
# ============================================================

ACCESS_DB_PATH = r"I:\ResEng\Tools\Programmers Paradise\GUI_WM\PCE_WM1.accdb"
ACCESS_TABLE_NAME = "PCE_WM"

# Toggle DB diagnostics to the terminal (no output when False)
DEBUG_DB = False
def log(*args, **kwargs):
    if DEBUG_DB:
        print(*args, **kwargs)

# Columns expected in Access (order matters for inserts)
TABLE_COLUMNS = [
    "ID",  # AutoNumber (never inserted)
    "GasIDREC",
    "PressuresIDREC",
    "Well Name",
    "Formation Producer",
    "Layer Producer",
    "Fault Block",
    "Pad Name",
    "Completions Technology",
    "Lateral Length",
    "Value Navigator UWI",
    "Composite name",
]

# ID is autonumber; never inserted/updated explicitly
AUTONUMBER_FIELD = "ID"

# Which fields are editable dropdowns (sourced from existing unique values)
DROPDOWN_FIELDS = ["Formation Producer", "Layer Producer", "Fault Block", "Completions Technology"]

# Which fields are plain text entries
ENTRY_FIELDS = ["Well Name", "Lateral Length", "Value Navigator UWI", "Pad Name"]

# Editable fields in the grid
EDITABLE_FIELDS = ENTRY_FIELDS + DROPDOWN_FIELDS

# Optional keyboard affordance: Space toggles ✓ for complete rows
ENABLE_SPACE_TOGGLE = True


# ============================================================
# DATA ACCESS LAYER
# ============================================================

def connect_access(db_path: str):
    """
    Create a pyodbc connection to a local Access database.
    """
    conn_str = (
        r"DRIVER={Microsoft Access Driver (*.mdb, *.accdb)};"
        rf"DBQ={db_path};"
        r"UID=Admin;PWD=;"
    )
    return pyodbc.connect(conn_str)


def load_access_table(db_path: str, table_name: str) -> pd.DataFrame:
    """
    Load the full Access table as a pandas DataFrame.
    Ordered by ID ascending (stable order, newest at bottom).
    """
    p = Path(db_path)
    if not p.exists():
        raise FileNotFoundError(f"DB path not found: {db_path}")
    try:
        from datetime import datetime
        log(f"[DB] Using: {p}  (modified: {datetime.fromtimestamp(p.stat().st_mtime)})")
    except Exception:
        pass

    with connect_access(db_path) as conn:
        try:
            cur = conn.cursor()
            cur.execute(f"SELECT COUNT(*) FROM [{table_name}]")
            total = cur.fetchone()[0]
            log(f"[DB] Row count in Access right now: {total}")
        except Exception:
            pass

        df = pd.read_sql(f"SELECT * FROM [{table_name}] ORDER BY ID ASC", conn)

        try:
            tail = df[["ID", "Well Name"]].tail(5)
            log("[DB] Tail IDs just loaded:\n" + tail.to_string(index=False))
        except Exception:
            pass

    return df


def get_unique_options(df: pd.DataFrame) -> dict:
    """
    Build unique sorted lists for each dropdown column from the current data.
    """
    options = {}
    for col in DROPDOWN_FIELDS:
        if col in df.columns:
            vals = (
                df[col]
                .dropna()
                .astype(str)
                .map(str.strip)
                .replace({"": None})
                .dropna()
                .unique()
            )
            options[col] = sorted(vals)
        else:
            options[col] = []
    return options


def insert_records(conn, table_name: str, rows: list[dict]):
    """
    Batch insert (ID omitted). Each row is a dict mapping column -> value.
    """
    insert_cols = [c for c in TABLE_COLUMNS if c != AUTONUMBER_FIELD]
    placeholders = ", ".join(["?"] * len(insert_cols))
    col_list = ", ".join([f"[{c}]" for c in insert_cols])
    sql = f"INSERT INTO [{table_name}] ({col_list}) VALUES ({placeholders})"

    cur = conn.cursor()
    params_batch = [[row.get(c, None) for c in insert_cols] for row in rows]
    cur.executemany(sql, params_batch)
    conn.commit()


def find_existing_id(conn, table_name: str, gas_id: str | None, pres_id: str | None):
    """
    Return the ID of a row that matches the provided identifiers.
    Prefer matching BOTH (GasIDREC AND PressuresIDREC) when both are given;
    fall back to a single-column match only if one is missing.
    """
    cur = conn.cursor()
    if gas_id and pres_id:
        cur.execute(
            f"SELECT ID FROM [{table_name}] WHERE GasIDREC = ? AND PressuresIDREC = ?",
            (gas_id, pres_id),
        )
    elif gas_id:
        cur.execute(f"SELECT ID FROM [{table_name}] WHERE GasIDREC = ?", (gas_id,))
    elif pres_id:
        cur.execute(f"SELECT ID FROM [{table_name}] WHERE PressuresIDREC = ?", (pres_id,))
    else:
        return None
    row = cur.fetchone()
    return row[0] if row else None


def update_record(conn, table_name: str, rec_id: int, payload: dict):
    """
    Update selected columns by ID. GasIDREC/PressuresIDREC remain unchanged.
    """
    updatable_cols = [c for c in TABLE_COLUMNS if c not in (AUTONUMBER_FIELD, "GasIDREC", "PressuresIDREC")]
    sets, params = [], []
    for c in updatable_cols:
        if c in payload:
            sets.append(f"[{c}] = ?")
            params.append(payload.get(c))
    if not sets:
        return
    params.append(rec_id)
    sql = f"UPDATE [{table_name}] SET {', '.join(sets)} WHERE ID = ?"
    cur = conn.cursor()
    cur.execute(sql, params)


def compose_name(well: str | None, layer: str | None, tech: str | None) -> str | None:
    """
    Return "Well - Layer - Tech" if all three are present; otherwise None.
    """
    w = (well or "").strip()
    l = (layer or "").strip()
    t = (tech or "").strip()
    if not (w and l and t):
        return None
    return f"{w} - {l} - {t}"


# ============================================================
# GUI HELPERS
# ============================================================

class XYScrollFrame(ttk.Frame):
    """
    A simple frame with a Canvas + interior frame that supports both
    vertical and horizontal scrolling. You can pack your own content
    inside self.viewPort.
    """
    def __init__(self, parent, *args, **kwargs):
        super().__init__(parent, *args, **kwargs)
        self.canvas = tk.Canvas(self, borderwidth=0, highlightthickness=0)
        self.viewPort = ttk.Frame(self.canvas)
        self.vsb = ttk.Scrollbar(self, orient="vertical")
        self.hsb = ttk.Scrollbar(self, orient="horizontal")

        self.canvas.configure(yscrollcommand=self.vsb.set, xscrollcommand=self.hsb.set)
        self.canvas.grid(row=0, column=0, sticky="nsew")
        self.vsb.grid(row=0, column=1, sticky="ns")
        self.hsb.grid(row=1, column=0, sticky="ew")
        self.grid_rowconfigure(0, weight=1)
        self.grid_columnconfigure(0, weight=1)

        self.canvas_window = self.canvas.create_window((0, 0), window=self.viewPort, anchor="nw")
        self.viewPort.bind("<Configure>", lambda e: self.canvas.configure(scrollregion=self.canvas.bbox("all")))


class CellEditor:
    """
    Detached editor in a tiny borderless Toplevel over the target cell.
    Combobox commits on selection/Enter/Tab (no focus-out commit).
    Entry commits on Enter/Tab/FocusOut. Escape always cancels.

    Takes focus + grab on open; releases on destroy so the UI never gets stuck.
    """
    def __init__(self, app, tree, item, col_name, bbox, options, current_val, on_commit):
        self.app = app
        self.tree = tree
        self.item = item
        self.col_name = col_name
        self.on_commit = on_commit

        x, y, w, h = bbox
        abs_x = tree.winfo_rootx() + x
        abs_y = tree.winfo_rooty() + y

        self.top = tk.Toplevel(tree)
        self.top.withdraw()
        self.top.overrideredirect(True)

        try:
            self.top.transient(tree.winfo_toplevel())
            self.top.lift()
            self.top.attributes("-topmost", True)
        except Exception:
            pass

        self.top.geometry(f"{w}x{h}+{abs_x}+{abs_y}")
        self.top.deiconify()

        self.is_combo = options is not None
        if self.is_combo:
            self.widget = ttk.Combobox(self.top, values=options, state="readonly")
            self.widget.set(current_val if current_val is not None else "")
        else:
            self.widget = ttk.Entry(self.top)
            if current_val is not None:
                self.widget.insert(0, str(current_val))

        self.widget.pack(fill="both", expand=True)

        # focus + grab
        try:
            self.top.grab_set()
        except Exception:
            pass
        try:
            self.top.focus_force()
        except Exception:
            pass
        try:
            self.widget.focus_force()
        except Exception:
            pass

        try:
            self.top.after(150, lambda: self.top.attributes("-topmost", False))
        except Exception:
            pass

        # Bindings
        self.widget.bind("<Return>", self._commit)
        self.widget.bind("<Tab>", self._commit)
        self.widget.bind("<Escape>", self._cancel)

        if self.is_combo:
            self.widget.bind("<<ComboboxSelected>>", self._commit)
            self.widget.after(10, lambda: self.widget.event_generate("<Alt-Down>"))
        else:
            self.widget.bind("<FocusOut>", self._commit)

        # Ensure app pointer is cleared if the window goes away
        self.top.bind("<Destroy>", lambda e: setattr(self.app, "_editor", None))

    def _commit(self, _=None):
        try:
            value = self.widget.get().strip()
        except Exception:
            value = ""
        self.destroy()
        self.on_commit(value)

    def _cancel(self, _=None):
        self.destroy()

    def destroy(self):
        try:
            self.top.grab_release()
        except Exception:
            pass
        try:
            self.top.destroy()
        except Exception:
            pass
        self.app._editor = None


# ============================================================
# MAIN APP
# ============================================================

class App(tk.Tk):
    """
    Two-tab Access table editor:
      - Current Wells: shows live table, allows inline editing of editable columns.
      - Add New Wells: lets you fill in details for rows that have ID pairs but blank Well Name.
    """
    def __init__(self):
        super().__init__()
        self.title("WM Updater — Gas/Pressure IDs to Access")
        self.geometry("1360x800")

        # --- Toolbar
        toolbar = ttk.Frame(self)
        toolbar.pack(fill="x", padx=8, pady=6)
        self.db_path_var = tk.StringVar(value=ACCESS_DB_PATH)
        self.table_var = tk.StringVar(value=ACCESS_TABLE_NAME)
        ttk.Label(toolbar, text="Access DB:").pack(side="left")
        ttk.Entry(toolbar, textvariable=self.db_path_var, width=80).pack(side="left", padx=4)
        ttk.Button(toolbar, text="…", command=self.pick_db).pack(side="left")
        ttk.Label(toolbar, text="  Table:").pack(side="left", padx=(12, 0))
        ttk.Entry(toolbar, textvariable=self.table_var, width=20).pack(side="left", padx=4)
        ttk.Button(toolbar, text="Reload", command=self.reload_all).pack(side="left", padx=8)

        # --- Notebook
        self.nb = ttk.Notebook(self)
        self.nb.pack(fill="both", expand=True)

        # ========== Tab 1: Current Wells
        self.tab_current = ttk.Frame(self.nb)
        self.nb.add(self.tab_current, text="Current Wells")

        tree_wrap = ttk.Frame(self.tab_current)
        tree_wrap.pack(fill="both", expand=True)

        self.tree = ttk.Treeview(tree_wrap, show="headings", selectmode="none")
        self._editor: CellEditor | None = None   # active cell editor (if any)
        self.columns_present: list[str] = []
        self._checked = set()
        self._pending_edits: dict[str, dict] = {}

        # Styling
        style = ttk.Style(self)
        try:
            style.theme_use("clam")
        except Exception:
            pass

        style.configure(
            "Treeview",
            borderwidth=1,
            relief="solid",
            rowheight=26,          # calmer, consistent row height
        )
        # If something manages to “select”, keep it visually neutral
        style.map(
            "Treeview",
            background=[("selected", "#ffffff")],
            foreground=[("selected", "#000000")],
        )

        style.configure(
            "Treeview.Heading",
            background="#f3f3f3",
            borderwidth=1,
            relief="solid",
        )
        style.map(
            "Treeview.Heading",
            relief=[("pressed", "sunken"), ("active", "raised")],
        )

        style.configure("Slim.TCheckbutton", padding=0)
        style.configure("TEntry", padding=(4, 2))
        style.configure("TCombobox", padding=(4, 2))

        # Scrollbars: close editor whenever you scroll via bars
        ys = ttk.Scrollbar(tree_wrap, orient="vertical")
        xs = ttk.Scrollbar(tree_wrap, orient="horizontal")
        ys.configure(command=lambda *a: (self._close_editor(False), self.tree.yview(*a)))
        xs.configure(command=lambda *a: (self._close_editor(False), self.tree.xview(*a)))
        self.tree.configure(yscrollcommand=ys.set, xscrollcommand=xs.set)

        self.tree.grid(row=0, column=0, sticky="nsew")
        ys.grid(row=0, column=1, sticky="ns")
        xs.grid(row=1, column=0, sticky="ew")
        tree_wrap.rowconfigure(0, weight=1)
        tree_wrap.columnconfigure(0, weight=1)

        # Bindings for editor lifecycle and interactions
        self.nb.bind("<<NotebookTabChanged>>", lambda e: self._close_editor(False))   # switching tabs
        self.tree.bind("<Unmap>",                lambda e: self._close_editor(False)) # hiding tree
        self.bind("<Unmap>",                     lambda e: self._close_editor(False)) # app minimized/unmapped
        self.tree.bind("<Configure>",            lambda e: self._close_editor(False), add="+")
        # Wheel scrolling (vertical; hold Shift for horizontal)
        self.tree.bind("<MouseWheel>",           self.on_mousewheel)   # Windows/macOS
        self.tree.bind("<Shift-MouseWheel>",     self.on_mousewheel)
        self.tree.bind("<Button-4>",             self.on_mousewheel)   # Linux up
        self.tree.bind("<Button-5>",             self.on_mousewheel)   # Linux down

        self.tree.bind("<Button-1>", self.on_tree_click)          # toggle Select when clicking column 1 cell
        self.tree.bind("<Double-1>", self.on_tree_double_click)   # start editor
        self.bind("<FocusOut>", lambda e: self._close_editor(False), add="+")
        self.tree.bind("<space>", self.on_space_toggle)

        # Footer (tab 1)
        current_footer = ttk.Frame(self.tab_current)
        current_footer.pack(fill="x", padx=8, pady=(4, 8))
        ttk.Button(current_footer, text="Save checked edits → Access", command=self.save_checked_edits).pack(side="right")

        # ========== Tab 2: Add New
        self.tab_add = ttk.Frame(self.nb)
        self.nb.add(self.tab_add, text="Add New Wells")

        head = ttk.Frame(self.tab_add)
        head.pack(fill="x", padx=8, pady=6)
        ttk.Label(head, text="Select rows to insert. Well Name is optional; other fields via dropdown.").pack(side="left")

        self.scroll = XYScrollFrame(self.tab_add)
        self.scroll.pack(fill="both", expand=True, padx=8, pady=4)
        self.scroll.vsb.configure(command=self.scroll.canvas.yview)
        self.scroll.hsb.configure(command=self.scroll.canvas.xview)

        # App footer
        footer = ttk.Frame(self)
        footer.pack(fill="x", padx=8, pady=10)
        self.count_label = ttk.Label(footer, text="Ready")
        self.count_label.pack(side="left")
        ttk.Button(footer, text="Update Selected → Access", command=self.do_update).pack(side="right")

        # Data caches
        self.df_current: pd.DataFrame | None = None
        self.dropdown_options: dict[str, list] = {}
        self.new_widgets: list[dict] = []

        self.new_ids: list[dict] = []
        self._staged_pairs: set[tuple] = set()  # (GasIDREC, PressuresIDREC)

        # Initial load
        self.reload_all()

    # ---------------- Toolbar actions ----------------

    def pick_db(self):
        path = filedialog.askopenfilename(filetypes=[("Access DB", "*.accdb;*.mdb"), ("All", "*.*")])
        if path:
            self.db_path_var.set(path)

    # ---------------- Editor lifecycle ----------------

    def _close_editor(self, commit: bool = False):
        """
        Close the active cell editor (if any).
        commit=False: discard, commit=True: commit current value.
        Ensures any grab held by the editor is released.
        """
        ed = self._editor
        if not ed:
            return
        try:
            if commit:
                ed._commit()
            else:
                ed.destroy()   # destroy() releases grab inside CellEditor
        except Exception:
            try:
                if getattr(ed, "top", None):
                    ed.top.grab_release()
            except Exception:
                pass
        finally:
            self._editor = None

    # ---------------- Data/UI load ----------------

    def reload_all(self):
        """
        Pull from Access, rebuild Current Wells grid with:
        - completed rows first
        - pending rows (blank Well Name) at the bottom, highlighted.
        Rebuild Add New using ONLY self.new_ids (staged by the user).
        """
        try:
            self.df_current = load_access_table(self.db_path_var.get(), self.table_var.get())
        except Exception as e:
            messagebox.showerror("Load Error", f"Failed to load Access table:\n{e}")
            return

        # Columns present in DB (keep order)
        cols_present = [c for c in TABLE_COLUMNS if c in self.df_current.columns]
        self.columns_present = ["Select"] + cols_present
        self.tree.configure(columns=self.columns_present)

        # Calmer, consistent widths; stretch long text columns
        col_widths = {
            "Select": 64,
            "ID": 40,
            "GasIDREC": 260,
            "PressuresIDREC": 260,
            "Well Name": 220,
            "Formation Producer": 160,
            "Layer Producer": 160,
            "Fault Block": 140,
            "Pad Name": 160,
            "Completions Technology": 180,
            "Lateral Length": 120,
            "Value Navigator UWI": 200,
            "Composite name": 260,
        }
        min_widths = {
            "Select": 32,
            "ID": 56,
            "GasIDREC": 220,
            "PressuresIDREC": 220,
            "Well Name": 160,
            "Formation Producer": 140,
            "Layer Producer": 140,
            "Fault Block": 120,
            "Pad Name": 140,
            "Completions Technology": 160,
            "Lateral Length": 96,
            "Value Navigator UWI": 160,
            "Composite name": 200,
        }
        stretch_cols = {"Well Name", "Pad Name", "Value Navigator UWI", "Composite name"}

        for c in self.columns_present:
            self.tree.heading(c, text=c, anchor="w" if c != "Select" else "center")
            self.tree.column(
                c,
                width=col_widths.get(c, 160),
                minwidth=min_widths.get(c, 100),
                anchor="w" if c != "Select" else "center",
                stretch=(c in stretch_cols),
            )

        # Reset UI + tags
        self.tree.delete(*self.tree.get_children())
        self._checked.clear()
        self._pending_edits.clear()
        self._close_editor(False)

        self.tree.tag_configure("even", background="#ffffff")
        self.tree.tag_configure("odd",  background="#f7f7f7")
        self.tree.tag_configure("pending", background="#fff4cc")  # light highlight for pending rows

        # Split complete vs pending (blank or NaN Well Name)
        mask_pending = (
            self.df_current["Well Name"].isnull()
            | (self.df_current["Well Name"].astype(str).str.strip() == "")
        )
        df_complete = self.df_current.loc[~mask_pending]
        df_pending  = self.df_current.loc[mask_pending]

        # Track which Treeview items are pending so we can move them when checked
        self._pending_row_ids = set()
        self._pending_iid_to_pair = {}

        # Stable zebra striping independent of DataFrame index
        rowno = 0

        # Insert COMPLETE rows first
        for idx, row in df_complete.iterrows():
            iid = str(row["ID"]) if "ID" in row and pd.notna(row["ID"]) else str(idx)
            values = ["☐"] + [row.get(c, "") for c in cols_present]
            tag = "odd" if (rowno % 2) else "even"
            self.tree.insert("", "end", iid=iid, values=values, tags=(tag,))
            rowno += 1

        # Then insert PENDING rows (at bottom), highlighted
        for idx, row in df_pending.iterrows():
            iid = str(row["ID"]) if "ID" in row and pd.notna(row["ID"]) else f"p_{idx}"
            values = ["☐"] + [row.get(c, "") for c in cols_present]
            tag = ("odd" if (rowno % 2) else "even", "pending")
            self.tree.insert("", "end", iid=iid, values=values, tags=tag)
            self._pending_row_ids.add(iid)
            self._pending_iid_to_pair[iid] = (
                str(row.get("GasIDREC") or ""),
                str(row.get("PressuresIDREC") or "")
            )
            rowno += 1

        # Dropdown choices from ALL data
        self.dropdown_options = get_unique_options(self.df_current)

        # Build Add New tab ONLY from staged rows (self.new_ids)
        self.build_add_rows()

        # Footer counts
        pending_ct = len(df_pending)
        staged_ct  = len(self.new_ids)
        self.count_label.config(
            text=f"Loaded {len(self.df_current)} rows • {pending_ct} pending (bottom) • Staged for Add New: {staged_ct}"
        )

    # ---------------- Add New tab ----------------

    def build_add_rows(self):
        """
        Build the grid of inputs for rows that have IDs but blank Well Name.
        """
        # Clear previous UI
        for child in list(self.scroll.viewPort.children.values()):
            child.destroy()
        self.new_widgets.clear()

        table = ttk.Frame(self.scroll.viewPort)
        table.pack(fill="both", expand=True, padx=8, pady=4)

        headers = ["", "GasIDREC", "PressuresIDREC", *ENTRY_FIELDS, *DROPDOWN_FIELDS, "Composite name"]
        col_widths = [36, 150, 150] + [200]*len(ENTRY_FIELDS) + [180]*len(DROPDOWN_FIELDS) + [240]

        # Header row (grid-style)
        for ci, title in enumerate(headers):
            text = "✓" if ci == 0 else title
            hdr = tk.Label(
                table, text=text, font=("Segoe UI", 9, "bold"),
                bg="#f3f3f3", bd=1, relief="solid", anchor="center"
            )
            hdr.grid(row=0, column=ci, sticky="nsew", padx=0, pady=0, ipadx=4, ipady=3)
            weight = 0 if ci in (0, 1, 2) else 1
            table.grid_columnconfigure(ci, minsize=col_widths[ci], weight=weight, uniform="addcols")

        def cell(parent, r, c):
            box = tk.Frame(parent, bd=1, relief="solid")
            box.grid(row=r, column=c, sticky="nsew", padx=0, pady=0)
            return box

        for ri, rec in enumerate(self.new_ids, start=1):
            # Select
            var_sel = tk.BooleanVar(value=True)
            ttk.Checkbutton(cell(table, ri, 0), variable=var_sel, style="Slim.TCheckbutton").pack(anchor="center")

            # IDs
            tk.Label(cell(table, ri, 1), text=str(rec.get("GasIDREC") or ""), anchor="w").pack(fill="x", padx=4, pady=2)
            tk.Label(cell(table, ri, 2), text=str(rec.get("PressuresIDREC") or ""), anchor="w").pack(fill="x", padx=4, pady=2)

            # Entries
            entry_vars = {}
            col_index = 3
            for col in ENTRY_FIELDS:
                v = tk.StringVar(value="")
                ttk.Entry(cell(table, ri, col_index), textvariable=v).pack(fill="x", expand=True, padx=4, pady=2)
                entry_vars[col] = v
                col_index += 1

            # Dropdowns
            dropdown_vars = {}
            for col in DROPDOWN_FIELDS:
                v = tk.StringVar(value="")
                ttk.Combobox(
                    cell(table, ri, col_index),
                    textvariable=v,
                    values=self.dropdown_options.get(col, []),
                    state="readonly"
                ).pack(fill="x", expand=True, padx=4, pady=2)
                dropdown_vars[col] = v
                col_index += 1

            # Composite
            comp_var = tk.StringVar(value="")
            ttk.Label(cell(table, ri, col_index), textvariable=comp_var).pack(fill="x", expand=True, padx=4, pady=2)

            # ---- Per-row callback with captured defaults (fixes late-binding bug) ----
            def _sync(*_,
                    entry_vars=entry_vars,
                    dropdown_vars=dropdown_vars,
                    comp_var=comp_var):
                wname = entry_vars.get("Well Name").get() if "Well Name" in entry_vars else ""
                layer = dropdown_vars.get("Layer Producer").get() if "Layer Producer" in dropdown_vars else ""
                tech  = dropdown_vars.get("Completions Technology").get() if "Completions Technology" in dropdown_vars else ""
                comp_var.set(compose_name(wname, layer, tech) or "")

            # Attach traces so any change recomputes the composite (for THIS row)
            if "Well Name" in entry_vars:
                entry_vars["Well Name"].trace_add("write", _sync)
            if "Layer Producer" in dropdown_vars:
                dropdown_vars["Layer Producer"].trace_add("write", _sync)
            if "Completions Technology" in dropdown_vars:
                dropdown_vars["Completions Technology"].trace_add("write", _sync)

            # Compute once initially
            _sync()

            # Stash row widgets/state
            self.new_widgets.append({
                "selected": var_sel,
                "gas": rec.get("GasIDREC"),
                "pres": rec.get("PressuresIDREC"),
                "entries": entry_vars,
                "dropdowns": dropdown_vars,
                "comp_var": comp_var,
            })

    # ---------------- Current Wells interactions ----------------

    def on_tree_click(self, event):
        """
        Click in column #1 toggles checkbox. For any other region (headers, separators,
        scrollbar, empty space), we close any editor and let Tk handle the default behavior.
        """
        region = self.tree.identify("region", event.x, event.y)   # 'cell', 'heading', 'separator', 'tree', 'nothing'
        column = self.tree.identify_column(event.x)               # '#1' = Select
        item   = self.tree.identify_row(event.y)

        # Toggle only when clicking the Select column on a valid row
        if region == "cell" and column == "#1" and item:
            self.tree.focus(item)
            self._toggle_item_checkbox(item)
            return "break"  # we handled it

        # For everything else, just close the editor and allow default behavior
        if self._editor:
            self._close_editor(False)
        # NOTE: intentionally no 'return "break"' here

    def on_tree_double_click(self, event):
        """
        Start a CellEditor over the double-clicked cell if it is editable.
        - You must check (✓) the row first.
        - Pending rows (blank Well Name, shown at bottom) are not editable here.
        """
        region = self.tree.identify("region", event.x, event.y)
        if region != "cell":
            return

        item = self.tree.identify_row(event.y)
        col_id = self.tree.identify_column(event.x)  # '#1', '#2', ...
        if not item or not col_id:
            return

        # Block editing for pending rows (they should be staged to Add New instead)
        if item in getattr(self, "_pending_row_ids", set()):
            return

        # Require the row to be check-marked before editing
        if item not in self._checked:
            try:
                self.bell()
            except Exception:
                pass
            return

        # Map '#n' -> column name
        try:
            col_index = int(col_id.replace("#", "")) - 1
        except Exception:
            return
        if col_index < 0 or col_index >= len(self.columns_present):
            return

        col_name = self.columns_present[col_index]
        # Not editable columns
        if col_name in ("Select", AUTONUMBER_FIELD, "GasIDREC", "PressuresIDREC"):
            return
        if col_name not in EDITABLE_FIELDS:
            return

        # Cell rectangle and current value
        bbox = self.tree.bbox(item, col_id)
        if not bbox:
            return
        x, y, w, h = bbox
        current_val = self.tree.set(item, col_name)
        options = self.dropdown_options.get(col_name) if col_name in self.dropdown_options else None

        # Close any previous editor
        self._close_editor(False)

        def _commit(value: str):
            # record change in grid
            self.tree.set(item, col_name, value)
            self._pending_edits.setdefault(item, {})[col_name] = (value if value != "" else None)
            # keep Composite name in sync
            if col_name in ("Well Name", "Layer Producer", "Completions Technology"):
                comp = compose_name(
                    self.tree.set(item, "Well Name"),
                    self.tree.set(item, "Layer Producer"),
                    self.tree.set(item, "Completions Technology"),
                )
                if "Composite name" in self.columns_present:
                    self.tree.set(item, "Composite name", comp or "")
                self._pending_edits.setdefault(item, {})["Composite name"] = comp

        # Create the detached editor window over the cell
        self._editor = CellEditor(
            app=self,
            tree=self.tree,
            item=item,
            col_name=col_name,
            bbox=(x, y, w, h),
            options=options,            # None => Entry, list => Combobox
            current_val=current_val,
            on_commit=_commit,
        )

    def _toggle_item_checkbox(self, item: str):
        """Shared logic to toggle the ✓ cell, including staging pending rows."""
        cur = self.tree.set(item, "Select")
        new = "☑" if cur != "☑" else "☐"
        self.tree.set(item, "Select", new)

        # Pending rows get staged to Add New when checked
        if item in getattr(self, "_pending_row_ids", set()):
            if new == "☑":
                gas = self.tree.set(item, "GasIDREC")
                prs = self.tree.set(item, "PressuresIDREC")
                pair = (str(gas or ""), str(prs or ""))
                if pair not in self._staged_pairs:
                    self._staged_pairs.add(pair)
                    self.new_ids.append({"GasIDREC": pair[0], "PressuresIDREC": pair[1]})
                    self.build_add_rows()
                # remove from Current Wells view
                try:
                    self.tree.delete(item)
                    self._pending_row_ids.discard(item)
                    self._pending_iid_to_pair.pop(item, None)
                except Exception:
                    pass
                self.nb.select(self.tab_add)
            return

        # Complete rows: maintain the checked set
        if new == "☑":
            self._checked.add(item)
        else:
            self._checked.discard(item)

    def on_space_toggle(self, event):
        """
        Optional keyboard affordance: space toggles ✓ for the focused row.
        To avoid accidental staging, we ignore pending rows here.
        """
        if not ENABLE_SPACE_TOGGLE:
            return "break"
        item = self.tree.focus()
        if not item:
            return "break"
        if item in getattr(self, "_pending_row_ids", set()):
            try:
                self.bell()  # gentle nudge
            except Exception:
                pass
            return "break"
        self._toggle_item_checkbox(item)
        return "break"

    def on_mousewheel(self, event):
        """
        Predictable scrolling that also closes any inline editor:
        - Vertical by default
        - Hold Shift for horizontal
        Works on Windows/macOS (<MouseWheel>) and Linux (<Button-4/5>).
        """
        # Always close the editor when scrolling
        self._close_editor(False)

        # Shift pressed?
        shift_held = bool(getattr(event, "state", 0) & 0x0001)

        # Windows/macOS path
        if hasattr(event, "delta") and event.delta:
            # Tk uses multiples of 120; normalize to +/-1 units
            units = -1 * (event.delta // 120 or (1 if event.delta < 0 else -1))
            if shift_held:
                self.tree.xview_scroll(units, "units")
            else:
                self.tree.yview_scroll(units, "units")
            return "break"

        # Linux X11: mouse wheel generates Button-4 (up) / Button-5 (down)
        if getattr(event, "num", None) in (4, 5):
            units = -1 if event.num == 4 else 1
            if shift_held:
                self.tree.xview_scroll(units, "units")
            else:
                self.tree.yview_scroll(units, "units")
            return "break"

    # ---------------- Save edits ----------------

    def save_checked_edits(self):
        """
        Apply pending edits for checked rows back into Access.
        """
        if not self._checked:
            messagebox.showinfo("No rows checked", "Check the rows you want to save, then try again.")
            return

        to_update = {iid: edits for iid, edits in self._pending_edits.items() if iid in self._checked and edits}
        if not to_update:
            messagebox.showinfo("Nothing to save", "No pending edits on checked rows.")
            return

        db_path = self.db_path_var.get()
        table = self.table_var.get()
        updated = 0
        failed = 0

        with connect_access(db_path) as conn:
            for iid, payload in to_update.items():
                # Find Access row ID
                try:
                    rec_id = int(iid)
                except Exception:
                    row_vals = {c: self.tree.set(iid, c) for c in self.columns_present if c != "Select"}
                    rec_id = find_existing_id(conn, table, row_vals.get("GasIDREC"), row_vals.get("PressuresIDREC"))

                if not rec_id:
                    failed += 1
                    continue

                # Only update editable columns (and Composite name if available)
                safe_payload = {k: v for k, v in payload.items() if k in EDITABLE_FIELDS}

                wn = payload.get("Well Name", self.tree.set(iid, "Well Name"))
                lp = payload.get("Layer Producer", self.tree.set(iid, "Layer Producer"))
                ct = payload.get("Completions Technology", self.tree.set(iid, "Completions Technology"))
                comp = compose_name(wn, lp, ct)
                if comp is not None:
                    safe_payload["Composite name"] = comp
                    if "Composite name" in self.columns_present:
                        self.tree.set(iid, "Composite name", comp)

                try:
                    update_record(conn, table, rec_id, safe_payload)
                    updated += 1
                except Exception as e:
                    failed += 1
                    messagebox.showerror("Update Error", f"Failed to update ID={rec_id}:\n{e}")
                    return
            conn.commit()

        # Clear pending edits for saved rows
        for iid in list(self._checked):
            self._pending_edits.pop(iid, None)

        messagebox.showinfo("Done", f"Updated: {updated}\nFailed: {failed}")
        self.reload_all()

    # ---------------- Add New: apply ----------------

    def do_update(self):
        """
        Apply updates/inserts for rows staged on the Add New tab.
        After success, remove processed rows from staging so the list stays clean.
        """
        rows = []
        staged_pairs_selected = []  # to track which staged rows were chosen this time
        for item in self.new_widgets:
            if not item["selected"].get():
                continue
            payload = {
                "GasIDREC": item["gas"],
                "PressuresIDREC": item["pres"],
            }
            for col, v in item["entries"].items():
                payload[col] = v.get().strip() or None
            for col, v in item["dropdowns"].items():
                payload[col] = v.get().strip() or None
            payload["Composite name"] = item["comp_var"].get() or None
            rows.append(payload)
            staged_pairs_selected.append((str(item["gas"] or ""), str(item["pres"] or "")))

        if not rows:
            messagebox.showinfo("Nothing selected", "No rows were checked. Tick the boxes to select rows to apply.")
            return

        db_path = self.db_path_var.get()
        table = self.table_var.get()

        updated = 0
        inserted = 0
        skipped = 0
        processed_ok_pairs: list[tuple] = []

        with connect_access(db_path) as conn:
            to_insert = []
            for r in rows:
                wn = r.get("Well Name")
                if wn:
                    cur = conn.cursor()
                    cur.execute(f"SELECT COUNT(1) FROM [{table}] WHERE [Well Name] = ?", (wn,))
                    if cur.fetchone()[0] > 0:
                        if not messagebox.askyesno("Duplicate Well Name", f"'{wn}' already exists. Continue?"):
                            skipped += 1
                            continue

                rec_id = find_existing_id(conn, table, r.get("GasIDREC"), r.get("PressuresIDREC"))
                pair = (str(r.get("GasIDREC") or ""), str(r.get("PressuresIDREC") or ""))
                if rec_id:
                    try:
                        update_record(conn, table, rec_id, r)
                        updated += 1
                        processed_ok_pairs.append(pair)
                    except Exception as e:
                        messagebox.showerror("Update Error", f"Failed to update ID={rec_id}:\n{e}")
                        return
                else:
                    to_insert.append((r, pair))

            if to_insert:
                try:
                    insert_records(conn, table, [r for r, _pair in to_insert])
                    inserted += len(to_insert)
                    processed_ok_pairs.extend([pair for _r, pair in to_insert])
                except Exception as e:
                    messagebox.showerror("Insert Error", f"Failed to insert records:\n{e}")
                    return
            conn.commit()

        messagebox.showinfo("Done", f"Updated: {updated}\nInserted: {inserted}\nSkipped: {skipped}")

        # Remove successfully processed pairs from staging (self.new_ids / self._staged_pairs)
        if processed_ok_pairs:
            keep = []
            processed_set = set(processed_ok_pairs)
            for rec in self.new_ids:
                pair = (str(rec.get("GasIDREC") or ""), str(rec.get("PressuresIDREC") or ""))
                if pair not in processed_set:
                    keep.append(rec)
                else:
                    # also drop from the staged set
                    self._staged_pairs.discard(pair)
            self.new_ids = keep

        # Rebuild both tabs
        self.reload_all()


if __name__ == "__main__":
    import warnings
    warnings.filterwarnings("ignore", category=UserWarning, module="pandas")
    app = App()
    app.mainloop()