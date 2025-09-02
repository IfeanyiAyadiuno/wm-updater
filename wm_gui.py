import tkinter as tk
from tkinter import ttk, messagebox, filedialog
import pandas as pd
import pyodbc
from pathlib import Path

# =========================
# CONFIG — EDIT THESE
# =========================
ACCESS_DB_PATH = r"I:\ResEng\Tools\Programmers Paradise\GUI_WM\PCE_WM1.accdb"  # <- your Access .accdb path
ACCESS_TABLE_NAME = "PCE_WM"                                                          # <- Access table name

# Columns expected in the Access table (keep order consistent with insert)
# Adjust names to match EXACT Access field names
TABLE_COLUMNS = [
    "ID",                 # AutoNumber in Access? If so, we will NOT insert it (set to None / omit in INSERT)
    "GasIDREC",
    "PressuresIDREC",
    "Well Name",
    "Formation Producer",
    "Layer Producer",
    "Fault Block",
    "Pad Name",
    "Inital Flow Date",
    "ES Well Name",
    "Completions Technology",
    "Lateral Length",
    "Value Navigator UWI",
    "Composite name",
]

# If ID is AutoNumber, remove it from INSERT/GUI targets
AUTONUMBER_FIELD = "ID"

# Which fields should be filled via dropdowns (values sourced from existing table uniques)
DROPDOWN_FIELDS = [
    "Formation Producer", "Layer Producer", "Fault Block", "Pad Name", "Completions Technology"
]

# Which fields are free-text/entry (kept simple)
ENTRY_FIELDS = [
    "Well Name", "Inital Flow Date", "ES Well Name", "Lateral Length", "Value Navigator UWI",
]

EDITABLE_FIELDS = ENTRY_FIELDS + DROPDOWN_FIELDS

# =========================
# DATA ACCESS LAYER
# =========================

def connect_access(db_path: str):
    conn_str = (
        r"DRIVER={Microsoft Access Driver (*.mdb, *.accdb)};"
        rf"DBQ={db_path};"
        r"UID=Admin;PWD=;"
    )
    return pyodbc.connect(conn_str)


def load_access_table(db_path: str, table_name: str) -> pd.DataFrame:
    with connect_access(db_path) as conn:
        df = pd.read_sql(f"SELECT * FROM [{table_name}]", conn)
    return df


def get_unique_options(df: pd.DataFrame) -> dict:
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


def record_exists(conn, table_name: str, gas_id: str, pres_id: str, well_name: str | None = None) -> bool:
    cur = conn.cursor()
    # Primary check: if GasIDREC or PressuresIDREC already exists
    cur.execute(
        f"SELECT COUNT(1) FROM [{table_name}] WHERE GasIDREC = ? OR PressuresIDREC = ?",
        (gas_id, pres_id)
    )
    count = cur.fetchone()[0]
    if count and count > 0:
        return True
    # Secondary check if user typed a Well Name
    if well_name:
        cur.execute(
            f"SELECT COUNT(1) FROM [{table_name}] WHERE [Well Name] = ?",
            (well_name,)
        )
        count2 = cur.fetchone()[0]
        if count2 and count2 > 0:
            return True
    return False


def insert_records(conn, table_name: str, rows: list[dict]):
    # Build parameter list dynamically, excluding AutoNumber field if needed
    insert_cols = [c for c in TABLE_COLUMNS if c != AUTONUMBER_FIELD]

    placeholders = ", ".join(["?"] * len(insert_cols))
    col_list = ", ".join([f"[{c}]" for c in insert_cols])
    sql = f"INSERT INTO [{table_name}] ({col_list}) VALUES ({placeholders})"

    cur = conn.cursor()
    params_batch = []
    for row in rows:
        params = [row.get(c, None) for c in insert_cols]
        params_batch.append(params)

    cur.executemany(sql, params_batch)
    conn.commit()


def find_existing_id(conn, table_name: str, gas_id: str, pres_id: str):
    """Return the ID of an existing record matching either GasIDREC or PressuresIDREC, else None."""
    cur = conn.cursor()
    cur.execute(
        f"SELECT ID FROM [{table_name}] WHERE GasIDREC = ? OR PressuresIDREC = ?",
        (gas_id, pres_id)
    )
    row = cur.fetchone()
    return row[0] if row else None


def update_record(conn, table_name: str, rec_id: int, payload: dict):
    """Dynamically UPDATE an existing row by ID. Leaves GasIDREC/PressuresIDREC unchanged."""
    updatable_cols = [c for c in TABLE_COLUMNS if c not in (AUTONUMBER_FIELD, "GasIDREC", "PressuresIDREC")]
    sets = []
    params = []
    for c in updatable_cols:
        if c in payload:
            sets.append(f"[{c}] = ?")
            params.append(payload.get(c))
    if not sets:
        return  # nothing to update
    params.append(rec_id)
    sql = f"UPDATE [{table_name}] SET {', '.join(sets)} WHERE ID = ?"
    cur = conn.cursor()
    cur.execute(sql, params)

def compose_name(well: str | None, layer: str | None, tech: str | None) -> str | None:
    w = (well or "").strip()
    l = (layer or "").strip()
    t = (tech or "").strip()
    if not (w and l and t):
        return None
    return f"{w} - {l} - {t}"
# =========================
# GUI
# =========================

class XYScrollFrame(ttk.Frame):
    def __init__(self, parent, *args, **kwargs):
        super().__init__(parent, *args, **kwargs)

        self.canvas = tk.Canvas(self, borderwidth=0, highlightthickness=0)
        self.viewPort = ttk.Frame(self.canvas)

        self.vsb = ttk.Scrollbar(self, orient="vertical", command=self.canvas.yview)
        self.hsb = ttk.Scrollbar(self, orient="horizontal", command=self.canvas.xview)

        self.canvas.configure(yscrollcommand=self.vsb.set, xscrollcommand=self.hsb.set)

        # Layout
        self.canvas.grid(row=0, column=0, sticky="nsew")
        self.vsb.grid(row=0, column=1, sticky="ns")
        self.hsb.grid(row=1, column=0, sticky="ew")
        self.grid_rowconfigure(0, weight=1)
        self.grid_columnconfigure(0, weight=1)

        # Create window for content
        self.canvas_window = self.canvas.create_window((0, 0), window=self.viewPort, anchor="nw")

        # Update scrollregion when content size changes
        def _on_configure(_):
            self.canvas.configure(scrollregion=self.canvas.bbox("all"))
        self.viewPort.bind("<Configure>", _on_configure)

        # Keep inner frame width in sync when canvas is resized (so vertical scroll works nicely)
        def _on_canvas_configure(event):
            # comment next line if you want true free horizontal expansion without stretching
            # self.canvas.itemconfigure(self.canvas_window, width=event.width)
            pass
        self.canvas.bind("<Configure>", _on_canvas_configure)

        # Mouse wheel (vertical)
        def _on_mousewheel(event):
            # Windows/Mac
            delta = int(-1*(event.delta/120))
            self.canvas.yview_scroll(delta, "units")
        self.canvas.bind_all("<MouseWheel>", _on_mousewheel)

        # Shift + wheel (horizontal)
        def _on_shift_wheel(event):
            delta = int(-1*(event.delta/120))
            self.canvas.xview_scroll(delta, "units")
        self.canvas.bind_all("<Shift-MouseWheel>", _on_shift_wheel)



class App(tk.Tk):
    def __init__(self):
        super().__init__()
        self.title("WM Updater — Gas/Pressure IDs to Access")
        self.geometry("1360x800")

        # Top toolbar
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

        # Notebook with two tabs
        self.nb = ttk.Notebook(self)
        self.nb.pack(fill="both", expand=True)

        
                # Tab 1: Current Wells
        self.tab_current = ttk.Frame(self.nb)
        self.nb.add(self.tab_current, text="Current Wells")

        # --- wrapper for tree + scrollbars
        tree_wrap = ttk.Frame(self.tab_current)
        tree_wrap.pack(fill="both", expand=True)

        # columns list gets set in reload_all; create an empty tree for now
        self.tree = ttk.Treeview(tree_wrap, show="headings")

        style = ttk.Style(self)
        style.theme_use("clam")  # clam supports simple borders nicely
        style.configure("Treeview", borderwidth=1, relief="solid")
        style.configure("Treeview.Heading", borderwidth=1, relief="solid")

        style.configure("Slim.TCheckbutton", padding=0)
        style.configure("TEntry", padding=(4, 2))
        style.configure("TCombobox", padding=(4, 2))

        # scrollbars
        ys = ttk.Scrollbar(tree_wrap, orient="vertical", command=self.tree.yview)
        xs = ttk.Scrollbar(tree_wrap, orient="horizontal", command=self.tree.xview)
        self.tree.configure(yscrollcommand=ys.set, xscrollcommand=xs.set)

        # place
        self.tree.grid(row=0, column=0, sticky="nsew")
        ys.grid(row=0, column=1, sticky="ns")
        xs.grid(row=1, column=0, sticky="ew")
        tree_wrap.rowconfigure(0, weight=1)
        tree_wrap.columnconfigure(0, weight=1)

        # inline edit + checkbox toggles
        self.tree.bind("<Button-1>", self.on_tree_click)           # toggle checkbox when clicking Select column
        self.tree.bind("<Double-1>", self.on_tree_double_click)    # start inline edit on double-click

        # footer for save action
        current_footer = ttk.Frame(self.tab_current)
        current_footer.pack(fill="x", padx=8, pady=(4, 8))
        ttk.Button(current_footer, text="Save checked edits → Access", command=self.save_checked_edits).pack(side="right")# Tab 1: Current Wells
        
        
        
        
        # Tab 2: Add New
        self.tab_add = ttk.Frame(self.nb)
        self.nb.add(self.tab_add, text="Add New IDs")

        head = ttk.Frame(self.tab_add)
        head.pack(fill="x", padx=8, pady=6)
        ttk.Label(head, text="Select rows to insert. Well Name is optional for now; other fields via dropdown.").pack(side="left")

        #self.scroll = XYScrollFrame(self.tab_add)
        #self.scroll.pack(fill="both", expand=True, padx=8, pady=4)

        self.scroll = XYScrollFrame(self.tab_add)
        self.scroll.pack(fill="both", expand=True, padx=8, pady=4)

        # Footer actions
        footer = ttk.Frame(self)
        footer.pack(fill="x", padx=8, pady=10)
        self.count_label = ttk.Label(footer, text="Ready")
        self.count_label.pack(side="left")
        ttk.Button(footer, text="Update Selected → Access", command=self.do_update).pack(side="right")

        # Data caches
        self.df_current: pd.DataFrame | None = None
        self.dropdown_options: dict[str, list] = {}
        self.new_widgets: list[dict] = []  # each row: {vars/widgets}

        self.reload_all()

    # --------------------- UI callbacks ---------------------
    def pick_db(self):
        path = filedialog.askopenfilename(filetypes=[("Access DB", "*.accdb;*.mdb"), ("All", "*.*")])
        if path:
            self.db_path_var.set(path)

    def reload_all(self):
        try:
            self.df_current = load_access_table(self.db_path_var.get(), self.table_var.get())
        except Exception as e:
            messagebox.showerror("Load Error", f"Failed to load Access table:\n{e}")
            return

        self.tree.delete(*self.tree.get_children())

        cols_present = [c for c in TABLE_COLUMNS if c in self.df_current.columns]
        self.columns_present = ["Select"] + cols_present
        self.tree.configure(columns=self.columns_present)

        for c in self.columns_present:
            self.tree.heading(c, text=c)
            self.tree.column(
                c,
                width=180 if c != "Select" else 70,
                anchor="w" if c != "Select" else "center",
                stretch=False,
            )

        self._checked = set()
        self._pending_edits = {}
        self._cell_editor = None

        # zebra striping (you already had this)
        self.tree.tag_configure("even", background="#ffffff")
        self.tree.tag_configure("odd",  background="#f7f7f7")

        for idx, row in self.df_current.iterrows():
            iid = str(row["ID"]) if "ID" in row and pd.notna(row["ID"]) else str(idx)
            values = ["☐"] + [row.get(c, "") for c in cols_present]
            tag = "odd" if (idx % 2) else "even"
            self.tree.insert("", "end", iid=iid, values=values, tags=(tag,))

        # dropdown options from existing data
        self.dropdown_options = get_unique_options(self.df_current)

        # build input rows for new Gas/Pressure IDs directly from Access (blank Well Names)
        pending = self.df_current[
            (self.df_current["Well Name"].isnull()) | (self.df_current["Well Name"].astype(str).str.strip() == "")
        ][["GasIDREC", "PressuresIDREC"]]

        self.new_ids = pending.to_dict(orient="records")
        self.build_add_rows()

        self.count_label.config(text=f"Loaded {len(self.df_current)} existing rows • {len(self.new_ids)} pending IDs")



    def build_add_rows(self):
        for child in list(self.scroll.viewPort.children.values()):
            child.destroy()
        self.new_widgets.clear()

        table = ttk.Frame(self.scroll.viewPort)
        table.pack(fill="both", expand=True, padx=8, pady=4)

        # Header: make Select super narrow; others reasonable
        headers = ["", "GasIDREC", "PressuresIDREC", *ENTRY_FIELDS, *DROPDOWN_FIELDS, "Composite name"]
        col_widths = (
            [36, 150, 150] +                 # Select, GasIDREC, PressuresIDREC
            [200] * len(ENTRY_FIELDS) +      # text entry fields
            [180] * len(DROPDOWN_FIELDS) +   # dropdown fields
            [240]                             # Composite
        )

        # header row — bordered labels so it looks like a grid
        for ci, title in enumerate(headers):
            text = "✓" if ci == 0 else title
            hdr = tk.Label(
                table, text=text, font=("Segoe UI", 9, "bold"),
                bg="#f3f3f3", bd=1, relief="solid", anchor="center"
            )
            hdr.grid(row=0, column=ci, sticky="nsew", padx=0, pady=0, ipadx=4, ipady=3)
            weight = 0 if ci in (0, 1, 2) else 1
            table.grid_columnconfigure(ci, minsize=col_widths[ci], weight=weight, uniform="addcols")

        # helper: make a bordered cell that holds a widget
        def wrap_cell(parent, row, col):
            wrapper = tk.Frame(parent, bd=1, relief="solid")
            wrapper.grid(row=row, column=col, sticky="nsew", padx=0, pady=0)
            return wrapper

        # rows
        for ri, rec in enumerate(self.new_ids, start=1):
            # Select checkbox (centered, slim)
            cell0 = wrap_cell(table, ri, 0)
            var_sel = tk.BooleanVar(value=True)
            chk = ttk.Checkbutton(cell0, variable=var_sel, style="Slim.TCheckbutton")
            chk.pack(anchor="center", padx=0, pady=0)

            # GasID / PressuresID — bordered labels
            tk.Label(wrap_cell(table, ri, 1), text=str(rec.get("GasIDREC") or ""), anchor="w").pack(
                fill="x", expand=True, padx=4, pady=2
            )
            tk.Label(wrap_cell(table, ri, 2), text=str(rec.get("PressuresIDREC") or ""), anchor="w").pack(
                fill="x", expand=True, padx=4, pady=2
            )

            # entry fields — ttk.Entry inside a bordered cell
            entry_vars = {}
            col_index = 3
            for col in ENTRY_FIELDS:
                v = tk.StringVar(value="")
                w = wrap_cell(table, ri, col_index)
                ttk.Entry(w, textvariable=v).pack(fill="x", expand=True, padx=4, pady=2)
                entry_vars[col] = v
                col_index += 1

            # dropdown fields — ttk.Combobox inside a bordered cell
            dropdown_vars = {}
            for col in DROPDOWN_FIELDS:
                v = tk.StringVar(value="")
                w = wrap_cell(table, ri, col_index)
                ttk.Combobox(w, textvariable=v, values=self.dropdown_options.get(col, []), state="readonly") \
                    .pack(fill="x", expand=True, padx=4, pady=2)
                dropdown_vars[col] = v
                col_index += 1

            # Composite (read-only) — label inside a bordered cell
            comp_var = tk.StringVar(value="")
            w_comp = wrap_cell(table, ri, col_index)
            ttk.Label(w_comp, textvariable=comp_var).pack(fill="x", expand=True, padx=4, pady=2)

            # keep Composite in sync with the three parts
            def _sync_comp_inner(*_):
                wname = entry_vars.get("Well Name").get() if "Well Name" in entry_vars else ""
                layer = dropdown_vars.get("Layer Producer").get() if "Layer Producer" in dropdown_vars else ""
                tech  = dropdown_vars.get("Completions Technology").get() if "Completions Technology" in dropdown_vars else ""
                comp_var.set(compose_name(wname, layer, tech) or "")

            if "Well Name" in entry_vars:
                entry_vars["Well Name"].trace_add("write", _sync_comp_inner)
            if "Layer Producer" in dropdown_vars:
                dropdown_vars["Layer Producer"].trace_add("write", _sync_comp_inner)
            if "Completions Technology" in dropdown_vars:
                dropdown_vars["Completions Technology"].trace_add("write", _sync_comp_inner)
            _sync_comp_inner()

            # stash for save
            self.new_widgets.append({
                "selected": var_sel,
                "gas": rec.get("GasIDREC"),
                "pres": rec.get("PressuresIDREC"),
                "entries": entry_vars,
                "dropdowns": dropdown_vars,
                "comp_var": comp_var,
            })



    
    def on_tree_click(self, event):
        """Toggle checkbox if user clicked the Select column; otherwise let Treeview handle selection."""
        region = self.tree.identify("region", event.x, event.y)
        if region != "cell":
            return
        column = self.tree.identify_column(event.x)  # '#1' = first displayed col
        if column != "#1":
            return  # only handle clicks in Select column
        item = self.tree.identify_row(event.y)
        if not item:
            return
        # toggle
        current = self.tree.set(item, "Select")
        new = "☑" if current != "☑" else "☐"
        self.tree.set(item, "Select", new)
        if new == "☑":
            self._checked.add(item)
        else:
            self._checked.discard(item)

    def on_tree_double_click(self, event):
        region = self.tree.identify("region", event.x, event.y)
        if region != "cell":
            return
        item = self.tree.identify_row(event.y)
        col_id = self.tree.identify_column(event.x)   # '#1', '#2', ...
        if not item or not col_id:
            return

        # map '#n' to column name
        try:
            col_index = int(col_id.replace("#", "")) - 1
        except Exception:
            return
        if col_index < 0 or col_index >= len(self.columns_present):
            return
        col_name = self.columns_present[col_index]

        # not editable?
        if col_name in ("Select", AUTONUMBER_FIELD, "GasIDREC", "PressuresIDREC"):
            return
        if col_name not in EDITABLE_FIELDS:
            return

        # get current cell bbox
        bbox = self.tree.bbox(item, col_id)
        if not bbox:
            return
        x, y, w, h = bbox

        # existing value
        current_val = self.tree.set(item, col_name)

        # destroy old editor if any
        if self._cell_editor is not None:
            try:
                self._cell_editor.destroy()
            except Exception:
                pass
            self._cell_editor = None

        # dropdowns use Combobox; others use Entry
        if col_name in self.dropdown_options:
            editor = ttk.Combobox(self.tree, values=self.dropdown_options[col_name], state="readonly")
            editor.set(current_val)
        else:
            editor = ttk.Entry(self.tree)
            editor.insert(0, str(current_val) if current_val is not None else "")

        # place over cell
        editor.place(x=x, y=y, width=w, height=h)
        editor.focus_set()
        self._cell_editor = editor

        def _finish(event=None):
            # pull value and remove the overlay editor
            val = editor.get().strip()
            try:
                editor.destroy()
            finally:
                self._cell_editor = None

            # write to the tree cell
            self.tree.set(item, col_name, val)
            # remember as a pending edit (None means clear)
            self._pending_edits.setdefault(item, {})[col_name] = (val if val != "" else None)

            # if a Composite component changed, recompute & queue it too
            if col_name in ("Well Name", "Layer Producer", "Completions Technology"):
                comp = compose_name(
                    self.tree.set(item, "Well Name"),
                    self.tree.set(item, "Layer Producer"),
                    self.tree.set(item, "Completions Technology"),
                )
                if "Composite name" in self.columns_present:
                    self.tree.set(item, "Composite name", comp or "")
                self._pending_edits.setdefault(item, {})["Composite name"] = comp

        # IMPORTANT: these binds must be OUTSIDE the _finish function
        editor.bind("<Return>", _finish)
        editor.bind("<Escape>", lambda e: (editor.destroy(), setattr(self, "_cell_editor", None)))
        editor.bind("<FocusOut>", _finish)
        editor.bind("<Tab>", _finish)  # optional: tab commits too




    def save_checked_edits(self):
        """Apply pending edits for checked rows to Access."""
        if not self._checked:
            messagebox.showinfo("No rows checked", "Check the rows you want to save, then try again.")
            return

        # collect updates: only for checked rows, only columns with pending edits
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
                # iid is Access ID if the item was inserted that way
                try:
                    rec_id = int(iid)
                except:
                    # fallback: try to resolve by IDs in the row
                    row_vals = {c: self.tree.set(iid, c) for c in self.columns_present if c != "Select"}
                    rec_id = find_existing_id(conn, table, row_vals.get("GasIDREC"), row_vals.get("PressuresIDREC"))

                if not rec_id:
                    failed += 1
                    continue

                # sanity: don’t try to update non-editables or ID columns
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

        # reset pending for saved rows
        for iid in list(self._checked):
            self._pending_edits.pop(iid, None)

        messagebox.showinfo("Done", f"Updated: {updated}\nFailed: {failed}")
        self.reload_all()

    def do_update(self):
        # collect selected payloads from the Add New tab
        rows = []
        for item in self.new_widgets:
            if not item["selected"].get():
                continue
            payload = {
                "GasIDREC": item["gas"],
                "PressuresIDREC": item["pres"],
            }
            # gather entries + dropdowns
            for col, v in item["entries"].items():
                payload[col] = v.get().strip() or None
            for col, v in item["dropdowns"].items():
                payload[col] = v.get().strip() or None
            
            payload["Composite name"] = item["comp_var"].get() or None
            
            rows.append(payload)

        if not rows:
            messagebox.showinfo("Nothing selected", "No rows were checked. Tick the boxes to select rows to apply.")
            return

        db_path = self.db_path_var.get()
        table = self.table_var.get()

        updated = 0
        inserted = 0
        skipped = 0

        with connect_access(db_path) as conn:
            to_insert = []
            for r in rows:
                # If Well Name provided and already exists on a different row, warn
                wn = r.get("Well Name")
                if wn:
                    cur = conn.cursor()
                    cur.execute(f"SELECT COUNT(1) FROM [{table}] WHERE [Well Name] = ?", (wn,))
                    if cur.fetchone()[0] > 0:
                        if not messagebox.askyesno("Duplicate Well Name", f"'{wn}' already exists. Continue?"):
                            skipped += 1
                            continue

                # Try to find existing row by IDs
                rec_id = find_existing_id(conn, table, r.get("GasIDREC"), r.get("PressuresIDREC"))
                if rec_id:
                    try:
                        update_record(conn, table, rec_id, r)
                        updated += 1
                    except Exception as e:
                        messagebox.showerror("Update Error", f"Failed to update ID={rec_id}:\n{e}")
                        return
                else:
                    # No existing row — stage for insert
                    to_insert.append(r)

            # Do any staged inserts in batch
            if to_insert:
                try:
                    insert_records(conn, table, to_insert)
                    inserted += len(to_insert)
                except Exception as e:
                    messagebox.showerror("Insert Error", f"Failed to insert records:\n{e}")
                    return

            conn.commit()

        messagebox.showinfo("Done", f"Updated: {updated}\nInserted: {inserted}\nSkipped: {skipped}")

        # Only rebuild the UI if we actually changed the DB.
        # If everything was skipped (e.g., you hit "No" on the duplicate prompt),
        # keep the current inputs so you can edit and try again.
        if (updated + inserted) > 0:
            self.reload_all()
        else:
            # No-op refresh: leave the current Add New rows exactly as they are
            # so you can change values without retyping.
            pass





if __name__ == "__main__":
    # Optional: quiet the pandas warning about non-SQLAlchemy connection
    import warnings
    warnings.filterwarnings("ignore", category=UserWarning, module="pandas")

    app = App()
    app.mainloop()
