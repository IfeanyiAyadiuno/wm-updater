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
    "Well Name", "Inital Flow Date", "ES Well Name", "Lateral Length", "Value Navigator UWI", "Composite name"
]

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


# =========================
# GUI
# =========================

class ScrollFrame(ttk.Frame):
    def __init__(self, parent, *args, **kwargs):
        super().__init__(parent, *args, **kwargs)
        canvas = tk.Canvas(self, borderwidth=0, highlightthickness=0)
        self.viewPort = ttk.Frame(canvas)
        vsb = ttk.Scrollbar(self, orient="vertical", command=canvas.yview)
        canvas.configure(yscrollcommand=vsb.set)

        vsb.pack(side="right", fill="y")
        canvas.pack(side="left", fill="both", expand=True)
        canvas.create_window((0, 0), window=self.viewPort, anchor="nw")

        self.viewPort.bind("<Configure>", lambda e: canvas.configure(scrollregion=canvas.bbox("all")))
        self.viewPort.bind_all("<MouseWheel>", lambda e: canvas.yview_scroll(int(-1*(e.delta/120)), "units"))


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

        # Tab 1: Current Table
        self.tab_current = ttk.Frame(self.nb)
        self.nb.add(self.tab_current, text="Current Table")

        self.tree = ttk.Treeview(self.tab_current, columns=TABLE_COLUMNS, show="headings")
        for col in TABLE_COLUMNS:
            self.tree.heading(col, text=col)
            self.tree.column(col, width=130, anchor="w")
        self.tree.pack(fill="both", expand=True)

        # Tab 2: Add New
        self.tab_add = ttk.Frame(self.nb)
        self.nb.add(self.tab_add, text="Add New IDs")

        head = ttk.Frame(self.tab_add)
        head.pack(fill="x", padx=8, pady=6)
        ttk.Label(head, text="Select rows to insert. Well Name is optional for now; other fields via dropdown.").pack(side="left")

        self.scroll = ScrollFrame(self.tab_add)
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

        # populate current table Treeview
        self.tree.delete(*self.tree.get_children())
        # align Treeview columns to available data
        cols_present = [c for c in TABLE_COLUMNS if c in self.df_current.columns]
        self.tree.configure(columns=cols_present)
        for c in cols_present:
            self.tree.heading(c, text=c)
        for _, row in self.df_current.iterrows():
            values = [row.get(c, "") for c in cols_present]
            self.tree.insert("", "end", values=values)

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
        # clear prior widgets
        for child in list(self.scroll.viewPort.children.values()):
            child.destroy()
        self.new_widgets.clear()

        # header row
        header = [
            "Select", "GasIDREC", "PressuresIDREC", *ENTRY_FIELDS, *DROPDOWN_FIELDS
        ]
        hdr = ttk.Frame(self.scroll.viewPort)
        hdr.pack(fill="x", pady=(0, 4))
        for i, t in enumerate(header):
            ttk.Label(hdr, text=t, font=("Segoe UI", 9, "bold")).grid(row=0, column=i, sticky="w", padx=6)

        # value rows
        for idx, rec in enumerate(self.new_ids):
            rowf = ttk.Frame(self.scroll.viewPort)
            rowf.pack(fill="x", pady=2)

            var_sel = tk.BooleanVar(value=True)
            chk = ttk.Checkbutton(rowf, variable=var_sel)
            chk.grid(row=0, column=0, padx=6, sticky="w")

            ttk.Label(rowf, text=str(rec.get("GasIDREC", ""))).grid(row=0, column=1, padx=6, sticky="w")
            ttk.Label(rowf, text=str(rec.get("PressuresIDREC", ""))).grid(row=0, column=2, padx=6, sticky="w")

            # entries for ENTRY_FIELDS
            entry_vars = {}
            base_col = 3
            for j, col in enumerate(ENTRY_FIELDS):
                v = tk.StringVar(value="" if col == "Well Name" else "")
                e = ttk.Entry(rowf, textvariable=v, width=22)
                e.grid(row=0, column=base_col + j, padx=6, sticky="w")
                entry_vars[col] = v

            # dropdowns for DROPDOWN_FIELDS
            dropdown_vars = {}
            base2 = base_col + len(ENTRY_FIELDS)
            for k, col in enumerate(DROPDOWN_FIELDS):
                v = tk.StringVar(value="")
                choices = self.dropdown_options.get(col, [])
                cb = ttk.Combobox(rowf, textvariable=v, values=choices, width=20, state="readonly")
                cb.grid(row=0, column=base2 + k, padx=6, sticky="w")
                dropdown_vars[col] = v

            self.new_widgets.append({
                "selected": var_sel,
                "gas": rec.get("GasIDREC"),
                "pres": rec.get("PressuresIDREC"),
                "entries": entry_vars,
                "dropdowns": dropdown_vars,
            })

    def do_update(self):
        # collect selected payloads from UI
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
        self.reload_all()




if __name__ == "__main__":
    # Optional: quiet the pandas warning about non-SQLAlchemy connection
    import warnings
    warnings.filterwarnings("ignore", category=UserWarning, module="pandas")

    app = App()
    app.mainloop()
