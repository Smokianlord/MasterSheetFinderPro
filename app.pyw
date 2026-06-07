import csv
import json
import os
import re
import tkinter as tk
from tkinter import filedialog, messagebox, ttk

try:
    import openpyxl
except Exception:
    openpyxl = None

APP_TITLE = "Master Sheet Finder Pro"
APP_VERSION = "v6"
SETTINGS_FILE = "master_sheet_finder_settings.json"

COLORS = {
    "navy": "#0f172a",
    "navy2": "#111827",
    "blue": "#2563eb",
    "purple": "#7c3aed",
    "cyan": "#0891b2",
    "green": "#16a34a",
    "orange": "#ea580c",
    "red": "#dc2626",
    "pink": "#db2777",
    "bg": "#eef2ff",
    "card": "#ffffff",
    "soft": "#f8fafc",
    "line": "#dbe3f0",
    "text": "#0f172a",
    "muted": "#64748b",
    "table_bg": "#ffffff",
    "table_alt": "#f8fafc",
    "table_selected": "#bfdbfe",
    "bad_bg": "#fee2e2",
    "ok_bg": "#dcfce7",
    "warn_bg": "#ffedd5",
}

SENSITIVE_HINTS = (
    "password",
    "pass",
    "token",
    "secret",
    "auth",
)

SKIN_SITES = [
    ("DMarket", ("dmarket", "d market")),
    ("KeyDrop", ("keydrop", "key drop", "key-drop")),
    ("ForceDrop", ("forcedrop", "force drop")),
    ("CSGOFast", ("csgofastx", "csgofast", "csgo fast")),
    ("Hellcase", ("hellcase",)),
    ("CSGO.Net", ("csgo.net", "csgonet", "csgo net")),
    ("SkinClub", ("skinclub", "skin club")),
    ("Juice.gg", ("juice.gg", "juicegg", "juice")),
    ("Casehug", ("casehug", "case hug")),
    ("DaddySkins", ("daddyskins", "daddy skins")),
]

SEARCH_FIELDS = [
    "All fields",
    "Profile Name",
    "Steam Name",
    "Steam Username",
    "Email",
    "SteamID64",
    "Tradelink",
] + [name for name, _aliases in SKIN_SITES]

GENERAL_FILTERS = [
    "Show all",
    "Has tradelink",
    "Missing tradelink",
    "Has email",
    "Missing email",
    "Has SteamID64",
    "Missing SteamID64",
    "Banned only",
    "Active only",
    "Has any skin site",
    "Missing all skin sites",
]

SITE_STATUS_FILTERS = [
    "Any status",
    "Has value",
    "Missing",
    "Banned",
    "Not banned",
]

COLUMN_VIEWS = [
    "Clean view",
    "All safe columns",
    "All columns",
]

BAD_SITE_VALUES = {
    "",
    "banned",
    "ban",
    "n/a",
    "na",
    "none",
    "null",
    "-",
    "--",
}

FIELD_ALIASES = {
    "Profile Name": ("profile name", "profile"),
    "Steam Name": ("steam name",),
    "Steam Username": ("steam username", "steam user", "username"),
    "Steam Password": ("steam password",),
    "Email": ("email",),
    "Email Password": ("email password",),
    "SteamID64": ("steamid64", "steam id64", "steam id"),
}

IMPORTANT_FIELDS = [
    "Profile Name",
    "Steam Name",
    "Steam Username",
    "Email",
    "SteamID64",
]


class SheetFinderPro(tk.Tk):
    def __init__(self):
        super().__init__()
        self.title(f"{APP_TITLE} {APP_VERSION}")
        self.geometry("1480x860")
        self.minsize(1160, 700)
        self.configure(bg=COLORS["bg"])

        self.rows = []
        self.headers = []
        self.filtered_rows = []
        self.file_path = None
        self.detail_widgets = []
        self.last_copied_value = ""

        self.hidden_sensitive = tk.BooleanVar(value=True)
        self.search_var = tk.StringVar()
        self.search_field_var = tk.StringVar(value="All fields")
        self.filter_var = tk.StringVar(value="Show all")
        self.site_var = tk.StringVar(value="All sites")
        self.site_status_var = tk.StringVar(value="Any status")
        self.site_copy_var = tk.StringVar(value="SkinClub")
        self.column_view_var = tk.StringVar(value="Clean view")
        self.status_var = tk.StringVar(value="Open your Master Sheet CSV/XLSX export to start.")

        self.stat_total_var = tk.StringVar(value="0")
        self.stat_showing_var = tk.StringVar(value="0")
        self.stat_trade_var = tk.StringVar(value="0")
        self.stat_skin_var = tk.StringVar(value="0")
        self.stat_banned_var = tk.StringVar(value="0")

        self.configure_icon()
        self.configure_style()
        self.build_ui()
        self.load_settings()

    def app_dir(self):
        return os.path.dirname(os.path.abspath(__file__))

    def configure_icon(self):
        icon_path = os.path.join(self.app_dir(), "app_icon.png")
        if os.path.exists(icon_path):
            try:
                self.icon_img = tk.PhotoImage(file=icon_path)
                self.iconphoto(True, self.icon_img)
            except Exception:
                pass

    def configure_style(self):
        style = ttk.Style(self)
        try:
            style.theme_use("clam")
        except Exception:
            pass

        self.option_add("*Font", "Arial 10")
        self.option_add("*TCombobox*Listbox.font", "Arial 10")

        style.configure("TFrame", background=COLORS["bg"])
        style.configure("Card.TFrame", background=COLORS["card"], relief="flat")
        style.configure("TLabel", background=COLORS["bg"], foreground=COLORS["text"])
        style.configure("Card.TLabel", background=COLORS["card"], foreground=COLORS["text"])
        style.configure("Muted.Card.TLabel", background=COLORS["card"], foreground=COLORS["muted"])
        style.configure("Title.Card.TLabel", background=COLORS["card"], foreground=COLORS["text"], font=("Arial", 13, "bold"))
        style.configure("Header.TLabel", background=COLORS["navy"], foreground="#ffffff")
        style.configure("HeaderTitle.TLabel", background=COLORS["navy"], foreground="#ffffff", font=("Arial", 20, "bold"))
        style.configure("HeaderSub.TLabel", background=COLORS["navy"], foreground="#cbd5e1", font=("Arial", 10))
        style.configure("TCheckbutton", background=COLORS["card"], foreground=COLORS["text"])

        style.configure(
            "Treeview",
            background=COLORS["table_bg"],
            fieldbackground=COLORS["table_bg"],
            foreground=COLORS["text"],
            bordercolor=COLORS["line"],
            rowheight=30,
            font=("Arial", 9),
        )
        style.configure(
            "Treeview.Heading",
            background=COLORS["navy2"],
            foreground="#ffffff",
            relief="flat",
            font=("Arial", 9, "bold"),
        )
        style.map("Treeview", background=[("selected", COLORS["table_selected"])] , foreground=[("selected", COLORS["text"])] )
        style.map("Treeview.Heading", background=[("active", COLORS["blue"])])

        style.configure("TCombobox", fieldbackground="#ffffff", background="#ffffff", foreground=COLORS["text"])
        style.configure("TEntry", fieldbackground="#ffffff", foreground=COLORS["text"])
        style.configure("Horizontal.TPanedwindow", background=COLORS["bg"])

    def build_ui(self):
        self.build_header()
        content = tk.Frame(self, bg=COLORS["bg"])
        content.pack(fill="both", expand=True)

        controls = self.card(content)
        controls.pack(fill="x", padx=14, pady=(14, 10))
        self.build_controls(controls)

        stats = tk.Frame(content, bg=COLORS["bg"])
        stats.pack(fill="x", padx=14, pady=(0, 10))
        self.build_stats(stats)

        body = tk.PanedWindow(content, orient="horizontal", sashwidth=8, bg=COLORS["bg"], bd=0, relief="flat")
        body.pack(fill="both", expand=True, padx=14, pady=(0, 10))

        left = self.card(body)
        right = self.card(body)
        body.add(left, minsize=700)
        body.add(right, minsize=380)

        self.build_results_panel(left)
        self.build_details_panel(right)

        status_bar = tk.Frame(self, bg=COLORS["navy2"], height=28)
        status_bar.pack(fill="x", side="bottom")
        self.status_label = tk.Label(
            status_bar,
            textvariable=self.status_var,
            bg=COLORS["navy2"],
            fg="#e2e8f0",
            anchor="w",
            padx=12,
            font=("Arial", 9),
        )
        self.status_label.pack(fill="both", expand=True)

    def build_header(self):
        header = tk.Frame(self, bg=COLORS["navy"], height=88)
        header.pack(fill="x")
        header.pack_propagate(False)

        left = tk.Frame(header, bg=COLORS["navy"])
        left.pack(side="left", fill="both", expand=True, padx=18, pady=13)

        tk.Label(left, text="Master Sheet Finder Pro", bg=COLORS["navy"], fg="#ffffff", font=("Arial", 22, "bold")).pack(anchor="w")
        tk.Label(left, text="Fast account search, skin-site filters, and one-click copy tools", bg=COLORS["navy"], fg="#cbd5e1", font=("Arial", 10)).pack(anchor="w", pady=(2, 0))

        right = tk.Frame(header, bg=COLORS["navy"])
        right.pack(side="right", padx=18, pady=18)
        self.header_file_label = tk.Label(right, text="No sheet loaded", bg=COLORS["navy"], fg="#cbd5e1", anchor="e", font=("Arial", 10))
        self.header_file_label.pack(anchor="e")
        tk.Label(right, text=APP_VERSION, bg=COLORS["blue"], fg="#ffffff", padx=10, pady=2, font=("Arial", 9, "bold")).pack(anchor="e", pady=(7, 0))

    def build_controls(self, parent):
        parent.columnconfigure(1, weight=1)
        parent.columnconfigure(4, weight=0)

        top = tk.Frame(parent, bg=COLORS["card"])
        top.pack(fill="x", padx=14, pady=(12, 8))
        top.columnconfigure(1, weight=1)

        self.color_button(top, "Open CSV / XLSX", self.open_file, COLORS["blue"]).grid(row=0, column=0, rowspan=2, padx=(0, 12), pady=2, sticky="nsw")

        tk.Label(top, text="Search", bg=COLORS["card"], fg=COLORS["muted"], font=("Arial", 9, "bold")).grid(row=0, column=1, sticky="w")
        search = tk.Entry(top, textvariable=self.search_var, relief="solid", bd=1, bg="#ffffff", fg=COLORS["text"], insertbackground=COLORS["text"], font=("Arial", 12))
        search.grid(row=1, column=1, padx=(0, 12), ipady=7, sticky="ew")
        search.bind("<KeyRelease>", lambda e: self.apply_filter())
        search.focus_set()

        tk.Label(top, text="Search in", bg=COLORS["card"], fg=COLORS["muted"], font=("Arial", 9, "bold")).grid(row=0, column=2, sticky="w")
        search_in = ttk.Combobox(top, textvariable=self.search_field_var, values=SEARCH_FIELDS, state="readonly", width=18)
        search_in.grid(row=1, column=2, padx=(0, 12), ipady=5, sticky="w")
        search_in.bind("<<ComboboxSelected>>", self.on_search_field_changed)

        self.color_button(top, "Clear", self.clear_search, COLORS["purple"]).grid(row=1, column=3, padx=(0, 8), sticky="w")
        self.color_button(top, "Save filtered CSV", self.export_filtered_csv, COLORS["cyan"]).grid(row=1, column=4, sticky="w")

        filters = tk.Frame(parent, bg=COLORS["card"])
        filters.pack(fill="x", padx=14, pady=(0, 12))
        for c in range(10):
            filters.columnconfigure(c, weight=0)
        filters.columnconfigure(9, weight=1)

        self.labeled_combo(filters, "General filter", self.filter_var, GENERAL_FILTERS, 0, 0, 20, lambda e: self.apply_filter())
        self.labeled_combo(filters, "Skin site", self.site_var, ["All sites"] + self.site_names(), 0, 2, 16, self.on_site_filter_changed)
        self.labeled_combo(filters, "Site status", self.site_status_var, SITE_STATUS_FILTERS, 0, 4, 14, lambda e: self.apply_filter())
        self.labeled_combo(filters, "Columns", self.column_view_var, COLUMN_VIEWS, 0, 6, 16, lambda e: self.apply_filter())

        check = tk.Checkbutton(
            filters,
            text="Hide passwords/tokens",
            variable=self.hidden_sensitive,
            command=self.apply_filter,
            bg=COLORS["card"],
            fg=COLORS["text"],
            activebackground=COLORS["card"],
            selectcolor="#ffffff",
            font=("Arial", 10),
        )
        check.grid(row=1, column=8, padx=(14, 0), sticky="w")

        selected_site_box = tk.Frame(filters, bg=COLORS["soft"], highlightbackground=COLORS["line"], highlightthickness=1)
        selected_site_box.grid(row=0, column=9, rowspan=2, padx=(14, 0), sticky="e")
        tk.Label(selected_site_box, text="Quick copy site", bg=COLORS["soft"], fg=COLORS["muted"], font=("Arial", 9, "bold")).grid(row=0, column=0, padx=8, pady=(4, 0), sticky="w")
        site_copy = ttk.Combobox(selected_site_box, textvariable=self.site_copy_var, values=self.site_names(), state="readonly", width=14)
        site_copy.grid(row=1, column=0, padx=8, pady=(2, 7), sticky="w")
        site_copy.bind("<<ComboboxSelected>>", lambda e: self.update_copy_site_button())
        self.copy_site_button = self.color_button(selected_site_box, "Copy SkinClub ID/link", self.copy_selected_site_value, COLORS["green"])
        self.copy_site_button.grid(row=1, column=1, padx=(0, 8), pady=(2, 7), sticky="e")

    def labeled_combo(self, parent, label, variable, values, row, col, width, callback):
        box = tk.Frame(parent, bg=COLORS["card"])
        box.grid(row=row, column=col, columnspan=2, padx=(0, 12), pady=2, sticky="w")
        tk.Label(box, text=label, bg=COLORS["card"], fg=COLORS["muted"], font=("Arial", 9, "bold")).pack(anchor="w")
        combo = ttk.Combobox(box, textvariable=variable, values=values, state="readonly", width=width)
        combo.pack(anchor="w", pady=(2, 0), ipady=4)
        combo.bind("<<ComboboxSelected>>", callback)
        return combo

    def build_stats(self, parent):
        cards = [
            ("Total rows", self.stat_total_var, COLORS["blue"]),
            ("Showing", self.stat_showing_var, COLORS["purple"]),
            ("With tradelink", self.stat_trade_var, COLORS["cyan"]),
            ("With skin site", self.stat_skin_var, COLORS["green"]),
            ("Banned", self.stat_banned_var, COLORS["red"]),
        ]
        for i, (title, var, color) in enumerate(cards):
            card = tk.Frame(parent, bg=COLORS["card"], highlightbackground=COLORS["line"], highlightthickness=1)
            card.pack(side="left", fill="x", expand=True, padx=(0 if i == 0 else 8, 0))
            stripe = tk.Frame(card, bg=color, height=4)
            stripe.pack(fill="x")
            tk.Label(card, text=title, bg=COLORS["card"], fg=COLORS["muted"], font=("Arial", 9, "bold")).pack(anchor="w", padx=12, pady=(8, 0))
            tk.Label(card, textvariable=var, bg=COLORS["card"], fg=COLORS["text"], font=("Arial", 20, "bold")).pack(anchor="w", padx=12, pady=(0, 8))

    def build_results_panel(self, parent):
        title = tk.Frame(parent, bg=COLORS["card"])
        title.pack(fill="x", padx=12, pady=(12, 8))
        tk.Label(title, text="Results", bg=COLORS["card"], fg=COLORS["text"], font=("Arial", 14, "bold")).pack(side="left")
        tk.Label(title, text="Double-click a row to copy visible row", bg=COLORS["card"], fg=COLORS["muted"], font=("Arial", 9)).pack(side="right")

        table_frame = tk.Frame(parent, bg=COLORS["card"])
        table_frame.pack(fill="both", expand=True, padx=12, pady=(0, 12))
        table_frame.rowconfigure(0, weight=1)
        table_frame.columnconfigure(0, weight=1)

        self.tree = ttk.Treeview(table_frame, show="headings", selectmode="browse")
        vsb = ttk.Scrollbar(table_frame, orient="vertical", command=self.tree.yview)
        hsb = ttk.Scrollbar(table_frame, orient="horizontal", command=self.tree.xview)
        self.tree.configure(yscrollcommand=vsb.set, xscrollcommand=hsb.set)
        self.tree.grid(row=0, column=0, sticky="nsew")
        vsb.grid(row=0, column=1, sticky="ns")
        hsb.grid(row=1, column=0, sticky="ew")
        self.tree.tag_configure("odd", background=COLORS["table_alt"])
        self.tree.tag_configure("banned", background=COLORS["bad_bg"])
        self.tree.tag_configure("no_site", background=COLORS["warn_bg"])
        self.tree.bind("<<TreeviewSelect>>", lambda e: self.show_details())
        self.tree.bind("<Double-1>", lambda e: self.copy_selected_row())

    def build_details_panel(self, parent):
        top = tk.Frame(parent, bg=COLORS["card"])
        top.pack(fill="x", padx=12, pady=(12, 8))
        tk.Label(top, text="Selected Account", bg=COLORS["card"], fg=COLORS["text"], font=("Arial", 14, "bold")).pack(side="left")
        self.selected_badge = tk.Label(top, text="No row selected", bg=COLORS["soft"], fg=COLORS["muted"], padx=10, pady=4, font=("Arial", 9, "bold"))
        self.selected_badge.pack(side="right")

        canvas_frame = tk.Frame(parent, bg=COLORS["card"])
        canvas_frame.pack(fill="both", expand=True, padx=12, pady=(0, 12))
        canvas_frame.rowconfigure(0, weight=1)
        canvas_frame.columnconfigure(0, weight=1)

        self.detail_canvas = tk.Canvas(canvas_frame, bg=COLORS["card"], highlightthickness=0)
        scrollbar = ttk.Scrollbar(canvas_frame, orient="vertical", command=self.detail_canvas.yview)
        self.detail_inner = tk.Frame(self.detail_canvas, bg=COLORS["card"])
        self.detail_window = self.detail_canvas.create_window((0, 0), window=self.detail_inner, anchor="nw")
        self.detail_canvas.configure(yscrollcommand=scrollbar.set)
        self.detail_canvas.grid(row=0, column=0, sticky="nsew")
        scrollbar.grid(row=0, column=1, sticky="ns")
        self.detail_inner.bind("<Configure>", self.on_detail_configure)
        self.detail_canvas.bind("<Configure>", self.on_detail_canvas_configure)
        self.detail_canvas.bind_all("<MouseWheel>", self.on_mousewheel)

        self.empty_details()

    def on_mousewheel(self, event):
        widget = self.focus_get()
        if widget is not None and str(widget).startswith(str(self.detail_canvas)):
            self.detail_canvas.yview_scroll(int(-1 * (event.delta / 120)), "units")

    def on_detail_configure(self, event=None):
        self.detail_canvas.configure(scrollregion=self.detail_canvas.bbox("all"))

    def on_detail_canvas_configure(self, event):
        self.detail_canvas.itemconfigure(self.detail_window, width=event.width)

    def card(self, parent):
        return tk.Frame(parent, bg=COLORS["card"], highlightbackground=COLORS["line"], highlightthickness=1)

    def color_button(self, parent, text, command, color, fg="#ffffff"):
        btn = tk.Button(
            parent,
            text=text,
            command=command,
            bg=color,
            fg=fg,
            activebackground=color,
            activeforeground=fg,
            relief="flat",
            bd=0,
            padx=12,
            pady=7,
            cursor="hand2",
            font=("Arial", 9, "bold"),
        )
        return btn

    def subtle_button(self, parent, text, command):
        return self.color_button(parent, text, command, COLORS["soft"], COLORS["text"])

    def site_names(self):
        return [name for name, _aliases in SKIN_SITES]

    def on_search_field_changed(self, event=None):
        field = self.search_field_var.get()
        if field in self.site_names():
            self.site_copy_var.set(field)
            self.update_copy_site_button()
        self.apply_filter()

    def on_site_filter_changed(self, event=None):
        site = self.site_var.get()
        if site in self.site_names():
            self.site_copy_var.set(site)
            self.update_copy_site_button()
        self.apply_filter()

    def update_copy_site_button(self):
        if hasattr(self, "copy_site_button"):
            site = self.site_copy_var.get() or "site"
            self.copy_site_button.configure(text=f"Copy {site} ID/link")

    def open_file(self):
        path = filedialog.askopenfilename(
            filetypes=[
                ("Spreadsheet files", "*.csv *.xlsx"),
                ("CSV", "*.csv"),
                ("Excel", "*.xlsx"),
                ("All files", "*.*"),
            ]
        )
        if not path:
            return
        try:
            self.load_file(path)
        except Exception as e:
            messagebox.showerror("Could not open file", str(e))

    def load_file(self, path):
        ext = os.path.splitext(path)[1].lower()
        if ext == ".csv":
            self.load_csv(path)
        elif ext == ".xlsx":
            self.load_xlsx(path)
        else:
            raise RuntimeError("Please choose a .csv or .xlsx file")

        self.file_path = path
        self.save_settings()
        self.header_file_label.configure(text=os.path.basename(path))
        self.apply_filter(select_first=True)
        self.status_var.set(f"Loaded {len(self.rows)} rows and {len(self.headers)} columns from {os.path.basename(path)}")

    def load_csv(self, path):
        with open(path, newline="", encoding="utf-8-sig") as f:
            reader = csv.reader(f)
            raw_headers = next(reader, [])
            self.headers = self.make_unique([self.clean_header(h) for h in raw_headers])
            self.rows = []
            for values in reader:
                values = values + [""] * max(0, len(self.headers) - len(values))
                row = dict(zip(self.headers, values[:len(self.headers)]))
                self.rows.append(self.normalize_row(row))

    def load_xlsx(self, path):
        if openpyxl is None:
            raise RuntimeError("Install openpyxl first for XLSX files: pip install openpyxl")
        wb = openpyxl.load_workbook(path, data_only=True)
        ws = wb.active
        raw_headers = [c.value if c.value is not None else f"Column {i + 1}" for i, c in enumerate(ws[1])]
        self.headers = self.make_unique([self.clean_header(h) for h in raw_headers])
        self.rows = []
        for row in ws.iter_rows(min_row=2, values_only=True):
            values = ["" if v is None else str(v) for v in row]
            values = values + [""] * max(0, len(self.headers) - len(values))
            self.rows.append(self.normalize_row(dict(zip(self.headers, values[:len(self.headers)]))))

    def clean_header(self, header):
        text = str(header).strip() if header is not None else "Column"
        return " ".join(text.split()) or "Column"

    def make_unique(self, headers):
        seen = {}
        out = []
        for h in headers:
            base = h or "Column"
            seen[base] = seen.get(base, 0) + 1
            out.append(base if seen[base] == 1 else f"{base} {seen[base]}")
        return out

    def normalize_row(self, row):
        return {h: "" if row.get(h) is None else str(row.get(h, "")).strip() for h in self.headers}

    def compact_text(self, text):
        return re.sub(r"[^a-z0-9]+", "", str(text).lower())

    def visible_headers(self):
        view = self.column_view_var.get()
        if view == "All columns":
            cols = self.headers[:]
        elif view == "All safe columns":
            cols = [h for h in self.headers if not self.is_sensitive_header(h)]
        else:
            cols = self.clean_view_headers()

        if self.hidden_sensitive.get() and view != "All columns":
            cols = [h for h in cols if not self.is_sensitive_header(h)]
        return cols

    def clean_view_headers(self):
        cols = []
        for field in IMPORTANT_FIELDS:
            h = self.first_header_for_field(field)
            if h and h not in cols:
                cols.append(h)

        selected_site = self.site_copy_var.get()
        search_field = self.search_field_var.get()
        site_filter = self.site_var.get()
        preferred_sites = []
        for s in [search_field, site_filter, selected_site]:
            if s in self.site_names() and s not in preferred_sites:
                preferred_sites.append(s)
        for site in preferred_sites:
            for h in self.skin_site_headers(site):
                if h not in cols:
                    cols.append(h)

        for site, _aliases in SKIN_SITES:
            for h in self.skin_site_headers(site):
                if h not in cols:
                    cols.append(h)

        return [h for h in cols if h in self.headers]

    def is_sensitive_header(self, header):
        h = header.lower()
        return any(s in h for s in SENSITIVE_HINTS)

    def apply_filter(self, select_first=False):
        q = self.search_var.get().strip().lower()
        general_filter = self.filter_var.get()
        site_name = self.site_var.get()
        site_status = self.site_status_var.get()

        rows = []
        for row in self.rows:
            if q and not self.row_matches_search(row, q):
                continue
            if not self.row_matches_general_filter(row, general_filter):
                continue
            if not self.row_matches_site_filter(row, site_name, site_status):
                continue
            rows.append(row)

        self.filtered_rows = rows
        self.refresh_table(select_first=select_first or bool(q))
        self.update_stats()
        self.status_var.set(f"Showing {len(self.filtered_rows)} of {len(self.rows)} rows")

    def row_matches_search(self, row, query):
        field = self.search_field_var.get()
        if field == "All fields":
            return any(query in str(v).lower() for v in row.values())
        return any(query in str(v).lower() for v in self.values_for_field(row, field))

    def values_for_field(self, row, field):
        if field == "Tradelink":
            return [self.extract_tradelink(row)]
        if field in self.site_names():
            return self.skin_site_values(row, field)

        aliases = FIELD_ALIASES.get(field, (field,))
        matches = []
        for h, v in row.items():
            hc = self.compact_text(h)
            for alias in aliases:
                ac = self.compact_text(alias)
                if ac and (hc == ac or ac in hc):
                    if field == "Email" and "pass" in hc:
                        continue
                    matches.append(v)
                    break
        return matches

    def first_header_for_field(self, field):
        aliases = FIELD_ALIASES.get(field, (field,))
        for h in self.headers:
            hc = self.compact_text(h)
            for alias in aliases:
                ac = self.compact_text(alias)
                if ac and (hc == ac or ac in hc):
                    if field == "Email" and "pass" in hc:
                        continue
                    return h
        return None

    def row_matches_general_filter(self, row, chosen_filter):
        if chosen_filter == "Show all":
            return True
        if chosen_filter == "Has tradelink":
            return bool(self.extract_tradelink(row))
        if chosen_filter == "Missing tradelink":
            return not bool(self.extract_tradelink(row))
        if chosen_filter == "Has email":
            return bool(self.first_field_value(row, "Email"))
        if chosen_filter == "Missing email":
            return not bool(self.first_field_value(row, "Email"))
        if chosen_filter == "Has SteamID64":
            return bool(self.first_field_value(row, "SteamID64"))
        if chosen_filter == "Missing SteamID64":
            return not bool(self.first_field_value(row, "SteamID64"))
        if chosen_filter == "Banned only":
            return self.has_banned_value(row)
        if chosen_filter == "Active only":
            return not self.has_banned_value(row)
        if chosen_filter == "Has any skin site":
            return self.has_any_skin_site(row)
        if chosen_filter == "Missing all skin sites":
            return not self.has_any_skin_site(row)
        return True

    def row_matches_site_filter(self, row, site_name, site_status):
        if site_status == "Any status":
            return True

        if site_name == "All sites":
            if site_status == "Has value":
                return self.has_any_skin_site(row)
            if site_status == "Missing":
                return not self.has_any_skin_site(row)
            if site_status == "Banned":
                return self.has_banned_skin_site(row)
            if site_status == "Not banned":
                return not self.has_banned_skin_site(row)
            return True

        if site_status == "Has value":
            return self.has_skin_site(row, site_name)
        if site_status == "Missing":
            return not self.has_skin_site(row, site_name) and not self.is_site_banned(row, site_name)
        if site_status == "Banned":
            return self.is_site_banned(row, site_name)
        if site_status == "Not banned":
            return not self.is_site_banned(row, site_name)
        return True

    def first_field_value(self, row, field):
        for v in self.values_for_field(row, field):
            s = str(v).strip()
            if s:
                return s
        return ""

    def has_banned_value(self, row):
        return any(str(v).strip().lower() == "banned" for v in row.values())

    def header_matches_alias(self, header, alias):
        h = self.compact_text(header)
        a = self.compact_text(alias)
        return a and a in h

    def skin_site_headers(self, site_name):
        aliases = ()
        for name, site_aliases in SKIN_SITES:
            if name == site_name:
                aliases = site_aliases
                break
        if not aliases:
            return []
        return [h for h in self.headers if any(self.header_matches_alias(h, alias) for alias in aliases)]

    def skin_site_values(self, row, site_name):
        values = []
        for h in self.skin_site_headers(site_name):
            s = str(row.get(h, "")).strip()
            if s:
                values.append(s)
        return values

    def is_usable_site_value(self, value):
        cleaned = str(value).strip().lower()
        return cleaned not in BAD_SITE_VALUES

    def has_skin_site(self, row, site_name):
        return any(self.is_usable_site_value(v) for v in self.skin_site_values(row, site_name))

    def is_site_banned(self, row, site_name):
        return any(str(v).strip().lower() == "banned" for v in self.skin_site_values(row, site_name))

    def has_any_skin_site(self, row):
        return any(self.has_skin_site(row, site_name) for site_name, _aliases in SKIN_SITES)

    def has_banned_skin_site(self, row):
        return any(self.is_site_banned(row, site_name) for site_name, _aliases in SKIN_SITES)

    def first_tradelink_header(self):
        for h in self.headers:
            hc = self.compact_text(h)
            if "tradelink" in hc or "tradeoffer" in hc or hc == "trade":
                return h
        return None

    def extract_tradelink(self, row):
        for h, v in row.items():
            hc = self.compact_text(h)
            s = str(v).strip()
            if not s:
                continue
            if "tradelink" in hc or "tradeoffer" in hc or hc == "trade":
                link = self.first_url(s)
                return link or s
        for v in row.values():
            s = str(v).strip()
            if "steamcommunity.com/tradeoffer" in s.lower():
                return self.first_url(s) or s
        return ""

    def first_url(self, text):
        match = re.search(r"https?://[^\s]+", text)
        return match.group(0) if match else ""

    def refresh_table(self, select_first=False):
        selected_index = None
        sel = self.tree.selection() if hasattr(self, "tree") else []
        if sel:
            try:
                selected_index = int(sel[0])
            except Exception:
                selected_index = None

        for c in self.tree.get_children():
            self.tree.delete(c)

        cols = self.visible_headers()
        self.tree["columns"] = cols

        for h in cols:
            self.tree.heading(h, text=h)
            hc = self.compact_text(h)
            width = 130
            if hc in ("profilename", "steamname", "steamusername", "email"):
                width = 175
            if "link" in hc or "trade" in hc:
                width = 280
            if any(self.header_matches_alias(h, alias) for _site, aliases in SKIN_SITES for alias in aliases):
                width = 125
            self.tree.column(h, width=width, anchor="w", stretch=True)

        for idx, row in enumerate(self.filtered_rows):
            values = [row.get(h, "") for h in cols]
            tags = ["odd" if idx % 2 else "even"]
            if self.has_banned_value(row):
                tags.append("banned")
            elif not self.has_any_skin_site(row):
                tags.append("no_site")
            self.tree.insert("", "end", iid=str(idx), values=values, tags=tags)

        if self.filtered_rows:
            target = 0 if select_first or selected_index is None or selected_index >= len(self.filtered_rows) else selected_index
            self.tree.selection_set(str(target))
            self.tree.focus(str(target))
            self.tree.see(str(target))
        self.show_details()

    def selected_row(self):
        sel = self.tree.selection()
        if not sel:
            return None
        try:
            return self.filtered_rows[int(sel[0])]
        except Exception:
            return None

    def show_details(self):
        row = self.selected_row()
        self.clear_detail_inner()
        if not row:
            self.empty_details()
            return

        display_name = self.first_field_value(row, "Profile Name") or self.first_field_value(row, "Steam Name") or "Selected row"
        status = "Banned" if self.has_banned_value(row) else "Active / not banned"
        badge_bg = COLORS["bad_bg"] if status == "Banned" else COLORS["ok_bg"]
        badge_fg = COLORS["red"] if status == "Banned" else COLORS["green"]
        self.selected_badge.configure(text=status, bg=badge_bg, fg=badge_fg)

        self.section_title("Profile summary", COLORS["blue"])
        summary = [
            ("Profile Name", self.first_field_value(row, "Profile Name")),
            ("Steam Name", self.first_field_value(row, "Steam Name")),
            ("Steam Username", self.first_field_value(row, "Steam Username")),
            ("Email", self.first_field_value(row, "Email")),
            ("SteamID64", self.first_field_value(row, "SteamID64")),
            ("Tradelink", self.extract_tradelink(row)),
        ]
        for label, value in summary:
            self.field_row(label, value, copy_label="Copy" if value else None, copy_command=lambda v=value, l=label: self.copy_to_clipboard(v, f"Copied {l}"))

        self.section_title("Quick copy", COLORS["purple"])
        quick = tk.Frame(self.detail_inner, bg=COLORS["card"])
        quick.pack(fill="x", pady=(0, 10))
        buttons = [
            ("Steam username", lambda: self.copy_named_field("Steam Username"), COLORS["blue"]),
            ("Steam password", lambda: self.copy_named_field("Steam Password"), COLORS["orange"]),
            ("Email", lambda: self.copy_named_field("Email"), COLORS["cyan"]),
            ("Email password", lambda: self.copy_named_field("Email Password"), COLORS["orange"]),
            ("SteamID64", lambda: self.copy_named_field("SteamID64"), COLORS["purple"]),
            ("Tradelink", self.copy_tradelink, COLORS["green"]),
            (f"{self.site_copy_var.get()} ID/link", self.copy_selected_site_value, COLORS["pink"]),
            ("All details", self.copy_details, COLORS["navy2"]),
        ]
        for i, (text, command, color) in enumerate(buttons):
            self.color_button(quick, text, command, color).grid(row=i // 2, column=i % 2, padx=(0, 8), pady=4, sticky="ew")
        quick.columnconfigure(0, weight=1)
        quick.columnconfigure(1, weight=1)

        self.section_title("Skin sites", COLORS["green"])
        any_site_row = False
        for site, _aliases in SKIN_SITES:
            values = [str(v).strip() for v in self.skin_site_values(row, site) if str(v).strip()]
            usable_values = [v for v in values if self.is_usable_site_value(v)]
            banned = any(v.lower() == "banned" for v in values)
            if usable_values:
                any_site_row = True
                value = usable_values[0]
                self.site_row(site, value, "Has value", COLORS["ok_bg"], COLORS["green"], copyable=True)
            elif banned:
                any_site_row = True
                self.site_row(site, "Banned", "Banned", COLORS["bad_bg"], COLORS["red"], copyable=False)
            else:
                self.site_row(site, "Missing", "Missing", COLORS["soft"], COLORS["muted"], copyable=False)
        if not any_site_row:
            tk.Label(self.detail_inner, text="No skin-site values found.", bg=COLORS["card"], fg=COLORS["muted"]).pack(anchor="w", pady=(0, 10))

        self.section_title("All non-empty fields", COLORS["cyan"])
        non_empty = 0
        for h in self.headers:
            if self.hidden_sensitive.get() and self.is_sensitive_header(h):
                continue
            v = row.get(h, "")
            if v:
                non_empty += 1
                self.field_row(h, v, copy_label="Copy", copy_command=lambda value=v, label=h: self.copy_to_clipboard(value, f"Copied {label}"))
        if non_empty == 0:
            tk.Label(self.detail_inner, text="No visible data in this row.", bg=COLORS["card"], fg=COLORS["muted"]).pack(anchor="w", pady=(0, 10))

        self.detail_canvas.yview_moveto(0)

    def clear_detail_inner(self):
        for child in self.detail_inner.winfo_children():
            child.destroy()

    def empty_details(self):
        self.selected_badge.configure(text="No row selected", bg=COLORS["soft"], fg=COLORS["muted"])
        tk.Label(self.detail_inner, text="No account selected", bg=COLORS["card"], fg=COLORS["text"], font=("Arial", 15, "bold")).pack(anchor="w", pady=(10, 4))
        tk.Label(self.detail_inner, text="Open your sheet, then select a row to see copy buttons and site IDs.", bg=COLORS["card"], fg=COLORS["muted"], wraplength=360, justify="left").pack(anchor="w")

    def section_title(self, text, color):
        wrap = tk.Frame(self.detail_inner, bg=COLORS["card"])
        wrap.pack(fill="x", pady=(12, 6))
        tk.Frame(wrap, bg=color, width=5, height=24).pack(side="left", padx=(0, 8))
        tk.Label(wrap, text=text, bg=COLORS["card"], fg=COLORS["text"], font=("Arial", 12, "bold")).pack(side="left")

    def field_row(self, label, value, copy_label=None, copy_command=None):
        row = tk.Frame(self.detail_inner, bg=COLORS["soft"], highlightbackground=COLORS["line"], highlightthickness=1)
        row.pack(fill="x", pady=3)
        row.columnconfigure(1, weight=1)
        tk.Label(row, text=label, bg=COLORS["soft"], fg=COLORS["muted"], anchor="w", width=16, font=("Arial", 9, "bold")).grid(row=0, column=0, padx=8, pady=7, sticky="nw")
        tk.Message(row, text=value or "-", bg=COLORS["soft"], fg=COLORS["text"], anchor="w", width=280, font=("Arial", 9)).grid(row=0, column=1, padx=(0, 8), pady=7, sticky="ew")
        if copy_label and copy_command:
            self.color_button(row, copy_label, copy_command, COLORS["blue"]).grid(row=0, column=2, padx=(0, 8), pady=5, sticky="e")

    def site_row(self, site, value, status, bg, fg, copyable):
        row = tk.Frame(self.detail_inner, bg=COLORS["card"], highlightbackground=COLORS["line"], highlightthickness=1)
        row.pack(fill="x", pady=3)
        row.columnconfigure(1, weight=1)
        tk.Label(row, text=site, bg=COLORS["card"], fg=COLORS["text"], width=12, anchor="w", font=("Arial", 9, "bold")).grid(row=0, column=0, padx=8, pady=7, sticky="w")
        tk.Message(row, text=value, bg=COLORS["card"], fg=COLORS["text"], width=220, font=("Arial", 9)).grid(row=0, column=1, padx=(0, 8), pady=7, sticky="ew")
        tk.Label(row, text=status, bg=bg, fg=fg, padx=8, pady=3, font=("Arial", 8, "bold")).grid(row=0, column=2, padx=(0, 8), pady=7, sticky="e")
        if copyable:
            self.color_button(row, "Copy", lambda v=value, s=site: self.copy_to_clipboard(v, f"Copied {s} ID/link"), COLORS["green"]).grid(row=0, column=3, padx=(0, 8), pady=5, sticky="e")

    def clear_search(self):
        self.search_var.set("")
        self.search_field_var.set("All fields")
        self.filter_var.set("Show all")
        self.site_var.set("All sites")
        self.site_status_var.set("Any status")
        self.site_copy_var.set("SkinClub")
        self.column_view_var.set("Clean view")
        self.update_copy_site_button()
        self.apply_filter(select_first=True)

    def copy_to_clipboard(self, text, message):
        text = str(text).strip()
        if not text:
            return False
        self.clipboard_clear()
        self.clipboard_append(text)
        self.update()
        self.last_copied_value = text
        self.status_var.set(message)
        self.flash_status(COLORS["green"])
        return True

    def flash_status(self, color):
        try:
            self.status_label.configure(bg=color, fg="#ffffff")
            self.after(800, lambda: self.status_label.configure(bg=COLORS["navy2"], fg="#e2e8f0"))
        except Exception:
            pass

    def copy_named_field(self, field):
        row = self.selected_row()
        if not row:
            messagebox.showinfo("No row selected", "Select an account first.")
            return
        value = self.first_field_value(row, field)
        if self.copy_to_clipboard(value, f"Copied {field}"):
            return
        messagebox.showinfo("No value", f"No {field} found in the selected row.")

    def copy_selected_site_value(self):
        row = self.selected_row()
        if not row:
            messagebox.showinfo("No row selected", "Select an account first.")
            return

        site = self.site_copy_var.get()
        values = [str(v).strip() for v in self.skin_site_values(row, site) if str(v).strip()]
        usable_values = [v for v in values if self.is_usable_site_value(v)]
        value = usable_values[0] if usable_values else ""

        if self.copy_to_clipboard(value, f"Copied {site} ID/link"):
            return
        messagebox.showinfo("No site ID/link", f"No {site} ID/link found in the selected row.")

    def copy_details(self):
        row = self.selected_row()
        if not row:
            messagebox.showinfo("No row selected", "Select an account first.")
            return
        lines = []
        for h in self.headers:
            if self.hidden_sensitive.get() and self.is_sensitive_header(h):
                continue
            v = str(row.get(h, "")).strip()
            if v:
                lines.append(f"{h}: {v}")
        if lines:
            self.copy_to_clipboard("\n".join(lines), "Copied selected details")

    def copy_tradelink(self):
        row = self.selected_row()
        if not row:
            messagebox.showinfo("No row selected", "Select an account first.")
            return
        tradelink = self.extract_tradelink(row)
        if self.copy_to_clipboard(tradelink, "Copied tradelink"):
            return
        messagebox.showinfo("No tradelink", "No tradelink found in the selected row.")

    def copy_selected_row(self):
        item = self.tree.focus()
        if not item:
            return
        values = self.tree.item(item, "values")
        if values:
            self.copy_to_clipboard("\t".join(values), "Copied visible row")

    def export_filtered_csv(self):
        if not self.filtered_rows:
            messagebox.showinfo("Nothing to export", "No filtered rows to export.")
            return
        path = filedialog.asksaveasfilename(
            defaultextension=".csv",
            filetypes=[("CSV", "*.csv"), ("All files", "*.*")],
            initialfile="filtered_master_sheet.csv",
        )
        if not path:
            return
        cols = self.visible_headers()
        with open(path, "w", newline="", encoding="utf-8-sig") as f:
            writer = csv.writer(f)
            writer.writerow(cols)
            for row in self.filtered_rows:
                writer.writerow([row.get(h, "") for h in cols])
        self.status_var.set(f"Exported {len(self.filtered_rows)} rows to {os.path.basename(path)}")
        messagebox.showinfo("Export complete", f"Exported {len(self.filtered_rows)} rows.")

    def update_stats(self):
        total = len(self.rows)
        showing = len(self.filtered_rows)
        trade = sum(1 for r in self.rows if self.extract_tradelink(r))
        skin = sum(1 for r in self.rows if self.has_any_skin_site(r))
        banned = sum(1 for r in self.rows if self.has_banned_value(r))
        self.stat_total_var.set(str(total))
        self.stat_showing_var.set(str(showing))
        self.stat_trade_var.set(str(trade))
        self.stat_skin_var.set(str(skin))
        self.stat_banned_var.set(str(banned))

    def settings_path(self):
        return os.path.join(self.app_dir(), SETTINGS_FILE)

    def save_settings(self):
        data = {"last_file": self.file_path}
        try:
            with open(self.settings_path(), "w", encoding="utf-8") as f:
                json.dump(data, f, indent=2)
        except Exception:
            pass

    def load_settings(self):
        try:
            with open(self.settings_path(), "r", encoding="utf-8") as f:
                data = json.load(f)
            last_file = data.get("last_file")
            if last_file and os.path.exists(last_file):
                self.status_var.set(f"Last sheet: {os.path.basename(last_file)}. Click Open CSV / XLSX to load it again.")
        except Exception:
            pass


if __name__ == "__main__":
    app = SheetFinderPro()
    app.mainloop()
