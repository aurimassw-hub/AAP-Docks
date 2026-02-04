# main.py  (UPDATED: Padalinys/Pareigos dropdown from Darbuotojai.xlsx + Issue date selector)
# Requires: openpyxl, python-docx (not "docx")
# Optional: none (no tkcalendar)

import importlib.util
import json
import os
import threading
from dataclasses import dataclass
from datetime import date, datetime
from pathlib import Path
import tkinter as tk
from tkinter import messagebox, ttk

from openpyxl import Workbook, load_workbook


# ---------------------------
# Excel headers
# ---------------------------

ISSUANCE_HEADER = [
    "Numeris",
    "Vardas Pavardė",
    "Tab. Nr",
    "Padalinys",
    "Pareigos",
    "Lytis",
    "Aprangos kodas",
    "Apranga",
    "Išduota",
    "Susidėvėjimas",
]

# Darbuotojai.xlsx (Vardas, Pavarde, TabNr, Pareigos, Padalinys, Lytis)
EMP_HEADER = ["Vardas", "Pavarde", "TabNr", "Pareigos", "Padalinys", "Lytis"]


# ---------------------------
# Models
# ---------------------------

@dataclass
class EmployeeInfo:
    tab_nr: str
    name: str
    department: str
    position: str
    gender: str


# ---------------------------
# Utils
# ---------------------------

def parse_date(value):
    if value is None:
        return None
    if isinstance(value, datetime):
        return value.date()
    if isinstance(value, date):
        return value
    for fmt in ("%Y-%m-%d", "%d.%m.%Y", "%Y/%m/%d"):
        try:
            return datetime.strptime(str(value), fmt).date()
        except ValueError:
            continue
    return None


def add_months(start_date, months):
    if start_date is None:
        return None
    months = int(months or 0)
    if months <= 0:
        return None
    month = start_date.month - 1 + months
    year = start_date.year + month // 12
    month = month % 12 + 1
    days_in_month = [
        31,
        29 if year % 4 == 0 and (year % 100 != 0 or year % 400 == 0) else 28,
        31, 30, 31, 30, 31, 31, 30, 31, 30, 31
    ][month - 1]
    day = min(start_date.day, days_in_month)
    return date(year, month, day)


def resolve_path(base_path: Path, target: str):
    target_path = Path(target)
    if target_path.is_absolute():
        return target_path
    return (base_path / target_path).resolve()


def sanitize_filename(name: str) -> str:
    bad = '<>:"/\\|?*'
    for ch in bad:
        name = name.replace(ch, "_")
    return " ".join(name.split()).strip()


def python_docx_available():
    return (
        importlib.util.find_spec("docx") is not None
        and importlib.util.find_spec("docx.api") is not None
    )


def strip_size_suffix(name: str) -> str:
    if not name:
        return ""
    if "(" in name and name.endswith(")"):
        return name[: name.rfind("(")].strip()
    return name


def ensure_size_suffix(name: str, size: str) -> str:
    if not size:
        return name
    base = strip_size_suffix(name)
    suffix = f"({size} dydis)"
    if base.endswith(suffix):
        return base
    return f"{base} {suffix}".strip()


def split_full_name(full_name: str):
    s = " ".join((full_name or "").split()).strip()
    if not s:
        return "", ""
    parts = s.split(" ")
    if len(parts) == 1:
        return parts[0], ""
    return " ".join(parts[:-1]), parts[-1]


# ---------------------------
# Settings
# ---------------------------

def load_settings(settings_path: str):
    p = Path(settings_path).resolve()
    base_dir = p.parent
    with open(p, "r", encoding="utf-8") as f:
        data = json.load(f)

    paths = data.get("paths", {})
    return {
        "template": resolve_path(base_dir, paths.get("template", "AAP kortelės_2025.docx")),
        "excel": resolve_path(base_dir, paths.get("excel", "AAP DB.xlsx")),
        "employees_excel": resolve_path(base_dir, paths.get("employees_excel", "Darbuotojai.xlsx")),
        "outputs": resolve_path(base_dir, paths.get("outputs", "sugeneruotos kortelės")),
        "gear": resolve_path(base_dir, paths.get("gear_json", "aprangos_kodai.json")),
        "ui": data.get("ui", {}),
    }


# ---------------------------
# Excel helpers
# ---------------------------

def ensure_issuance_excel(excel_path: Path):
    if excel_path.exists():
        return
    wb = Workbook()
    ws = wb.active
    ws.title = "AAP DB"
    ws.append(ISSUANCE_HEADER)
    wb.save(excel_path)


def ensure_employees_excel(employees_path: Path):
    if employees_path.exists():
        return
    wb = Workbook()
    ws = wb.active
    ws.title = "Darbuotojai"
    ws.append(EMP_HEADER)
    wb.save(employees_path)


def read_employees(employees_path: Path) -> list[dict]:
    ensure_employees_excel(employees_path)
    wb = load_workbook(employees_path)
    ws = wb.active

    out = []
    for row in ws.iter_rows(min_row=2, values_only=True):
        if not any(row):
            continue
        d = dict(zip(EMP_HEADER, row))
        tab = str(d.get("TabNr", "") or "").strip()
        if not tab:
            continue
        vardas = str(d.get("Vardas", "") or "").strip()
        pav = str(d.get("Pavarde", "") or "").strip()
        full = " ".join([x for x in [vardas, pav] if x]).strip()

        out.append({
            "tab_nr": tab,
            "name": full,
            "department": str(d.get("Padalinys", "") or "").strip(),
            "position": str(d.get("Pareigos", "") or "").strip(),
            "gender": str(d.get("Lytis", "") or "").strip(),
        })
    return out


def unique_sorted_values(employees_path: Path, key: str) -> list[str]:
    """
    key: 'Padalinys' or 'Pareigos'
    """
    ensure_employees_excel(employees_path)
    wb = load_workbook(employees_path)
    ws = wb.active

    idx = EMP_HEADER.index(key) + 1
    vals = set()
    for r in range(2, ws.max_row + 1):
        v = str(ws.cell(r, idx).value or "").strip()
        if v:
            vals.add(v)
    return sorted(vals, key=lambda s: s.lower())


def upsert_employee(employees_path: Path, emp: EmployeeInfo):
    ensure_employees_excel(employees_path)
    wb = load_workbook(employees_path)
    ws = wb.active

    target_row = None
    for r in range(2, ws.max_row + 1):
        tab = str(ws.cell(r, 3).value or "").strip()  # TabNr col=3
        if tab == emp.tab_nr:
            target_row = r
            break

    vardas, pavarde = split_full_name(emp.name)

    if target_row is None:
        ws.append([vardas, pavarde, emp.tab_nr, emp.position, emp.department, emp.gender])
    else:
        ws.cell(target_row, 1, vardas)
        ws.cell(target_row, 2, pavarde)
        ws.cell(target_row, 3, emp.tab_nr)
        ws.cell(target_row, 4, emp.position)
        ws.cell(target_row, 5, emp.department)
        ws.cell(target_row, 6, emp.gender)

    wb.save(employees_path)


def read_issuance_rows(excel_path: Path) -> list[dict]:
    ensure_issuance_excel(excel_path)
    wb = load_workbook(excel_path)
    ws = wb.active

    rows = []
    for row in ws.iter_rows(min_row=2, values_only=True):
        if not any(row):
            continue
        rows.append(dict(zip(ISSUANCE_HEADER, row)))
    return rows


def next_document_number(excel_path: Path) -> int:
    rows = read_issuance_rows(excel_path)
    mx = 0
    for r in rows:
        v = r.get("Numeris")
        if v in (None, ""):
            continue
        try:
            mx = max(mx, int(v))
        except Exception:
            continue
    return mx + 1


def insert_issuance_rows(excel_path: Path, emp: EmployeeInfo, items: list[dict], doc_number: int, issued_date: date):
    ensure_issuance_excel(excel_path)
    wb = load_workbook(excel_path)
    ws = wb.active

    issued_str = issued_date.isoformat()

    for it in reversed(items):
        ws.insert_rows(2)
        ws.cell(2, 1, doc_number)
        ws.cell(2, 2, emp.name)
        ws.cell(2, 3, emp.tab_nr)
        ws.cell(2, 4, emp.department)
        ws.cell(2, 5, emp.position)
        ws.cell(2, 6, emp.gender)
        ws.cell(2, 7, it["code"])
        ws.cell(2, 8, it["name"])
        ws.cell(2, 9, issued_str)
        ws.cell(2, 10, it["months"])

    wb.save(excel_path)


# ---------------------------
# Gear codes JSON
# ---------------------------

def load_gear_codes(path: Path) -> dict:
    if not path.exists():
        return {}
    with open(path, "r", encoding="utf-8") as f:
        data = json.load(f)
    return data if isinstance(data, dict) else {}


def save_gear_codes(path: Path, codes: dict):
    path.parent.mkdir(parents=True, exist_ok=True)
    with open(path, "w", encoding="utf-8") as f:
        json.dump(codes, f, ensure_ascii=False, indent=2)


# ---------------------------
# Word
# ---------------------------

def replace_placeholders(doc, mapping: dict):
    for paragraph in doc.paragraphs:
        for run in paragraph.runs:
            t = run.text
            if not t:
                continue
            for k, v in mapping.items():
                if k in t:
                    t = t.replace(k, v)
            run.text = t

    for table in doc.tables:
        for row in table.rows:
            for cell in row.cells:
                for paragraph in cell.paragraphs:
                    for run in paragraph.runs:
                        t = run.text
                        if not t:
                            continue
                        for k, v in mapping.items():
                            if k in t:
                                t = t.replace(k, v)
                        run.text = t


def pick_gear_table(doc):
    return doc.tables[0] if doc.tables else None


def generate_word(template_path: Path, outputs_dir: Path, emp: EmployeeInfo, items: list[dict], doc_number: int, issued_date: date):
    if not python_docx_available():
        messagebox.showerror(
            "Trūksta priklausomybių",
            "Nerastas python-docx. Įdiek 'python-docx' ir pašalink paketą 'docx'."
        )
        return None

    from docx import Document

    outputs_dir.mkdir(parents=True, exist_ok=True)

    doc = Document(str(template_path))

    replace_placeholders(doc, {
        "{Employee}": emp.name,
        "{Emploee}": emp.name,
        "{Departament}": emp.department,
        "{Position}": emp.position,
    })

    tbl = pick_gear_table(doc)
    if tbl is not None:
        while len(tbl.rows) > 1:
            tbl._tbl.remove(tbl.rows[1]._tr)

        for it in items[:14]:
            r = tbl.add_row()
            c = r.cells
            change_date = add_months(issued_date, it["months"])
            c[0].text = issued_date.strftime("%Y-%m-%d")
            c[1].text = str(it["code"])
            c[2].text = str(it["name"])
            c[3].text = str(it["months"])
            c[4].text = change_date.strftime("%Y-%m-%d") if change_date else ""
            c[5].text = "__"

    filename = sanitize_filename(f"AAP {doc_number} {emp.name}.docx")
    out_path = outputs_dir / filename
    doc.save(str(out_path))
    return out_path


# ---------------------------
# UI widgets: searchable combobox
# ---------------------------

class SearchableCombobox(ttk.Frame):
    """
    Entry + Listbox dropdown (searchable).
    Use get() / set() like a combobox.
    """
    def __init__(self, parent, values=None, width=42):
        super().__init__(parent)
        self.values = values or []
        self.var = tk.StringVar()

        self.entry = ttk.Entry(self, textvariable=self.var, width=width)
        self.entry.pack(fill="x")

        self.popup = None
        self.listbox = None

        self.entry.bind("<KeyRelease>", self._on_key)
        self.entry.bind("<Button-1>", self._show_popup)
        self.entry.bind("<Down>", self._show_popup)
        self.entry.bind("<FocusOut>", self._on_focus_out)

    def set_values(self, values):
        self.values = values or []
        self._refresh_list(self.var.get())

    def get(self):
        return self.var.get().strip()

    def set(self, value):
        self.var.set(value or "")

    def _on_focus_out(self, _ev):
        # close popup if click outside
        self.after(150, self._hide_popup)

    def _on_key(self, _ev):
        self._show_popup()
        self._refresh_list(self.var.get())

    def _show_popup(self, _ev=None):
        if self.popup and self.popup.winfo_exists():
            return

        self.popup = tk.Toplevel(self)
        self.popup.wm_overrideredirect(True)
        self.popup.attributes("-topmost", True)

        x = self.entry.winfo_rootx()
        y = self.entry.winfo_rooty() + self.entry.winfo_height()
        self.popup.geometry(f"{self.entry.winfo_width()}x200+{x}+{y}")

        self.listbox = tk.Listbox(self.popup, height=10)
        self.listbox.pack(fill="both", expand=True)

        self.listbox.bind("<ButtonRelease-1>", self._choose)
        self.listbox.bind("<Return>", self._choose)

        self._refresh_list(self.var.get())

    def _hide_popup(self):
        if self.popup and self.popup.winfo_exists():
            self.popup.destroy()
        self.popup = None
        self.listbox = None

    def _refresh_list(self, query):
        if not (self.popup and self.listbox):
            return
        q = (query or "").strip().lower()
        self.listbox.delete(0, tk.END)
        for v in self.values:
            if not q or q in v.lower():
                self.listbox.insert(tk.END, v)

    def _choose(self, _ev=None):
        if not self.listbox:
            return
        sel = self.listbox.curselection()
        if not sel:
            return
        value = self.listbox.get(sel[0])
        self.var.set(value)
        self._hide_popup()
        self.entry.focus_set()


# ---------------------------
# App
# ---------------------------

class AAPApp(tk.Tk):
    def __init__(self, settings_path: str):
        super().__init__()

        self.settings = load_settings(settings_path)
        ensure_issuance_excel(self.settings["excel"])
        ensure_employees_excel(self.settings["employees_excel"])
        self.gear_codes = load_gear_codes(self.settings["gear"])

        self.current_employee: EmployeeInfo | None = None
        self.change_mode = False
        self.issue_date_for_new_employee: date | None = None

        self.title(self.settings.get("ui", {}).get("title", "AAP Issuance"))
        w = int(self.settings.get("ui", {}).get("width", 900))
        h = int(self.settings.get("ui", {}).get("height", 650))
        self.geometry(f"{w}x{h}")
        self.resizable(False, False)

        if self.settings.get("ui", {}).get("center_on_screen", True):
            self.center_window()

        container = ttk.Frame(self)
        container.pack(fill="both", expand=True)

        self.pages = {}
        for PageClass in (EmployeeSelectPage, EmployeeInfoPage, CurrentGearPage, NewGearPage):
            page = PageClass(container, self)
            self.pages[PageClass.__name__] = page
            page.grid(row=0, column=0, sticky="nsew")

        self.show_page("EmployeeSelectPage")

    def center_window(self):
        self.update_idletasks()
        width = self.winfo_width()
        height = self.winfo_height()
        x = (self.winfo_screenwidth() // 2) - (width // 2)
        y = (self.winfo_screenheight() // 2) - (height // 2)
        self.geometry(f"{width}x{height}+{x}+{y}")

    def show_page(self, name: str):
        page = self.pages[name]
        if hasattr(page, "refresh"):
            page.refresh()
        page.tkraise()

    # Data wrappers
    def list_employees(self) -> list[dict]:
        return read_employees(self.settings["employees_excel"])

    def upsert_employee(self, emp: EmployeeInfo):
        upsert_employee(self.settings["employees_excel"], emp)

    def list_departments(self) -> list[str]:
        return unique_sorted_values(self.settings["employees_excel"], "Padalinys")

    def list_positions(self) -> list[str]:
        return unique_sorted_values(self.settings["employees_excel"], "Pareigos")

    def issuance_rows_for(self, tab_nr: str) -> list[dict]:
        rows = read_issuance_rows(self.settings["excel"])
        return [r for r in rows if str(r.get("Tab. Nr", "")).strip() == tab_nr]

    def get_next_doc_number(self) -> int:
        return next_document_number(self.settings["excel"])

    def save_issuance(self, emp: EmployeeInfo, items: list[dict], doc_number: int, issued_date: date):
        insert_issuance_rows(self.settings["excel"], emp, items, doc_number, issued_date)


# ---------------------------
# Pages
# ---------------------------

class EmployeeSelectPage(ttk.Frame):
    def __init__(self, parent, app: AAPApp):
        super().__init__(parent)
        self.app = app

        ttk.Label(self, text="Darbuotojo paieška", font=("Arial", 16)).pack(pady=10)

        top = ttk.Frame(self)
        top.pack(fill="x", padx=20)

        ttk.Label(top, text="Paieška (TabNr / vardas / pavardė):").pack(side="left")
        self.search_var = tk.StringVar()
        self.search_var.trace_add("write", lambda *_: self.refresh_list())
        ttk.Entry(top, textvariable=self.search_var, width=40).pack(side="left", padx=8)

        list_frame = ttk.Frame(self)
        list_frame.pack(fill="both", expand=True, padx=20, pady=10)

        self.listbox = tk.Listbox(list_frame, height=18)
        self.listbox.pack(side="left", fill="both", expand=True)
        sb = ttk.Scrollbar(list_frame, command=self.listbox.yview)
        sb.pack(side="right", fill="y")
        self.listbox.config(yscrollcommand=sb.set)

        btns = ttk.Frame(self)
        btns.pack(pady=10)

        ttk.Button(btns, text="Toliau", command=self.on_next).pack(side="left", padx=5)
        ttk.Button(btns, text="Naujas darbuotojas", command=self.on_new).pack(side="left", padx=5)
        ttk.Button(btns, text="Keisti padalinį", command=self.on_change).pack(side="left", padx=5)

        self.items = []

    def refresh(self):
        self.refresh_list()

    def refresh_list(self):
        self.listbox.delete(0, tk.END)
        self.items = []

        emps = self.app.list_employees()
        q = self.search_var.get().strip().lower()

        for e in sorted(emps, key=lambda x: (x["name"].lower(), x["tab_nr"])):
            tab = e["tab_nr"]
            nm = e["name"]
            if q and (q not in tab.lower() and q not in nm.lower()):
                continue
            self.items.append(e)
            self.listbox.insert(tk.END, f"{tab} — {nm}")

    def selected(self):
        sel = self.listbox.curselection()
        if not sel:
            return None
        return self.items[sel[0]]

    def on_next(self):
        s = self.selected()
        if not s:
            messagebox.showwarning("Pasirinkimas", "Pasirinkite darbuotoją iš sąrašo.")
            return

        self.app.current_employee = EmployeeInfo(
            tab_nr=s["tab_nr"],
            name=s["name"],
            department=s.get("department", ""),
            position=s.get("position", ""),
            gender=s.get("gender", "") or "Vyras",
        )
        self.app.change_mode = False
        self.app.issue_date_for_new_employee = None
        self.app.show_page("CurrentGearPage")

    def on_new(self):
        self.app.current_employee = None
        self.app.change_mode = False
        self.app.issue_date_for_new_employee = date.today()
        self.app.show_page("EmployeeInfoPage")

    def on_change(self):
        s = self.selected()
        if not s:
            messagebox.showwarning("Pasirinkimas", "Pasirinkite darbuotoją iš sąrašo.")
            return

        self.app.current_employee = EmployeeInfo(
            tab_nr=s["tab_nr"],
            name=s["name"],
            department=s.get("department", ""),
            position=s.get("position", ""),
            gender=s.get("gender", "") or "Vyras",
        )
        self.app.change_mode = True
        self.app.issue_date_for_new_employee = None
        self.app.show_page("EmployeeInfoPage")


class EmployeeInfoPage(ttk.Frame):
    def __init__(self, parent, app: AAPApp):
        super().__init__(parent)
        self.app = app

        ttk.Label(self, text="Darbuotojo informacija", font=("Arial", 16)).pack(pady=10)

        form = ttk.Frame(self)
        form.pack(pady=10)

        self.tab_var = tk.StringVar()
        self.name_var = tk.StringVar()
        self.gender_var = tk.StringVar(value="Vyras")

        # searchable dropdowns (from Darbuotojai.xlsx)
        self.dept_picker = SearchableCombobox(form, values=[], width=45)
        self.pos_picker = SearchableCombobox(form, values=[], width=45)

        # issue date selection (only for NEW employee)
        self.issue_year = tk.StringVar()
        self.issue_month = tk.StringVar()
        self.issue_day = tk.StringVar()
        self.issue_date_frame = ttk.Frame(form)

        ttk.Label(form, text="Tab. Nr:").grid(row=0, column=0, sticky="e", padx=5, pady=6)
        self.tab_entry = ttk.Entry(form, textvariable=self.tab_var, width=48)
        self.tab_entry.grid(row=0, column=1, sticky="w", padx=5, pady=6)

        ttk.Label(form, text="Vardas Pavardė:").grid(row=1, column=0, sticky="e", padx=5, pady=6)
        self.name_entry = ttk.Entry(form, textvariable=self.name_var, width=48)
        self.name_entry.grid(row=1, column=1, sticky="w", padx=5, pady=6)

        ttk.Label(form, text="Lytis:").grid(row=2, column=0, sticky="e", padx=5, pady=6)
        self.gender_combo = ttk.Combobox(
            form,
            textvariable=self.gender_var,
            values=["Vyras", "Moteris"],
            state="readonly",
            width=45
        )
        self.gender_combo.grid(row=2, column=1, sticky="w", padx=5, pady=6)

        ttk.Label(form, text="Padalinys:").grid(row=3, column=0, sticky="e", padx=5, pady=6)
        self.dept_picker.grid(row=3, column=1, sticky="w", padx=5, pady=6)

        ttk.Label(form, text="Pareigos:").grid(row=4, column=0, sticky="e", padx=5, pady=6)
        self.pos_picker.grid(row=4, column=1, sticky="w", padx=5, pady=6)

        # issue date row (row=5) created once
        ttk.Label(form, text="Išdavimo data:").grid(row=5, column=0, sticky="e", padx=5, pady=6)
        self.issue_date_frame.grid(row=5, column=1, sticky="w", padx=5, pady=6)

        self._build_issue_date_controls()

        btns = ttk.Frame(self)
        btns.pack(pady=10)

        ttk.Button(btns, text="Atgal", command=lambda: self.app.show_page("EmployeeSelectPage")).pack(side="left", padx=5)
        self.next_btn = ttk.Button(btns, text="Toliau", command=self.on_next)
        self.next_btn.pack(side="left", padx=5)

    def _build_issue_date_controls(self):
        # YYYY / MM / DD comboboxes
        today = date.today()
        years = [str(y) for y in range(today.year - 2, today.year + 3)]
        months = [f"{m:02d}" for m in range(1, 13)]
        days = [f"{d:02d}" for d in range(1, 32)]

        self.issue_year.set(str(today.year))
        self.issue_month.set(f"{today.month:02d}")
        self.issue_day.set(f"{today.day:02d}")

        self.year_combo = ttk.Combobox(self.issue_date_frame, textvariable=self.issue_year, values=years, width=6, state="readonly")
        self.month_combo = ttk.Combobox(self.issue_date_frame, textvariable=self.issue_month, values=months, width=4, state="readonly")
        self.day_combo = ttk.Combobox(self.issue_date_frame, textvariable=self.issue_day, values=days, width=4, state="readonly")

        self.year_combo.pack(side="left")
        ttk.Label(self.issue_date_frame, text="-").pack(side="left", padx=2)
        self.month_combo.pack(side="left")
        ttk.Label(self.issue_date_frame, text="-").pack(side="left", padx=2)
        self.day_combo.pack(side="left")

        ttk.Button(self.issue_date_frame, text="Šiandien", command=self._set_today).pack(side="left", padx=8)

    def _set_today(self):
        t = date.today()
        self.issue_year.set(str(t.year))
        self.issue_month.set(f"{t.month:02d}")
        self.issue_day.set(f"{t.day:02d}")

    def _get_issue_date(self) -> date:
        try:
            y = int(self.issue_year.get())
            m = int(self.issue_month.get())
            d = int(self.issue_day.get())
            return date(y, m, d)
        except Exception:
            return date.today()

    def refresh(self):
        # refresh dropdown values from Darbuotojai.xlsx each time
        depts = self.app.list_departments()
        poss = self.app.list_positions()
        self.dept_picker.set_values(depts)
        self.pos_picker.set_values(poss)

        info = self.app.current_employee
        if info:
            self.tab_var.set(info.tab_nr)
            self.name_var.set(info.name)
            self.gender_var.set(info.gender or "Vyras")
            self.dept_picker.set(info.department or "")
            self.pos_picker.set(info.position or "")
        else:
            self.tab_var.set("")
            self.name_var.set("")
            self.gender_var.set("Vyras")
            self.dept_picker.set("")
            self.pos_picker.set("")

        if self.app.change_mode:
            self.next_btn.config(text="Išsaugoti")
            self.tab_entry.config(state="disabled")
            self.name_entry.config(state="disabled")
            # issue date hidden when editing existing employee
            self.issue_date_frame.grid_remove()
        else:
            self.next_btn.config(text="Toliau")
            self.tab_entry.config(state="normal")
            self.name_entry.config(state="normal")
            # issue date visible for NEW employee creation
            self.issue_date_frame.grid()

            # if app has pre-set date, apply it
            if self.app.issue_date_for_new_employee:
                t = self.app.issue_date_for_new_employee
                self.issue_year.set(str(t.year))
                self.issue_month.set(f"{t.month:02d}")
                self.issue_day.set(f"{t.day:02d}")

    def on_next(self):
        tab = self.tab_var.get().strip()
        name = " ".join(self.name_var.get().split()).strip()
        if not tab or not name:
            messagebox.showwarning("Trūksta duomenų", "Įveskite Tab. Nr ir Vardas Pavardė.")
            return

        dept = self.dept_picker.get()
        pos = self.pos_picker.get()
        gen = self.gender_var.get().strip()

        if not dept or not pos:
            messagebox.showwarning("Trūksta duomenų", "Pasirinkite Padalinį ir Pareigas.")
            return

        emp = EmployeeInfo(tab_nr=tab, name=name, department=dept, position=pos, gender=gen)
        self.app.upsert_employee(emp)
        self.app.current_employee = emp

        if self.app.change_mode:
            self.app.change_mode = False
            self.app.show_page("EmployeeSelectPage")
        else:
            # store chosen issuance date for subsequent steps
            self.app.issue_date_for_new_employee = self._get_issue_date()
            self.app.show_page("CurrentGearPage")


class CurrentGearPage(ttk.Frame):
    def __init__(self, parent, app: AAPApp):
        super().__init__(parent)
        self.app = app

        ttk.Label(self, text="Dabartinė apranga", font=("Arial", 16)).pack(pady=10)

        self.tree = ttk.Treeview(
            self,
            columns=("Apranga", "Kodas", "Išduota", "Susidėvėjimas", "Keisti iki", "Likę (d.)"),
            show="headings",
            height=15,
        )
        for col in self.tree["columns"]:
            self.tree.heading(col, text=col)
            self.tree.column(col, width=140)
        self.tree.pack(padx=20, pady=10, fill="x")

        btns = ttk.Frame(self)
        btns.pack(pady=10)

        ttk.Button(btns, text="Atgal", command=lambda: self.app.show_page("EmployeeSelectPage")).pack(side="left", padx=5)
        ttk.Button(btns, text="Toliau → (nauja įranga)", command=lambda: self.app.show_page("NewGearPage")).pack(side="left", padx=5)

        self.tree.tag_configure("warn", foreground="red")

    def refresh(self):
        for r in self.tree.get_children():
            self.tree.delete(r)

        emp = self.app.current_employee
        if emp is None:
            return

        rows = self.app.issuance_rows_for(emp.tab_nr)
        filtered = [
            r for r in rows
            if str(r.get("Padalinys", "") or "").strip() == emp.department
            and str(r.get("Pareigos", "") or "").strip() == emp.position
        ]

        latest_by_item = {}
        for r in filtered:
            gear_name = str(r.get("Apranga", "") or "").strip()
            if not gear_name:
                continue
            key = strip_size_suffix(gear_name)
            issued = parse_date(r.get("Išduota"))
            if issued is None:
                continue
            cur = latest_by_item.get(key)
            if cur is None or issued > cur["issued"]:
                latest_by_item[key] = {"row": r, "issued": issued}

        today = date.today()
        for v in latest_by_item.values():
            r = v["row"]
            issued = v["issued"]
            months = int(r.get("Susidėvėjimas") or 0)
            change_date = add_months(issued, months)
            remaining = (change_date - today).days if change_date else ""

            item_id = self.tree.insert("", tk.END, values=(
                str(r.get("Apranga", "") or ""),
                str(r.get("Aprangos kodas", "") or ""),
                issued.strftime("%Y-%m-%d"),
                months,
                change_date.strftime("%Y-%m-%d") if change_date else "",
                remaining,
            ))
            if isinstance(remaining, int) and remaining < 7:
                self.tree.item(item_id, tags=("warn",))


class NewGearPage(ttk.Frame):
    def __init__(self, parent, app: AAPApp):
        super().__init__(parent)
        self.app = app

        ttk.Label(self, text="Nauja apranga (Excel tipo lentelė)", font=("Arial", 16)).pack(pady=10)

        header = ttk.Frame(self)
        header.pack()
        ttk.Label(header, text="Kodas", width=20).grid(row=0, column=0, padx=2)
        ttk.Label(header, text="Apranga", width=50).grid(row=0, column=1, padx=2)
        ttk.Label(header, text="Susidėvėjimas (mėn.)", width=20).grid(row=0, column=2, padx=2)

        grid = ttk.Frame(self)
        grid.pack(pady=6)

        self.entries = []
        self.after_ids = {}

        for i in range(14):
            code_var = tk.StringVar()
            name_var = tk.StringVar()
            months_var = tk.StringVar()

            e_code = ttk.Entry(grid, textvariable=code_var, width=20)
            e_name = ttk.Entry(grid, textvariable=name_var, width=50)
            e_mon = ttk.Entry(grid, textvariable=months_var, width=20)

            e_code.grid(row=i, column=0, padx=2, pady=2)
            e_name.grid(row=i, column=1, padx=2, pady=2)
            e_mon.grid(row=i, column=2, padx=2, pady=2)

            e_code.bind("<KeyRelease>", lambda _ev, idx=i: self.on_code_change(idx))

            self.entries.append({"code": code_var, "name": name_var, "months": months_var})

        btns = ttk.Frame(self)
        btns.pack(pady=10)

        ttk.Button(btns, text="Atgal", command=lambda: self.app.show_page("CurrentGearPage")).pack(side="left", padx=5)
        self.btn_gen_new = ttk.Button(btns, text="Generuoti ir naujas įrašas", command=self.on_generate_new)
        self.btn_gen_new.pack(side="left", padx=5)
        self.btn_gen_close = ttk.Button(btns, text="Generuoti ir uždaryti", command=self.on_generate_close)
        self.btn_gen_close.pack(side="left", padx=5)

    def refresh(self):
        for r in self.entries:
            r["code"].set("")
            r["name"].set("")
            r["months"].set("")

    def on_code_change(self, idx):
        if idx in self.after_ids:
            self.after_cancel(self.after_ids[idx])
        self.after_ids[idx] = self.after(120, lambda: self.apply_code(idx))

    def apply_code(self, idx):
        r = self.entries[idx]
        raw = r["code"].get().strip()
        if not raw:
            return

        if "-" in raw:
            base, size = raw.split("-", 1)
        else:
            base, size = raw, ""

        base = base.strip()
        size = size.strip()

        info = self.app.gear_codes.get(base)
        if info:
            name = str(info.get("name", "") or "").strip()
            months = info.get("months", "")
            r["name"].set(ensure_size_suffix(name, size))
            r["months"].set(str(months))
        else:
            if size and r["name"].get().strip():
                r["name"].set(ensure_size_suffix(r["name"].get().strip(), size))

    def collect_items(self) -> list[dict]:
        items = []
        for r in self.entries:
            raw = r["code"].get().strip()
            name = r["name"].get().strip()
            months = r["months"].get().strip()

            if not raw and not name and not months:
                continue

            if "-" in raw:
                base, size = raw.split("-", 1)
            else:
                base, size = raw, ""

            base = base.strip()
            size = size.strip()

            final_name = ensure_size_suffix(name, size)
            try:
                m = int(months)
            except Exception:
                m = 0

            items.append({"code": base, "name": final_name, "months": m})
        return items

    def update_gear_codes(self, items: list[dict]):
        updated = False
        for it in items:
            base = it["code"]
            if base and base not in self.app.gear_codes:
                self.app.gear_codes[base] = {
                    "name": strip_size_suffix(it["name"]),
                    "months": it["months"],
                }
                updated = True
        if updated:
            save_gear_codes(self.app.settings["gear"], self.app.gear_codes)

    def toggle_buttons(self, state: str):
        self.btn_gen_new.config(state=state)
        self.btn_gen_close.config(state=state)

    def _do_generate(self, close_after: bool):
        emp = self.app.current_employee
        if emp is None:
            messagebox.showwarning("Klaida", "Nėra pasirinkto darbuotojo.")
            self.toggle_buttons("normal")
            return

        items = self.collect_items()
        if not items:
            messagebox.showwarning("Trūksta duomenų", "Įveskite bent vieną aprangos eilutę.")
            self.toggle_buttons("normal")
            return

        self.update_gear_codes(items)

        doc_number = self.app.get_next_doc_number()
        issued_date = self.app.issue_date_for_new_employee or date.today()

        self.app.save_issuance(emp, items, doc_number, issued_date)

        out_path = generate_word(
            template_path=self.app.settings["template"],
            outputs_dir=self.app.settings["outputs"],
            emp=emp,
            items=items,
            doc_number=doc_number,
            issued_date=issued_date
        )

        if out_path:
            messagebox.showinfo("Atlikta", f"Sugeneruota:\n{out_path}")

        if close_after:
            self.app.destroy()
        else:
            self.app.show_page("EmployeeSelectPage")

        self.toggle_buttons("normal")

    def on_generate_new(self):
        self.toggle_buttons("disabled")
        threading.Thread(target=self._do_generate, args=(False,), daemon=True).start()

    def on_generate_close(self):
        self.toggle_buttons("disabled")
        threading.Thread(target=self._do_generate, args=(True,), daemon=True).start()


if __name__ == "__main__":
    settings_file = os.environ.get("AAP_SETTINGS", "settings.json")
    app = AAPApp(settings_file)
    app.mainloop()
