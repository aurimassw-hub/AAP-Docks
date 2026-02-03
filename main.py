import json
import os
import threading
from dataclasses import dataclass
from datetime import date, datetime
from pathlib import Path
import tkinter as tk
from tkinter import ttk, messagebox

from docx import Document
from openpyxl import Workbook, load_workbook


HEADER_COLUMNS = [
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


@dataclass
class EmployeeInfo:
    tab_nr: str
    name: str
    department: str
    position: str
    gender: str


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
    month = start_date.month - 1 + months
    year = start_date.year + month // 12
    month = month % 12 + 1
    day = min(start_date.day, [31, 29 if year % 4 == 0 and (year % 100 != 0 or year % 400 == 0) else 28, 31, 30, 31, 30, 31, 31, 30, 31, 30, 31][month - 1])
    return date(year, month, day)


def resolve_path(base_path, target):
    target_path = Path(target)
    if target_path.is_absolute():
        return target_path
    return (base_path / target_path).resolve()


def load_settings(settings_path):
    with open(settings_path, "r", encoding="utf-8") as file:
        data = json.load(file)
    base_dir = Path(settings_path).resolve().parent
    paths = data.get("paths", {})
    return {
        "template": resolve_path(base_dir, paths.get("template", "template.docx")),
        "excel": resolve_path(base_dir, paths.get("excel", "AAP DB.xlsx")),
        "outputs": resolve_path(base_dir, paths.get("outputs", "sugeneruotos kortelės")),
        "workplaces": resolve_path(base_dir, paths.get("workplaces_json", "darbo_vietos.json")),
        "gear": resolve_path(base_dir, paths.get("gear_json", "aprangos_kodai.json")),
    }


def ensure_excel_file(excel_path):
    if excel_path.exists():
        return
    workbook = Workbook()
    sheet = workbook.active
    sheet.append(HEADER_COLUMNS)
    workbook.save(excel_path)


def load_workplaces(path):
    if not path.exists():
        return {}
    with open(path, "r", encoding="utf-8") as file:
        data = json.load(file)
    if isinstance(data, dict):
        return data
    return {}


def load_gear_codes(path):
    if not path.exists():
        return {}
    with open(path, "r", encoding="utf-8") as file:
        data = json.load(file)
    if isinstance(data, dict):
        return data
    return {}


def save_gear_codes(path, codes):
    path.parent.mkdir(parents=True, exist_ok=True)
    with open(path, "w", encoding="utf-8") as file:
        json.dump(codes, file, ensure_ascii=False, indent=2)


def strip_size_suffix(name):
    if not name:
        return ""
    if "(" in name and name.endswith(")"):
        return name[: name.rfind("(")].strip()
    return name


def ensure_size_suffix(name, size):
    if not size:
        return name
    base = strip_size_suffix(name)
    suffix = f"({size} dydis)"
    if base.endswith(suffix):
        return base
    return f"{base} {suffix}".strip()


def replace_placeholders(doc, mapping):
    for paragraph in doc.paragraphs:
        for key, value in mapping.items():
            if key in paragraph.text:
                for run in paragraph.runs:
                    run.text = run.text.replace(key, value)
    for table in doc.tables:
        for row in table.rows:
            for cell in row.cells:
                for paragraph in cell.paragraphs:
                    for key, value in mapping.items():
                        if key in paragraph.text:
                            for run in paragraph.runs:
                                run.text = run.text.replace(key, value)


class AAPApp(tk.Tk):
    def __init__(self, settings_path):
        super().__init__()
        self.settings = load_settings(settings_path)
        ensure_excel_file(self.settings["excel"])
        self.workplaces = load_workplaces(self.settings["workplaces"])
        self.gear_codes = load_gear_codes(self.settings["gear"])

        self.current_employee = None
        self.change_mode = False

        self.title("AAP Issuance")
        self.geometry("900x650")
        self.resizable(False, False)
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

    def show_page(self, page_name):
        page = self.pages[page_name]
        if hasattr(page, "refresh"):
            page.refresh()
        page.tkraise()

    def load_employee_rows(self):
        workbook = load_workbook(self.settings["excel"])
        sheet = workbook.active
        rows = []
        for row in sheet.iter_rows(min_row=2, values_only=True):
            if not any(row):
                continue
            row_data = dict(zip(HEADER_COLUMNS, row))
            rows.append(row_data)
        return rows

    def find_latest_employee_info(self, tab_nr):
        rows = self.load_employee_rows()
        filtered = [row for row in rows if str(row.get("Tab. Nr", "")).strip() == tab_nr]
        latest = None
        latest_date = None
        for row in filtered:
            issued = parse_date(row.get("Išduota"))
            if issued and (latest_date is None or issued > latest_date):
                latest_date = issued
                latest = row
        if latest is None:
            return None
        return EmployeeInfo(
            tab_nr=str(latest.get("Tab. Nr", "")),
            name=str(latest.get("Vardas Pavardė", "")),
            department=str(latest.get("Padalinys", "")),
            position=str(latest.get("Pareigos", "")),
            gender=str(latest.get("Lytis", "")),
        )

    def insert_workplace_change(self, employee_info):
        workbook = load_workbook(self.settings["excel"])
        sheet = workbook.active
        sheet.insert_rows(2)
        values = [
            "",
            employee_info.name,
            employee_info.tab_nr,
            employee_info.department,
            employee_info.position,
            employee_info.gender,
            "",
            "",
            date.today().isoformat(),
            "",
        ]
        for col, value in enumerate(values, start=1):
            sheet.cell(row=2, column=col, value=value)
        workbook.save(self.settings["excel"])

    def get_next_document_number(self):
        rows = self.load_employee_rows()
        numbers = []
        for row in rows:
            value = row.get("Numeris")
            if value is None or value == "":
                continue
            try:
                numbers.append(int(value))
            except (TypeError, ValueError):
                continue
        return max(numbers, default=0) + 1

    def append_issuance_rows(self, employee_info, issued_items, document_number):
        workbook = load_workbook(self.settings["excel"])
        sheet = workbook.active
        issued_date = date.today().isoformat()
        for item in reversed(issued_items):
            sheet.insert_rows(2)
            values = [
                document_number,
                employee_info.name,
                employee_info.tab_nr,
                employee_info.department,
                employee_info.position,
                employee_info.gender,
                item["code"],
                item["name"],
                issued_date,
                item["months"],
            ]
            for col, value in enumerate(values, start=1):
                sheet.cell(row=2, column=col, value=value)
        workbook.save(self.settings["excel"])

    def generate_word_doc(self, employee_info, issued_items, document_number):
        template_path = self.settings["template"]
        output_dir = self.settings["outputs"]
        output_dir.mkdir(parents=True, exist_ok=True)
        doc = Document(template_path)
        replace_placeholders(
            doc,
            {
                "{Employee}": employee_info.name,
                "{Emploee}": employee_info.name,
                "{Departament}": employee_info.department,
                "{Position}": employee_info.position,
            },
        )
        if doc.tables:
            table = doc.tables[0]
            while len(table.rows) > 1:
                table._tbl.remove(table.rows[1]._tr)
            for item in issued_items[:14]:
                row = table.add_row()
                issued_date = date.today()
                change_date = add_months(issued_date, int(item["months"]))
                cells = row.cells
                cells[0].text = issued_date.strftime("%Y-%m-%d")
                cells[1].text = item["code"]
                cells[2].text = item["name"]
                cells[3].text = str(item["months"])
                cells[4].text = change_date.strftime("%Y-%m-%d") if change_date else ""
                cells[5].text = "__"
        filename = f"AAP {document_number} {employee_info.name}.docx"
        output_path = output_dir / filename
        doc.save(output_path)


class EmployeeSelectPage(ttk.Frame):
    def __init__(self, parent, controller):
        super().__init__(parent)
        self.controller = controller

        ttk.Label(self, text="Darbuotojo paieška", font=("Arial", 16)).pack(pady=10)

        search_frame = ttk.Frame(self)
        search_frame.pack(fill="x", padx=20)
        ttk.Label(search_frame, text="Paieška:").pack(side="left")
        self.search_var = tk.StringVar()
        self.search_var.trace_add("write", lambda *_: self.refresh_list())
        ttk.Entry(search_frame, textvariable=self.search_var, width=40).pack(side="left", padx=5)

        list_frame = ttk.Frame(self)
        list_frame.pack(fill="both", expand=True, padx=20, pady=10)
        self.listbox = tk.Listbox(list_frame, height=18)
        self.listbox.pack(side="left", fill="both", expand=True)
        scrollbar = ttk.Scrollbar(list_frame, command=self.listbox.yview)
        scrollbar.pack(side="right", fill="y")
        self.listbox.config(yscrollcommand=scrollbar.set)

        button_frame = ttk.Frame(self)
        button_frame.pack(pady=10)
        ttk.Button(button_frame, text="Toliau", command=self.on_next).pack(side="left", padx=5)
        ttk.Button(button_frame, text="Naujas darbuotojas", command=self.on_new).pack(side="left", padx=5)
        ttk.Button(button_frame, text="Keisti padalinį", command=self.on_change).pack(side="left", padx=5)

        self.employees = []

    def refresh(self):
        self.refresh_list()

    def refresh_list(self):
        rows = self.controller.load_employee_rows()
        latest_by_tab = {}
        for row in rows:
            tab_nr = str(row.get("Tab. Nr", "")).strip()
            if not tab_nr:
                continue
            issued = parse_date(row.get("Išduota"))
            existing = latest_by_tab.get(tab_nr)
            if existing is None or (issued and issued > existing["date"]):
                latest_by_tab[tab_nr] = {"row": row, "date": issued or date.min}

        search_text = self.search_var.get().strip().lower()
        self.employees = []
        self.listbox.delete(0, tk.END)
        for tab_nr, info in sorted(latest_by_tab.items(), key=lambda x: x[0]):
            row = info["row"]
            name = str(row.get("Vardas Pavardė", "")).strip()
            display = f"{tab_nr} — {name}"
            if search_text and search_text not in tab_nr.lower() and search_text not in name.lower():
                continue
            self.employees.append({"tab_nr": tab_nr, "name": name})
            self.listbox.insert(tk.END, display)

    def get_selected_employee(self):
        selection = self.listbox.curselection()
        if not selection:
            return None
        return self.employees[selection[0]]

    def on_next(self):
        selected = self.get_selected_employee()
        if not selected:
            messagebox.showwarning("Pasirinkimas", "Pasirinkite darbuotoją iš sąrašo.")
            return
        info = self.controller.find_latest_employee_info(selected["tab_nr"])
        if info is None:
            messagebox.showwarning("Klaida", "Nepavyko rasti darbuotojo informacijos.")
            return
        self.controller.current_employee = info
        self.controller.change_mode = False
        self.controller.show_page("CurrentGearPage")

    def on_new(self):
        self.controller.current_employee = None
        self.controller.change_mode = False
        self.controller.show_page("EmployeeInfoPage")

    def on_change(self):
        selected = self.get_selected_employee()
        if not selected:
            messagebox.showwarning("Pasirinkimas", "Pasirinkite darbuotoją iš sąrašo.")
            return
        info = self.controller.find_latest_employee_info(selected["tab_nr"])
        if info is None:
            messagebox.showwarning("Klaida", "Nepavyko rasti darbuotojo informacijos.")
            return
        self.controller.current_employee = info
        self.controller.change_mode = True
        self.controller.show_page("EmployeeInfoPage")


class EmployeeInfoPage(ttk.Frame):
    def __init__(self, parent, controller):
        super().__init__(parent)
        self.controller = controller

        ttk.Label(self, text="Darbuotojo informacija", font=("Arial", 16)).pack(pady=10)

        form_frame = ttk.Frame(self)
        form_frame.pack(pady=10)

        self.tab_var = tk.StringVar()
        self.name_var = tk.StringVar()
        self.gender_var = tk.StringVar()
        self.dept_var = tk.StringVar()
        self.pos_var = tk.StringVar()

        self.create_field(form_frame, "Tab. Nr:", self.tab_var, 0)
        self.create_field(form_frame, "Vardas Pavardė:", self.name_var, 1)
        self.create_combo(form_frame, "Lytis:", self.gender_var, ["Vyras", "Moteris"], 2)
        self.create_combo(form_frame, "Padalinys:", self.dept_var, list(self.controller.workplaces.keys()), 3)
        self.create_combo(form_frame, "Pareigos:", self.pos_var, [], 4)

        self.dept_var.trace_add("write", lambda *_: self.update_positions())

        button_frame = ttk.Frame(self)
        button_frame.pack(pady=10)
        ttk.Button(button_frame, text="Atgal", command=lambda: controller.show_page("EmployeeSelectPage")).pack(side="left", padx=5)
        self.next_button = ttk.Button(button_frame, text="Toliau", command=self.on_next)
        self.next_button.pack(side="left", padx=5)

    def create_field(self, parent, label, var, row):
        ttk.Label(parent, text=label).grid(row=row, column=0, sticky="e", pady=5, padx=5)
        entry = ttk.Entry(parent, textvariable=var, width=40)
        entry.grid(row=row, column=1, sticky="w", pady=5)

    def create_combo(self, parent, label, var, values, row):
        ttk.Label(parent, text=label).grid(row=row, column=0, sticky="e", pady=5, padx=5)
        combo = ttk.Combobox(parent, textvariable=var, values=values, state="readonly", width=37)
        combo.grid(row=row, column=1, sticky="w", pady=5)

    def update_positions(self):
        dept = self.dept_var.get()
        positions = self.controller.workplaces.get(dept, [])
        pos_combo = self.children["!frame"].children["!combobox2"]
        pos_combo["values"] = positions
        if positions:
            self.pos_var.set(positions[0])
        else:
            self.pos_var.set("")

    def refresh(self):
        info = self.controller.current_employee
        if info:
            self.tab_var.set(info.tab_nr)
            self.name_var.set(info.name)
            self.gender_var.set(info.gender or "Vyras")
            self.dept_var.set(info.department)
            self.pos_var.set(info.position)
        else:
            self.tab_var.set("")
            self.name_var.set("")
            self.gender_var.set("Vyras")
            self.dept_var.set("")
            self.pos_var.set("")

        if self.controller.change_mode:
            self.next_button.config(text="Išsaugoti")
            self.children["!frame"].children["!entry"].config(state="disabled")
            self.children["!frame"].children["!entry2"].config(state="disabled")
        else:
            self.next_button.config(text="Toliau")
            self.children["!frame"].children["!entry"].config(state="normal")
            self.children["!frame"].children["!entry2"].config(state="normal")

    def on_next(self):
        tab_nr = self.tab_var.get().strip()
        name = self.name_var.get().strip()
        if not tab_nr or not name:
            messagebox.showwarning("Trūksta duomenų", "Įveskite Tab. Nr ir vardą pavardę.")
            return
        employee_info = EmployeeInfo(
            tab_nr=tab_nr,
            name=name,
            department=self.dept_var.get().strip(),
            position=self.pos_var.get().strip(),
            gender=self.gender_var.get().strip(),
        )
        self.controller.current_employee = employee_info
        if self.controller.change_mode:
            self.controller.insert_workplace_change(employee_info)
            self.controller.change_mode = False
            self.controller.show_page("EmployeeSelectPage")
        else:
            self.controller.show_page("CurrentGearPage")


class CurrentGearPage(ttk.Frame):
    def __init__(self, parent, controller):
        super().__init__(parent)
        self.controller = controller

        ttk.Label(self, text="Dabartinė apranga", font=("Arial", 16)).pack(pady=10)

        self.tree = ttk.Treeview(
            self,
            columns=("Apranga", "Kodas", "Išduota", "Susidėvėjimas", "Keisti iki", "Likę"),
            show="headings",
            height=15,
        )
        for col in self.tree["columns"]:
            self.tree.heading(col, text=col)
            self.tree.column(col, width=130)
        self.tree.pack(padx=20, pady=10, fill="x")

        button_frame = ttk.Frame(self)
        button_frame.pack(pady=10)
        ttk.Button(button_frame, text="Atgal", command=lambda: controller.show_page("EmployeeSelectPage")).pack(side="left", padx=5)
        ttk.Button(button_frame, text="Toliau → (nauja įranga)", command=lambda: controller.show_page("NewGearPage")).pack(side="left", padx=5)

    def refresh(self):
        for row in self.tree.get_children():
            self.tree.delete(row)
        employee = self.controller.current_employee
        if employee is None:
            return
        rows = self.controller.load_employee_rows()
        relevant = [row for row in rows if str(row.get("Tab. Nr", "")).strip() == employee.tab_nr]
        latest_info = None
        latest_date = None
        for row in relevant:
            issued = parse_date(row.get("Išduota"))
            if issued and (latest_date is None or issued > latest_date):
                latest_date = issued
                latest_info = row
        if latest_info:
            employee.department = str(latest_info.get("Padalinys", ""))
            employee.position = str(latest_info.get("Pareigos", ""))
            employee.gender = str(latest_info.get("Lytis", ""))

        filtered = [
            row
            for row in relevant
            if str(row.get("Padalinys", "")) == employee.department
            and str(row.get("Pareigos", "")) == employee.position
        ]

        latest_by_item = {}
        for row in filtered:
            name = str(row.get("Apranga", ""))
            if not name:
                continue
            base_name = strip_size_suffix(name)
            issued = parse_date(row.get("Išduota"))
            if issued is None:
                continue
            existing = latest_by_item.get(base_name)
            if existing is None or issued > existing["issued"]:
                latest_by_item[base_name] = {"row": row, "issued": issued}

        today = date.today()
        for item in latest_by_item.values():
            row = item["row"]
            issued = parse_date(row.get("Išduota"))
            months = int(row.get("Susidėvėjimas") or 0)
            change_date = add_months(issued, months)
            remaining = (change_date - today).days if change_date else ""
            values = (
                row.get("Apranga", ""),
                row.get("Aprangos kodas", ""),
                issued.strftime("%Y-%m-%d") if issued else "",
                months,
                change_date.strftime("%Y-%m-%d") if change_date else "",
                remaining,
            )
            item_id = self.tree.insert("", tk.END, values=values)
            if isinstance(remaining, int) and remaining < 7:
                self.tree.item(item_id, tags=("warn",))
        self.tree.tag_configure("warn", foreground="red")


class NewGearPage(ttk.Frame):
    def __init__(self, parent, controller):
        super().__init__(parent)
        self.controller = controller
        ttk.Label(self, text="Nauja apranga", font=("Arial", 16)).pack(pady=10)

        header = ttk.Frame(self)
        header.pack()
        ttk.Label(header, text="Kodas", width=20).grid(row=0, column=0)
        ttk.Label(header, text="Apranga", width=40).grid(row=0, column=1)
        ttk.Label(header, text="Susidėvėjimas (mėn.)", width=20).grid(row=0, column=2)

        self.entries = []
        self.after_ids = {}
        grid = ttk.Frame(self)
        grid.pack(pady=5)

        for i in range(14):
            code_var = tk.StringVar()
            name_var = tk.StringVar()
            months_var = tk.StringVar()
            code_entry = ttk.Entry(grid, textvariable=code_var, width=20)
            name_entry = ttk.Entry(grid, textvariable=name_var, width=40)
            months_entry = ttk.Entry(grid, textvariable=months_var, width=20)
            code_entry.grid(row=i, column=0, padx=2, pady=2)
            name_entry.grid(row=i, column=1, padx=2, pady=2)
            months_entry.grid(row=i, column=2, padx=2, pady=2)
            code_entry.bind("<KeyRelease>", lambda event, idx=i: self.on_code_change(idx))
            self.entries.append({
                "code": code_var,
                "name": name_var,
                "months": months_var,
                "code_entry": code_entry,
            })

        button_frame = ttk.Frame(self)
        button_frame.pack(pady=10)
        ttk.Button(button_frame, text="Atgal", command=lambda: controller.show_page("CurrentGearPage")).pack(side="left", padx=5)
        self.generate_close = ttk.Button(button_frame, text="Generuoti ir uždaryti", command=self.on_generate_close)
        self.generate_close.pack(side="left", padx=5)
        self.generate_new = ttk.Button(button_frame, text="Generuoti ir naujas įrašas", command=self.on_generate_new)
        self.generate_new.pack(side="left", padx=5)

    def refresh(self):
        for row in self.entries:
            row["code"].set("")
            row["name"].set("")
            row["months"].set("")

    def on_code_change(self, idx):
        if idx in self.after_ids:
            self.after_cancel(self.after_ids[idx])
        self.after_ids[idx] = self.after(200, lambda: self.apply_code(idx))

    def apply_code(self, idx):
        entry = self.entries[idx]
        code = entry["code"].get().strip()
        if not code:
            entry["name"].set("")
            entry["months"].set("")
            return
        if "-" in code:
            base, size = code.split("-", 1)
        else:
            base, size = code, ""
        base = base.strip()
        size = size.strip()
        gear_info = self.controller.gear_codes.get(base)
        if gear_info:
            name = gear_info.get("name", "")
            name = ensure_size_suffix(name, size)
            entry["name"].set(name)
            entry["months"].set(str(gear_info.get("months", "")))
        else:
            entry["name"].set(entry["name"].get())

    def collect_items(self):
        items = []
        for row in self.entries:
            code = row["code"].get().strip()
            name = row["name"].get().strip()
            months = row["months"].get().strip()
            if not code and not name and not months:
                continue
            if "-" in code:
                base, size = code.split("-", 1)
            else:
                base, size = code, ""
            base = base.strip()
            size = size.strip()
            final_name = ensure_size_suffix(name, size)
            try:
                months_value = int(months)
            except ValueError:
                months_value = 0
            items.append({"code": base, "name": final_name, "months": months_value})
        return items

    def update_gear_codes(self, items):
        updated = False
        for item in items:
            base = item["code"]
            if base and base not in self.controller.gear_codes:
                self.controller.gear_codes[base] = {
                    "name": strip_size_suffix(item["name"]),
                    "months": item["months"],
                }
                updated = True
        if updated:
            save_gear_codes(self.controller.settings["gear"], self.controller.gear_codes)

    def toggle_buttons(self, state):
        self.generate_close.config(state=state)
        self.generate_new.config(state=state)

    def run_generation(self, close_after):
        items = self.collect_items()
        if not items:
            messagebox.showwarning("Trūksta duomenų", "Įveskite bent vieną aprangos eilutę.")
            self.toggle_buttons("normal")
            return
        employee = self.controller.current_employee
        document_number = self.controller.get_next_document_number()
        self.update_gear_codes(items)
        self.controller.append_issuance_rows(employee, items, document_number)
        self.controller.generate_word_doc(employee, items, document_number)
        if close_after:
            self.controller.destroy()
        else:
            self.controller.show_page("EmployeeSelectPage")
        self.toggle_buttons("normal")

    def on_generate_close(self):
        self.toggle_buttons("disabled")
        threading.Thread(target=self.run_generation, args=(True,), daemon=True).start()

    def on_generate_new(self):
        self.toggle_buttons("disabled")
        threading.Thread(target=self.run_generation, args=(False,), daemon=True).start()


if __name__ == "__main__":
    settings_file = os.environ.get("AAP_SETTINGS", "settings.json")
    app = AAPApp(settings_file)
    app.mainloop()
