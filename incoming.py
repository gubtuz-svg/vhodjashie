# -*- coding: utf-8 -*-
import os
import sqlite3
import tkinter as tk
from tkinter import ttk, messagebox, filedialog
from datetime import date

DB_PATH = os.path.join(os.path.dirname(os.path.abspath(__file__)), "incoming.db")

def get_conn():
    conn = sqlite3.connect(DB_PATH)
    conn.execute("""
        CREATE TABLE IF NOT EXISTS documents (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            in_number TEXT,
            in_date TEXT,
            sender TEXT,
            subject TEXT,
            case_name TEXT,
            note TEXT
        )
    """)
    conn.commit()
    return conn

class App(tk.Tk):
    def __init__(self):
        super().__init__()
        self.title("Регистрация входящей документации")
        self.geometry("1000x600")
        self.conn = get_conn()
        self.selected_id = None
        self._build_ui()
        self.load_data()

    def _build_ui(self):
        # --- Форма ввода ---
        form = ttk.LabelFrame(self, text="Новый документ / редактирование")
        form.pack(fill="x", padx=10, pady=5)

        self.vars = {}
        fields = [
            ("Входящий №:", "in_number"),
            ("Дата (ГГГГ-ММ-ДД):", "in_date"),
            ("От кого:", "sender"),
            ("Для чего / кратко:", "subject"),
            ("Дело (папка/номенклатура):", "case_name"),
        ]
        for i, (label, key) in enumerate(fields):
            ttk.Label(form, text=label).grid(row=i // 3, column=(i % 3) * 2, sticky="e", padx=4, pady=4)
            var = tk.StringVar()
            self.vars[key] = var
            ttk.Entry(form, textvariable=var, width=28).grid(row=i // 3, column=(i % 3) * 2 + 1, sticky="w", padx=4)

        ttk.Label(form, text="Примечание:").grid(row=2, column=0, sticky="e", padx=4, pady=4)
        self.vars["note"] = tk.StringVar()
        ttk.Entry(form, textvariable=self.vars["note"], width=60).grid(row=2, column=1, columnspan=5, sticky="w", padx=4)

        btns = ttk.Frame(form)
        btns.grid(row=3, column=0, columnspan=6, pady=6)
        ttk.Button(btns, text="Добавить", command=self.add_record).pack(side="left", padx=3)
        ttk.Button(btns, text="Сохранить изменения", command=self.update_record).pack(side="left", padx=3)
        ttk.Button(btns, text="Очистить форму", command=self.clear_form).pack(side="left", padx=3)
        ttk.Button(btns, text="Удалить выбранное", command=self.delete_record).pack(side="left", padx=3)

        # --- Поиск ---
        search = ttk.Frame(self)
        search.pack(fill="x", padx=10, pady=2)
        ttk.Label(search, text="Поиск:").pack(side="left")
        self.search_var = tk.StringVar()
        ttk.Entry(search, textvariable=self.search_var, width=40).pack(side="left", padx=5)
        ttk.Button(search, text="Найти", command=self.load_data).pack(side="left", padx=2)
        ttk.Button(search, text="Сброс", command=self.reset_search).pack(side="left", padx=2)

        # --- Таблица ---
        cols = ("id", "in_number", "in_date", "sender", "subject", "case_name", "note")
        headers = ("ID", "Вх. №", "Дата", "От кого", "Для чего", "Дело", "Примечание")
        widths = (40, 80, 90, 180, 220, 160, 140)
        self.tree = ttk.Treeview(self, columns=cols, show="headings")
        for c, h, w in zip(cols, headers, widths):
            self.tree.heading(c, text=h)
            self.tree.column(c, width=w, anchor="w")
        self.tree.pack(fill="both", expand=True, padx=10, pady=5)
        self.tree.bind("<<TreeviewSelect>>", self.on_select)

        # --- Кнопки печати ---
        pbar = ttk.Frame(self)
        pbar.pack(fill="x", padx=10, pady=5)
        ttk.Button(pbar, text="🖨 Печать выбранного документа", command=self.print_selected).pack(side="left", padx=3)
        ttk.Button(pbar, text="🖨 Печать всего журнала", command=self.print_all).pack(side="left", padx=3)
        ttk.Label(pbar, text="(откроется Word — нажмите Ctrl+P)").pack(side="left", padx=10)

    # ---------- Данные ----------
    def load_data(self, *args):
        q = self.search_var.get().strip() if hasattr(self, "search_var") else ""
        for i in self.tree.get_children():
            self.tree.delete(i)
        if q:
            rows = self.conn.execute(
                "SELECT * FROM documents WHERE in_number LIKE ? OR sender LIKE ? OR subject LIKE ? OR case_name LIKE ? ORDER BY in_date DESC, id DESC",
                (f"%{q}%",) * 4,
            ).fetchall()
        else:
            rows = self.conn.execute("SELECT * FROM documents ORDER BY in_date DESC, id DESC").fetchall()
        for r in rows:
            self.tree.insert("", "end", values=r)

    def reset_search(self):
        self.search_var.set("")
        self.load_data()

    def add_record(self):
        v = {k: s.get().strip() for k, s in self.vars.items()}
        if not v["in_number"] or not v["in_date"]:
            messagebox.showwarning("Внимание", "Заполните входящий номер и дату.")
            return
        self.conn.execute(
            "INSERT INTO documents (in_number, in_date, sender, subject, case_name, note) VALUES (?,?,?,?,?,?)",
            (v["in_number"], v["in_date"], v["sender"], v["subject"], v["case_name"], v["note"]),
        )
        self.conn.commit()
        self.load_data()
        self.clear_form()

    def update_record(self):
        if self.selected_id is None:
            messagebox.showinfo("Инфо", "Сначала выберите запись в таблице.")
            return
        v = {k: s.get().strip() for k, s in self.vars.items()}
        self.conn.execute(
            "UPDATE documents SET in_number=?, in_date=?, sender=?, subject=?, case_name=?, note=? WHERE id=?",
            (v["in_number"], v["in_date"], v["sender"], v["subject"], v["case_name"], v["note"], self.selected_id),
        )
        self.conn.commit()
        self.load_data()

    def delete_record(self):
        if self.selected_id is None:
            messagebox.showinfo("Инфо", "Выберите запись для удаления.")
            return
        if messagebox.askyesno("Подтверждение", "Удалить выбранную запись?"):
            self.conn.execute("DELETE FROM documents WHERE id=?", (self.selected_id,))
            self.conn.commit()
            self.clear_form()
            self.load_data()

    def on_select(self, event):
        sel = self.tree.selection()
        if not sel:
            return
        vals = self.tree.item(sel[0], "values")
        self.selected_id = vals[0]
        keys = ["in_number", "in_date", "sender", "subject", "case_name", "note"]
        for k, val in zip(keys, vals[1:]):
            self.vars[k].set(val)

    def clear_form(self):
        self.selected_id = None
        for s in self.vars.values():
            s.set("")
        for i in self.tree.selection():
            self.tree.selection_remove(i)

    # ---------- Печать ----------
    def _make_word_file(self, rows, title, one=False):
        try:
            from docx import Document
            from docx.shared import Pt
        except ImportError:
            messagebox.showerror(
                "Нужен модуль python-docx",
                "Для печати установите библиотеку:\n\npip install python-docx\n\n(один раз в командной строке)"
            )
            return

        doc = Document()
        doc.add_heading(title, level=1)

        if one:
            r = rows[0]
            pairs = [
                ("Входящий номер", r[1]), ("Дата", r[2]), ("От кого", r[3]),
                ("Для чего / кратко", r[4]), ("Дело", r[5]), ("Примечание", r[6]),
            ]
            for k, v in pairs:
                p = doc.add_paragraph()
                p.add_run(f"{k}: ").bold = True
                p.add_run(str(v) if v else "—")
        else:
            table = doc.add_table(rows=1, cols=7)
            table.style = "Table Grid"
            hdr = ["№", "Вх. №", "Дата", "От кого", "Для чего", "Дело", "Примечание"]
            for i, h in enumerate(hdr):
                cell = table.rows[0].cells[i]
                cell.text = h
                cell.paragraphs[0].runs[0].bold = True
            for idx, r in enumerate(rows, 1):
                cells = table.add_row().cells
                cells[0].text = str(idx)
                for i, val in enumerate(r[1:], 1):
                    cells[i].text = str(val) if val else ""
            for row in table.rows:
                for cell in row.cells:
                    for p in cell.paragraphs:
                        for run in p.runs:
                            run.font.size = Pt(9)

        out = filedialog.asksaveasfilename(
            defaultextension=".docx",
            filetypes=[("Word документ", "*.docx")],
            initialfile="Печать_входящих.docx",
        )
        if not out:
            return
        doc.save(out)
        messagebox.showinfo("Готово", f"Файл сохранён:\n{out}\n\nОткройте его и нажмите Ctrl+P для печати.")
        try:
            os.startfile(out)
        except Exception:
            pass

    def print_selected(self):
        sel = self.tree.selection()
        if not sel:
            messagebox.showinfo("Инфо", "Выберите запись в таблице.")
            return
        vals = self.tree.item(sel[0], "values")
        self._make_word_file([vals], "Входящий документ", one=True)

    def print_all(self):
        rows = self.conn.execute("SELECT * FROM documents ORDER BY in_date DESC, id DESC").fetchall()
        if not rows:
            messagebox.showinfo("Инфо", "Журнал пуст.")
            return
        self._make_word_file(rows, "Журнал входящей документации")

if __name__ == "__main__":
    app = App()
    app.mainloop()
