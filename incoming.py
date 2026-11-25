# -*- coding: utf-8 -*-
import os
import sys
import glob
import sqlite3
import tkinter as tk
from tkinter import ttk, messagebox, filedialog
from datetime import datetime
from PIL import Image, ImageTk

# ---------- Пути: рядом с .exe (если собран) или рядом с .py ----------
if getattr(sys, "frozen", False):
    BASE_DIR = os.path.dirname(sys.executable)
else:
    BASE_DIR = os.path.dirname(os.path.abspath(__file__))

DB_PATH = os.path.join(BASE_DIR, "incoming.db")
SCANS_DIR = os.path.join(BASE_DIR, "scans")
os.makedirs(SCANS_DIR, exist_ok=True)

def _clean_filename(s: str) -> str:
    bad = '\\/:*?"<>|'
    for ch in bad:
        s = s.replace(ch, "-")
    return s.strip() or "БезНомера"

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
    cols = [r[1] for r in conn.execute("PRAGMA table_info(documents)").fetchall()]
    if "scan_path" not in cols:
        conn.execute("ALTER TABLE documents ADD COLUMN scan_path TEXT")
    conn.commit()
    return conn

# ---------- Сканирование (JPG, с опцией Ч/Б) ----------
def save_jpg(img: Image.Image, path: str, bw: bool):
    if bw:
        img = img.convert("L")
    else:
        img = img.convert("RGB")
    img.save(path, "JPEG", quality=85)

def scan_twain_jpg(prefix: str, bw: bool, dpi=300):
    files = []
    try:
        import twain
        sm = twain.SourceManager(0)
        sources = sm.GetSourceList()
        if not sources:
            sm.destroy(); return files
        epson = next((s for s in sources if "epson" in s.lower()), sources[0])
        src = sm.OpenSource(epson)
        if not src:
            sm.destroy(); return files
        src.SetCapability(twain.ICAP_XRESOLUTION, twain.TWTY_FIX32, float(dpi))
        src.SetCapability(twain.ICAP_YRESOLUTION, twain.TWTY_FIX32, float(dpi))
        src.SetCapability(twain.ICAP_PIXELTYPE, twain.TWTY_UINT16, 0 if bw else 2)
        src.RequestAcquire(0, 0)
        page = 1
        while True:
            rv = src.XferImageNatively()
            if rv[0] != twain.TWRC_SUCCESS:
                break
            handle = rv[1]
            tmp_bmp = os.path.join(SCANS_DIR, f"tmp_{datetime.now().timestamp()}_{page}.bmp")
            twain.DIBToBMFile(handle, tmp_bmp)
            jpg = f"{prefix}_{page}.jpg"
            try:
                save_jpg(Image.open(tmp_bmp), jpg, bw)
                files.append(jpg)
            finally:
                try: os.remove(tmp_bmp)
                except: pass
            page += 1
        src.CloseSource()
        sm.destroy()
    except Exception as e:
        print("TWAIN error:", e)
    return files

def scan_wia_jpg(prefix: str, bw: bool):
    files = []
    try:
        import win32com.client
        WIA = win32com.client.Dispatch("WIA.CommonDialog")
        img = WIA.ShowAcquireImage()
        if not img:
            return files
        tmp_bmp = os.path.join(SCANS_DIR, f"tmp_wia_{datetime.now().timestamp()}.bmp")
        img.SaveFile(tmp_bmp)
        jpg = f"{prefix}_1.jpg"
        save_jpg(Image.open(tmp_bmp), jpg, bw)
        files.append(jpg)
        try: os.remove(tmp_bmp)
        except: pass
    except Exception as e:
        print("WIA error:", e)
    return files

def do_scan_jpg(prefix: str, bw: bool):
    files = scan_twain_jpg(prefix, bw)
    if not files:
        files = scan_wia_jpg(prefix, bw)
    return files


# ---------- Окно просмотра скана ----------
class ScanViewer(tk.Toplevel):
    def __init__(self, master, files):
        super().__init__(master)
        self.title("Просмотр скана")
        self.geometry("960x740")
        self.files = [f for f in files if os.path.exists(f)]
        self.idx = 0
        self.zoom = 1.0
        self.rotation = 0
        self._photo = None

        top = ttk.Frame(self); top.pack(fill="x", padx=6, pady=4)
        ttk.Button(top, text="◀ Пред.", command=self.prev).pack(side="left", padx=2)
        ttk.Button(top, text="След. ▶", command=self.next).pack(side="left", padx=2)
        ttk.Separator(top, orient="vertical").pack(side="left", fill="y", padx=6)
        ttk.Button(top, text="↺ 90°", command=lambda: self.rotate(-90)).pack(side="left", padx=2)
        ttk.Button(top, text="↻ 90°", command=lambda: self.rotate(90)).pack(side="left", padx=2)
        ttk.Button(top, text="💾 Сохранить поворот", command=self.save_rotation).pack(side="left", padx=2)
        ttk.Separator(top, orient="vertical").pack(side="left", fill="y", padx=6)
        ttk.Button(top, text="−", command=lambda: self.change_zoom(0.9)).pack(side="left")
        self.lbl_zoom = ttk.Label(top, text="100%"); self.lbl_zoom.pack(side="left", padx=2)
        ttk.Button(top, text="+", command=lambda: self.change_zoom(1.1)).pack(side="left")
        ttk.Button(top, text="Сброс", command=self.reset_zoom).pack(side="left", padx=2)
        ttk.Button(top, text="🖨 Печать скана", command=self.print_scan).pack(side="left", padx=8)

        self.lbl_page = ttk.Label(top, text=""); self.lbl_page.pack(side="right", padx=6)
        ttk.Button(top, text="Закрыть", command=self.destroy).pack(side="right", padx=2)

        self.canvas = tk.Canvas(self, bg="#2b2b2b")
        self.canvas.pack(fill="both", expand=True)
        self.canvas.bind("<MouseWheel>", self.on_wheel)
        self.canvas.bind("<Configure>", lambda e: self.render())

        if not self.files:
            messagebox.showinfo("Просмотр", "Файлы скана не найдены.")
            self.destroy(); return
        self.render()

    def change_zoom(self, k):
        self.zoom = max(0.1, min(4.0, self.zoom * k))
        self.lbl_zoom.config(text=f"{int(self.zoom*100)}%")
        self.render()

    def reset_zoom(self):
        self.zoom = 1.0
        self.lbl_zoom.config(text="100%")
        self.render()

    def prev(self):
        if self.idx > 0:
            self.idx -= 1; self.rotation = 0; self.render()

    def next(self):
        if self.idx < len(self.files) - 1:
            self.idx += 1; self.rotation = 0; self.render()

    def on_wheel(self, e):
        self.change_zoom(1.1 if e.delta > 0 else 0.9)

    def rotate(self, deg):
        self.rotation = (self.rotation + deg) % 360
        self.render()

    def save_rotation(self):
        if not self.files:
            return
        f = self.files[self.idx]
        if self.rotation == 0:
            messagebox.showinfo("Поворот", "Поворот не задан.")
            return
        try:
            img = Image.open(f)
            img = img.rotate(-self.rotation, expand=True)
            img.save(f, "JPEG", quality=85)
            self.rotation = 0
            self.render()
            messagebox.showinfo("Готово", "Поворот сохранён в файл.")
        except Exception as e:
            messagebox.showerror("Ошибка", f"Не удалось сохранить: {e}")

    def print_scan(self):
        if not self.files:
            return
        try:
            from docx import Document
            from docx.shared import Cm
        except ImportError:
            messagebox.showerror("Нужен модуль python-docx", "Выполните: pip install python-docx")
            return
        doc = Document()
        for section in doc.sections:
            section.top_margin = Cm(1)
            section.bottom_margin = Cm(1)
            section.left_margin = Cm(1)
            section.right_margin = Cm(1)
        for f in self.files:
            try:
                img = Image.open(f)
                if f == self.files[self.idx] and self.rotation:
                    img = img.rotate(-self.rotation, expand=True)
                tmp = os.path.join(SCANS_DIR, f"_print_{datetime.now().timestamp()}.jpg")
                img.convert("RGB").save(tmp, "JPEG", quality=90)
                doc.add_picture(tmp, width=Cm(17))
                try: os.remove(tmp)
                except: pass
                doc.add_page_break()
            except Exception as e:
                print("print err:", e)
        out = filedialog.asksaveasfilename(
            defaultextension=".docx",
            filetypes=[("Word документ", "*.docx")],
            initialfile="Скан_для_печати.docx",
        )
        if not out:
            return
        doc.save(out)
        messagebox.showinfo("Готово", f"Файл сохранён:\n{out}\n\nОткройте и нажмите Ctrl+P.")
        try: os.startfile(out)
        except: pass

    def render(self):
        self.canvas.delete("all")
        if not self.files:
            return
        f = self.files[self.idx]
        self.lbl_page.config(text=f"Стр. {self.idx+1} из {len(self.files)}  —  {os.path.basename(f)}")
        try:
            img = Image.open(f)
        except Exception as e:
            self.canvas.create_text(20, 20, anchor="nw", fill="white", text=f"Ошибка: {e}")
            return
        if self.rotation:
            img = img.rotate(-self.rotation, expand=True)
        cw = max(self.canvas.winfo_width(), 200)
        ch = max(self.canvas.winfo_height(), 200)
        w, h = img.size
        fit = min(cw / w, ch / h) if w and h else 1
        scale = fit * self.zoom
        nw, nh = max(1, int(w * scale)), max(1, int(h * scale))
        img = img.resize((nw, nh), Image.LANCZOS)
        self._photo = ImageTk.PhotoImage(img)
        self.canvas.create_image(cw // 2, ch // 2, image=self._photo, anchor="center")


class App(tk.Tk):
    def __init__(self):
        super().__init__()
        self.title("Регистрация входящей документации")
        self.geometry("1200x700")
        self.conn = get_conn()
        self.selected_id = None
        self._build_ui()
        self.load_data()

    def _build_ui(self):
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

        scan_frame = ttk.Frame(form)
        scan_frame.grid(row=4, column=0, columnspan=6, pady=4, sticky="w")
        ttk.Button(scan_frame, text="📠 Сканировать (JPG)", command=self.scan_current).pack(side="left", padx=3)
        self.bw_var = tk.BooleanVar(value=False)
        ttk.Checkbutton(scan_frame, text="Ч/Б (экономия места)", variable=self.bw_var).pack(side="left", padx=10)

        search = ttk.Frame(self)
        search.pack(fill="x", padx=10, pady=2)
        ttk.Label(search, text="Поиск:").pack(side="left")
        self.search_var = tk.StringVar()
        ttk.Entry(search, textvariable=self.search_var, width=40).pack(side="left", padx=5)
        ttk.Button(search, text="Найти", command=self.load_data).pack(side="left", padx=2)
        ttk.Button(search, text="Сброс", command=self.reset_search).pack(side="left", padx=2)

        cols = ("id", "in_number", "in_date", "sender", "subject", "case_name", "note", "scan")
        headers = ("ID", "Вх. №", "Дата", "От кого", "Для чего", "Дело", "Примечание", "Скан")
        widths = (40, 80, 90, 170, 210, 150, 130, 60)
        self.tree = ttk.Treeview(self, columns=cols, show="headings")
        for c, h, w in zip(cols, headers, widths):
            self.tree.heading(c, text=h)
            self.tree.column(c, width=w, anchor="w")
        self.tree.pack(fill="both", expand=True, padx=10, pady=5)
        self.tree.bind("<<TreeviewSelect>>", self.on_select)
        self.tree.bind("<Double-1>", lambda e: self.open_scan())

        pbar = ttk.Frame(self)
        pbar.pack(fill="x", padx=10, pady=5)
        ttk.Button(pbar, text="🖨 Печать выбранного", command=self.print_selected).pack(side="left", padx=3)
        ttk.Button(pbar, text="🖨 Печать журнала", command=self.print_all).pack(side="left", padx=3)
        ttk.Button(pbar, text="👁 Просмотр скана", command=self.open_scan).pack(side="left", padx=12)
        ttk.Button(pbar, text="🗑 Удалить скан", command=self.delete_scan).pack(side="left", padx=3)
        ttk.Button(pbar, text="📁 Открыть папку сканов", command=self.open_scans_folder).pack(side="left", padx=12)
        ttk.Label(pbar, text="(двойной клик по строке — просмотр)").pack(side="left", padx=6)

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
            scan_ok = "✔" if r[7] and any(os.path.exists(p) for p in r[7].split(";") if p) else ""
            self.tree.insert("", "end", values=(*r[:7], scan_ok))

    def reset_search(self):
        self.search_var.set("")
        self.load_data()

    def add_record(self):
        v = {k: s.get().strip() for k, s in self.vars.items()}
        if not v["in_number"] or not v["in_date"]:
            messagebox.showwarning("Внимание", "Заполните входящий номер и дату.")
            return
        cur = self.conn.execute(
            "INSERT INTO documents (in_number, in_date, sender, subject, case_name, note) VALUES (?,?,?,?,?,?)",
            (v["in_number"], v["in_date"], v["sender"], v["subject"], v["case_name"], v["note"]),
        )
        self.conn.commit()
        self.selected_id = cur.lastrowid
        self.load_data()
        self.clear_form()
        messagebox.showinfo("Готово", "Запись добавлена. Нажмите «📠 Сканировать (JPG)», чтобы прикрепить скан.")

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
        if messagebox.askyesno("Подтверждение", "Удалить запись? Файлы сканов останутся в папке scans."):
            self.conn.execute("DELETE FROM documents WHERE id=?", (self.selected_id,))
            self.conn.commit()
            self.clear_form()
            self.load_data()

    def on_select(self, event):
        sel = self.tree.selection()
        if not sel:
            return
        vals = self.tree.item(sel[0], "values")
        self.selected_id = int(vals[0])
        keys = ["in_number", "in_date", "sender", "subject", "case_name", "note"]
        for k, val in zip(keys, vals[1:7]):
            self.vars[k].set(val)

    def clear_form(self):
        self.selected_id = None
        for s in self.vars.values():
            s.set("")
        for i in self.tree.selection():
            self.tree.selection_remove(i)

    # ---------- Сканирование ----------
    def scan_current(self):
        if self.selected_id is None:
            messagebox.showinfo("Инфо", "Сначала выберите запись (или добавьте её), к которой прикрепить скан.")
            return

        row = self.conn.execute("SELECT in_number, in_date FROM documents WHERE id=?", (self.selected_id,)).fetchone()
        raw_num = (row[0] or "БезНомера").strip()
        raw_date = (row[1] or "безДаты").strip()

        safe_num = _clean_filename(raw_num)
        safe_date = _clean_filename(raw_date)

        prefix = os.path.join(SCANS_DIR, f"{safe_num}_{safe_date}")

        # удаляем старые JPG с этим же префиксом
        for old in glob.glob(prefix + "_*.jpg"):
            try:
                os.remove(old)
            except Exception:
                pass

        bw = self.bw_var.get()
        messagebox.showinfo("Сканирование",
                            f"Откроется окно сканера (Epson Scan / WIA).\n"
                            f"Режим: {'Ч/Б' if bw else 'Цветной'}\n"
                            f"Имя файла: {os.path.basename(prefix)}_N.jpg\n"
                            f"Отсканируйте страницы и подтвердите.")
        files = do_scan_jpg(prefix, bw)
        if files:
            self.conn.execute("UPDATE documents SET scan_path=? WHERE id=?", (";".join(files), self.selected_id))
            self.conn.commit()
            self.load_data()
            if messagebox.askyesno("Готово", f"Сохранено страниц: {len(files)}.\n\nОткрыть просмотр сейчас?"):
                self.open_scan()
        else:
            messagebox.showerror("Ошибка", "Сканирование не выполнено.\nПроверьте драйвер Epson / TWAIN / WIA.")

    def open_scan(self):
        if self.selected_id is None:
            messagebox.showinfo("Инфо", "Выберите запись в таблице.")
            return
        row = self.conn.execute("SELECT scan_path FROM documents WHERE id=?", (self.selected_id,)).fetchone()
        if not row or not row[0]:
            messagebox.showinfo("Инфо", "У этой записи нет скана. Нажмите «📠 Сканировать (JPG)».")
            return
        files = [p for p in row[0].split(";") if p and os.path.exists(p)]
        if not files:
            messagebox.showinfo("Инфо", "Файлы скана не найдены на диске (возможно, удалены).")
            return
        ScanViewer(self, files)

    def delete_scan(self):
        if self.selected_id is None:
            messagebox.showinfo("Инфо", "Выберите запись в таблице.")
            return
        row = self.conn.execute("SELECT scan_path FROM documents WHERE id=?", (self.selected_id,)).fetchone()
        if not row or not row[0]:
            messagebox.showinfo("Инфо", "У этой записи нет скана.")
            return
        files = [p for p in row[0].split(";") if p]
        if not messagebox.askyesno("Удаление скана",
                                   f"Удалить скан у записи?\nФайлов: {len(files)}\n\n"
                                   f"Файлы будут удалены с диска безвозвратно."):
            return
        deleted, errors = 0, []
        for f in files:
            try:
                if os.path.exists(f):
                    os.remove(f)
                    deleted += 1
            except Exception as e:
                errors.append(f"{os.path.basename(f)}: {e}")
        self.conn.execute("UPDATE documents SET scan_path=NULL WHERE id=?", (self.selected_id,))
        self.conn.commit()
        self.load_data()
        msg = f"Удалено файлов: {deleted}"
        if errors:
            msg += "\n\nОшибки:\n" + "\n".join(errors)
        messagebox.showinfo("Готово", msg)

    def open_scans_folder(self):
        os.startfile(SCANS_DIR)

    # ---------- Печать ----------
    def _make_word_file(self, rows, title, one=False):
        try:
            from docx import Document
            from docx.shared import Pt
        except ImportError:
            messagebox.showerror("Нужен модуль python-docx", "Выполните: pip install python-docx")
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
            scan = r[7] if len(r) > 7 else ""
            if scan:
                p = doc.add_paragraph()
                p.add_run("Скан: ").bold = True
                p.add_run(scan)
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
                for i, val in enumerate(r[1:7], 1):
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
        messagebox.showinfo("Готово", f"Файл сохранён:\n{out}\n\nОткройте и нажмите Ctrl+P.")
        try: os.startfile(out)
        except: pass

    def print_selected(self):
        sel = self.tree.selection()
        if not sel:
            messagebox.showinfo("Инфо", "Выберите запись.")
            return
        vals = self.tree.item(sel[0], "values")
        row = self.conn.execute("SELECT * FROM documents WHERE id=?", (int(vals[0]),)).fetchone()
        self._make_word_file([row], "Входящий документ", one=True)

    def print_all(self):
        rows = self.conn.execute("SELECT * FROM documents ORDER BY in_date DESC, id DESC").fetchall()
        if not rows:
            messagebox.showinfo("Инфо", "Журнал пуст.")
            return
        self._make_word_file(rows, "Журнал входящей документации")

if __name__ == "__main__":
    app = App()
    app.mainloop()
