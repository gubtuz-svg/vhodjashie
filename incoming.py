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

# ---------- Цвета / стиль ----------
C_BG        = "#F4F6FB"   # фон окна
C_CARD      = "#FFFFFF"   # карточки
C_ACCENT    = "#4A6CF7"   # акцент (кнопки)
C_ACCENT_H  = "#3A56C8"   # акцент при наведении
C_TEXT      = "#1F2937"   # текст
C_MUTED     = "#6B7280"   # подписи
C_ROW_ALT   = "#F9FAFF"   # чередование строк
C_DANGER    = "#E5484D"   # удаление

FONT_MAIN   = ("Segoe UI", 10)
FONT_BOLD   = ("Segoe UI", 10, "bold")
FONT_TITLE  = ("Segoe UI", 16, "bold")
FONT_SMALL  = ("Segoe UI", 9)

# ---------- Утилиты ----------
def _clean_filename(s: str) -> str:
    bad = '\\/:*?"<>|'
    for ch in bad:
        s = s.replace(ch, "-")
    s = s.strip().strip(".")
    return s or "БезНазвания"

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

# ---------- Сканирование ----------
def save_jpg(img: Image.Image, path: str, bw: bool):
    img = img.convert("L") if bw else img.convert("RGB")
    img.save(path, "JPEG", quality=85)

# Соответствие «размер листа» -> (ширина_дюймы, высота_дюймы) для TWAIN
PAPER_SIZES = {
    "A4 (210×297)":      (8.27, 11.69),
    "A5 (148×210)":      (5.83, 8.27),
    "Letter (8.5×11)":   (8.5, 11.0),
    "Legal (8.5×14)":    (8.5, 14.0),
    "Авто (по краю)":    None,
}

def scan_twain_jpg(prefix: str, bw: bool, dpi: int, paper: str):
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

        # Размер листа
        size = PAPER_SIZES.get(paper)
        if size:
            try:
                w_in, h_in = size
                src.SetCapability(twain.ICAP_SUPPORTEDSIZES, twain.TWTY_UINT16, 1)  # 1 = user-defined
                # в TWAIN единицы — дюймы * 65536 (fix32)
                src.SetCapability(twain.ICAP_FRAMES, twain.TWTY_FRAME,
                                  (0, 0, int(w_in * 65536), int(h_in * 65536)))
            except Exception as e:
                print("paper size not applied:", e)

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

def do_scan_jpg(prefix: str, bw: bool, dpi: int, paper: str):
    files = scan_twain_jpg(prefix, bw, dpi, paper)
    if not files:
        files = scan_wia_jpg(prefix, bw)
    return files


# ---------- Просмотр скана ----------
class ScanViewer(tk.Toplevel):
    def __init__(self, master, files):
        super().__init__(master)
        self.title("Просмотр скана")
        self.geometry("960x740")
        self.configure(bg=C_BG)
        self.files = [f for f in files if os.path.exists(f)]
        self.idx = 0
        self.zoom = 1.0
        self.rotation = 0
        self._photo = None

        bar = tk.Frame(self, bg=C_CARD); bar.pack(fill="x", padx=8, pady=6)
        def btn(txt, cmd, bg=C_ACCENT):
            return tk.Button(bar, text=txt, command=cmd, bg=bg, fg="white",
                             activebackground=C_ACCENT_H, activeforeground="white",
                             relief="flat", bd=0, padx=10, pady=4, font=FONT_SMALL,
                             cursor="hand2")
        btn("◀", self.prev).pack(side="left", padx=2)
        btn("▶", self.next).pack(side="left", padx=2)
        btn("↺ 90°", lambda: self.rotate(-90)).pack(side="left", padx=6)
        btn("↻ 90°", lambda: self.rotate(90)).pack(side="left", padx=2)
        btn("💾 Поворот", self.save_rotation).pack(side="left", padx=6)
        btn("−", lambda: self.change_zoom(0.9)).pack(side="left", padx=2)
        self.lbl_zoom = tk.Label(bar, text="100%", bg=C_CARD, fg=C_TEXT, font=FONT_SMALL)
        self.lbl_zoom.pack(side="left", padx=2)
        btn("+", lambda: self.change_zoom(1.1)).pack(side="left", padx=2)
        btn("Сброс", self.reset_zoom, bg=C_MUTED).pack(side="left", padx=6)
        btn("🖨 Печать", self.print_scan, bg="#16A34A").pack(side="left", padx=6)
        btn("Закрыть", self.destroy, bg=C_DANGER).pack(side="right", padx=2)
        self.lbl_page = tk.Label(bar, text="", bg=C_CARD, fg=C_MUTED, font=FONT_SMALL)
        self.lbl_page.pack(side="right", padx=8)

        self.canvas = tk.Canvas(self, bg="#1F2937", highlightthickness=0)
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
        if not self.files: return
        f = self.files[self.idx]
        if self.rotation == 0:
            messagebox.showinfo("Поворот", "Поворот не задан."); return
        try:
            img = Image.open(f).rotate(-self.rotation, expand=True)
            img.save(f, "JPEG", quality=85)
            self.rotation = 0; self.render()
            messagebox.showinfo("Готово", "Поворот сохранён.")
        except Exception as e:
            messagebox.showerror("Ошибка", f"Не удалось сохранить: {e}")

    def print_scan(self):
        if not self.files: return
        try:
            from docx import Document
            from docx.shared import Cm
        except ImportError:
            messagebox.showerror("Нужен модуль python-docx", "Выполните: pip install python-docx"); return
        doc = Document()
        for s in doc.sections:
            s.top_margin = s.bottom_margin = s.left_margin = s.right_margin = Cm(1)
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
        out = filedialog.asksaveasfilename(defaultextension=".docx",
            filetypes=[("Word документ", "*.docx")], initialfile="Скан_для_печати.docx")
        if not out: return
        doc.save(out)
        messagebox.showinfo("Готово", f"Файл сохранён:\n{out}\n\nОткройте и нажмите Ctrl+P.")
        try: os.startfile(out)
        except: pass

    def render(self):
        self.canvas.delete("all")
        if not self.files: return
        f = self.files[self.idx]
        self.lbl_page.config(text=f"Стр. {self.idx+1} из {len(self.files)}  —  {os.path.basename(f)}")
        try:
            img = Image.open(f)
        except Exception as e:
            self.canvas.create_text(20, 20, anchor="nw", fill="white", text=f"Ошибка: {e}"); return
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


# ---------- Главное окно ----------
class App(tk.Tk):
    def __init__(self):
        super().__init__()
        self.title("Регистрация входящей документации")
        self.geometry("1280x760")
        self.minsize(1100, 680)
        self.configure(bg=C_BG)
        self._apply_style()
        self.conn = get_conn()
        self.selected_id = None
        self.sort_col = "in_date"
        self.sort_desc = True
        self._build_ui()
        self.load_data()

    def _apply_style(self):
        style = ttk.Style(self)
        try:
            style.theme_use("clam")
        except Exception:
            pass
        style.configure(".", font=FONT_MAIN, background=C_BG, foreground=C_TEXT)
        style.configure("TFrame", background=C_BG)
        style.configure("Card.TFrame", background=C_CARD)
        style.configure("TLabel", background=C_BG, foreground=C_TEXT, font=FONT_MAIN)
        style.configure("Muted.TLabel", background=C_BG, foreground=C_MUTED, font=FONT_SMALL)
        style.configure("Title.TLabel", background=C_BG, foreground=C_TEXT, font=FONT_TITLE)
        style.configure("TEntry", fieldbackground="white", bordercolor="#D1D5DB",
                        lightcolor="#D1D5DB", darkcolor="#D1D5DB", padding=6)
        style.configure("TCombobox", fieldbackground="white", padding=5)
        style.configure("TCheckbutton", background=C_CARD, foreground=C_TEXT, font=FONT_MAIN)
        style.configure("TLabelframe", background=C_CARD, bordercolor="#E5E7EB",
                        relief="solid", borderwidth=1)
        style.configure("TLabelframe.Label", background=C_CARD, foreground=C_ACCENT, font=FONT_BOLD)

        # Кнопки
        style.configure("Accent.TButton", background=C_ACCENT, foreground="white",
                        borderwidth=0, focusthickness=0, padding=(14, 8), font=FONT_BOLD)
        style.map("Accent.TButton",
                  background=[("active", C_ACCENT_H), ("pressed", C_ACCENT_H)])
        style.configure("Ghost.TButton", background="#EEF2FF", foreground=C_ACCENT,
                        borderwidth=0, padding=(12, 8), font=FONT_MAIN)
        style.map("Ghost.TButton", background=[("active", "#E0E7FF")])
        style.configure("Danger.TButton", background="#FEE2E2", foreground=C_DANGER,
                        borderwidth=0, padding=(12, 8), font=FONT_MAIN)
        style.map("Danger.TButton", background=[("active", "#FECACA")])
        style.configure("Success.TButton", background="#DCFCE7", foreground="#166534",
                        borderwidth=0, padding=(12, 8), font=FONT_MAIN)
        style.map("Success.TButton", background=[("active", "#BBF7D0")])

        # Таблица
        style.configure("Treeview", background=C_CARD, fieldbackground=C_CARD,
                        foreground=C_TEXT, rowheight=28, borderwidth=0, font=FONT_MAIN)
        style.configure("Treeview.Heading", background="#EEF2FF", foreground=C_TEXT,
                        relief="flat", font=FONT_BOLD, padding=(8, 6))
        style.map("Treeview.Heading", background=[("active", "#E0E7FF")])
        style.map("Treeview", background=[("selected", "#DBEAFE")],
                  foreground=[("selected", C_TEXT)])

    def _build_ui(self):
        # Верхняя панель с заголовком
        header = tk.Frame(self, bg=C_BG)
        header.pack(fill="x", padx=16, pady=(12, 4))
        tk.Label(header, text="📥  Входящая документация", bg=C_BG, fg=C_TEXT,
                 font=FONT_TITLE).pack(side="left")
        tk.Label(header, text=f"Папка сканов: {SCANS_DIR}", bg=C_BG, fg=C_MUTED,
                 font=FONT_SMALL).pack(side="right")

        # Форма
        form = ttk.LabelFrame(self, text="  Новый документ / редактирование  ")
        form.pack(fill="x", padx=16, pady=6)
        inner = tk.Frame(form, bg=C_CARD)
        inner.pack(fill="x", padx=12, pady=10)

        self.vars = {}
        fields = [
            ("Входящий №", "in_number", 0, 0),
            ("Дата (ГГГГ-ММ-ДД)", "in_date", 0, 2),
            ("От кого (организация)", "sender", 0, 4),
            ("Для чего / кратко", "subject", 1, 0),
            ("Дело (папка/номенклатура)", "case_name", 1, 2),
            ("Примечание", "note", 1, 4),
        ]
        for label, key, r, c in fields:
            tk.Label(inner, text=label, bg=C_CARD, fg=C_MUTED, font=FONT_SMALL).grid(
                row=r, column=c, sticky="w", padx=(6, 4), pady=4)
            var = tk.StringVar()
            self.vars[key] = var
            e = ttk.Entry(inner, textvariable=var, width=28)
            e.grid(row=r, column=c+1, sticky="ew", padx=(0, 12), pady=4)
        for i in range(6):
            inner.grid_columnconfigure(i, weight=1 if i % 2 else 0)

        # Кнопки управления записью
        row_btns = tk.Frame(inner, bg=C_CARD)
        row_btns.grid(row=2, column=0, columnspan=6, sticky="w", pady=(8, 0))
        ttk.Button(row_btns, text="➕ Добавить", style="Accent.TButton",
                   command=self.add_record).pack(side="left", padx=4)
        ttk.Button(row_btns, text="💾 Сохранить изменения", style="Ghost.TButton",
                   command=self.update_record).pack(side="left", padx=4)
        ttk.Button(row_btns, text="🧹 Очистить форму", style="Ghost.TButton",
                   command=self.clear_form).pack(side="left", padx=4)
        ttk.Button(row_btns, text="🗑 Удалить запись", style="Danger.TButton",
                   command=self.delete_record).pack(side="left", padx=4)

        # Сканирование
        scan_frame = ttk.LabelFrame(self, text="  Сканирование  ")
        scan_frame.pack(fill="x", padx=16, pady=6)
        s_inner = tk.Frame(scan_frame, bg=C_CARD)
        s_inner.pack(fill="x", padx=12, pady=10)

        ttk.Button(s_inner, text="📠 Сканировать", style="Accent.TButton",
                   command=self.scan_current).pack(side="left", padx=(0, 12))

        tk.Label(s_inner, text="Размер листа:", bg=C_CARD, fg=C_MUTED,
                 font=FONT_SMALL).pack(side="left", padx=(0, 4))
        self.paper_var = tk.StringVar(value="A4 (210×297)")
        ttk.Combobox(s_inner, textvariable=self.paper_var, values=list(PAPER_SIZES.keys()),
                     state="readonly", width=18).pack(side="left", padx=(0, 12))

        tk.Label(s_inner, text="DPI:", bg=C_CARD, fg=C_MUTED,
                 font=FONT_SMALL).pack(side="left", padx=(0, 4))
        self.dpi_var = tk.StringVar(value="300")
        ttk.Combobox(s_inner, textvariable=self.dpi_var, values=["150", "200", "300", "400", "600"],
                     state="readonly", width=6).pack(side="left", padx=(0, 12))

        self.bw_var = tk.BooleanVar(value=False)
        ttk.Checkbutton(s_inner, text="Ч/Б (экономия места)",
                        variable=self.bw_var).pack(side="left", padx=(0, 12))

        ttk.Button(s_inner, text="👁 Просмотр", style="Ghost.TButton",
                   command=self.open_scan).pack(side="left", padx=4)
        ttk.Button(s_inner, text="🗑 Удалить скан", style="Danger.TButton",
                   command=self.delete_scan).pack(side="left", padx=4)
        ttk.Button(s_inner, text="📁 Папка сканов", style="Ghost.TButton",
                   command=self.open_scans_folder).pack(side="left", padx=4)

        # Поиск и сортировка
        tools = tk.Frame(self, bg=C_BG)
        tools.pack(fill="x", padx=16, pady=(6, 2))
        tk.Label(tools, text="🔍 Поиск:", bg=C_BG, fg=C_MUTED, font=FONT_MAIN).pack(side="left")
        self.search_var = tk.StringVar()
        ttk.Entry(tools, textvariable=self.search_var, width=36).pack(side="left", padx=6)
        ttk.Button(tools, text="Найти", style="Ghost.TButton",
                   command=self.load_data).pack(side="left", padx=2)
        ttk.Button(tools, text="Сброс", style="Ghost.TButton",
                   command=self.reset_search).pack(side="left", padx=2)

        tk.Label(tools, text="Сортировать по:", bg=C_BG, fg=C_MUTED,
                 font=FONT_MAIN).pack(side="left", padx=(24, 4))
        self.sort_field_var = tk.StringVar(value="Дата")
        sort_combo = ttk.Combobox(tools, textvariable=self.sort_field_var,
            values=["Дата", "Вх. №", "От кого", "Дело"], state="readonly", width=14)
        sort_combo.pack(side="left", padx=4)
        sort_combo.bind("<<ComboboxSelected>>", lambda e: self._apply_sort_from_combo())

        self.sort_dir_var = tk.StringVar(value="▼ по убыванию")
        dir_combo = ttk.Combobox(tools, textvariable=self.sort_dir_var,
            values=["▲ по возрастанию", "▼ по убыванию"], state="readonly", width=18)
        dir_combo.pack(side="left", padx=4)
        dir_combo.bind("<<ComboboxSelected>>", lambda e: self._apply_sort_from_combo())

        # Таблица
        tree_frame = tk.Frame(self, bg=C_BG)
        tree_frame.pack(fill="both", expand=True, padx=16, pady=(4, 6))
        cols = ("id", "in_number", "in_date", "sender", "subject", "case_name", "note", "scan")
        headers = ("ID", "Вх. № ▲▼", "Дата ▲▼", "От кого ▲▼", "Для чего", "Дело ▲▼", "Примечание", "Скан")
        widths = (50, 100, 110, 220, 260, 180, 160, 60)
        self.tree = ttk.Treeview(tree_frame, columns=cols, show="headings")
        for c, h, w in zip(cols, headers, widths):
            self.tree.heading(c, text=h, command=lambda cc=c: self.sort_by_click(cc))
            self.tree.column(c, width=w, anchor="w", stretch=(c in ("subject", "sender")))
        vsb = ttk.Scrollbar(tree_frame, orient="vertical", command=self.tree.yview)
        hsb = ttk.Scrollbar(tree_frame, orient="horizontal", command=self.tree.xview)
        self.tree.configure(yscrollcommand=vsb.set, xscrollcommand=hsb.set)
        self.tree.grid(row=0, column=0, sticky="nsew")
        vsb.grid(row=0, column=1, sticky="ns")
        hsb.grid(row=1, column=0, sticky="ew")
        tree_frame.grid_rowconfigure(0, weight=1)
        tree_frame.grid_columnconfigure(0, weight=1)
        self.tree.tag_configure("odd", background=C_ROW_ALT)
        self.tree.bind("<<TreeviewSelect>>", self.on_select)
        self.tree.bind("<Double-1>", lambda e: self.open_scan())

        # Нижняя панель печати
        pbar = tk.Frame(self, bg=C_BG)
        pbar.pack(fill="x", padx=16, pady=(2, 12))
        ttk.Button(pbar, text="🖨 Печать выбранного", style="Success.TButton",
                   command=self.print_selected).pack(side="left", padx=4)
        ttk.Button(pbar, text="🖨 Печать журнала", style="Success.TButton",
                   command=self.print_all).pack(side="left", padx=4)
        tk.Label(pbar, text="Двойной клик по строке — просмотр скана",
                 bg=C_BG, fg=C_MUTED, font=FONT_SMALL).pack(side="right")

    # ---------- Сортировка ----------
    def sort_by_click(self, col):
        key = {"in_number": "in_number", "in_date": "in_date",
               "sender": "sender", "case_name": "case_name"}.get(col)
        if not key:
            return
        if self.sort_col == key:
            self.sort_desc = not self.sort_desc
        else:
            self.sort_col = key
            self.sort_desc = False
        self.load_data()

    def _apply_sort_from_combo(self):
        m = {"Дата": "in_date", "Вх. №": "in_number", "От кого": "sender", "Дело": "case_name"}
        self.sort_col = m.get(self.sort_field_var.get(), "in_date")
        self.sort_desc = self.sort_dir_var.get().startswith("▼")
        self.load_data()

    # ---------- Данные ----------
    def load_data(self, *args):
        q = self.search_var.get().strip() if hasattr(self, "search_var") else ""
        for i in self.tree.get_children():
            self.tree.delete(i)
        order = f"{self.sort_col} {'DESC' if self.sort_desc else 'ASC'}"
        if q:
            rows = self.conn.execute(
                f"SELECT * FROM documents WHERE in_number LIKE ? OR sender LIKE ? "
                f"OR subject LIKE ? OR case_name LIKE ? ORDER BY {order}, id DESC",
                (f"%{q}%",) * 4,
            ).fetchall()
        else:
            rows = self.conn.execute(f"SELECT * FROM documents ORDER BY {order}, id DESC").fetchall()
        for i, r in enumerate(rows):
            scan_ok = "✔" if r[7] and any(os.path.exists(p) for p in r[7].split(";") if p) else ""
            tag = "odd" if i % 2 else ""
            self.tree.insert("", "end", values=(*r[:7], scan_ok), tags=(tag,))

    def reset_search(self):
        self.search_var.set("")
        self.load_data()

    def add_record(self):
        v = {k: s.get().strip() for k, s in self.vars.items()}
        if not v["in_number"] or not v["in_date"]:
            messagebox.showwarning("Внимание", "Заполните входящий номер и дату."); return
        cur = self.conn.execute(
            "INSERT INTO documents (in_number, in_date, sender, subject, case_name, note) VALUES (?,?,?,?,?,?)",
            (v["in_number"], v["in_date"], v["sender"], v["subject"], v["case_name"], v["note"]),
        )
        self.conn.commit()
        self.selected_id = cur.lastrowid
        self.load_data()
        self.clear_form()
        messagebox.showinfo("Готово", "Запись добавлена. Нажмите «📠 Сканировать».")

    def update_record(self):
        if self.selected_id is None:
            messagebox.showinfo("Инфо", "Сначала выберите запись в таблице."); return
        v = {k: s.get().strip() for k, s in self.vars.items()}
        self.conn.execute(
            "UPDATE documents SET in_number=?, in_date=?, sender=?, subject=?, case_name=?, note=? WHERE id=?",
            (v["in_number"], v["in_date"], v["sender"], v["subject"], v["case_name"], v["note"], self.selected_id),
        )
        self.conn.commit()
        self.load_data()

    def delete_record(self):
        if self.selected_id is None:
            messagebox.showinfo("Инфо", "Выберите запись для удаления."); return
        if messagebox.askyesno("Подтверждение", "Удалить запись? Файлы сканов останутся в папке scans."):
            self.conn.execute("DELETE FROM documents WHERE id=?", (self.selected_id,))
            self.conn.commit()
            self.clear_form()
            self.load_data()

    def on_select(self, event):
        sel = self.tree.selection()
        if not sel: return
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
            messagebox.showinfo("Инфо", "Сначала выберите запись (или добавьте её)."); return

        row = self.conn.execute(
            "SELECT in_number, in_date, sender FROM documents WHERE id=?",
            (self.selected_id,)).fetchone()
        safe_num  = _clean_filename(row[0] or "БезНомера")
        safe_date = _clean_filename(row[1] or "безДаты")
        safe_org  = _clean_filename(row[2] or "БезОрганизации")

        # Папка по организации
        org_dir = os.path.join(SCANS_DIR, safe_org)
        os.makedirs(org_dir, exist_ok=True)
        prefix = os.path.join(org_dir, f"{safe_num}_{safe_date}")

        # чистим старые
        for old in glob.glob(prefix + "_*.jpg"):
            try: os.remove(old)
            except: pass

        bw = self.bw_var.get()
        dpi = int(self.dpi_var.get())
        paper = self.paper_var.get()

        messagebox.showinfo(
            "Сканирование",
            f"Организация: {safe_org}\n"
            f"Размер листа: {paper}\n"
            f"DPI: {dpi}   Режим: {'Ч/Б' if bw else 'Цветной'}\n"
            f"Имя файла: {safe_num}_{safe_date}_N.jpg\n\n"
            f"Откроется окно сканера — отсканируйте и подтвердите."
        )
        files = do_scan_jpg(prefix, bw, dpi, paper)
        if files:
            self.conn.execute("UPDATE documents SET scan_path=? WHERE id=?",
                              (";".join(files), self.selected_id))
            self.conn.commit()
            self.load_data()
            if messagebox.askyesno("Готово", f"Сохранено страниц: {len(files)}.\n\nОткрыть прос
