# -*- coding: utf-8 -*-
import os
import sys
import csv
import json
import glob
import zipfile
import sqlite3
import calendar
import shutil
import tkinter as tk
from tkinter import ttk, messagebox, filedialog
from datetime import datetime, date
from PIL import Image, ImageTk

# ---------- Пути ----------
if getattr(sys, "frozen", False):
    BASE_DIR = os.path.dirname(sys.executable)
else:
    BASE_DIR = os.path.dirname(os.path.abspath(__file__))

DB_PATH = os.path.join(BASE_DIR, "incoming.db")
SCANS_DIR = os.path.join(BASE_DIR, "scans")
SETTINGS_PATH = os.path.join(BASE_DIR, "settings.json")
os.makedirs(SCANS_DIR, exist_ok=True)

# ---------- OCR ----------
def find_tesseract():
    p = shutil.which("tesseract")
    if p: return p
    for cand in [r"C:\Program Files\Tesseract-OCR\tesseract.exe",
                 r"C:\Program Files (x86)\Tesseract-OCR\tesseract.exe"]:
        if os.path.exists(cand): return cand
    return None
TESSERACT_PATH = find_tesseract()

def ocr_image(img_path, lang="rus+eng"):
    if not TESSERACT_PATH: return ""
    try:
        import pytesseract
        pytesseract.pytesseract.tesseract_cmd = TESSERACT_PATH
        return pytesseract.image_to_string(Image.open(img_path), lang=lang).strip()
    except Exception as e:
        print("OCR error:", e)
        return ""

# ---------- Настройки ----------
DEFAULT_SETTINGS = {"theme":"light","paper":"A4 (210×297)","dpi":"300",
                    "bw":False,"win_geom":"1280x760","last_dir":"",
                    "autobackup_days":7,"scan_format":"JPG"}
def load_settings():
    if os.path.exists(SETTINGS_PATH):
        try:
            with open(SETTINGS_PATH,"r",encoding="utf-8") as f:
                d = json.load(f)
            for k,v in DEFAULT_SETTINGS.items(): d.setdefault(k,v)
            return d
        except: pass
    return dict(DEFAULT_SETTINGS)
def save_settings(d):
    try:
        with open(SETTINGS_PATH,"w",encoding="utf-8") as f:
            json.dump(d, f, ensure_ascii=False, indent=2)
    except: pass
SETTINGS = load_settings()

# ---------- Темы ----------
THEMES = {
    "light": {"bg":"#F4F6FB","card":"#FFFFFF","accent":"#4A6CF7","accent_h":"#3A56C8",
              "text":"#1F2937","muted":"#6B7280","row_alt":"#F9FAFF","danger":"#E5484D",
              "heading":"#EEF2FF","sel":"#DBEAFE","border":"#E5E7EB"},
    "dark":  {"bg":"#1E1E2A","card":"#2A2A3A","accent":"#6C8EFF","accent_h":"#5575E6",
              "text":"#E5E7EB","muted":"#9CA3AF","row_alt":"#2F2F42","danger":"#F87171",
              "heading":"#33334D","sel":"#3B4A7A","border":"#3A3A52"},
}
C = THEMES.get(SETTINGS.get("theme","light"), THEMES["light"])
FONT_MAIN=("Segoe UI",10); FONT_BOLD=("Segoe UI",10,"bold")
FONT_TITLE=("Segoe UI",16,"bold"); FONT_SMALL=("Segoe UI",9)

def _clean_filename(s):
    for ch in '\\/:*?"<>|': s = s.replace(ch,"-")
    return s.strip().strip(".") or "БезНазвания"

# ---------- БД ----------
def get_conn():
    conn = sqlite3.connect(DB_PATH, timeout=10)
    conn.execute("PRAGMA journal_mode=WAL")
    conn.execute("""CREATE TABLE IF NOT EXISTS documents (
        id INTEGER PRIMARY KEY AUTOINCREMENT,
        in_number TEXT, in_date TEXT, sender TEXT, subject TEXT,
        case_name TEXT, note TEXT,
        status TEXT DEFAULT 'Новый', executor TEXT, remind_date TEXT,
        created_at TEXT, scan_path TEXT, deleted INTEGER DEFAULT 0)""")
    cols = [r[1] for r in conn.execute("PRAGMA table_info(documents)").fetchall()]
    for col, defv in [("scan_path","TEXT"),("status","TEXT DEFAULT 'Новый'"),
                      ("executor","TEXT"),("remind_date","TEXT"),
                      ("created_at","TEXT"),("deleted","INTEGER DEFAULT 0"),
                      ("ocr_text","TEXT")]:
        if col not in cols:
            conn.execute(f"ALTER TABLE documents ADD COLUMN {col} {defv}")
    conn.commit()
    return conn

# ---------- Календарь ----------
RU_MONTHS = ["Январь","Февраль","Март","Апрель","Май","Июнь",
             "Июль","Август","Сентябрь","Октябрь","Ноябрь","Декабрь"]
RU_DAYS = ["Пн","Вт","Ср","Чт","Пт","Сб","Вс"]

class DatePicker(tk.Toplevel):
    def __init__(self, master, var):
        super().__init__(master)
        self.var = var; self.title("Выбор даты"); self.resizable(False, False)
        self.configure(bg=C["card"]); self.transient(master)
        try: cur = datetime.strptime(var.get(),"%Y-%m-%d").date()
        except: cur = date.today()
        self.year, self.month = cur.year, cur.month
        self.day_selected = cur
        self.top = tk.Frame(self, bg=C["card"]); self.top.pack(pady=8, padx=10, fill="x")
        tk.Button(self.top, text="◀", command=self.prev_m, bg=C["accent"], fg="white",
                  relief="flat", width=3, cursor="hand2").pack(side="left")
        self.lbl = tk.Label(self.top, text="", bg=C["card"], fg=C["text"], font=FONT_BOLD)
        self.lbl.pack(side="left", expand=True, fill="x")
        tk.Button(self.top, text="▶", command=self.next_m, bg=C["accent"], fg="white",
                  relief="flat", width=3, cursor="hand2").pack(side="right")
        self.grid_frame = tk.Frame(self, bg=C["card"]); self.grid_frame.pack(padx=10, pady=4)
        bot = tk.Frame(self, bg=C["card"]); bot.pack(pady=8, padx=10, fill="x")
        tk.Button(bot, text="Сегодня", command=self.today, bg=C["accent"], fg="white",
                  relief="flat", padx=10, cursor="hand2").pack(side="left")
        tk.Button(bot, text="Очистить", command=self.clear, bg=C["muted"], fg="white",
                  relief="flat", padx=10, cursor="hand2").pack(side="left", padx=6)
        tk.Button(bot, text="Закрыть", command=self.destroy, bg=C["danger"], fg="white",
                  relief="flat", padx=10, cursor="hand2").pack(side="right")
        self.render()
        self.after(50, lambda: self._focus())
    def _focus(self):
        try: self.grab_set(); self.focus_force()
        except: pass
    def render(self):
        for w in self.grid_frame.winfo_children(): w.destroy()
        self.lbl.config(text=f"{RU_MONTHS[self.month-1]} {self.year}")
        for i,d in enumerate(RU_DAYS):
            tk.Label(self.grid_frame, text=d, bg=C["card"], fg=C["muted"],
                     font=FONT_SMALL, width=4).grid(row=0, column=i, padx=2, pady=3)
        for r, week in enumerate(calendar.monthcalendar(self.year, self.month), 1):
            for c, day in enumerate(week):
                if day == 0: continue
                is_today = (date.today() == date(self.year, self.month, day))
                is_sel = (self.day_selected == date(self.year, self.month, day))
                bg = C["accent"] if (is_today or is_sel) else C["card"]
                fg = "white" if (is_today or is_sel) else C["text"]
                b = tk.Button(self.grid_frame, text=str(day), width=4, relief="flat",
                              bg=bg, fg=fg, cursor="hand2",
                              command=lambda d=day: self.pick(d))
                b.grid(row=r, column=c, padx=1, pady=1)
    def prev_m(self):
        self.month -= 1
        if self.month < 1: self.month = 12; self.year -= 1
        self.render()
    def next_m(self):
        self.month += 1
        if self.month > 12: self.month = 1; self.year += 1
        self.render()
    def pick(self, d):
        self.var.set(date(self.year, self.month, d).strftime("%Y-%m-%d")); self.destroy()
    def today(self): self.var.set(date.today().strftime("%Y-%m-%d")); self.destroy()
    def clear(self): self.var.set(""); self.destroy()

# ---------- Сканирование ----------
def save_jpg(img, path, bw):
    img = img.convert("L") if bw else img.convert("RGB")
    img.save(path, "JPEG", quality=85)
PAPER_SIZES = {"A4 (210×297)":(8.27,11.69),"A5 (148×210)":(5.83,8.27),
               "Letter (8.5×11)":(8.5,11.0),"Legal (8.5×14)":(8.5,14.0),
               "Авто (по краю)":None}

def scan_twain_jpg(prefix, bw, dpi, paper):
    files = []
    try:
        import twain
        sm = twain.SourceManager(0); sources = sm.GetSourceList()
        if not sources: sm.destroy(); return files
        epson = next((s for s in sources if "epson" in s.lower()), sources[0])
        src = sm.OpenSource(epson)
        if not src: sm.destroy(); return files
        src.SetCapability(twain.ICAP_XRESOLUTION, twain.TWTY_FIX32, float(dpi))
        src.SetCapability(twain.ICAP_YRESOLUTION, twain.TWTY_FIX32, float(dpi))
        src.SetCapability(twain.ICAP_PIXELTYPE, twain.TWTY_UINT16, 0 if bw else 2)
        size = PAPER_SIZES.get(paper)
        if size:
            try: src.SetCapability(twain.ICAP_SUPPORTEDSIZES, twain.TWTY_UINT16, 1)
            except Exception as e: print("size1:", e)
            try:
                w_in, h_in = size
                src.SetCapability(twain.ICAP_FRAMES, twain.TWTY_FRAME,
                                  (0,0,int(w_in*65536),int(h_in*65536)))
            except Exception as e: print("size2:", e)
        src.RequestAcquire(0, 0); page = 1
        while True:
            rv = src.XferImageNatively()
            if rv[0] != twain.TWRC_SUCCESS: break
            tmp = os.path.join(SCANS_DIR, f"tmp_{datetime.now().timestamp()}_{page}.bmp")
            twain.DIBToBMFile(rv[1], tmp)
            jpg = f"{prefix}_{page}.jpg"
            try: save_jpg(Image.open(tmp), jpg, bw); files.append(jpg)
            finally:
                try: os.remove(tmp)
                except: pass
            page += 1
        src.CloseSource(); sm.destroy()
    except Exception as e: print("TWAIN error:", e)
    return files

def scan_wia_jpg(prefix, bw):
    files = []
    try:
        import win32com.client
        WIA = win32com.client.Dispatch("WIA.CommonDialog")
        img = WIA.ShowAcquireImage()
        if not img: return files
        tmp = os.path.join(SCANS_DIR, f"tmp_wia_{datetime.now().timestamp()}.bmp")
        img.SaveFile(tmp); jpg = f"{prefix}_1.jpg"
        save_jpg(Image.open(tmp), jpg, bw); files.append(jpg)
        try: os.remove(tmp)
        except: pass
    except Exception as e: print("WIA error:", e)
    return files

def do_scan(prefix, bw, dpi, paper, fmt="JPG"):
    files = scan_twain_jpg(prefix, bw, dpi, paper)
    if not files: files = scan_wia_jpg(prefix, bw)
    if not files: return []
    if fmt in ("PDF","PDF + JPG"):
        pdf_path = f"{prefix}.pdf"
        try:
            imgs = [Image.open(f).convert("RGB") for f in files]
            imgs[0].save(pdf_path, save_all=True, append_images=imgs[1:], resolution=dpi)
        except Exception as e:
            print("PDF error:", e); pdf_path = None
        if fmt == "PDF":
            for f in files:
                try: os.remove(f)
                except: pass
            return [pdf_path] if pdf_path else []
        else:
            return files + ([pdf_path] if pdf_path else [])
    return files

# ---------- Просмотр ----------
class ScanViewer(tk.Toplevel):
    def __init__(self, master, files):
        super().__init__(master)
        self.title("Просмотр скана"); self.geometry("960x740"); self.configure(bg=C["bg"])
        self._photo = None; self.idx = 0; self.zoom = 1.0; self.rotation = 0
        # Поддержка PDF
        self.files = []
        for f in files:
            if f.lower().endswith(".pdf"):
                try:
                    from pdf2image import convert_from_path
                    pages = convert_from_path(f, dpi=150)
                    base = os.path.splitext(f)[0]
                    for i, pg in enumerate(pages, 1):
                        jpg = f"{base}_pdfpage_{i}.jpg"
                        pg.save(jpg, "JPEG", quality=85)
                        self.files.append(jpg)
                except Exception as e:
                    print("PDF preview error:", e)
                    messagebox.showerror("PDF","Для просмотра PDF установите Poppler:\n"
                        "https://github.com/oschwartz10612/poppler-windows/releases\n\n"
                        f"Ошибка: {e}")
            else:
                self.files.append(f)
        self.files = [f for f in self.files if os.path.exists(f)]
        bar = tk.Frame(self, bg=C["card"]); bar.pack(fill="x", padx=8, pady=6)
        def btn(txt, cmd, bg=None):
            return tk.Button(bar, text=txt, command=cmd, bg=bg or C["accent"], fg="white",
                             activebackground=C["accent_h"], activeforeground="white",
                             relief="flat", bd=0, padx=10, pady=4, font=FONT_SMALL,
                             cursor="hand2")
        btn("◀", self.prev).pack(side="left", padx=2)
        btn("▶", self.next).pack(side="left", padx=2)
        btn("↺ 90°", lambda: self.rotate(-90)).pack(side="left", padx=6)
        btn("↻ 90°", lambda: self.rotate(90)).pack(side="left", padx=2)
        btn("💾 Поворот", self.save_rotation).pack(side="left", padx=6)
        btn("−", lambda: self.change_zoom(0.9)).pack(side="left", padx=2)
        self.lbl_zoom = tk.Label(bar, text="100%", bg=C["card"], fg=C["text"], font=FONT_SMALL)
        self.lbl_zoom.pack(side="left", padx=2)
        btn("+", lambda: self.change_zoom(1.1)).pack(side="left", padx=2)
        btn("Сброс", self.reset_zoom, C["muted"]).pack(side="left", padx=6)
        btn("🖨 Печать", self.print_scan, "#16A34A").pack(side="left", padx=6)
        btn("Закрыть", self.destroy, C["danger"]).pack(side="right", padx=2)
        self.lbl_page = tk.Label(bar, text="", bg=C["card"], fg=C["muted"], font=FONT_SMALL)
        self.lbl_page.pack(side="right", padx=8)
        self.canvas = tk.Canvas(self, bg="#1F2937", highlightthickness=0)
        self.canvas.pack(fill="both", expand=True)
        self.canvas.bind("<MouseWheel>", self.on_wheel)
        self.canvas.bind("<Configure>", lambda e: self.render())
        if not self.files:
            messagebox.showinfo("Просмотр","Файлы не найдены."); self.destroy(); return
        self.render()
    def change_zoom(self, k):
        self.zoom = max(0.1, min(4.0, self.zoom*k))
        self.lbl_zoom.config(text=f"{int(self.zoom*100)}%"); self.render()
    def reset_zoom(self): self.zoom = 1.0; self.lbl_zoom.config(text="100%"); self.render()
    def prev(self):
        if self.idx > 0: self.idx -= 1; self.rotation = 0; self.render()
    def next(self):
        if self.idx < len(self.files)-1: self.idx += 1; self.rotation = 0; self.render()
    def on_wheel(self, e): self.change_zoom(1.1 if e.delta > 0 else 0.9)
    def rotate(self, deg): self.rotation = (self.rotation+deg)%360; self.render()
    def save_rotation(self):
        if not self.files: return
        f = self.files[self.idx]
        if self.rotation == 0: messagebox.showinfo("Поворот","Поворот не задан."); return
        try:
            Image.open(f).rotate(-self.rotation, expand=True).save(f,"JPEG",quality=85)
            self.rotation = 0; self.render(); messagebox.showinfo("Готово","Поворот сохранён.")
        except Exception as e: messagebox.showerror("Ошибка", f"Не удалось: {e}")
    def print_scan(self):
        if not self.files: return
        try:
            from docx import Document; from docx.shared import Cm
        except ImportError:
            messagebox.showerror("Нужен python-docx","pip install python-docx"); return
        doc = Document()
        for s in doc.sections:
            s.top_margin = s.bottom_margin = s.left_margin = s.right_margin = Cm(1)
        for f in self.files:
            try:
                img = Image.open(f)
                if f == self.files[self.idx] and self.rotation:
                    img = img.rotate(-self.rotation, expand=True)
                tmp = os.path.join(SCANS_DIR, f"_print_{datetime.now().timestamp()}.jpg")
                img.convert("RGB").save(tmp,"JPEG",quality=90)
                doc.add_picture(tmp, width=Cm(17))
                try: os.remove(tmp)
                except: pass
                doc.add_page_break()
            except Exception as e: print("print:", e)
        out = filedialog.asksaveasfilename(defaultextension=".docx",
            filetypes=[("Word","*.docx")], initialfile="Скан_для_печати.docx")
        if not out: return
        doc.save(out); messagebox.showinfo("Готово", f"Сохранено:\n{out}\n\nCtrl+P.")
        try: os.startfile(out)
        except: pass
    def render(self):
        self.canvas.delete("all")
        if not self.files: return
        f = self.files[self.idx]
        self.lbl_page.config(text=f"Стр. {self.idx+1} из {len(self.files)} — {os.path.basename(f)}")
        try: img = Image.open(f)
        except Exception as e:
            self.canvas.create_text(20,20,anchor="nw",fill="white",text=f"Ошибка: {e}"); return
        if self.rotation: img = img.rotate(-self.rotation, expand=True)
        cw = max(self.canvas.winfo_width(), 200); ch = max(self.canvas.winfo_height(), 200)
        w, h = img.size; fit = min(cw/w, ch/h) if w and h else 1
        scale = fit*self.zoom; nw, nh = max(1,int(w*scale)), max(1,int(h*scale))
        img = img.resize((nw, nh), Image.LANCZOS)
        self._photo = ImageTk.PhotoImage(img)
        self.canvas.create_image(cw//2, ch//2, image=self._photo, anchor="center")# ---------- Главное окно ----------
class App(tk.Tk):
    def __init__(self):
        super().__init__()
        self.title("Регистрация входящей документации")
        self.geometry(SETTINGS.get("win_geom", "1280x760"))
        self.minsize(1150, 720)
        self.configure(bg=C["bg"])
        self._apply_style()
        self.conn = get_conn()
        self.selected_id = None
        self.sort_col = "in_date"
        self.sort_desc = True
        self.filter_noscan = tk.BooleanVar(value=False)
        self._autocomplete_widgets = {}
        self._build_ui()
        self._bind_hotkeys()
        self.load_data()
        self._check_reminders()
        self._autobackup()
        if not TESSERACT_PATH:
            self.after(800, lambda: messagebox.showwarning(
                "OCR не настроен",
                "Tesseract OCR не найден.\nПоиск по содержимому скана будет недоступен.\n"
                "Установите Tesseract: https://github.com/UB-Mannheim/tesseract/wiki"))
        self.protocol("WM_DELETE_WINDOW", self._on_close)

    def _apply_style(self):
        style = ttk.Style(self)
        try: style.theme_use("clam")
        except: pass
        style.configure(".", font=FONT_MAIN, background=C["bg"], foreground=C["text"])
        style.configure("TFrame", background=C["bg"])
        style.configure("TLabel", background=C["bg"], foreground=C["text"], font=FONT_MAIN)
        style.configure("Muted.TLabel", background=C["bg"], foreground=C["muted"], font=FONT_SMALL)
        style.configure("TEntry", fieldbackground=C["card"], foreground=C["text"],
                        bordercolor=C["border"], lightcolor=C["border"], darkcolor=C["border"], padding=6)
        style.configure("TCombobox", fieldbackground=C["card"], foreground=C["text"], padding=5)
        style.configure("TCheckbutton", background=C["card"], foreground=C["text"], font=FONT_MAIN)
        style.configure("TLabelframe", background=C["card"], bordercolor=C["border"],
                        relief="solid", borderwidth=1)
        style.configure("TLabelframe.Label", background=C["card"], foreground=C["accent"], font=FONT_BOLD)
        style.configure("Accent.TButton", background=C["accent"], foreground="white",
                        borderwidth=0, padding=(14,8), font=FONT_BOLD)
        style.map("Accent.TButton", background=[("active", C["accent_h"]), ("pressed", C["accent_h"])])
        style.configure("Ghost.TButton", background=C["heading"], foreground=C["accent"],
                        borderwidth=0, padding=(12,8), font=FONT_MAIN)
        style.map("Ghost.TButton", background=[("active", C["heading"])])
        danger_bg = "#FEE2E2" if SETTINGS["theme"]=="light" else "#4A1D1D"
        style.configure("Danger.TButton", background=danger_bg, foreground=C["danger"],
                        borderwidth=0, padding=(12,8), font=FONT_MAIN)
        success_bg = "#DCFCE7" if SETTINGS["theme"]=="light" else "#1A3A24"
        success_fg = "#166534" if SETTINGS["theme"]=="light" else "#86EFAC"
        style.configure("Success.TButton", background=success_bg, foreground=success_fg,
                        borderwidth=0, padding=(12,8), font=FONT_MAIN)
        style.configure("Treeview", background=C["card"], fieldbackground=C["card"],
                        foreground=C["text"], rowheight=28, borderwidth=0, font=FONT_MAIN)
        style.configure("Treeview.Heading", background=C["heading"], foreground=C["text"],
                        relief="flat", font=FONT_BOLD, padding=(8,6))
        style.map("Treeview.Heading", background=[("active", C["heading"])])
        style.map("Treeview", background=[("selected", C["sel"])], foreground=[("selected", C["text"])])

    def _build_ui(self):
        header = tk.Frame(self, bg=C["bg"]); header.pack(fill="x", padx=16, pady=(12,4))
        tk.Label(header, text="📥  Входящая документация", bg=C["bg"], fg=C["text"],
                 font=FONT_TITLE).pack(side="left")
        ttk.Button(header, text="⚙ Настройки", style="Ghost.TButton",
                   command=self.open_settings).pack(side="right", padx=4)
        ttk.Button(header, text="🌓 Тема", style="Ghost.TButton",
                   command=self.toggle_theme).pack(side="right", padx=4)
        ttk.Button(header, text="💾 Бэкап", style="Ghost.TButton",
                   command=self.make_backup).pack(side="right", padx=4)
        ttk.Button(header, text="📤 Экспорт CSV", style="Ghost.TButton",
                   command=self.export_csv).pack(side="right", padx=4)
        ttk.Button(header, text="📥 Импорт CSV", style="Ghost.TButton",
                   command=self.import_csv).pack(side="right", padx=4)

        form = ttk.LabelFrame(self, text="  Новый документ / редактирование  ")
        form.pack(fill="x", padx=16, pady=6)
        inner = tk.Frame(form, bg=C["card"]); inner.pack(fill="x", padx=12, pady=10)
        self.vars = {}
        fields = [
            ("Входящий №", "in_number", 0, 0, 22),
            ("Дата", "in_date", 0, 2, 18),
            ("От кого (организация)", "sender", 0, 4, 26),
            ("Для чего / кратко", "subject", 1, 0, 50),
            ("Дело (папка/номенклатура)", "case_name", 1, 2, 26),
            ("Примечание", "note", 1, 4, 26),
            ("Исполнитель", "executor", 2, 0, 22),
            ("Напомнить до", "remind_date", 2, 2, 18),
        ]
        for label, key, r, c, w in fields:
            tk.Label(inner, text=label, bg=C["card"], fg=C["muted"], font=FONT_SMALL).grid(
                row=r, column=c, sticky="w", padx=(6,4), pady=4)
            var = tk.StringVar(); self.vars[key] = var
            if key in ("sender", "case_name"):
                cb = ttk.Combobox(inner, textvariable=var, width=w)
                cb.grid(row=r, column=c+1, sticky="ew", padx=(0,8), pady=4)
                cb.bind("<KeyRelease>", lambda e, k=key: self._update_autocomplete(k))
                self._autocomplete_widgets[key] = cb
            else:
                e = ttk.Entry(inner, textvariable=var, width=w)
                e.grid(row=r, column=c+1, sticky="ew", padx=(0,8), pady=4)
            if key in ("in_date", "remind_date"):
                tk.Button(inner, text="📅", command=lambda v=var: DatePicker(self, v),
                          bg=C["accent"], fg="white", relief="flat", width=3,
                          cursor="hand2").grid(row=r, column=c+2, sticky="w", padx=(0,6))

        tk.Label(inner, text="Статус", bg=C["card"], fg=C["muted"], font=FONT_SMALL).grid(
            row=2, column=4, sticky="w", padx=(6,4), pady=4)
        self.status_var = tk.StringVar(value="Новый")
        ttk.Combobox(inner, textvariable=self.status_var,
                     values=["Новый","В работе","Исполнен","В архиве"],
                     state="readonly", width=18).grid(row=2, column=5, sticky="ew", padx=(0,12))
        for i in range(7): inner.grid_columnconfigure(i, weight=1 if i%2 else 0)

        row_btns = tk.Frame(inner, bg=C["card"])
        row_btns.grid(row=3, column=0, columnspan=7, sticky="w", pady=(8,0))
        ttk.Button(row_btns, text="➕ Добавить (Ctrl+N)", style="Accent.TButton",
                   command=self.add_record).pack(side="left", padx=4)
        ttk.Button(row_btns, text="💾 Сохранить (Ctrl+S)", style="Ghost.TButton",
                   command=self.update_record).pack(side="left", padx=4)
        ttk.Button(row_btns, text="🧹 Очистить", style="Ghost.TButton",
                   command=self.clear_form).pack(side="left", padx=4)
        ttk.Button(row_btns, text="🗑 В корзину", style="Danger.TButton",
                   command=self.delete_record).pack(side="left", padx=4)

        scan_frame = ttk.LabelFrame(self, text="  Сканирование  ")
        scan_frame.pack(fill="x", padx=16, pady=6)
        s_inner = tk.Frame(scan_frame, bg=C["card"]); s_inner.pack(fill="x", padx=12, pady=10)
        self.scan_btn = ttk.Button(s_inner, text="📠 Сканировать", style="Accent.TButton",
                                   command=self.scan_current)
        self.scan_btn.pack(side="left", padx=(0,12))
        tk.Label(s_inner, text="Формат:", bg=C["card"], fg=C["muted"], font=FONT_SMALL).pack(side="left")
        self.format_var = tk.StringVar(value=SETTINGS.get("scan_format","JPG"))
        ttk.Combobox(s_inner, textvariable=self.format_var,
                     values=["JPG","PDF","PDF + JPG"], state="readonly", width=11).pack(side="left", padx=(0,10))
        tk.Label(s_inner, text="Размер:", bg=C["card"], fg=C["muted"], font=FONT_SMALL).pack(side="left")
        self.paper_var = tk.StringVar(value=SETTINGS.get("paper","A4 (210×297)"))
        ttk.Combobox(s_inner, textvariable=self.paper_var, values=list(PAPER_SIZES.keys()),
                     state="readonly", width=15).pack(side="left", padx=(0,10))
        tk.Label(s_inner, text="DPI:", bg=C["card"], fg=C["muted"], font=FONT_SMALL).pack(side="left")
        self.dpi_var = tk.StringVar(value=SETTINGS.get("dpi","300"))
        ttk.Combobox(s_inner, textvariable=self.dpi_var, values=["150","200","300","400","600"],
                     state="readonly", width=6).pack(side="left", padx=(0,10))
        self.bw_var = tk.BooleanVar(value=SETTINGS.get("bw", False))
        ttk.Checkbutton(s_inner, text="Ч/Б", variable=self.bw_var).pack(side="left", padx=(0,10))
        ttk.Button(s_inner, text="👁 Просмотр", style="Ghost.TButton",
                   command=self.open_scan).pack(side="left", padx=4)
        ttk.Button(s_inner, text="🗑 Удалить скан", style="Danger.TButton",
                   command=self.delete_scan).pack(side="left", padx=4)
        ttk.Button(s_inner, text="🔎 OCR", style="Ghost.TButton",
                   command=self.run_ocr_selected).pack(side="left", padx=4)
        ttk.Button(s_inner, text="📁 Папка", style="Ghost.TButton",
                   command=self.open_scans_folder).pack(side="left", padx=4)
        self.progress = ttk.Progressbar(s_inner, mode="indeterminate", length=120)
        self.progress.pack(side="left", padx=10)

        tools = tk.Frame(self, bg=C["bg"]); tools.pack(fill="x", padx=16, pady=(6,2))
        tk.Label(tools, text="🔍 Поиск:", bg=C["bg"], fg=C["muted"], font=FONT_MAIN).pack(side="left")
        self.search_var = tk.StringVar()
        self.search_entry = ttk.Entry(tools, textvariable=self.search_var, width=30)
        self.search_entry.pack(side="left", padx=6)
        self.search_entry.bind("<Return>", lambda e: self.load_data())
        ttk.Button(tools, text="Найти", style="Ghost.TButton", command=self.load_data).pack(side="left", padx=2)
        ttk.Button(tools, text="Сброс", style="Ghost.TButton", command=self.reset_search).pack(side="left", padx=2)
        ttk.Checkbutton(tools, text="Только без скана", variable=self.filter_noscan,
                        command=self.load_data).pack(side="left", padx=12)
        ttk.Button(tools, text="🗑 Корзина", style="Danger.TButton",
                   command=self.open_trash).pack(side="right", padx=2)
        tk.Label(tools, text="Сортировать по:", bg=C["bg"], fg=C["muted"], font=FONT_MAIN).pack(side="left", padx=(20,4))
        self.sort_field_var = tk.StringVar(value="Дата")
        sc = ttk.Combobox(tools, textvariable=self.sort_field_var,
            values=["Дата","Вх. №","От кого","Дело","Статус"], state="readonly", width=12)
        sc.pack(side="left", padx=2); sc.bind("<<ComboboxSelected>>", lambda e: self._apply_sort_from_combo())
        self.sort_dir_var = tk.StringVar(value="▼ убыв.")
        dc = ttk.Combobox(tools, textvariable=self.sort_dir_var,
            values=["▲ возр.","▼ убыв."], state="readonly", width=10)
        dc.pack(side="left", padx=2); dc.bind("<<ComboboxSelected>>", lambda e: self._apply_sort_from_combo())

        tree_frame = tk.Frame(self, bg=C["bg"])
        tree_frame.pack(fill="both", expand=True, padx=16, pady=(4,6))
        cols = ("id","in_number","in_date","sender","subject","case_name","status","executor","remind_date","scan")
        headers = ("ID","Вх. № ▲▼","Дата ▲▼","От кого ▲▼","Для чего","Дело ▲▼","Статус","Исполнитель","Напомнить","Скан")
        widths = (45,90,100,190,220,150,90,110,90,50)
        self.tree = ttk.Treeview(tree_frame, columns=cols, show="headings")
        for c,h,w in zip(cols,headers,widths):
            self.tree.heading(c, text=h, command=lambda cc=c: self.sort_by_click(cc))
            self.tree.column(c, width=w, anchor="w", stretch=(c in ("subject","sender")))
        vsb = ttk.Scrollbar(tree_frame, orient="vertical", command=self.tree.yview)
        hsb = ttk.Scrollbar(tree_frame, orient="horizontal", command=self.tree.xview)
        self.tree.configure(yscrollcommand=vsb.set, xscrollcommand=hsb.set)
        self.tree.grid(row=0,column=0,sticky="nsew"); vsb.grid(row=0,column=1,sticky="ns"); hsb.grid(row=1,column=0,sticky="ew")
        tree_frame.grid_rowconfigure(0, weight=1); tree_frame.grid_columnconfigure(0, weight=1)
        self.tree.tag_configure("odd", background=C["row_alt"])
        self.tree.tag_configure("new", foreground="#2563EB")
        self.tree.tag_configure("work", foreground="#CA8A04")
        self.tree.tag_configure("done", foreground="#16A34A")
        self.tree.tag_configure("arch", foreground=C["muted"])
        self.tree.bind("<<TreeviewSelect>>", self.on_select)
        self.tree.bind("<Double-1>", lambda e: self.open_scan())

        bottom = tk.Frame(self, bg=C["bg"]); bottom.pack(fill="x", padx=16, pady=(2,10))
        self.counter = tk.Label(bottom, text="", bg=C["bg"], fg=C["muted"], font=FONT_SMALL)
        self.counter.pack(side="left")
        ttk.Button(bottom, text="🖨 Печать выбранного", style="Success.TButton",
                   command=self.print_selected).pack(side="right", padx=4)
        ttk.Button(bottom, text="🖨 Печать журнала", style="Success.TButton",
                   command=self.print_all).pack(side="right", padx=4)

    def _bind_hotkeys(self):
        self.bind("<Control-n>", lambda e: self.add_record())
        self.bind("<Control-s>", lambda e: self.update_record())
        self.bind("<Control-f>", lambda e: self.search_entry.focus_set())
        self.bind("<Control-p>", lambda e: self.print_selected())
        self.bind("<F5>", lambda e: self.load_data())
        self.bind("<Delete>", lambda e: self.delete_record())

    def _get_distinct(self, column):
        rows = self.conn.execute(
            f"SELECT DISTINCT {column} FROM documents WHERE deleted=0 AND {column} IS NOT NULL AND {column}<>'' ORDER BY {column}"
        ).fetchall()
        return [r[0] for r in rows]

    def _update_autocomplete(self, key):
        cb = self._autocomplete_widgets.get(key)
        if not cb: return
        text = self.vars[key].get().strip().lower()
        values = self._get_distinct(key)
        filtered = [v for v in values if text in v.lower()] if text else values
        cb["values"] = filtered
        if filtered and text and filtered[0].lower() != text:
            cb.event_generate("<Down>")

    def toggle_theme(self):
        SETTINGS["theme"] = "dark" if SETTINGS.get("theme","light")=="light" else "light"
        save_settings(SETTINGS)
        messagebox.showinfo("Тема","Тема применена. Перезапустите программу.")    # ---------- Настройки ----------
    def open_settings(self):
        w = tk.Toplevel(self); w.title("Настройки"); w.geometry("520x320")
        w.configure(bg=C["bg"]); w.transient(self); w.grab_set()
        frm = tk.Frame(w, bg=C["card"]); frm.pack(fill="both", expand=True, padx=12, pady=12)
        tk.Label(frm, text="Папка сканов:", bg=C["card"], fg=C["muted"]).pack(anchor="w")
        scan_var = tk.StringVar(value=SCANS_DIR)
        ttk.Entry(frm, textvariable=scan_var, width=60).pack(fill="x", pady=2)
        tk.Label(frm, text="(изменение требует перезапуска)", bg=C["card"],
                 fg=C["muted"], font=FONT_SMALL).pack(anchor="w", pady=(0,8))
        tk.Label(frm, text="Авто-бэкап каждые N дней (0 — выкл.):", bg=C["card"], fg=C["muted"]).pack(anchor="w")
        ab_var = tk.StringVar(value=str(SETTINGS.get("autobackup_days",7)))
        ttk.Entry(frm, textvariable=ab_var, width=8).pack(anchor="w", pady=2)
        tk.Label(frm, text="Путь к Tesseract (для OCR):", bg=C["card"], fg=C["muted"]).pack(anchor="w", pady=(8,0))
        tess_var = tk.StringVar(value=TESSERACT_PATH or "")
        ttk.Entry(frm, textvariable=tess_var, width=60).pack(fill="x", pady=2)
        def save():
            try: SETTINGS["autobackup_days"] = int(ab_var.get())
            except: SETTINGS["autobackup_days"] = 7
            save_settings(SETTINGS); w.destroy()
            messagebox.showinfo("Готово","Настройки сохранены.")
        ttk.Button(frm, text="Сохранить", style="Accent.TButton",
                   command=save).pack(pady=12, side="left")
        ttk.Button(frm, text="Отмена", style="Ghost.TButton",
                   command=w.destroy).pack(pady=12, side="left", padx=6)

    # ---------- Сортировка ----------
    def sort_by_click(self, col):
        key = {"in_number":"in_number","in_date":"in_date","sender":"sender",
               "case_name":"case_name","status":"status"}.get(col)
        if not key: return
        if self.sort_col == key: self.sort_desc = not self.sort_desc
        else: self.sort_col = key; self.sort_desc = False
        self.load_data()

    def _apply_sort_from_combo(self):
        m = {"Дата":"in_date","Вх. №":"in_number","От кого":"sender","Дело":"case_name","Статус":"status"}
        self.sort_col = m.get(self.sort_field_var.get(),"in_date")
        self.sort_desc = self.sort_dir_var.get().startswith("▼")
        self.load_data()

    # ---------- Загрузка ----------
    def load_data(self, *args):
        q = self.search_var.get().strip() if hasattr(self, "search_var") else ""
        for i in self.tree.get_children(): self.tree.delete(i)
        order = f"{self.sort_col} {'DESC' if self.sort_desc else 'ASC'}"
        where = ["deleted=0"]; params = []
        if q:
            where.append("(in_number LIKE ? OR sender LIKE ? OR subject LIKE ? OR case_name LIKE ? OR executor LIKE ? OR ocr_text LIKE ?)")
            params += [f"%{q}%"]*6
        if self.filter_noscan.get():
            where.append("(scan_path IS NULL OR scan_path='')")
        sql = f"SELECT * FROM documents WHERE {' AND '.join(where)} ORDER BY {order}, id DESC"
        rows = self.conn.execute(sql, params).fetchall()
        total = self.conn.execute("SELECT COUNT(*) FROM documents WHERE deleted=0").fetchone()[0]
        withscan = self.conn.execute("SELECT COUNT(*) FROM documents WHERE deleted=0 AND scan_path IS NOT NULL AND scan_path<>''").fetchone()[0]
        for i, r in enumerate(rows):
            scan_ok = "✔" if r[11] and any(os.path.exists(p) for p in r[11].split(";") if p) else ""
            tags = ["odd"] if i%2 else []
            st = r[7] or "Новый"
            tags.append({"Новый":"new","В работе":"work","Исполнен":"done","В архиве":"arch"}.get(st,"new"))
            self.tree.insert("", "end", values=(r[0],r[1],r[2],r[3],r[4],r[5],r[7] or "",r[8] or "",r[9] or "",scan_ok), tags=tuple(tags))
        self.counter.config(text=f"Всего: {total}   •   Со сканами: {withscan}   •   Без скана: {total-withscan}   •   Показано: {len(rows)}")
        for k in ("sender", "case_name"):
            if k in self._autocomplete_widgets:
                self._autocomplete_widgets[k]["values"] = self._get_distinct(k)

    def reset_search(self):
        self.search_var.set(""); self.filter_noscan.set(False); self.load_data()

    # ---------- CRUD ----------
    def add_record(self):
        v = {k: s.get().strip() for k, s in self.vars.items()}
        if not v["in_number"] or not v["in_date"]:
            messagebox.showwarning("Внимание","Заполните номер и дату."); return
        dup = self.conn.execute("SELECT id FROM documents WHERE in_number=? AND in_date=? AND deleted=0",
                                (v["in_number"], v["in_date"])).fetchone()
        if dup:
            if not messagebox.askyesno("Дубль", f"Запись с №{v['in_number']} от {v['in_date']} уже есть (ID {dup[0]}).\nВсё равно добавить?"):
                return
        cur = self.conn.execute("""INSERT INTO documents
            (in_number,in_date,sender,subject,case_name,note,status,executor,remind_date,created_at)
            VALUES (?,?,?,?,?,?,?,?,?,?)""",
            (v["in_number"],v["in_date"],v["sender"],v["subject"],v["case_name"],v["note"],
             self.status_var.get(), v["executor"], v["remind_date"], datetime.now().isoformat(timespec="seconds")))
        self.conn.commit(); self.selected_id = cur.lastrowid; self.load_data(); self.clear_form()
        messagebox.showinfo("Готово","Запись добавлена. Можно прикрепить скан.")

    def update_record(self):
        if self.selected_id is None:
            messagebox.showinfo("Инфо","Выберите запись."); return
        v = {k: s.get().strip() for k, s in self.vars.items()}
        self.conn.execute("""UPDATE documents SET in_number=?,in_date=?,sender=?,subject=?,case_name=?,
            note=?,status=?,executor=?,remind_date=? WHERE id=?""",
            (v["in_number"],v["in_date"],v["sender"],v["subject"],v["case_name"],v["note"],
             self.status_var.get(), v["executor"], v["remind_date"], self.selected_id))
        self.conn.commit(); self.load_data(); messagebox.showinfo("Готово","Изменения сохранены.")

    def delete_record(self):
        if self.selected_id is None:
            messagebox.showinfo("Инфо","Выберите запись."); return
        if messagebox.askyesno("В корзину","Переместить запись в корзину?"):
            self.conn.execute("UPDATE documents SET deleted=1 WHERE id=?", (self.selected_id,))
            self.conn.commit(); self.clear_form(); self.load_data()

    def on_select(self, event):
        sel = self.tree.selection()
        if not sel: return
        vals = self.tree.item(sel[0],"values"); self.selected_id = int(vals[0])
        for k, val in zip(["in_number","in_date","sender","subject","case_name"], vals[1:6]):
            self.vars[k].set(val)
        self.vars["executor"].set(vals[7] if len(vals)>7 else "")
        self.vars["remind_date"].set(vals[8] if len(vals)>8 else "")
        self.status_var.set(vals[6] if len(vals)>6 else "Новый")
        row = self.conn.execute("SELECT note FROM documents WHERE id=?", (self.selected_id,)).fetchone()
        self.vars["note"].set(row[0] if row and row[0] else "")

    def clear_form(self):
        self.selected_id = None
        for s in self.vars.values(): s.set("")
        self.status_var.set("Новый")
        for i in self.tree.selection(): self.tree.selection_remove(i)

    # ---------- Корзина ----------
    def open_trash(self):
        w = tk.Toplevel(self); w.title("Корзина"); w.geometry("900x500")
        w.configure(bg=C["bg"]); w.transient(self)
        cols = ("id","in_number","in_date","sender","subject")
        t = ttk.Treeview(w, columns=cols, show="headings")
        for c,h,wd in zip(cols,("ID","Вх. №","Дата","От кого","Для чего"),(60,100,110,250,350)):
            t.heading(c, text=h); t.column(c, width=wd)
        t.pack(fill="both", expand=True, padx=10, pady=10)
        rows = self.conn.execute("SELECT id,in_number,in_date,sender,subject FROM documents WHERE deleted=1 ORDER BY id DESC").fetchall()
        for r in rows: t.insert("", "end", values=r)
        bar = tk.Frame(w, bg=C["bg"]); bar.pack(fill="x", padx=10, pady=(0,10))
        def restore():
            sel = t.selection()
            if not sel: return
            for s in sel:
                self.conn.execute("UPDATE documents SET deleted=0 WHERE id=?", (int(t.item(s,"values")[0]),))
            self.conn.commit(); w.destroy(); self.load_data()
        def purge():
            if not messagebox.askyesno("Очистить","Удалить ВСЕ записи из корзины безвозвратно?"): return
            self.conn.execute("DELETE FROM documents WHERE deleted=1"); self.conn.commit()
            w.destroy(); self.load_data()
        ttk.Button(bar, text="♻ Восстановить", style="Success.TButton", command=restore).pack(side="left", padx=4)
        ttk.Button(bar, text="🗑 Очистить корзину", style="Danger.TButton", command=purge).pack(side="left", padx=4)
        ttk.Button(bar, text="Закрыть", style="Ghost.TButton", command=w.destroy).pack(side="right", padx=4)

    # ---------- Сканирование ----------
    def scan_current(self):
        if self.selected_id is None:
            messagebox.showinfo("Инфо","Сначала выберите запись."); return
        row = self.conn.execute("SELECT in_number,in_date,sender,scan_path FROM documents WHERE id=?",
                                (self.selected_id,)).fetchone()
        if row[3] and not messagebox.askyesno("Пересканировать?",
                "У записи уже есть скан. Заменить его новым?\nСтарые файлы будут удалены."):
            return
        safe_num = _clean_filename(row[0] or "БезНомера")
        safe_date = _clean_filename(row[1] or "безДаты")
        safe_org = _clean_filename(row[2] or "БезОрганизации")
        org_dir = os.path.join(SCANS_DIR, safe_org); os.makedirs(org_dir, exist_ok=True)
        prefix = os.path.join(org_dir, f"{safe_num}_{safe_date}")
        for old in glob.glob(prefix + "_*.jpg"):
            try: os.remove(old)
            except: pass
        for old in glob.glob(prefix + ".pdf"):
            try: os.remove(old)
            except: pass
        bw = self.bw_var.get(); dpi = int(self.dpi_var.get())
        paper = self.paper_var.get(); fmt = self.format_var.get()
        SETTINGS["paper"]=paper; SETTINGS["dpi"]=str(dpi)
        SETTINGS["bw"]=bw; SETTINGS["scan_format"]=fmt; save_settings(SETTINGS)
        self.scan_btn.config(state="disabled"); self.progress.start(12); self.update()
        try:
            files = do_scan(prefix, bw, dpi, paper, fmt)
        finally:
            self.progress.stop(); self.scan_btn.config(state="normal")
        if files:
            self.conn.execute("UPDATE documents SET scan_path=? WHERE id=?", (";".join(files), self.selected_id))
            self.conn.commit()
            self._run_ocr_for(self.selected_id, files)
            self.load_data()
            if messagebox.askyesno("Готово", f"Сохранено файлов: {len(files)}.\n\nОткрыть просмотр?"):
                self.open_scan()
        else:
            messagebox.showerror("Ошибка","Сканирование не выполнено.\nПроверьте Epson / TWAIN / WIA.")

    def _run_ocr_for(self, doc_id, files):
        if not TESSERACT_PATH:
            messagebox.showwarning("OCR не настроен",
                "Tesseract OCR не найден.\nУстановите его с https://github.com/UB-Mannheim/tesseract/wiki\n"
                "Поиск по содержимому скана будет недоступен.")
            return
        jpgs = []
        for f in files:
            if f.lower().endswith(".jpg"):
                jpgs.append(f)
            elif f.lower().endswith(".pdf"):
                try:
                    from pdf2image import convert_from_path
                    pages = convert_from_path(f, dpi=200)
                    base = os.path.splitext(f)[0]
                    for i, pg in enumerate(pages, 1):
                        tmp = f"{base}_ocr_{i}.jpg"
                        pg.save(tmp, "JPEG", quality=85)
                        jpgs.append(tmp)
                except Exception as e:
                    print("PDF->OCR error:", e)
        if not jpgs: return
        self.progress.start(12); self.update()
        try:
            texts = []
            for f in jpgs:
                t = ocr_image(f)
                if t: texts.append(t)
            if texts:
                self.conn.execute("UPDATE documents SET ocr_text=? WHERE id=?",
                                  ("\n---\n".join(texts), doc_id))
                self.conn.commit()
        finally:
            self.progress.stop()

    def run_ocr_selected(self):
        if self.selected_id is None:
            messagebox.showinfo("Инфо","Выберите запись."); return
        row = self.conn.execute("SELECT scan_path FROM documents WHERE id=?", (self.selected_id,)).fetchone()
        if not row or not row[0]:
            messagebox.showinfo("Инфо","У записи нет скана."); return
        files = [p for p in row[0].split(";") if p and os.path.exists(p)]
        if not files:
            messagebox.showinfo("Инфо","Файлы скана не найдены."); return
        self._run_ocr_for(self.selected_id, files)
        messagebox.showinfo("Готово","OCR выполнен. Поиск теперь ищет по содержимому.")

    def open_scan(self):
        if self.selected_id is None:
            messagebox.showinfo("Инфо","Выберите запись."); return
        row = self.conn.execute("SELECT scan_path FROM documents WHERE id=?", (self.selected_id,)).fetchone()
        if not row or not row[0]:
            if messagebox.askyesno("Нет скана","У записи нет скана. Отсканировать сейчас?"):
                self.scan_current()
            return
        files = [p for p in row[0].split(";") if p and os.path.exists(p)]
        if not files:
            messagebox.showinfo("Инфо","Файлы не найдены."); return
        ScanViewer(self, files)

    def delete_scan(self):
        if self.selected_id is None:
            messagebox.showinfo("Инфо","Выберите запись."); return
        row = self.conn.execute("SELECT scan_path FROM documents WHERE id=?", (self.selected_id,)).fetchone()
        if not row or not row[0]:
            messagebox.showinfo("Инфо","У записи нет скана."); return
        files = [p for p in row[0].split(";") if p]
        if not messagebox.askyesno("Удаление", f"Удалить скан? Файлов: {len(files)}"): return
        deleted, errors = 0, []
        for f in files:
            try:
                if os.path.exists(f): os.remove(f); deleted += 1
            except Exception as e: errors.append(f"{os.path.basename(f)}: {e}")
        self.conn.execute("UPDATE documents SET scan_path=NULL, ocr_text=NULL WHERE id=?", (self.selected_id,))
        self.conn.commit(); self.load_data()
        msg = f"Удалено: {deleted}"
        if errors: msg += "\n\nОшибки:\n" + "\n".join(errors)
        messagebox.showinfo("Готово", msg)

    def open_scans_folder(self): os.startfile(SCANS_DIR)

    # ---------- Напоминания ----------
    def _check_reminders(self):
        today = date.today().isoformat()
        rows = self.conn.execute("SELECT in_number,in_date,remind_date FROM documents WHERE deleted=0 AND remind_date<>'' AND remind_date<=? ORDER BY remind_date", (today,)).fetchall()
        if rows:
            txt = "\n".join(f"• №{r[0]} от {r[1]} — до {r[2]}" for r in rows[:10])
            messagebox.showwarning("⏰ Напоминания", f"Документы с истекающим сроком:\n\n{txt}")

    # ---------- Бэкап ----------
    def _autobackup(self):
        days = SETTINGS.get("autobackup_days", 0)
        if not days: return
        backups = sorted(glob.glob(os.path.join(BASE_DIR, "backup_*.zip")))
        if backups:
            last = datetime.fromtimestamp(os.path.getmtime(backups[-1]))
            if (datetime.now() - last).days < days: return
        self.make_backup(silent=True)

    def make_backup(self, silent=False):
        default = os.path.join(BASE_DIR, f"backup_{datetime.now().strftime('%Y-%m-%d_%H%M')}.zip")
        path = filedialog.asksaveasfilename(defaultextension=".zip",
            filetypes=[("ZIP","*.zip")], initialfile=os.path.basename(default))
        if not path: return
        try:
            with zipfile.ZipFile(path, "w", zipfile.ZIP_DEFLATED) as z:
                if os.path.exists(DB_PATH): z.write(DB_PATH, "incoming.db")
                for root, _, files in os.walk(SCANS_DIR):
                    for f in files:
                        full = os.path.join(root, f)
                        z.write(full, os.path.relpath(full, BASE_DIR))
            if not silent:
                messagebox.showinfo("Готово", f"Бэкап создан:\n{path}")
        except Exception as e:
            messagebox.showerror("Ошибка", f"Не удалось создать бэкап: {e}")

    # ---------- CSV ----------
    def export_csv(self):
        path = filedialog.asksaveasfilename(defaultextension=".csv",
            filetypes=[("CSV","*.csv")], initialfile=f"входящие_{date.today().isoformat()}.csv")
        if not path: return
        try:
            with open(path, "w", newline="", encoding="utf-8-sig") as f:
                w = csv.writer(f, delimiter=";")
                w.writerow(["ID","Вх. №","Дата","От кого","Для чего","Дело","Примечание","Статус","Исполнитель","Напомнить","Создано","Скан","OCR"])
                for r in self.conn.execute("SELECT * FROM documents WHERE deleted=0 ORDER BY id").fetchall():
                    w.writerow([r[0],r[1],r[2],r[3],r[4],r[5],r[6],r[7],r[8],r[9],r[10],r[11],r[12] if len(r)>12 else ""])
            messagebox.showinfo("Готово", f"Экспортировано:\n{path}")
        except Exception as e:
            messagebox.showerror("Ошибка", f"Экспорт не удался: {e}")

    def import_csv(self):
        path = filedialog.askopenfilename(filetypes=[("CSV","*.csv")])
        if not path: return
        try:
            count = 0
            with open(path, "r", encoding="utf-8-sig") as f:
                reader = csv.reader(f, delimiter=";")
                next(reader, None)
                for row in reader:
                    if len(row) < 3: continue
                    self.conn.execute("""INSERT INTO documents
                        (in_number,in_date,sender,subject,case_name,note,status,executor,remind_date,created_at)
                        VALUES (?,?,?,?,?,?,?,?,?,?)""",
                        (row[1] if len(row)>1 else "", row[2] if len(row)>2 else "",
                         row[3] if len(row)>3 else "", row[4] if len(row)>4 else "",
                         row[5] if len(row)>5 else "", row[6] if len(row)>6 else "",
                         row[7] if len(row)>7 and row[7] else "Новый",
                         row[8] if len(row)>8 else "", row[9] if len(row)>9 else "",
                         datetime.now().isoformat(timespec="seconds")))
                    count += 1
            self.conn.commit(); self.load_data()
            messagebox.showinfo("Готово", f"Импортировано записей: {count}")
        except Exception as e:
            messagebox.showerror("Ошибка", f"Импорт не удался: {e}")

    # ---------- Печать ----------
    def _make_word_file(self, rows, title, one=False):
        try:
            from docx import Document; from docx.shared import Pt
        except ImportError:
            messagebox.showerror("Нужен python-docx","pip install python-docx"); return
        doc = Document(); doc.add_heading(title, level=1)
        if one:
            r = rows[0]
            for k, v in [("Входящий номер",r[1]),("Дата",r[2]),("От кого",r[3]),
                         ("Для чего",r[4]),("Дело",r[5]),("Примечание",r[6]),
                         ("Статус",r[7]),("Исполнитель",r[8]),("Напомнить до",r[9])]:
                p = doc.add_paragraph(); p.add_run(f"{k}: ").bold = True
                p.add_run(str(v) if v else "—")
            if r[11]:
                p = doc.add_paragraph(); p.add_run("Скан: ").bold = True; p.add_run(r[11])
        else:
            table = doc.add_table(rows=1, cols=9); table.style = "Table Grid"
            for i,h in enumerate(["№","Вх. №","Дата","От кого","Для чего","Дело","Статус","Исполнитель","Напомнить"]):
                cell = table.rows[0].cells[i]; cell.text = h; cell.paragraphs[0].runs[0].bold = True
            for idx, r in enumerate(rows, 1):
                cells = table.add_row().cells; cells[0].text = str(idx)
                for i, val in enumerate([r[1],r[2],r[3],r[4],r[5],r[7],r[8],r[9]], 1):
                    cells[i].text = str(val) if val else ""
            for row in table.rows:
                for cell in row.cells:
                    for p in cell.paragraphs:
                        for run in p.runs: run.font.size = Pt(9)
        out = filedialog.asksaveasfilename(defaultextension=".docx",
            filetypes=[("Word","*.docx")], initialfile="Печать_входящих.docx")
        if not out: return
        doc.save(out); messagebox.showinfo("Готово", f"Сохранено:\n{out}\n\nОткройте и Ctrl+P.")
        try: os.startfile(out)
        except: pass

    def print_selected(self):
        sel = self.tree.selection()
        if not sel: messagebox.showinfo("Инфо","Выберите запись."); return
        vals = self.tree.item(sel[0],"values")
        row = self.conn.execute("SELECT * FROM documents WHERE id=?", (int(vals[0]),)).fetchone()
        self._make_word_file([row],"Входящий документ",one=True)

    def print_all(self):
        rows = self.conn.execute(f"SELECT * FROM documents WHERE deleted=0 ORDER BY {self.sort_col} "
                                 f"{'DESC' if self.sort_desc else 'ASC'}, id DESC").fetchall()
        if not rows: messagebox.showinfo("Инфо","Журнал пуст."); return
        self._make_word_file(rows,"Журнал входящей документации")

    # ---------- Закрытие ----------
    def _on_close(self):
        SETTINGS["win_geom"] = self.geometry(); save_settings(SETTINGS)
        self.destroy()

if __name__ == "__main__":
    app = App()
    app.mainloop()
