import os
import sys
import json
import shutil
import logging
import glob
import webbrowser
import tempfile
import subprocess

# =====================================================================
# REMOTE ON/OFF SWITCH
# Agar kabhi kisi client ka access rokna ho, bas neeche "ACCESS_ENABLED"
# ko False kar ke GitHub par commit kar dena — poora asal code (neeche)
# safe rahega, delete karne ki zaroorat nahi. Jab wapas chalana ho, sirf
# True kar dena aur commit kar dena.
# =====================================================================
ACCESS_ENABLED = True
BLOCK_MESSAGE = "Softwar Access Denied.\nContact admin for support."

if not ACCESS_ENABLED:
    import tkinter as _tk
    from tkinter import messagebox as _mb
    _root = _tk.Tk()
    _root.withdraw()
    _mb.showerror("Access Suspended", BLOCK_MESSAGE)
    _root.destroy()
    sys.exit()
# =====================================================================


def open_file_with_default_app(file_path):
    """Preview file (PDF/Excel) ko default app mein kholta hai — har platform
    ka apna sabse pakka/native tareeka (webbrowser.open se zyada reliable,
    khaas kar Mac par .app bundle ke andar se chalne par)."""
    try:
        if sys.platform == "win32":
            os.startfile(file_path)
        elif sys.platform == "darwin":
            subprocess.run(["open", file_path], check=True)
        else:
            subprocess.run(["xdg-open", file_path], check=True)
    except Exception as e:
        messagebox.showerror("Open Error", f"File ban gayi hai lekin khul nahi saki:\n{file_path}\n\nError: {e}")

# Sidebar footer mein dikhane ke liye — launcher (ap2.py) is number ko GitHub
# ke version.txt/release se compare karta hai, yahan sirf display ke liye hai.
CURRENT_VERSION = "1.0"
from datetime import datetime
import calendar as calendar_module
import customtkinter as ctk
from tkinter import messagebox, filedialog, ttk, simpledialog
from PIL import Image, ImageTk, ImageEnhance

# ReportLab core imports fix for NumberedCanvas
from reportlab.pdfgen import canvas
from reportlab.lib.pagesizes import letter
from reportlab.lib import colors
from reportlab.platypus import SimpleDocTemplate, Paragraph, Spacer, Table, TableStyle, PageBreak
from reportlab.lib.styles import ParagraphStyle

# openpyxl imports for Excel engine
from openpyxl import Workbook
from openpyxl.styles import Font as XLFont, PatternFill, Alignment, Border, Side

ctk.set_appearance_mode("System")
ctk.set_default_color_theme("blue")

ADMIN_USER = "admin"
ADMIN_PASS = "1234"

# - - - - PERMANENT STORAGE SETUP (Cross-Platform) - - - -
# Windows par pehle jaisa C:\TrutzschlerData use hota hai; Mac/Linux par
# hardcoded "C:\..." path exist hi nahi karta (aur likhne ki permission bhi
# nahi hoti), is liye us par har user ke apne Home folder ke andar
# ~/TrutzschlerData use hota hai.
if sys.platform == "win32":
    APP_DATA_DIR = "C:\\TrutzschlerData"
else:
    APP_DATA_DIR = os.path.join(os.path.expanduser("~"), "TrutzschlerData")

BASE_HISTORY_DIR = os.path.join(APP_DATA_DIR, "System_History")
DB_FILE_PATH = os.path.join(APP_DATA_DIR, "inventory_database.json")
# NAYA: Client Factories (Name + Address) ka permanent master record.
# Ek baar factory ka naam aur address save karein, phir hamesha dropdown
# se select karke dono fields automatically bhar jayenge.
FACTORY_DB_PATH = os.path.join(APP_DATA_DIR, "factory_database.json")

# NAYA: Automatic backups folder — har save se pehle purani DB file ka
# ek timestamped safety copy yahan rakha jata hai (crash/corruption se
# bachne ke liye). Sirf latest MAX_BACKUPS copies rakhi jati hain.
BACKUP_DIR = os.path.join(APP_DATA_DIR, "Backups")
MAX_BACKUPS_PER_FILE = 15

for directory in [APP_DATA_DIR, BASE_HISTORY_DIR, BACKUP_DIR]:
    if not os.path.exists(directory):
        os.makedirs(directory)

BG_IMAGE_PATH = os.path.join(APP_DATA_DIR, "trutzschler_bg.png")

# --- NAYA: CENTRAL LOGGING SETUP ---
# Pehle silent "print()" ka use ho raha tha jiski wajah se agar koi save
# ya update-check fail hota tha to woh sirf terminal (jo .exe run karte
# waqt normally dikhta hi nahi) me chala jata tha — user ko pata hi nahi
# chalta tha kuch ghalat hua. Ab har error ek permanent log file
# (app_log.txt) me bhi likha jata hai taake baad me issue diagnose kiya
# ja sake.
LOG_FILE_PATH = os.path.join(APP_DATA_DIR, "app_log.txt")
logging.basicConfig(
    filename=LOG_FILE_PATH,
    level=logging.INFO,
    format="%(asctime)s | %(levelname)s | %(message)s",
    encoding="utf-8"
)
app_logger = logging.getLogger("TrutzschlerWorkspace")


def log_error(context, exc):
    """Har jagah print(f'Error ...: {e}') ki jagah yeh use karo —
    terminal par bhi dikhega aur permanent log file me bhi save hoga."""
    message = f"{context}: {exc}"
    print(message)
    try:
        app_logger.error(message)
    except Exception:
        pass


def create_timestamped_backup(source_path):
    """Kisi bhi database file (inventory / factory) ka overwrite se
    THEEK PEHLE ek timestamped copy Backups/ folder me bana deta hai.
    Agar naya save corrupt ho jaye ya galat data save ho jaye, to yahan
    se purani copy wapas restore ki ja sakti hai. Har file ke liye sirf
    latest MAX_BACKUPS_PER_FILE copies rakhi jati hain, purani khud
    delete ho jati hain."""
    if not os.path.exists(source_path):
        return
    try:
        base_name = os.path.splitext(os.path.basename(source_path))[0]
        stamp = datetime.now().strftime("%Y%m%d_%H%M%S")
        backup_name = f"{base_name}_{stamp}.json"
        shutil.copy2(source_path, os.path.join(BACKUP_DIR, backup_name))

        # Purani backups saaf karna (sirf isi file ki, naam se match karke)
        existing = sorted(
            glob.glob(os.path.join(BACKUP_DIR, f"{base_name}_*.json")),
            key=os.path.getmtime
        )
        for old_file in existing[:-MAX_BACKUPS_PER_FILE]:
            try:
                os.remove(old_file)
            except Exception:
                pass
    except Exception as e:
        log_error("Backup creation skipped", e)


def atomic_json_write(target_path, data, **dump_kwargs):
    """Direct 'open(path,\"w\")' se file ko overwrite karna khatarnak
    hai — agar save ke beech me hi software crash ho jaye ya bijli
    chali jaye, to JSON file aadhi likhi reh jati hai aur poori tarah
    CORRUPT ho jati hai (sara data loss). Is liye pehle ek temp file me
    likha jata hai, aur sirf successful likhne ke baad hi asli file ki
    jagah rakha jata hai (os.replace atomic operation hai)."""
    create_timestamped_backup(target_path)
    temp_path = target_path + ".tmp"
    with open(temp_path, "w", encoding="utf-8") as f:
        json.dump(data, f, **dump_kwargs)
    os.replace(temp_path, target_path)

# --- Automatic Background Image Downloader ---
if not os.path.exists(BG_IMAGE_PATH):
    try:
        IMAGE_URL = "https://images.unsplash.com/photo-1581091226825-a6a2a5aee158?q=80&w=1280&auto=format&fit=crop"
        headers = {'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64)'}
        req = urllib.request.Request(IMAGE_URL, headers=headers)
        # Speed fix: timeout laga diya hai taake yeh one-time background image
        # download bhi slow network par software ko atka na de.
        with urllib.request.urlopen(req, timeout=5) as response, open(BG_IMAGE_PATH, 'wb') as out_file:
            out_file.write(response.read())
    except Exception as e:
        print(f"Network asset sync skipped: {e}")

# --- DEFAULT INVENTORY DATA WITH PERMANENT BASE PRICES ---
DEFAULT_INVENTORY = {
    "Blowroom": [
        {"id": "1.1.1.1", "name": "Condenser", "code": "BR-COI", "specs": "Standard condenser for material transport and air separation", "price": "12500"},
        {"id": "1.1.1.2", "name": "Stock Trunk", "code": "FD-S 1200", "specs": "High capacity stock trunk for material buffering", "price": "18400"},
        {"id": "1.1.1.3", "name": "2 Way Distribution", "code": "BR-2W", "specs": "Pneumatic distribution system with active change boxes", "price": "9200"},
        {"id": "1.1.1.4", "name": "Automatic Change Box", "code": "BR-AC", "specs": "High performance control change box for multi-line distribution", "price": "14500"}
    ],
    "Card": [
        {"id": "1.2.1.1", "name": "Card Clothing Wire", "code": "CR-WR", "specs": "High durability cylinder and doffer wire set", "price": "4200"},
        {"id": "1.2.1.2", "name": "Flat Tops", "code": "CR-FT", "specs": "Precision engineered flat tops for enhanced trash removal", "price": "6800"}
    ],
    "Drawing": [
        {"id": "1.3.1.1", "name": "Top Roller", "code": "DR-TR", "specs": "Precision top roller with heavy duty bearing alignments", "price": "1500"},
        {"id": "1.3.1.2", "name": "Bottom Fluted Roller", "code": "DR-BR", "specs": "High tolerance bottom fluted mechanical roller", "price": "2900"}
    ],
    "Comber": [
        {"id": "1.4.1.1", "name": "Top Comb", "code": "CM-TC", "specs": "Premium pinning density top comb segment", "price": "2100"},
        {"id": "1.4.1.2", "name": "Circular Comb", "code": "CM-CC", "specs": "Graded combing profile circular comb block assembly", "price": "5400"}
    ]
}
# --- DEPENDENCY-FREE CALENDAR POPUP (Dashboard From/To date filter) ---
# Built using only the standard library's "calendar" module + customtkinter,
# so it works even on machines where nothing extra (like tkcalendar) has
# been installed or bundled into the .exe build.
class MiniCalendarPopup(ctk.CTkToplevel):
    def __init__(self, parent, target_entry, on_pick=None):
        super().__init__(parent)
        self.title("Select Date")
        self.resizable(False, False)
        self.transient(parent)
        self.grab_set()
        self.lift()

        self.target_entry = target_entry
        self.on_pick = on_pick

        today = datetime.now()
        try:
            existing = datetime.strptime(target_entry.get().strip(), "%d.%m.%Y")
            self.cur_year, self.cur_month = existing.year, existing.month
        except Exception:
            self.cur_year, self.cur_month = today.year, today.month

        nav = ctk.CTkFrame(self, fg_color="transparent")
        nav.pack(fill="x", padx=10, pady=(10, 4))

        ctk.CTkButton(nav, text="◀", width=32, command=self.prev_month).pack(side="left")
        self.lbl_month = ctk.CTkLabel(nav, text="", font=("Arial", 13, "bold"))
        self.lbl_month.pack(side="left", expand=True, fill="x")
        ctk.CTkButton(nav, text="▶", width=32, command=self.next_month).pack(side="left")

        self.grid_frame = ctk.CTkFrame(self, fg_color="transparent")
        self.grid_frame.pack(padx=10, pady=(0, 6))

        ctk.CTkButton(self, text="Clear Date", width=100, fg_color="#94a3b8",
                      hover_color="#64748b", command=self.clear_date).pack(pady=(0, 10))

        self.render_calendar()

        self.update_idletasks()
        x = parent.winfo_x() + (parent.winfo_width() // 2) - (self.winfo_width() // 2)
        y = parent.winfo_y() + (parent.winfo_height() // 2) - (self.winfo_height() // 2)
        self.geometry(f"+{max(x,0)}+{max(y,0)}")

    def render_calendar(self):
        for w in self.grid_frame.winfo_children():
            w.destroy()

        self.lbl_month.configure(text=f"{calendar_module.month_name[self.cur_month]} {self.cur_year}")

        days = ["Mo", "Tu", "We", "Th", "Fr", "Sa", "Su"]
        for i, d in enumerate(days):
            ctk.CTkLabel(self.grid_frame, text=d, font=("Arial", 11, "bold"), width=34).grid(row=0, column=i, padx=1, pady=1)

        cal = calendar_module.Calendar(firstweekday=0)
        month_days = cal.monthdayscalendar(self.cur_year, self.cur_month)
        today = datetime.now()
        for r, week in enumerate(month_days, start=1):
            for c, day in enumerate(week):
                if day == 0:
                    ctk.CTkLabel(self.grid_frame, text="", width=34, height=28).grid(row=r, column=c, padx=1, pady=1)
                else:
                    is_today = (day == today.day and self.cur_month == today.month and self.cur_year == today.year)
                    btn = ctk.CTkButton(
                        self.grid_frame, text=str(day), width=34, height=28,
                        fg_color="#228B22" if is_today else "#1f538d",
                        hover_color="#006400" if is_today else "#153a63",
                        font=("Arial", 11),
                        command=lambda d=day: self.pick_day(d)
                    )
                    btn.grid(row=r, column=c, padx=1, pady=1)

    def prev_month(self):
        self.cur_month -= 1
        if self.cur_month < 1:
            self.cur_month = 12
            self.cur_year -= 1
        self.render_calendar()

    def next_month(self):
        self.cur_month += 1
        if self.cur_month > 12:
            self.cur_month = 1
            self.cur_year += 1
        self.render_calendar()

    def pick_day(self, day):
        picked = f"{day:02d}.{self.cur_month:02d}.{self.cur_year}"
        self.target_entry.delete(0, "end")
        self.target_entry.insert(0, picked)
        if self.on_pick:
            self.on_pick()
        self.destroy()

    def clear_date(self):
        self.target_entry.delete(0, "end")
        if self.on_pick:
            self.on_pick()
        self.destroy()


# --- HEADING ADD KARNE KA POPUP DIALOG ---
def ask_heading_dialog(parent):
    dialog = ctk.CTkToplevel(parent)
    dialog.title("Add Line Heading")
    dialog.geometry("380x180")
    dialog.resizable(False, False)
    dialog.grab_set()
    
    dialog.update_idletasks()
    x = parent.winfo_x() + (parent.winfo_width() // 2) - 190
    y = parent.winfo_y() + (parent.winfo_height() // 2) - 90
    dialog.geometry(f"+{x}+{y}")

    ctk.CTkLabel(dialog, text="Enter Line / Section Heading Title:", font=("Arial", 13, "bold")).pack(pady=(20, 10))
    
    ent_title = ctk.CTkEntry(dialog, width=300, placeholder_text="e.g. Line 1: Blowroom Main Section")
    ent_title.pack(pady=5)
    ent_title.focus()

    res = {"title": None}

    def on_save():
        val = ent_title.get().strip()
        if val:
            res["title"] = val
            dialog.destroy()

    ctk.CTkButton(dialog, text="➕ Add Heading Banner", fg_color="#8b5cf6", hover_color="#7c3aed", command=on_save).pack(pady=15)
    parent.wait_window(dialog)
    return res["title"]

def load_inventory_from_db():
    # Cross-platform: Windows par C:\TrutzschlerData, Mac/Linux par
    # ~/TrutzschlerData — dono APP_DATA_DIR se aate hain
    target_folder = APP_DATA_DIR
    target_path = DB_FILE_PATH
    
    if os.path.exists(target_path):
        try:
            with open(target_path, "r", encoding="utf-8") as f:
                data = json.load(f)
                for cat in data:
                    for item in data[cat]:
                        if "price" not in item:
                            item["price"] = ""
                return data
        except Exception:
            return DEFAULT_INVENTORY
    else:
        # Agar folder nahi bana hua toh pehle folder banayein
        if not os.path.exists(target_folder):
            try:
                os.makedirs(target_folder)
            except:
                pass
        try:
            atomic_json_write(target_path, DEFAULT_INVENTORY, indent=4)
        except Exception as e:
            log_error("Could not create initial inventory database", e)
        return DEFAULT_INVENTORY

def save_inventory_to_db(data):
    target_folder = APP_DATA_DIR
    target_path = DB_FILE_PATH
    
    # Save karne se pehle check karna ke folder maujood hai ya nahi
    if not os.path.exists(target_folder):
        try:
            os.makedirs(target_folder)
        except:
            pass

    # ASLI FIX: "➕ Dup" button se banaye gaye per-quotation duplicate
    # instances ("is_duplicate_instance": True) kabhi bhi master product
    # catalog file me permanently save NAHI hone chahiye — warna woh
    # hamesha ke liye asli catalog ka hissa ban ke ganda kar dete. Yeh
    # filter har save par lagta hai, chahe save kahin se bhi call ho.
    filtered_data = {
        cat: [p for p in items if not p.get("is_duplicate_instance", False)]
        for cat, items in data.items()
    }

    try:
        atomic_json_write(target_path, filtered_data, indent=4)
    except Exception as e:
        log_error("Error saving database", e)


# =====================================================================
# CLIENT FACTORY MASTER RECORD (NAAM + ADDRESS) — LOAD / SAVE
# =====================================================================
def load_factories_from_db():
    """Saved client factories ki list return karta hai.
    Format: [{"name": "AGI Denim", "address": "Plot 1, Karachi"}, ...]"""
    if not os.path.exists(FACTORY_DB_PATH):
        return []
    try:
        with open(FACTORY_DB_PATH, "r", encoding="utf-8") as f:
            data = json.load(f)
        clean = []
        if isinstance(data, list):
            for row in data:
                if isinstance(row, dict) and str(row.get("name", "")).strip():
                    clean.append({
                        "name": str(row.get("name", "")).strip(),
                        "address": str(row.get("address", "")).strip()
                    })
        return clean
    except Exception:
        return []


def save_factories_to_db(data):
    if not os.path.exists(APP_DATA_DIR):
        try:
            os.makedirs(APP_DATA_DIR)
        except Exception:
            pass
    try:
        atomic_json_write(FACTORY_DB_PATH, data, indent=4, ensure_ascii=False)
    except Exception as e:
        log_error("Error saving factory database", e)


class FactoryEntryDialog(ctk.CTkToplevel):
    """Chhota popup jis se Factory ka Naam aur poora Address add/edit hota hai."""

    def __init__(self, parent, title="Add New Client Factory", initial=None):
        super().__init__(parent)
        self.title(title)
        self.geometry("480x400")
        self.resizable(False, False)
        self.transient(parent)
        self.lift()

        self.result = None
        initial = initial or {}

        ctk.CTkLabel(self, text=title, font=("Arial", 16, "bold"),
                     text_color="#1f538d").pack(pady=(18, 12))

        ctk.CTkLabel(self, text="Client Factory Name", font=("Arial", 12, "bold"),
                     anchor="w").pack(fill="x", padx=30, pady=(4, 2))
        self.ent_name = ctk.CTkEntry(self, placeholder_text="e.g. AGI Denim (Pvt) Ltd",
                                     height=36, font=("Arial", 13))
        self.ent_name.pack(fill="x", padx=30)
        if initial.get("name"):
            self.ent_name.insert(0, initial["name"])

        ctk.CTkLabel(self, text="Full Factory Address Location", font=("Arial", 12, "bold"),
                     anchor="w").pack(fill="x", padx=30, pady=(14, 2))
        self.txt_address = ctk.CTkTextbox(self, height=110, font=("Arial", 13),
                                          border_width=2, border_color="#cbd5e1")
        self.txt_address.pack(fill="x", padx=30)
        if initial.get("address"):
            self.txt_address.insert("1.0", initial["address"])

        btn_frame = ctk.CTkFrame(self, fg_color="transparent")
        btn_frame.pack(fill="x", padx=30, pady=20)
        ctk.CTkButton(btn_frame, text="Cancel", width=110, fg_color="#64748b",
                      hover_color="#475569", command=self.destroy).pack(side="left")
        ctk.CTkButton(btn_frame, text="💾 Save Factory", width=150, fg_color="#228B22",
                      hover_color="#006400", command=self.validate_and_save).pack(side="right")

        self.update_idletasks()
        x = (self.winfo_screenwidth() // 2) - (self.winfo_width() // 2)
        y = (self.winfo_screenheight() // 2) - (self.winfo_height() // 2)
        self.geometry(f"+{x}+{y}")

        self.focus_force()
        self.grab_set()
        self.ent_name.after(200, lambda: self.ent_name.focus_set())

    def validate_and_save(self):
        name = self.ent_name.get().strip()
        address = self.txt_address.get("1.0", "end").strip()
        if not name:
            messagebox.showwarning("Field Required",
                                   "Client Factory Name is mandatory.", parent=self)
            return
        self.result = {"name": name, "address": address}
        self.destroy()


class NumberedCanvas(canvas.Canvas):
    def __init__(self, *args, **kwargs):
        super().__init__(*args, **kwargs)
        self._saved_page_states = []
    def showPage(self):
        self._saved_page_states.append(dict(self.__dict__))
        self._startPage()
    def save(self):
        num_pages = len(self._saved_page_states)
        for state in self._saved_page_states:
            self.__dict__.update(state)
            self.draw_decorations(num_pages)
            super().showPage()
        super().save()
    def draw_decorations(self, total_pages):
        self.saveState()
        
        self.setFont("Helvetica-Bold", 14)
        self.setFillColor(colors.HexColor("#003399"))
        self.drawString(54, 755, "TRÜTZSCHLER")
        
        self.setFont("Helvetica-Bold", 14)
        self.setFillColor(colors.HexColor("#7f8c8d"))
        self.drawString(162, 755, "SPINNING")
        
        self.setFont("Helvetica", 7.5)
        self.setFillColor(colors.HexColor("#555555"))
        self.drawString(54, 743, "Trützschler Group SE  PO Box 410164  D-41241 Mönchengladbach")
        
        self.setStrokeColor(colors.HexColor("#003399"))
        self.setLineWidth(0.75)
        self.line(54, 732, 558, 732)
        
        self.setFont("Helvetica", 6.5)
        self.setFillColor(colors.HexColor("#333333"))
        self.setStrokeColor(colors.HexColor("#cccccc"))
        self.setLineWidth(0.5)
        self.line(54, 75, 558, 75)
        
        col1_y = 65
        self.drawString(54, col1_y, "Board of Directors:")
        self.drawString(54, col1_y - 9, "Heinrich Krull")
        self.drawString(54, col1_y - 18, "Florian Rück")
        self.drawString(54, col1_y - 27, "Alexander Stampfer")
        
        col2_x = 160
        self.drawString(col2_x, col1_y, "Chairman of the Supervisory Board:")
        self.drawString(col2_x, col1_y - 9, "Dr.-Ing. Roland Münch")
        self.drawString(col2_x, col1_y - 18, "Trützschler Group SE, AG Mönchengladbach HRB 20308")
        self.drawString(col2_x, col1_y - 27, "Headquarters: Mönchengladbach Deutschland")
        
        col3_x = 345
        self.drawString(col3_x, col1_y, "Deutsche Bank AG, Mönchengladbach")
        self.drawString(col3_x, col1_y - 9, "IBAN: DE 85 3107 0001 0720 9000 00")
        self.drawString(col3_x, col1_y - 18, "BIC: DEUTDEDD310")
        self.drawString(col3_x, col1_y - 27, "VAT ID No.: DE 291 704 386")
        
        col4_x = 465
        self.drawString(col4_x, col1_y, "Commerzbank AG")
        self.drawString(col4_x, col1_y - 9, "Mönchengladbach")
        self.drawString(col4_x, col1_y - 18, "IBAN: DE 66 3104 0015 0363 6008 00")
        self.drawString(col4_x, col1_y - 27, "BIC: COBADEFF310")
        
        self.setFont("Helvetica", 8)
        self.drawRightString(558, 30, f"Page {self._pageNumber} of {total_pages}")
        self.restoreState()


class LoginWindow(ctk.CTkFrame):
    def __init__(self, master, on_success):
        super().__init__(master, fg_color="#0b0f19")
        self.on_success = on_success
        self.pack(expand=True, fill="both")
        
        self.bg_label = ctk.CTkLabel(self, text="")
        self.bg_label.place(x=0, y=0, relwidth=1, relheight=1)
        self.bind("<Configure>", self.resize_background)
        
        self.box = ctk.CTkFrame(self, width=420, height=480, corner_radius=24, fg_color="#121826", border_width=1.5, border_color="#2a364f")
        self.box.place(relx=0.5, rely=0.5, anchor="center")
        self.box.pack_propagate(False) 
        
        ctk.CTkLabel(self.box, text="TRÜTZSCHLER", font=("Segoe UI", 32, "bold"), text_color="#3b82f6").pack(pady=(45, 2))
        ctk.CTkLabel(self.box, text="GERMAN ENGINEERING WORKSPACE", font=("Segoe UI", 10, "bold"), text_color="#64748b").pack(pady=(0, 35))
        
        self.ent_user = ctk.CTkEntry(self.box, placeholder_text="Username", width=320, height=44, font=("Segoe UI", 14), corner_radius=10, fg_color="#090d16", border_color="#1e293b", text_color="#f8fafc", placeholder_text_color="#475569")
        self.ent_user.pack(pady=12)
        
        self.ent_pass = ctk.CTkEntry(self.box, placeholder_text="Password", show="*", width=320, height=44, font=("Segoe UI", 14), corner_radius=10, fg_color="#090d16", border_color="#1e293b", text_color="#f8fafc", placeholder_text_color="#475569")
        self.ent_pass.pack(pady=12)
        
        self.ent_pass.bind("<Return>", lambda event: self.check_credentials())
        self.ent_user.bind("<Return>", lambda event: self.check_credentials())

        self.btn_access = ctk.CTkButton(self.box, text="SECURE ACCESS SYSTEM", font=("Segoe UI", 13, "bold"), width=320, height=46, corner_radius=10, fg_color="#3b82f6", hover_color="#2563eb", text_color="#ffffff")
        self.btn_access.configure(command=self.check_credentials)
        self.btn_access.pack(pady=(35, 20))

    def resize_background(self, event=None):
        if os.path.exists(BG_IMAGE_PATH):
            try:
                window_width = self.winfo_width()
                window_height = self.winfo_height()
                if window_width > 10 and window_height > 10:
                    # Speed fix: agar size pichli baar jaisa hi hai to dobara
                    # image open/resize/enhance na karein — yeh bhaari
                    # operation hai aur <Configure> event baar baar (khaaskar
                    # window ke shuru mein) fire hota hai.
                    last_size = getattr(self, "_bg_last_size", None)
                    if last_size == (window_width, window_height):
                        return
                    self._bg_last_size = (window_width, window_height)

                    img = Image.open(BG_IMAGE_PATH)
                    img = img.resize((window_width, window_height), Image.Resampling.LANCZOS)
                    enhancer = ImageEnhance.Brightness(img)
                    img = enhancer.enhance(0.45) 
                    self.bg_img_ctk = ctk.CTkImage(light_image=img, dark_image=img, size=(window_width, window_height))
                    self.bg_label.configure(image=self.bg_img_ctk)
            except Exception as e:
                print(f"Sizing hook feedback: {e}")
        else:
            self.bg_label.configure(fg_color="#0b0f19")

    def check_credentials(self):
        if self.ent_user.get().strip() == ADMIN_USER and self.ent_pass.get().strip() == ADMIN_PASS:
            self.pack_forget()
            self.on_success()
        else:
            messagebox.showerror("Access Denied", "Invalid Credentials. Please try again.")
            
    def clear_fields(self):
        self.ent_user.delete(0, "end")
        self.ent_pass.delete(0, "end")
        self.ent_user.focus()


class CustomInputDialog(ctk.CTkToplevel):
    def __init__(self, parent, title, prompts, initial_values=None):
        super().__init__(parent)
        self.title(title)
        self.geometry("400x420") 
        self.resizable(False, False)
        
        # Windows system configurations for stable focus mapping
        self.transient(parent)
        self.lift()
        
        self.result = None
        self.entries = {}
        
        ctk.CTkLabel(self, text=title, font=("Arial", 16, "bold"), text_color="#1f538d").pack(pady=15)
        
        first_ent = None
        for placeholder, key in prompts:
            frame = ctk.CTkFrame(self, fg_color="transparent")
            frame.pack(fill="x", padx=30, pady=5)
            
            # FIXED: Explicit normal state to guarantee layout usability
            ent = ctk.CTkEntry(frame, placeholder_text=placeholder, width=340, height=35, state="normal")
            ent.pack()
            
            if initial_values and key in initial_values:
                ent.insert(0, initial_values[key])
                
            self.entries[key] = ent
            if first_ent is None:
                first_ent = ent
            
        btn_frame = ctk.CTkFrame(self, fg_color="transparent")
        btn_frame.pack(fill="x", padx=30, pady=20)
        
        ctk.CTkButton(btn_frame, text="Cancel", width=100, fg_color="#64748b", command=self.destroy).pack(side="left", padx=5)
        ctk.CTkButton(btn_frame, text="Save Item", width=100, fg_color="#228B22", command=self.validate_and_save).pack(side="right", padx=5)
        
        self.center_window()
        
        # FIXED: Removed the buggy "<Button-1>" dialog focus-stealing binding entirely.
        # Enforcing immediate structural focus on window generation
        self.focus_force()
        self.grab_set()
        if first_ent:
            first_ent.after(200, lambda: first_ent.focus_set())

    def center_window(self):
        self.update_idletasks()
        width = self.winfo_width()
        height = self.winfo_height()
        x = (self.winfo_screenwidth() // 2) - (width // 2)
        y = (self.winfo_screenheight() // 2) - (height // 2)
        self.geometry(f'{width}x{height}+{x}+{y}')

    def validate_and_save(self):
        res = {k: v.get().strip() for k, v in self.entries.items()}
        if not res["name"]:
            messagebox.showwarning("Fields Required", "Item Name is a mandatory field.")
            return
        if not res["specs"]:
            res["specs"] = "No additional technical specification provided."
        self.result = res
        self.destroy()


class ToastNotification(ctk.CTkToplevel):
    """NAYA: Modern, non-blocking success/info popup.
    Native tkinter messagebox jarring lagta hai (system beep, "OK" click
    karna parta hai, poora screen block ho jata hai). Yeh instead bottom
    right corner me chhota, auto-vanish (2.2 sec) card dikhata hai — jaisa
    modern desktop software (VS Code, Slack, etc.) me hota hai. Koi click
    zaroori nahi, kaam mein rukawat nahi aati."""
    def __init__(self, parent, message, kind="success"):
        super().__init__(parent)
        self.overrideredirect(True)
        self.attributes("-topmost", True)
        try:
            self.attributes("-alpha", 0.0)
        except Exception:
            pass

        accent = "#228B22" if kind == "success" else ("#cc3333" if kind == "error" else "#1f538d")
        icon = "✅" if kind == "success" else ("⚠️" if kind == "error" else "ℹ️")

        card = ctk.CTkFrame(self, corner_radius=10, fg_color=("#f2f2f2", "#242424"),
                             border_width=2, border_color=accent)
        card.pack(fill="both", expand=True)
        ctk.CTkLabel(card, text=f"{icon}  {message}", font=("Arial", 13, "bold"),
                     wraplength=320, justify="left").pack(padx=18, pady=14)

        self.update_idletasks()
        w, h = 360, self.winfo_reqheight() + 6
        try:
            px = parent.winfo_x() + parent.winfo_width() - w - 30
            py = parent.winfo_y() + parent.winfo_height() - h - 40
        except Exception:
            px, py = 100, 100
        self.geometry(f"{w}x{h}+{px}+{py}")

        self._fade_in()
        self.after(2200, self._fade_out)

    def _fade_in(self, alpha=0.0):
        try:
            alpha = min(alpha + 0.15, 0.96)
            self.attributes("-alpha", alpha)
            if alpha < 0.96:
                self.after(15, lambda: self._fade_in(alpha))
        except Exception:
            pass

    def _fade_out(self, alpha=0.96):
        try:
            alpha = max(alpha - 0.12, 0.0)
            self.attributes("-alpha", alpha)
            if alpha > 0:
                self.after(15, lambda: self._fade_out(alpha))
            else:
                self.destroy()
        except Exception:
            try:
                self.destroy()
            except Exception:
                pass


class MasterSystemApp(ctk.CTk):
    def __init__(self):
        super().__init__()
        self.all_heading_cards = []
        self.title("Trutzschler Proforma & Quotation Matrix Workspace")
        self.geometry("1320x900")
        # Chhoti aur badi dono screens ke liye ek reasonable minimum rakha hai —
        # itna bada nahi ke chhoti screen par window screen se bahar chala jaye.
        self.minsize(1150, 700)
        self.selection_counter = 0  # Dynamic click tracking ke liye
        
        self.inventory_db = load_inventory_from_db()
        
        self.sidebar = None
        self.main_workspace = None
        self.grid_inputs = {}
        self.sub_scroll_frames = {}
        self.app_loaded = False  

        # Tracks which history record (if any) is currently loaded into the
        # workspace, so that when the user saves again we can ask whether
        # to update that SAME existing file or create a brand new record.
        self.loaded_record_path = None
        self.loaded_record_ref = None

        # Safe binding global click handler
        self.bind("<Button-1>", self.remove_global_cursor_focus)
        
        self.login_view = LoginWindow(self, self.load_main_application)
        # Heading switching tracking variables
        # Class ke init me yeh list lazmi add karein
        
        
        self._switching_lines = False

    def remove_global_cursor_focus(self, event):
        if not self.app_loaded:
            return
            
        try:
            # FIXED: Do not alter global workspace structural focus if a top-level popup is active
            for child in self.winfo_children():
                if isinstance(child, ctk.CTkToplevel) and child.winfo_exists():
                    return

            widget_under_mouse = self.winfo_containing(event.x_root, event.y_root)
            widget_str = str(widget_under_mouse).lower()
            if "entry" in widget_str or "text" in widget_str:
                return  
                
            self.focus_set()
        except Exception:
            pass

    def load_main_application(self):
        self.app_loaded = True 
        self.grid_columnconfigure(0, weight=0, minsize=290)
        self.grid_columnconfigure(1, weight=1)
        self.grid_rowconfigure(0, weight=1)

        self.sidebar = ctk.CTkFrame(self, width=290, corner_radius=0)
        self.sidebar.grid(row=0, column=0, sticky="nsew")
        
        self.main_workspace = ctk.CTkFrame(self, fg_color="transparent")
        self.main_workspace.grid(row=0, column=1, sticky="nsew", padx=20, pady=20)
        
        self.sidebar_setup()
        self.workspace_setup()
        self.toggle_dynamic_fields()

    def show_toast(self, message, kind="success"):
        """Helper: kahin se bhi self.show_toast('...') call karke modern
        auto-vanish notification dikhayi ja sakti hai, jarring messagebox
        ki jagah."""
        try:
            ToastNotification(self, message, kind=kind)
        except Exception as e:
            log_error("Toast display failed", e)

    def sidebar_setup(self):
        ctk.CTkLabel(self.sidebar, text="TRÜTZSCHLER", font=("Arial", 22, "bold"), text_color="#1f538d").pack(pady=(20, 2))
        ctk.CTkLabel(self.sidebar, text="Enterprise Tracking Core", font=("Arial", 11, "italic"), text_color="gray").pack(pady=(0, 15))
        ctk.CTkLabel(self.sidebar, text="📋 DOCUMENT CONFIGURATION", font=("Arial", 11, "bold")).pack(anchor="w", padx=20, pady=(10, 2))
        
        self.doc_type_var = ctk.StringVar(value="Quotation")
        combo_type = ctk.CTkComboBox(self.sidebar, values=["Quotation", "Proforma"], variable=self.doc_type_var, width=250, font=("Arial", 12, "bold"), command=self.toggle_dynamic_fields)
        combo_type.pack(padx=20, pady=5)

        # --- CLIENT FACTORY SELECTOR ROW ---
        # Pehle yahan sirf ek khaali input tha jisme har baar poora naam aur
        # address haath se likhna parta tha. Ab: saved factories ki searchable
        # dropdown list (▼) + naya factory add karne ka (➕) button. Factory
        # select karte hi neeche address khud ba khud bhar jata hai.
        self.factory_row = ctk.CTkFrame(self.sidebar, fg_color="transparent")
        self.factory_row.pack(padx=20, pady=4, fill="x")

        self.ent_factory = ctk.CTkEntry(
            self.factory_row,
            placeholder_text="Client Factory Name (AGI Denim)",
            width=180
        )
        self.ent_factory.pack(side="left", fill="x", expand=True)

        self.btn_factory_dropdown = ctk.CTkButton(
            self.factory_row, text="▼", width=32, height=28,
            font=("Arial", 12, "bold"), fg_color="#1f538d", hover_color="#153a63",
            command=self.open_factory_selector_popup
        )
        self.btn_factory_dropdown.pack(side="left", padx=(4, 0))

        self.btn_factory_add = ctk.CTkButton(
            self.factory_row, text="➕", width=32, height=28,
            font=("Arial", 12, "bold"), fg_color="#228B22", hover_color="#006400",
            command=self.open_add_factory_dialog
        )
        self.btn_factory_add.pack(side="left", padx=(4, 0))

        self.ent_address = ctk.CTkEntry(self.sidebar, placeholder_text="Full Factory Address Location", width=250)
        self.ent_address.pack(padx=20, pady=4)
        self.ent_offer = ctk.CTkEntry(self.sidebar, placeholder_text="Offer / Ref (e.g., 02-013106-00)", width=250)
        self.ent_offer.pack(padx=20, pady=4)
        self.ent_project = ctk.CTkEntry(self.sidebar, placeholder_text="Project Name (Draw Frames)", width=250)
        self.ent_project.pack(padx=20, pady=4)

        self.lbl_previews = ctk.CTkLabel(self.sidebar, text="⚙️ WORKING ENGINE PREVIEWS", font=("Arial", 11, "bold"))
        self.lbl_previews.pack(anchor="w", padx=20, pady=(15, 2))
        
        self.btn_preview_pdf = ctk.CTkButton(self.sidebar, text="🔍 Preview Quotation (PDF)", font=("Arial", 13, "bold"), fg_color="#ffcc00", text_color="black", hover_color="#e6b800", height=38, command=self.trigger_pdf_preview)
        self.btn_preview_excel = ctk.CTkButton(self.sidebar, text="🔍 Preview Spreadsheet (Excel)", font=("Arial", 13, "bold"), fg_color="#ffcc00", text_color="black", hover_color="#e6b800", height=38, command=self.trigger_excel_preview)

        self.lbl_exports = ctk.CTkLabel(self.sidebar, text="📦 EXPORT & STATE TRACKING", font=("Arial", 11, "bold"))
        self.lbl_exports.pack(anchor="w", padx=20, pady=(20, 2))
        
        self.btn_export_excel = ctk.CTkButton(self.sidebar, text="📊 Save & Export Quotation Excel", font=("Arial", 13, "bold"), fg_color="#1f538d", height=38, command=self.generate_excel_format)
        self.btn_export_pdf = ctk.CTkButton(self.sidebar, text="📜 Save & Export Quotation PDF", font=("Arial", 13, "bold"), fg_color="#228B22", hover_color="#006400", height=38, command=self.generate_pdf_proforma)

        ctk.CTkLabel(self.sidebar, text="🔒 SECURITY SESSIONS", font=("Arial", 11, "bold")).pack(anchor="w", padx=20, pady=(25, 2))
        btn_logout = ctk.CTkButton(self.sidebar, text="🔴 Logout System", font=("Arial", 13, "bold"), fg_color="#cc3333", hover_color="#992222", height=35, command=self.trigger_system_logout)
        btn_logout.pack(padx=20, pady=4, fill="x")

        # NAYA: Version footer — professional software me hamesha version
        # number dikhta hai taake support/debugging ke waqt pata chal sake
        # user kaunsa build use kar raha hai.
        ctk.CTkLabel(
            self.sidebar, text=f"v{CURRENT_VERSION}  •  Trutzschler Workspace",
            font=("Arial", 9), text_color="gray"
        ).pack(side="bottom", pady=10)

    # =================================================================
    # CLIENT FACTORY MANAGER (Dropdown + Search + Add / Edit / Delete)
    # =================================================================
    def apply_factory_selection(self, factory):
        """Selected factory ka naam aur address dono inputs me daal deta hai."""
        try:
            self.ent_factory.delete(0, "end")
            self.ent_factory.insert(0, factory.get("name", ""))
            self.ent_address.delete(0, "end")
            self.ent_address.insert(0, factory.get("address", ""))
            self.focus_set()
        except Exception as e:
            print(f"Factory selection error: {e}")

    def open_add_factory_dialog(self, refresh_callback=None):
        """Naya factory (naam + address) permanently save karta hai."""
        dialog = FactoryEntryDialog(self, title="Add New Client Factory")
        self.wait_window(dialog)
        if not dialog.result:
            return

        factories = load_factories_from_db()
        new_name = dialog.result["name"]
        if any(f["name"].lower() == new_name.lower() for f in factories):
            messagebox.showwarning(
                "Already Exists",
                f"'{new_name}' is already saved in your factory list."
            )
            return

        factories.append(dialog.result)
        factories.sort(key=lambda f: f["name"].lower())
        save_factories_to_db(factories)

        # Naya factory add hotay hi seedha select bhi ho jata hai
        self.apply_factory_selection(dialog.result)

        if refresh_callback:
            refresh_callback()

    def edit_saved_factory(self, old_factory, refresh_callback=None):
        dialog = FactoryEntryDialog(self, title="Edit Client Factory", initial=old_factory)
        self.wait_window(dialog)
        if not dialog.result:
            return

        factories = load_factories_from_db()
        for i, f in enumerate(factories):
            if f["name"] == old_factory["name"] and f.get("address", "") == old_factory.get("address", ""):
                factories[i] = dialog.result
                break
        factories.sort(key=lambda f: f["name"].lower())
        save_factories_to_db(factories)

        # Agar wohi factory abhi inputs me selected tha to usay update kar dein
        if self.ent_factory.get().strip() == old_factory["name"]:
            self.apply_factory_selection(dialog.result)

        if refresh_callback:
            refresh_callback()

    def delete_saved_factory(self, factory, refresh_callback=None):
        confirm = messagebox.askyesno(
            "Confirm Delete",
            f"Remove '{factory['name']}' from your saved factory list?\n\n"
            "(This only removes it from the saved list — no existing document will be deleted.)"
        )
        if not confirm:
            return

        factories = load_factories_from_db()
        factories = [
            f for f in factories
            if not (f["name"] == factory["name"] and f.get("address", "") == factory.get("address", ""))
        ]
        save_factories_to_db(factories)

        if refresh_callback:
            refresh_callback()

    def open_factory_selector_popup(self):
        """Saved factories ki searchable list — select karte hi naam + address bhar jate hain."""
        popup = ctk.CTkToplevel(self)
        popup.title("Select Client Factory")
        popup.geometry("560x520")
        popup.transient(self)
        popup.lift()

        ctk.CTkLabel(popup, text="🏭 Saved Client Factories", font=("Arial", 16, "bold"),
                     text_color="#1f538d").pack(pady=(15, 5))
        ctk.CTkLabel(popup, text="Select a factory — the name and address will fill in automatically.",
                     font=("Arial", 11), text_color="gray").pack(pady=(0, 10))

        head_bar = ctk.CTkFrame(popup, fg_color="transparent")
        head_bar.pack(fill="x", padx=18, pady=(0, 8))

        search_wrap = ctk.CTkFrame(head_bar, fg_color="#ffffff", corner_radius=10,
                                   border_width=2, border_color="#1f538d", height=36)
        search_wrap.pack(side="left", fill="x", expand=True)
        search_wrap.pack_propagate(False)

        ctk.CTkLabel(search_wrap, text="🔍", font=("Segoe UI Emoji", 14),
                     text_color="#1f538d", width=24).pack(side="left", padx=(8, 0))

        search_var = ctk.StringVar()
        ent_search = ctk.CTkEntry(search_wrap, textvariable=search_var,
                                  placeholder_text="Search Factory",
                                  border_width=0, fg_color="transparent",
                                  placeholder_text_color="#94a3b8", text_color="#0f172a",
                                  font=("Segoe UI", 13))
        ent_search.pack(side="left", fill="x", expand=True, padx=(2, 8))

        list_frame = ctk.CTkScrollableFrame(popup, fg_color="#f8fafc")
        list_frame.pack(fill="both", expand=True, padx=18, pady=5)

        def _grab_safe(action):
            """Child dialog kholne se pehle popup ka grab chhorna zaroori hai,
            warna Windows par confirm/input box freeze ho jata hai."""
            try:
                popup.grab_release()
            except Exception:
                pass
            action()
            try:
                if popup.winfo_exists():
                    popup.lift()
                    popup.grab_set()
            except Exception:
                pass

        def render_list(*_):
            for w in list_frame.winfo_children():
                w.destroy()

            query = search_var.get().strip().lower()
            factories = load_factories_from_db()
            filtered = [
                f for f in factories
                if query in f["name"].lower() or query in f.get("address", "").lower()
            ]

            if not factories:
                ctk.CTkLabel(list_frame,
                             text="No factory has been saved yet.\n\n"
                                  "Use '➕ Add New Factory' below to save your\n"
                                  "first factory's name and address.",
                             font=("Arial", 12), text_color="gray",
                             justify="center").pack(pady=40)
                return

            if not filtered:
                ctk.CTkLabel(list_frame, text=f"No factory matched '{search_var.get().strip()}'",
                             font=("Arial", 12), text_color="gray").pack(pady=40)
                return

            for fac in filtered:
                row = ctk.CTkFrame(list_frame, fg_color="#ffffff", corner_radius=8,
                                   border_width=1, border_color="#e2e8f0")
                row.pack(fill="x", padx=5, pady=4)

                info = ctk.CTkFrame(row, fg_color="transparent")
                info.pack(side="left", fill="x", expand=True, padx=10, pady=8)

                ctk.CTkLabel(info, text=fac["name"], font=("Arial", 13, "bold"),
                             text_color="#0f172a", anchor="w").pack(fill="x")
                addr_preview = fac.get("address", "") or "— no address saved —"
                if len(addr_preview) > 70:
                    addr_preview = addr_preview[:70] + "..."
                ctk.CTkLabel(info, text=addr_preview, font=("Arial", 11),
                             text_color="#64748b", anchor="w").pack(fill="x")

                def do_select(f=fac):
                    self.apply_factory_selection(f)
                    popup.destroy()

                ctk.CTkButton(row, text="🗑️", width=36, height=30, font=("Arial", 12),
                              fg_color="#cc3333", hover_color="#992222",
                              command=lambda f=fac: _grab_safe(lambda: self.delete_saved_factory(f, render_list))
                              ).pack(side="right", padx=(2, 8), pady=8)
                ctk.CTkButton(row, text="✏️", width=36, height=30, font=("Arial", 12),
                              fg_color="#64748b", hover_color="#475569",
                              command=lambda f=fac: _grab_safe(lambda: self.edit_saved_factory(f, render_list))
                              ).pack(side="right", padx=2, pady=8)
                ctk.CTkButton(row, text="✔ Select", width=80, height=30,
                              font=("Arial", 12, "bold"), fg_color="#228B22",
                              hover_color="#006400", command=do_select
                              ).pack(side="right", padx=2, pady=8)

        search_var.trace_add("write", render_list)

        bottom_bar = ctk.CTkFrame(popup, fg_color="transparent")
        bottom_bar.pack(fill="x", padx=18, pady=(5, 15))

        ctk.CTkButton(bottom_bar, text="➕ Add New Factory", font=("Arial", 13, "bold"),
                      fg_color="#228B22", hover_color="#006400", height=36,
                      command=lambda: _grab_safe(lambda: self.open_add_factory_dialog(render_list))
                      ).pack(side="left")
        ctk.CTkButton(bottom_bar, text="Close", font=("Arial", 13, "bold"), width=100,
                      fg_color="#64748b", hover_color="#475569", height=36,
                      command=popup.destroy).pack(side="right")

        render_list()

        popup.update_idletasks()
        x = self.winfo_x() + (self.winfo_width() // 2) - 280
        y = self.winfo_y() + (self.winfo_height() // 2) - 260
        popup.geometry(f"+{max(x, 0)}+{max(y, 0)}")

        popup.focus_force()
        popup.grab_set()
        ent_search.after(200, lambda: ent_search.focus_set())

    def toggle_dynamic_fields(self, *args):
        doc_type = self.doc_type_var.get()
        if doc_type == "Quotation":
            self.ent_address.pack_forget()
            self.ent_project.pack_forget()
            self.ent_offer.pack(padx=20, pady=4, after=self.factory_row)
            self.btn_preview_pdf.pack_forget()
            self.btn_export_pdf.pack_forget()
            self.btn_preview_excel.pack(padx=20, pady=4, fill="x", after=self.lbl_previews)
            self.btn_export_excel.pack(padx=20, pady=4, fill="x", after=self.lbl_exports)
            self.btn_preview_excel.configure(text="🔍 Preview Spreadsheet (Excel)")
            self.btn_export_excel.configure(text="📊 Save & Export Quotation Excel")
        else:
            self.ent_address.pack_forget()
            self.ent_offer.pack_forget()
            self.ent_project.pack_forget()
            self.ent_address.pack(padx=20, pady=4, after=self.factory_row)
            self.ent_offer.pack(padx=20, pady=4, after=self.ent_address)
            self.ent_project.pack(padx=20, pady=4, after=self.ent_offer)
            self.btn_preview_excel.pack_forget()
            self.btn_export_excel.pack_forget()
            self.btn_preview_pdf.pack(padx=20, pady=4, fill="x", after=self.lbl_previews)
            self.btn_export_pdf.pack(padx=20, pady=4, fill="x", after=self.lbl_exports)
            self.btn_preview_pdf.configure(text="🔍 Preview Proforma (PDF)")
            self.btn_export_pdf.configure(text="📜 Save & Export Proforma PDF")

    def workspace_setup(self):
        self.tabs_container = ctk.CTkTabview(self.main_workspace, height=800)
        self.tabs_container.pack(fill="both", expand=True)

        tab_matrix = self.tabs_container.add("Matrix Selection")
        
        top_action_bar = ctk.CTkFrame(tab_matrix, fg_color="#eef2f7", height=50, corner_radius=8)
        top_action_bar.pack(fill="x", padx=5, pady=(5, 10))
        top_action_bar.pack_propagate(False)

        self.matrix_search_var = ctk.StringVar()
        self.matrix_search_var.trace_add("write", self.filter_matrix_by_product_name)
        
        # FIX: pehle sirf placeholder text tha jo cursor andar jatay hi ghaib
        # ho jata tha — is liye pata hi nahi chalta tha ke yeh search box hai.
        # Ab ek wrapper frame ke andar hamesha nazar aane wala 🔍 icon aur
        # "Search Products" label lagaya gaya hai jo kabhi ghaib nahi hota.
        search_wrap_matrix = ctk.CTkFrame(
            top_action_bar,
            fg_color="#ffffff",
            corner_radius=10,
            border_width=2,
            border_color="#3b82f6",
            height=35
        )
        # fill/expand taake yeh search box screen ke size ke hisaab se khud
        # chhota-bada ho jaye, aur dosre buttons kabhi overlap na hon chahe
        # window kitni bhi choti ya badi ho.
        search_wrap_matrix.pack(side="left", fill="x", expand=True, padx=15, pady=7)
        search_wrap_matrix.pack_propagate(False)

        ctk.CTkLabel(
            search_wrap_matrix, text="🔍", font=("Segoe UI Emoji", 14),
            text_color="#3b82f6", width=22
        ).pack(side="left", padx=(10, 2))

        ent_matrix_search = ctk.CTkEntry(
            search_wrap_matrix,
            textvariable=self.matrix_search_var,
            placeholder_text="Search Products",
            height=29,
            corner_radius=6,
            border_width=0,
            fg_color="transparent",
            placeholder_text_color="#94a3b8",
            text_color="#0f172a",
            font=("Segoe UI", 13)
        )
        ent_matrix_search.pack(side="left", fill="x", expand=True, padx=(2, 10))

        btn_delete_item = ctk.CTkButton(top_action_bar, text="❌ Delete Selected", font=("Arial", 12, "bold"), fg_color="#cc3333", hover_color="#992222", width=130, command=self.delete_item_global_trigger)
        btn_delete_item.pack(side="right", padx=10, pady=10)
        
        

        btn_edit_item = ctk.CTkButton(top_action_bar, text="📝 Edit Selected", font=("Arial", 12, "bold"), fg_color="#1f538d", hover_color="#153a63", width=120, command=self.edit_item_global_trigger)
        btn_edit_item.pack(side="right", padx=10, pady=10)
        
        btn_add_item = ctk.CTkButton(top_action_bar, text="➕ Add New Item", font=("Arial", 12, "bold"), fg_color="#228B22", hover_color="#006400", width=120, command=self.add_item_global_trigger)
        btn_add_item.pack(side="right", padx=10, pady=10)

        # --- MANAGE DUPLICATES BUTTON ---
        # ASLI FIX: pehle isay search bar ke paas "left" me laga diya tha
        # jisse Refresh System button ke saath overlap ho raha tha. Ab isay
        # baaki utility buttons ke sath sahi tarah "right" side ke group me
        # rakha hai — Refresh System ke bilkul saath, koi overlap nahi.
        btn_manage_dups = ctk.CTkButton(
            top_action_bar,
            text="🔁 Manage Duplicates",
            font=("Arial", 12, "bold"),
            fg_color="#0ea5e9",
            hover_color="#0284c7",
            width=170,
            command=self.open_manage_duplicates_popup
        )
        btn_manage_dups.pack(side="right", padx=10, pady=10)

        # Refresh Workspace Button setup
        btn_refresh_system = ctk.CTkButton(
            top_action_bar, 
            text="🔄 Refresh System", 
            font=("Arial", 12, "bold"), 
            fg_color="#5a6268",        # Slate gray color takay positive actions se different lagay
            hover_color="#495057", 
            width=130, 
            command=self.refresh_entire_workspace
        )
        btn_refresh_system.pack(side="right", padx=10, pady=10)

        


        

        self.scroll_grid = ctk.CTkScrollableFrame(tab_matrix, fg_color="white", corner_radius=8, border_width=1, border_color="#dbdbdb")
        self.scroll_grid.pack(fill="both", expand=True, padx=5, pady=5)

        self.tabview_sub = ctk.CTkTabview(self.scroll_grid, height=590, command=self.reset_matrix_search_on_tab_change)
        self.tabview_sub.pack(fill="both", expand=True)

        

        for cat in self.inventory_db.keys():
            tab = self.tabview_sub.add(cat)

            # --- TOP ROW: Add Heading Button (Placed directly above Qty column) ---
            action_row = ctk.CTkFrame(tab, fg_color="transparent")
            action_row.pack(fill="x", padx=10, pady=(2, 4))

            btn_add_heading_row = ctk.CTkButton(
                action_row, 
                text="📑 Add Line Heading", 
                font=("Arial", 11, "bold"), 
                fg_color="#8b5cf6", 
                hover_color="#7c3aed", 
                height=26,
                width=130, 
                command=lambda: self.add_heading_trigger(self.tabview_sub.get())
            )
            btn_add_heading_row.pack(side="right", padx=(0, 5))
            
            frame_th = ctk.CTkFrame(tab, fg_color="transparent")
            frame_th.pack(fill="x", padx=10, pady=6)
            
            ctk.CTkLabel(frame_th, text="Sel", width=45, font=("Arial", 12, "bold"), anchor="center").pack(side="left")
            ctk.CTkLabel(frame_th, text="Items ", width=340, anchor="w", font=("Arial", 12, "bold")).pack(side="left", padx=10)
            ctk.CTkLabel(frame_th, text="Code Configuration", width=170, font=("Arial", 12, "bold"), anchor="w").pack(side="left", padx=5)
            ctk.CTkLabel(frame_th, text="Unit Price (€)", width=160, font=("Arial", 12, "bold"), anchor="w").pack(side="left", padx=5) 
            ctk.CTkLabel(frame_th, text="Qty", width=65, font=("Arial", 12, "bold"), anchor="center").pack(side="left", padx=5)
            
            sub_scroll = ctk.CTkScrollableFrame(tab, height=440)
            sub_scroll.pack(fill="both", expand=True)
            
            self.sub_scroll_frames[cat] = sub_scroll
            self.grid_inputs[cat] = []
            
            # Speed fix: pehle yahan HAR category ki product rows turant ban
            # jati thin (login ke foran baad), chahe user ne wo tab kabhi
            # khola bhi na ho. Ab sirf pehli/currently visible tab turant
            # banti hai; baaki categories "lazy" load hoti hain — jab user
            # us tab par click karta hai tab (reset_matrix_search_on_tab_change
            # mein) banti hain. Data/functionality bilkul same rehta hai,
            # sirf ban'ne ka waqt badalta hai.
            if not hasattr(self, '_populated_cats'):
                self._populated_cats = set()
            if not self._populated_cats:
                self.populate_category_rows(cat)
                self._populated_cats.add(cat)

        self.tab_dashboard = self.tabs_container.add("📊 Dashboard & History Tracking")
        self.setup_dashboard_view()

    def on_checkbox_toggle(self):
        """Saves product selections to the currently active line card when checked/unchecked."""
        if getattr(self, '_switching_lines', False):
            return

        if hasattr(self, 'active_heading_card') and self.active_heading_card:
            self.save_current_line_state(self.active_heading_card)

    def save_current_line_state(self, card_widget):
        """Active heading line ke selected products aur unke states ko save karta hai"""
        if not card_widget or not hasattr(self, 'grid_inputs'):
            return

        current_state = {}
        for cat, items in self.grid_inputs.items():
            for item in items:
                key = item.get("id")
                # Sirf tab save karein jab checkbox ON (Checked) ho
                if item.get("checkbox") and item["checkbox"].get() == 1:
                    c_val = item["code_show"].get() == 1 if "code_show" in item else True
                    p_val = item["price_show"].get() == 1 if "price_show" in item else True
                    q_val = item["qty"].get().strip() if "qty" in item else "1"
                    # ASLI FIX: pehle yahan "price" bilkul save nahi hota tha,
                    # is liye PDF/Excel hamesha sirf master-catalog price
                    # istemaal karte the — kisi bhi per-line custom price
                    # (jaise duplicate wale dialog se diya hua) ko yeh line
                    # kabhi save hi nahi karti thi.
                    pr_val = item["price"].get().strip() if "price" in item else ""

                    current_state[key] = {
                        "selected": True,
                        "code_show": c_val,
                        "price_show": p_val,
                        "qty": str(q_val),
                        "price": str(pr_val),
                        "order": item["select_order"].get() if "select_order" in item else 0
                    }

        # Heading Card ke andar isolated data save kar diya
        card_widget.saved_selections = current_state

    def reset_all_product_rows(self):
        """Screen se saaray tick marks aur input fields ko zabardasti BLANK karta hai"""
        if hasattr(self, 'grid_inputs'):
            for cat, items in self.grid_inputs.items():
                for item in items:
                    # Uncheck main selection
                    if "checkbox" in item and item["checkbox"]:
                        item["checkbox"].deselect()
                    
                    # Reset Code & Price checkboxes
                    if "code_show" in item and item["code_show"]:
                        item["code_show"].select()
                    if "price_show" in item and item["price_show"]:
                        item["price_show"].select()
                    
                    # Reset Qty Field
                    if "qty" in item and item["qty"]:
                        item["qty"].delete(0, "end")
                        item["qty"].insert(0, "1")
                    
                    # Reset Selection Order Variable
                    if "select_order" in item and hasattr(item["select_order"], 'set'):
                        item["select_order"].set(0)

    def activate_heading_line(self, selected_card, parent_container, clear_all=False):
        """Line switch hone par purani line ka data save karke nayi line ka fresh state load karta hai"""
        # Lock lagayein taake checkbox events trigger na hon
        self._switching_lines = True

        try:
            # 1. Purani active line ka state save karein
            if hasattr(self, 'active_heading_card') and self.active_heading_card and self.active_heading_card != selected_card:
                self.save_current_line_state(self.active_heading_card)

            # 2. Visual Color Highlight Reset
            for child in parent_container.winfo_children():
                if getattr(child, 'is_section_heading', False):
                    child.configure(fg_color="#1e3a8a")  # Inactive Dark Blue
            
            selected_card.configure(fg_color="#2563eb")  # Active Bright Blue
            self.active_heading_card = selected_card

            if not hasattr(selected_card, 'saved_selections'):
                selected_card.saved_selections = {}

            # 3. Screen ke saare product rows ko ZABARDASTI BLANK (Clear) karein
            self.reset_all_product_rows()

            # 4. Agar selected line par pehle se koi items saved hain, sirf unhi ko TICK karein
            if not clear_all and hasattr(self, 'grid_inputs'):
                saved = getattr(selected_card, 'saved_selections', {})
                for cat, items in self.grid_inputs.items():
                    for item in items:
                        key = item.get("id")
                        if key in saved and saved[key].get("selected", False):
                            prod_data = saved[key]
                            
                            # Main Checkbox ON
                            if item.get("checkbox"):
                                item["checkbox"].select()
                            
                            # Code Show/Hide Checkbox
                            if item.get("code_show"):
                                item["code_show"].select() if prod_data.get("code_show", True) else item["code_show"].deselect()
                            
                            # Price Show/Hide Checkbox
                            if item.get("price_show"):
                                item["price_show"].select() if prod_data.get("price_show", True) else item["price_show"].deselect()
                            
                            # Quantity
                            if item.get("qty"):
                                item["qty"].delete(0, "end")
                                item["qty"].insert(0, str(prod_data.get("qty", "1")))

                            # ASLI FIX: Price bhi isi line ki apni saved value se
                            # restore karein (agar is line ke liye custom price
                            # save hua tha), warna jo currently field me hai
                            # wahi rehne dein (purane/backward-compatible records)
                            if item.get("price") and "price" in prod_data and prod_data.get("price", "") != "":
                                item["price"].delete(0, "end")
                                item["price"].insert(0, str(prod_data.get("price")))

            self.update_idletasks()

        finally:
            # Switch complete hone par lock khol dein
            self._switching_lines = False

    def create_heading_card(self, category, line_number, title, saved_selections=None, auto_activate=True):
        """Reusable Line/Heading card builder. add_heading_trigger (manual, ek
        naya heading) aur load_selected_record (Dashboard se ek poora saved
        record rebuild karte waqt, har line ke liye) dono isi function ko
        call karte hain, taake widget-creation logic sirf ek jagah rahe."""
        current_tab = self.tabview_sub.tab(category)

        scroll_frame = None
        for child in current_tab.winfo_children():
            if isinstance(child, ctk.CTkScrollableFrame):
                scroll_frame = child
                break

        target_parent = scroll_frame if scroll_frame else current_tab

        heading_card = ctk.CTkFrame(target_parent, fg_color="#1e3a8a", corner_radius=6, cursor="hand2")

        children = target_parent.winfo_children()
        if children:
            heading_card.pack(fill="x", padx=5, pady=(10, 5), before=children[0])
        else:
            heading_card.pack(fill="x", padx=5, pady=(10, 5))

        heading_card.is_section_heading = True
        heading_card.raw_title = title
        heading_card.line_no = line_number
        # Line kis category-tab ke andar banayi gayi thi, yeh yaad rakhna zaroori
        # hai taake Dashboard se load karte waqt yeh sahi tab me dobara ban sake
        heading_card.owner_category = category
        # Agar purani saved selections di gayi hain (record load) to wohi
        # isolated storage bano, warna bilkul nayi khaali line
        heading_card.saved_selections = dict(saved_selections) if saved_selections else {}

        # Permanent registry mein save karein taake PDF/Excel generation ke waqt
        # yeh line HAMESHA mile, chahe live widget-scan miss kar jaye
        if not hasattr(self, 'all_heading_cards'):
            self.all_heading_cards = []
        self.all_heading_cards.append(heading_card)

        lbl_title = ctk.CTkLabel(
            heading_card, 
            text=f"{line_number}  {title}", 
            font=("Arial", 12, "bold"), 
            text_color="white"
        )
        lbl_title.pack(side="left", padx=12, pady=6)
        heading_card.lbl_title = lbl_title

        # Lambda Binding Fix for Selection
        heading_card.bind("<Button-1>", lambda e, h=heading_card, p=target_parent: self.activate_heading_line(h, p))
        lbl_title.bind("<Button-1>", lambda e, h=heading_card, p=target_parent: self.activate_heading_line(h, p))

        actions_frame = ctk.CTkFrame(heading_card, fg_color="transparent")
        actions_frame.pack(side="right", padx=5, pady=4)

        btn_edit = ctk.CTkButton(
            actions_frame, 
            text="✏️ Edit", 
            width=65, 
            height=24, 
            fg_color="#3b82f6", 
            hover_color="#2563eb",
            font=("Arial", 10, "bold"),
            command=lambda h=heading_card: self.edit_heading_title(h)
        )
        btn_edit.pack(side="left", padx=3)

        btn_del = ctk.CTkButton(
            actions_frame, 
            text="✕ Delete", 
            width=65, 
            height=24, 
            fg_color="#ef4444", 
            hover_color="#dc2626",
            font=("Arial", 10, "bold"),
            command=lambda h=heading_card: self.delete_heading_line(h)
        )
        btn_del.pack(side="left", padx=3)

        if auto_activate:
            # Fresh manual line -> screen bilkul blank; restored line -> apni
            # saved ticks ke sath load ho
            self.activate_heading_line(heading_card, target_parent, clear_all=(not saved_selections))

        return heading_card

    def add_heading_trigger(self, current_category):
        """Nayi Section Line Add karne ka complete logic"""

        dialog = ctk.CTkInputDialog(
            text=f"Enter Line Heading for '{current_category}':", 
            title="Add Custom Section Line"
        )
        heading_text = dialog.get_input()
        
        if heading_text and heading_text.strip():
            current_tab = self.tabview_sub.tab(current_category)
            
            scroll_frame = None
            for child in current_tab.winfo_children():
                if isinstance(child, ctk.CTkScrollableFrame):
                    scroll_frame = child
                    break
            
            target_parent = scroll_frame if scroll_frame else current_tab

            existing_headings = [
                w for w in target_parent.winfo_children() 
                if getattr(w, 'is_section_heading', False)
            ]
            
            cat_index = list(self.inventory_db.keys()).index(current_category) + 1 if hasattr(self, 'inventory_db') else 1
            line_number = f"{cat_index}.{len(existing_headings) + 1}"

            # Fresh Line Add hote hi Screen ko Pura Blank (Reset) karke banayein
            self.create_heading_card(current_category, line_number, heading_text.strip())

    def edit_heading_title(self, heading_card):
        """Heading Title Edit Window"""
        edit_win = ctk.CTkToplevel(self)
        edit_win.title("Edit Section Heading")
        edit_win.geometry("380x160")
        edit_win.resizable(False, False)
        edit_win.attributes("-topmost", True)

        lbl = ctk.CTkLabel(edit_win, text="Update Heading Title:", font=("Arial", 12, "bold"))
        lbl.pack(pady=(15, 5))

        entry = ctk.CTkEntry(edit_win, width=300)
        entry.pack(pady=5)
        entry.insert(0, heading_card.raw_title)
        entry.focus()

        def save_and_close():
            new_text = entry.get().strip()
            if new_text:
                heading_card.raw_title = new_text
                heading_card.lbl_title.configure(text=f"{heading_card.line_no}  {new_text}")
            edit_win.destroy()

        btn_save = ctk.CTkButton(edit_win, text="Save Changes", width=120, command=save_and_close)
        btn_save.pack(pady=12)

    def delete_heading_line(self, heading_card):
        """Line Delete Logic"""
        if hasattr(self, 'active_heading_card') and self.active_heading_card == heading_card:
            self.active_heading_card = None
            self._switching_lines = True
            self.reset_all_product_rows()
            self._switching_lines = False

        # Permanent registry se bhi nikal dein
        if hasattr(self, 'all_heading_cards') and heading_card in self.all_heading_cards:
            self.all_heading_cards.remove(heading_card)

        heading_card.destroy()

    def populate_category_rows(self, cat):
        for widget in self.sub_scroll_frames[cat].winfo_children():
            widget.destroy()
        self.grid_inputs[cat].clear()

        products = self.inventory_db.get(cat, [])
        for prod in products:
            is_dup = prod.get("is_duplicate_instance", False)

            # ASLI FIX: Duplicate row ko halke neele background se turant
            # pehchana ja sake, aur baaki normal catalog rows bilkul waise
            # hi rahein jaise pehle thin
            row = ctk.CTkFrame(self.sub_scroll_frames[cat], fg_color=("#e6f0ff" if is_dup else "transparent"), height=45)
            row.pack(fill="x", padx=5, pady=4)
            row.pack_propagate(False)

            chk_frame = ctk.CTkFrame(row, fg_color="transparent", width=45, height=45)
            chk_frame.pack(side="left")
            chk_frame.pack_propagate(False)

            chk_order_var = ctk.IntVar(value=0)
            chk = ctk.CTkCheckBox(chk_frame, text="", width=24, height=24)
            
            # Replace lines 610-618 with this:
            def track_click_selection(c=chk, v=chk_order_var):
                if c.get() == 1:
                    if not hasattr(self, 'selection_counter'):
                        self.selection_counter = 0
                    self.selection_counter += 1
                    v.set(self.selection_counter)
                else:
                    v.set(0)
                
                # Active line ke andar instant save karne ke liye call karein:
                self.on_checkbox_toggle()

            chk.configure(command=track_click_selection)
            chk.pack(expand=True)

            # ASLI FIX: Duplicate row par "🔁 Duplicate" badge dikhayein taake
            # yeh original product se hamesha alag pehchana ja sake, chahe
            # yeh list ke aakhir me kahin bhi ho
            item_label_text = f"🔁 {prod['name']}  (Duplicate)" if is_dup else prod["name"]
            lbl_item = ctk.CTkLabel(
                row, text=item_label_text, width=340, height=45, anchor="w",
                font=("Arial", 13, "bold" if is_dup else "normal"),
                text_color=("#1d4ed8" if is_dup else None)
            )
            lbl_item.pack(side="left", padx=10)

            code_compound_frame = ctk.CTkFrame(row, fg_color="transparent", width=170, height=45)
            code_compound_frame.pack(side="left", padx=5)
            code_compound_frame.pack_propagate(False)
            
            chk_code_show = ctk.CTkCheckBox(code_compound_frame, text="", width=20, height=20)
            chk_code_show.select()
            chk_code_show.pack(side="left", padx=(0, 5))

            ent_code = ctk.CTkEntry(code_compound_frame, width=115, height=30, font=("Arial", 12, "bold"), text_color="#1e293b", fg_color="#f1f5f9")
            ent_code.insert(0, prod["code"])
            ent_code.configure(state="normal") 
            ent_code.pack(side="left", expand=True, fill="x")

            price_compound_frame = ctk.CTkFrame(row, fg_color="transparent", width=160, height=45)
            price_compound_frame.pack(side="left", padx=5)
            price_compound_frame.pack_propagate(False)

            chk_price_show = ctk.CTkCheckBox(price_compound_frame, text="", width=20, height=20)
            chk_price_show.select() 
            chk_price_show.pack(side="left", padx=(0, 5))

            ent_price = ctk.CTkEntry(price_compound_frame, width=110, height=30, font=("Arial", 12, "bold"), justify="center", text_color="#1e293b", fg_color="#f1f5f9")
            ent_price.insert(0, prod.get("price", ""))
            ent_price.configure(state="normal")
            ent_price.pack(side="left", expand=True, fill="x")

            qty_frame = ctk.CTkFrame(row, fg_color="transparent", width=65, height=45)
            qty_frame.pack(side="left", padx=5)
            qty_frame.pack_propagate(False)
            ent_qty = ctk.CTkEntry(qty_frame, width=55, height=30, placeholder_text="1", font=("Arial", 12), justify="center")
            ent_qty.pack(expand=True)

            # --- DUPLICATE INSTANCE BUTTON ---
            dup_frame = ctk.CTkFrame(row, fg_color="transparent", width=80, height=45)
            dup_frame.pack(side="left", padx=5)
            dup_frame.pack_propagate(False)

            btn_dup = ctk.CTkButton(
                dup_frame, 
                text="➕ Dup", 
                width=65, 
                height=28, 
                fg_color="#3b82f6", 
                hover_color="#2563eb",
                font=("Arial", 11, "bold"),
                command=lambda c=cat, p=prod: self.duplicate_product_instance(c, p)
            )
            btn_dup.pack(expand=True)

            # --- REMOVE BUTTON: SIRF DUPLICATE ROWS PAR NAZAR AATA HAI ---
            # ASLI FIX (issue #2): pehle duplicate add hone ke baad usay
            # dobara dhoondhna, edit ya delete karna mushkil tha. Ab yeh
            # button sirf duplicate row par dikhta hai aur sirf usi ek
            # duplicate ko hatata hai — original product/catalog bilkul
            # mehfooz rehta hai.
            if is_dup:
                rm_frame = ctk.CTkFrame(row, fg_color="transparent", width=90, height=45)
                rm_frame.pack(side="left", padx=5)
                rm_frame.pack_propagate(False)

                btn_rm = ctk.CTkButton(
                    rm_frame,
                    text="✕ Remove",
                    width=80,
                    height=28,
                    fg_color="#ef4444",
                    hover_color="#dc2626",
                    font=("Arial", 11, "bold"),
                    command=lambda c=cat, pid=prod.get("id"): self.remove_duplicate_instance(c, pid)
                )
                btn_rm.pack(expand=True)

            # --- DYNAMIC TOGGLE LOGIC FOR ENTRY BOXES ---
            def toggle_code_entry(e_code=ent_code, c_show=chk_code_show):
                if c_show.get() == 1:
                    e_code.configure(state="normal", fg_color="#f1f5f9", text_color="#1e293b")
                else:
                    e_code.configure(state="disabled", fg_color="#e2e8f0", text_color="#94a3b8")

            def toggle_price_entry(e_price=ent_price, e_qty=ent_qty, p_show=chk_price_show):
                if p_show.get() == 1:
                    e_price.configure(state="normal", fg_color="#f1f5f9", text_color="#1e293b")
                    e_qty.configure(state="normal", fg_color="#ffffff")
                else:
                    e_price.configure(state="disabled", fg_color="#e2e8f0", text_color="#94a3b8")
                    e_qty.configure(state="disabled", fg_color="#e2e8f0")

            chk_code_show.configure(command=toggle_code_entry)
            chk_price_show.configure(command=toggle_price_entry)

            # ONLY ONE CLEAN APPEND HERE
            self.grid_inputs[cat].append({
                "id": prod.get("id", ""),
                "name": prod["name"],
                "specs": prod.get("specs", ""),
                "code": prod["code"],
                "checkbox": chk,
                "select_order": chk_order_var,
                "code_entry": ent_code,
                "code_show": chk_code_show,
                "qty": ent_qty,
                "price": ent_price,
                "price_show": chk_price_show,
                "row_frame": row,
                # ASLI FIX: is_duplicate_instance/source_name yahan tag kar
                # dete hain taake "🔁 Manage Duplicates" popup ko har baar
                # inventory_db dobara scan na karna pade
                "is_duplicate_instance": is_dup,
                "source_name": prod.get("source_name", "")
            })
    def duplicate_product_instance(self, cat, prod_data):
        try:
            # 1. PRICE AUR QUANTITY DIALOG POPUP
            dialog = ctk.CTkToplevel(self)
            dialog.title("Duplicate Configuration")
            dialog.geometry("380x250")
            dialog.resizable(False, False)
            dialog.grab_set()
            
            dialog.update_idletasks()
            x = self.winfo_x() + (self.winfo_width() // 2) - 190
            y = self.winfo_y() + (self.winfo_height() // 2) - 125
            dialog.geometry(f"+{x}+{y}")

            ctk.CTkLabel(dialog, text=f"Duplicate: {prod_data['name']}", font=("Arial", 13, "bold"), wraplength=340).pack(pady=(15, 10))

            p_frame = ctk.CTkFrame(dialog, fg_color="transparent")
            p_frame.pack(fill="x", padx=20, pady=5)
            ctk.CTkLabel(p_frame, text="Unit Price (€):", width=110, anchor="w").pack(side="left")
            ent_p = ctk.CTkEntry(p_frame, width=200)
            ent_p.insert(0, str(prod_data.get("price", "")))
            ent_p.pack(side="right")

            q_frame = ctk.CTkFrame(dialog, fg_color="transparent")
            q_frame.pack(fill="x", padx=20, pady=5)
            ctk.CTkLabel(q_frame, text="Quantity:", width=110, anchor="w").pack(side="left")
            ent_q = ctk.CTkEntry(q_frame, width=200)
            ent_q.insert(0, "1")
            ent_q.pack(side="right")

            result = {"submitted": False, "price": "", "qty": "1"}

            def on_confirm():
                result["submitted"] = True
                result["price"] = ent_p.get().strip()
                result["qty"] = ent_q.get().strip()
                dialog.destroy()

            ctk.CTkButton(dialog, text="Add Duplicate", fg_color="#22c55e", hover_color="#16a34a", command=on_confirm).pack(pady=15)
            
            self.wait_window(dialog)

            if not result["submitted"]:
                return

            # 2. CURRENT UI STATES SAVE KAREIN (taake neeche poori category
            # dobara populate hone ke baad screen ki baaki saari rows ki
            # current ticks/values wapis load ho sakein)
            saved_states = []
            max_order = 0
            if cat in self.grid_inputs:
                for item in self.grid_inputs[cat]:
                    raw_order = item.get("select_order")
                    if hasattr(raw_order, 'get'):
                        order_val = raw_order.get()
                    else:
                        try:
                            order_val = int(raw_order) if raw_order is not None else 0
                        except (ValueError, TypeError):
                            order_val = 0

                    if order_val > max_order:
                        max_order = order_val

                    # ASLI FIX: pehle yahan galat dictionary keys use ho rahi
                    # thin ("ent_price", "ent_qty", "ent_code", "chk_price_show",
                    # "chk_code_show") jo populate_category_rows me kabhi
                    # banayi hi nahi jaati — is liye har .get() hamesha None
                    # deta tha aur neeche ka "if ent_p_obj:" block kabhi chalta
                    # hi nahi tha. Sahi keys hain: "price", "qty", "code_entry",
                    # "price_show", "code_show".
                    saved_states.append({
                        "id": item.get("id"),
                        "sel": item["checkbox"].get() if item.get("checkbox") else 0,
                        "code_show": item["code_show"].get() if item.get("code_show") else 1,
                        "code_val": item["code_entry"].get() if item.get("code_entry") else "",
                        "price_show": item["price_show"].get() if item.get("price_show") else 1,
                        "price_val": item["price"].get() if item.get("price") else "",
                        "qty_val": item["qty"].get() if item.get("qty") else "1",
                        "select_order_val": order_val
                    })

            # 3. NAYA DUPLICATE INSTANCE BANAYEIN
            # "is_duplicate_instance" flag se mark karte hain taake: (a) yeh
            # screen par 🔁 badge + halke neele background ke saath saaf
            # pehchana ja sake, (b) master product catalog file par kabhi
            # permanently save na ho (dekhein save_inventory_to_db calls).
            # POSITION: chahe "Dup" button kisi bhi product par dabaya jaye
            # (upar ho ya kahin bhi), duplicate hamesha list ke SABSE AAKHIR
            # (bottom) me hi add hota hai — bilkul waisa jaisa pehle kaam
            # kar raha tha — taake yeh humesha current selection-sequence ke
            # "aakhir" me aaye, na ke original product ke turant neeche.
            new_id = f"{prod_data.get('id', 'prod')}_copy_{int(datetime.now().timestamp() * 1000)}"
            new_prod = prod_data.copy()
            new_prod["id"] = new_id
            new_prod["is_duplicate_instance"] = True
            new_prod["source_id"] = prod_data.get("id")
            new_prod["source_name"] = prod_data.get("name", "")

            if cat in self.inventory_db:
                self.inventory_db[cat].append(new_prod)
                
            # 4. RE-POPULATE CATEGORY ROWS
            self.populate_category_rows(cat)

            # 5. RESTORE STATES & APPLY CUSTOM VALUES (ab sahi keys ke saath)
            if cat in self.grid_inputs:
                for item in self.grid_inputs[cat]:
                    for state in saved_states:
                        if item.get("id") == state["id"]:
                            if item.get("checkbox"):
                                if state["sel"] == 1: item["checkbox"].select()
                                else: item["checkbox"].deselect()
                            
                            if item.get("code_show"):
                                if state["code_show"] == 1: item["code_show"].select()
                                else: item["code_show"].deselect()
                            
                            if item.get("price_show"):
                                if state["price_show"] == 1: item["price_show"].select()
                                else: item["price_show"].deselect()
                            
                            if item.get("code_entry"):
                                item["code_entry"].delete(0, "end")
                                item["code_entry"].insert(0, state["code_val"])
                                
                            if item.get("price"):
                                item["price"].delete(0, "end")
                                item["price"].insert(0, state["price_val"])
                                
                            if item.get("qty"):
                                item["qty"].delete(0, "end")
                                item["qty"].insert(0, state["qty_val"])
                                
                            if item.get("select_order") is not None and hasattr(item["select_order"], 'set'):
                                item["select_order"].set(state["select_order_val"])
                            elif item.get("select_order") is not None:
                                item["select_order"] = ctk.IntVar(value=state["select_order_val"])
                            break
                    
                    # Naye Duplicate Instance ko Configure Karein
                    if item.get("id") == new_id:
                        if item.get("price"):
                            item["price"].delete(0, "end")
                            item["price"].insert(0, result["price"])
                            
                        if item.get("qty"):
                            item["qty"].delete(0, "end")
                            item["qty"].insert(0, result["qty"])
                        
                        if item.get("checkbox"):
                            item["checkbox"].select()
                            
                        next_order = max_order + 1
                        if item.get("select_order") is not None and hasattr(item["select_order"], 'set'):
                            item["select_order"].set(next_order)
                        else:
                            item["select_order"] = ctk.IntVar(value=next_order)

            # Naye duplicate ki tick/price/qty ko turant currently active
            # line ke andar save kar dein, warna woh line switch hote hi
            # (ya save/PDF banate waqt) is naye row ko miss kar sakti hai
            self.on_checkbox_toggle()

            # 6. SUCCESS CONFIRMATION POPUP
            messagebox.showinfo(
                "Duplicate Success", 
                f"✅ A duplicate of '{prod_data['name']}' has been added to the end (bottom) of the list!\n\n"
                f"• Price: €{result['price']}\n• Quantity: {result['qty']}\n• Sequence Order: #{max_order + 1}\n\n"
                f"How to spot it: this row has a 🔁 badge and a light blue background. You can edit the Price/Qty directly in the row, "
                f"and remove it anytime using that row's '✕ Remove' button."
            )

        except Exception as e:
            messagebox.showerror("Error", f"Failed to duplicate item: {str(e)}")

    def remove_duplicate_instance(self, cat, dup_id, ask_confirm=True, refresh_popup=None):
        """Sirf ek specific duplicate line ko hamesha ke liye hata deta hai.
        ASLI FIX: pehle yeh poori category ke saare rows dobara populate kar
        deta tha, jisse screen "poora refresh" jaisa lagta tha aur kisi bhi
        row me abhi type ki gayi (lekin abhi tak save na hui) price/qty
        value zaya ho jaati thi. Ab yeh SIRF isi ek duplicate ki row ko
        surgically destroy karta hai — baaki kisi bhi row/category ko bilkul
        chhoo tak nahi ta. 'Refresh System' button hi ab poora software
        refresh karta hai, yeh function nahi."""
        if ask_confirm and not messagebox.askyesno(
            "Remove Duplicate",
            "Do you want to remove this duplicate line?\n\nThis will only remove this duplicate — the original product and everything else will remain untouched."
        ):
            return

        if cat in self.inventory_db:
            self.inventory_db[cat] = [p for p in self.inventory_db[cat] if p.get("id") != dup_id]

        if cat in self.grid_inputs:
            target_item = None
            for item in self.grid_inputs[cat]:
                if item.get("id") == dup_id:
                    target_item = item
                    break
            if target_item:
                row_frame = target_item.get("row_frame")
                if row_frame:
                    try:
                        row_frame.destroy()
                    except Exception:
                        pass
                self.grid_inputs[cat].remove(target_item)

        # Agar yeh duplicate kisi custom Line Heading ke andar tick/save tha,
        # to wahan se bhi uska reference saaf kar dein taake woh kabhi PDF/
        # Excel me orphan entry ki tarah na aaye
        if hasattr(self, 'all_heading_cards'):
            for card in self.all_heading_cards:
                try:
                    if card.winfo_exists() and hasattr(card, 'saved_selections'):
                        card.saved_selections.pop(dup_id, None)
                except Exception:
                    pass

        # Agar "Manage Duplicates" popup se call hua ho, to usay bhi turant
        # refresh kar dein taake list se yeh entry gayab ho jaye
        if refresh_popup:
            try:
                refresh_popup()
            except Exception:
                pass

    def open_manage_duplicates_popup(self):
        """Sirf CURRENTLY ACTIVE line ke duplicates ek jagah list karta hai
        (har line ke apne products hote hain, is liye duplicates bhi sirf
        usi line ke dikhne chahiye) — taake baar baar scroll karke dhoondhna
        na pade. Yahin se price/qty edit aur remove dono ho sakte hain."""
        popup = ctk.CTkToplevel(self)
        popup.title("Manage Duplicates")
        popup.geometry("720x480")
        popup.grab_set()

        popup.update_idletasks()
        x = self.winfo_x() + (self.winfo_width() // 2) - 360
        y = self.winfo_y() + (self.winfo_height() // 2) - 240
        popup.geometry(f"+{x}+{y}")

        header = ctk.CTkLabel(popup, text="🔁 Duplicate Lines", font=("Arial", 16, "bold"))
        header.pack(pady=(15, 5))

        sub_header = ctk.CTkLabel(
            popup,
            text="Edit the price/qty from here, or remove any duplicate.",
            font=("Arial", 11), text_color="#64748b"
        )
        sub_header.pack(pady=(0, 10))

        list_frame = ctk.CTkScrollableFrame(popup, width=680, height=360)
        list_frame.pack(fill="both", expand=True, padx=15, pady=(0, 10))

        def render_list():
            for w in list_frame.winfo_children():
                w.destroy()

            # ASLI FIX: pehle SAARI lines ke duplicates ek sath mix hoke
            # dikh rahe the, jisse pata nahi chalta tha ke kis line ka hai.
            # Har line ke checkbox-ticks sirf usi line ke active hone par
            # screen par dikhte hain (baaki lines switch hote hi reset ho
            # jati hain), is liye "abhi jo checkbox ON hai" = "isi active
            # line ka duplicate hai" — yehi filter yahan use kiya hai.
            active_card = getattr(self, 'active_heading_card', None)
            if active_card is not None:
                line_label = f"{getattr(active_card, 'line_no', '')} {getattr(active_card, 'raw_title', '')}".strip()
                header.configure(text=f"🔁 Duplicates in: {line_label}" if line_label else "🔁 Duplicates in current line")
            else:
                header.configure(text="🔁 Duplicate Lines")

            dup_rows = []
            if hasattr(self, 'grid_inputs'):
                for cat, items in self.grid_inputs.items():
                    for item in items:
                        if item.get("is_duplicate_instance") and item.get("checkbox") and item["checkbox"].get() == 1:
                            dup_rows.append((cat, item))

            if not dup_rows:
                lbl = ctk.CTkLabel(list_frame, text="There are no duplicates in this line yet.", font=("Arial", 12), text_color="#64748b")
                lbl.pack(pady=30)
                return

            for cat, item in dup_rows:
                row = ctk.CTkFrame(list_frame, fg_color="#f1f5f9", corner_radius=8)
                row.pack(fill="x", padx=5, pady=5)

                info_txt = f"🔁 {item.get('name', '')}"
                if item.get("source_name"):
                    info_txt += f"   (of: {item.get('source_name')})"
                info_txt += f"   —  [{cat}]"

                ctk.CTkLabel(row, text=info_txt, font=("Arial", 12, "bold"), anchor="w", wraplength=280).pack(side="left", padx=10, pady=8)

                ctk.CTkLabel(row, text="Price:", font=("Arial", 11)).pack(side="left", padx=(10, 2))
                ent_p = ctk.CTkEntry(row, width=90)
                ent_p.insert(0, item["price"].get() if item.get("price") else "")
                ent_p.pack(side="left", padx=(0, 10))

                ctk.CTkLabel(row, text="Qty:", font=("Arial", 11)).pack(side="left", padx=(0, 2))
                ent_q = ctk.CTkEntry(row, width=50)
                ent_q.insert(0, item["qty"].get() if item.get("qty") else "1")
                ent_q.pack(side="left", padx=(0, 10))

                def apply_change(cat=cat, item=item, ent_p=ent_p, ent_q=ent_q):
                    if item.get("price"):
                        item["price"].delete(0, "end")
                        item["price"].insert(0, ent_p.get().strip())
                    if item.get("qty"):
                        item["qty"].delete(0, "end")
                        item["qty"].insert(0, ent_q.get().strip())
                    # Turant active line me save kar dein taake PDF/Excel me
                    # yeh naya price/qty foran reflect ho
                    self.on_checkbox_toggle()
                    messagebox.showinfo("Updated", "Price/Qty updated successfully.", parent=popup)

                ctk.CTkButton(
                    row, text="✓ Apply", width=70, height=26, fg_color="#22c55e", hover_color="#16a34a",
                    font=("Arial", 11, "bold"), command=apply_change
                ).pack(side="left", padx=5)

                ctk.CTkButton(
                    row, text="✕ Remove", width=80, height=26, fg_color="#ef4444", hover_color="#dc2626",
                    font=("Arial", 11, "bold"),
                    command=lambda cat=cat, dup_id=item.get("id"): self.remove_duplicate_instance(cat, dup_id, refresh_popup=render_list)
                ).pack(side="left", padx=(5, 10))

        render_list()

        ctk.CTkButton(popup, text="Close", width=100, command=popup.destroy).pack(pady=(0, 15))

    def add_item_global_trigger(self):
        cat = self.tabview_sub.get()
        if cat: self.add_item_trigger(cat)

    def edit_item_global_trigger(self):
        cat = self.tabview_sub.get()
        if cat: self.edit_item_trigger(cat)

    def delete_item_global_trigger(self):
        cat = self.tabview_sub.get()
        if cat: self.delete_item_trigger(cat)

    def filter_matrix_by_product_name(self, *args):
        active_cat = self.tabview_sub.get()
        if not active_cat or active_cat not in self.grid_inputs:
            return
    
            
        search_query = self.matrix_search_var.get().strip().lower()
        search_tokens = search_query.split()
        
        # STEP 1: Pehle tamam frames ko screen se forget (hide) kar dein
        for item in self.grid_inputs[active_cat]:
            item["row_frame"].pack_forget()

        # STEP 2: Ab original list order ke hisab se wapis pack karein
        for item in self.grid_inputs[active_cat]:
            product_name = item["name"].lower()
            product_code = item["code"].lower()
            
            matches_all_tokens = True
            for token in search_tokens:
                if token not in product_name and token not in product_code:
                    matches_all_tokens = False
                    break
            
            if matches_all_tokens:
                item["row_frame"].pack(fill="x", padx=5, pady=4)

    def reset_matrix_search_on_tab_change(self):
        self.matrix_search_var.set("")
        # Speed fix: lazy-loading — is tab ki product rows agar abhi tak
        # nahi bani (kyunke pehli baar khula gaya hai) to yahan bana dein.
        try:
            current_cat = self.tabview_sub.get()
            if (hasattr(self, '_populated_cats')
                    and current_cat in self.sub_scroll_frames
                    and current_cat not in self._populated_cats):
                self.populate_category_rows(current_cat)
                self._populated_cats.add(current_cat)
        except Exception:
            pass

    def add_item_trigger(self, cat):
            dialog = CustomInputDialog(self, f"Add Product to {cat}", [
                ("Item Nomenclature Name", "name"),
                ("Item Reference Code (e.g., BR-NEW)", "code"),
                ("Base Price (€)", "price"),
                ("Detailed Specification Description (PDF Only)", "specs")
            ])
            self.wait_window(dialog)
            
            # 1. Check karein agar user ne cancel ya cross (X) kiya ho
            if not dialog.result:
                return
                
            new_name = dialog.result.get("name", "").strip()
            new_code = dialog.result.get("code", "").strip()

            # 2. Validation: Name khali nahi hona chahiye
            if not new_name:
                messagebox.showerror("Validation Error", "Product Name is required! An empty product cannot be added.")
                return

            # 3. STRICT DUPLICATION CHECK (Case-insensitive aur Whitespace-proof)
            if cat in self.inventory_db:
                for item in self.inventory_db[cat]:
                    existing_name = str(item.get("name", "")).strip().lower()
                    existing_code = str(item.get("code", "")).strip().lower()
                    
                    # Name duplicate check
                    if existing_name == new_name.lower():
                        messagebox.showerror("Duplicate Product", f"Problem! '{new_name}' Product already exist.")
                        return  # Code yahi ruk jayega, aage nahi barhega

                    # Code duplicate check (Agar code field khali na ho)
                    if new_code and existing_code == new_code.lower():
                        messagebox.showerror("Duplicate Code", f"Problem! '{new_code}' Product already exist.")
                        return  # Code yahi ruk jayega, aage nahi barhega

            # 4. AGAR APNE CHECKS PASS HO JAYEN TOU SAVE KAREIN
            new_item = {
                "id": f"gen.{datetime.now().timestamp()}",
                "name": new_name,
                "code": new_code,
                "price": dialog.result.get("price", "").strip(),
                "specs": dialog.result.get("specs", "").strip()
            }
            
            if cat not in self.inventory_db:
                self.inventory_db[cat] = []
                
            self.inventory_db[cat].append(new_item)
            
            # State aur GUI synchronize karein
            save_inventory_to_db(self.inventory_db)
            self.populate_category_rows(cat)
            
            messagebox.showinfo("Matrix Synchronized", f"'{new_name}' safely injected into {cat} database records.")

    def edit_item_trigger(self, cat):
        selected_index = None
        for idx, item_ui in enumerate(self.grid_inputs[cat]):
            if item_ui["checkbox"].get() == 1:
                if selected_index is not None:
                    messagebox.showwarning("Multiple Selection", "Please check/select only ONE item to edit at a time.")
                    return
                selected_index = idx
                
        if selected_index is None:
            messagebox.showwarning("Selection Empty", "Please check/select the checkbox of the single item you want to edit.")
            return

        current_item = self.inventory_db[cat][selected_index]
        initial_vals = {
            "name": current_item["name"],
            "code": current_item["code"],
            "price": current_item.get("price", ""),
            "specs": current_item["specs"]
        }

        dialog = CustomInputDialog(self, f"Modify Record inside {cat}", [
            ("Item Nomenclature Name", "name"),
            ("Item Reference Code", "code"),
            ("Base Price (€)", "price"),
            ("Detailed Specification Description", "specs")
        ], initial_values=initial_vals)
        self.wait_window(dialog)

        if dialog.result:
            self.inventory_db[cat][selected_index]["name"] = dialog.result["name"]
            self.inventory_db[cat][selected_index]["code"] = dialog.result["code"]
            self.inventory_db[cat][selected_index]["price"] = dialog.result["price"]
            self.inventory_db[cat][selected_index]["specs"] = dialog.result["specs"]
            
            save_inventory_to_db(self.inventory_db)
            self.populate_category_rows(cat)
            messagebox.showinfo("Success", "Core structural specification item records updated successfully.")

    def delete_item_trigger(self, cat):
        items_to_remove = []
        indices_to_remove = []
        
        for idx, item_ui in enumerate(self.grid_inputs[cat]):
            if item_ui["checkbox"].get() == 1:
                items_to_remove.append(item_ui["name"])
                indices_to_remove.append(idx)
                
        if not items_to_remove:
            messagebox.showwarning("Selection Empty", "Please check/select the checkbox of items you intend to erase.")
            return
            
        confirm = messagebox.askyesno("Confirm Purge", f"Are you sure you want to permanently erase the following {len(items_to_remove)} items?\n\n" + ", ".join(items_to_remove))
        if confirm:
            for idx in sorted(indices_to_remove, reverse=True):
                del self.inventory_db[cat][idx]
            save_inventory_to_db(self.inventory_db)
            self.populate_category_rows(cat)
            messagebox.showinfo("Purge Success", "Selected machine models deleted.")

    def setup_dashboard_view(self):
        frame_stats = ctk.CTkFrame(self.tab_dashboard, fg_color="transparent")
        frame_stats.pack(fill="x", padx=15, pady=10)
        
        self.card_quotations = ctk.CTkButton(frame_stats, text="Quotations Mapped\n0 Documents", font=("Arial", 14, "bold"), fg_color="#1f538d", height=70, state="disabled")
        self.card_quotations.pack(side="left", fill="x", expand=True, padx=5)

        self.card_proformas = ctk.CTkButton(frame_stats, text="Proformas Executed\n0 Documents", font=("Arial", 14, "bold"), fg_color="#228B22", height=70, state="disabled")
        self.card_proformas.pack(side="left", fill="x", expand=True, padx=5)

        log_header_panel = ctk.CTkFrame(self.tab_dashboard, fg_color="transparent", height=45)
        log_header_panel.pack(fill="x", padx=15, pady=(5, 0))

        self.db_search_var = ctk.StringVar()
        self.db_search_var.trace_add("write", self.filter_dashboard_logs_realtime)

        # FIX: dashboard search box par bhi hamesha nazar aane wala icon aur
        # "Search History" label laga diya hai taake saaf pata chale ke yeh
        # purane records dhoondhne ka box hai.
        search_wrap_db = ctk.CTkFrame(
            log_header_panel,
            fg_color="#ffffff",
            corner_radius=10,
            border_width=2,
            border_color="#228B22",
            height=35,
            width=360
        )
        search_wrap_db.pack(side="right", padx=10, pady=5)
        search_wrap_db.pack_propagate(False)

        ctk.CTkLabel(
            search_wrap_db, text="🔍", font=("Segoe UI Emoji", 14),
            text_color="#228B22", width=22
        ).pack(side="left", padx=(10, 2))

        ent_db_search = ctk.CTkEntry(
            search_wrap_db,
            textvariable=self.db_search_var,
            placeholder_text="Search History",
            height=29,
            corner_radius=6,
            border_width=0,
            fg_color="transparent",
            placeholder_text_color="#94a3b8",
            text_color="#0f172a",
            font=("Segoe UI", 12)
        )
        ent_db_search.pack(side="left", fill="x", expand=True, padx=(2, 10))

        # --- DATE FILTER PANEL: Year / Month dropdowns + From-To calendar range ---
        # These can be combined freely: pick just a Year, just a Month, both
        # together, a From-To date range, or any mix of the above.
        date_filter_panel = ctk.CTkFrame(self.tab_dashboard, fg_color="#f8fafc", corner_radius=8,
                                          border_width=1, border_color="#dbe3ea")
        date_filter_panel.pack(fill="x", padx=15, pady=(8, 0))

        ctk.CTkLabel(date_filter_panel, text="📅 Filter:", font=("Arial", 12, "bold"),
                     text_color="#1f538d").pack(side="left", padx=(10, 8), pady=8)

        ctk.CTkLabel(date_filter_panel, text="Year", font=("Arial", 11)).pack(side="left", padx=(4, 3))
        self.filter_year_var = ctk.StringVar(value="All Years")
        current_year = datetime.now().year
        default_years = [str(y) for y in range(current_year - 5, current_year + 6)]
        self.combo_filter_year = ctk.CTkComboBox(
            date_filter_panel, values=["All Years"] + default_years, variable=self.filter_year_var,
            width=105, command=lambda *_: self.apply_dashboard_filters()
        )
        self.combo_filter_year.pack(side="left", padx=(0, 10))

        ctk.CTkLabel(date_filter_panel, text="Month", font=("Arial", 11)).pack(side="left", padx=(4, 3))
        self.filter_month_var = ctk.StringVar(value="All Months")
        month_values = ["All Months", "January", "February", "March", "April", "May", "June",
                         "July", "August", "September", "October", "November", "December"]
        self.combo_filter_month = ctk.CTkComboBox(
            date_filter_panel, values=month_values, variable=self.filter_month_var,
            width=130, command=lambda *_: self.apply_dashboard_filters()
        )
        self.combo_filter_month.pack(side="left", padx=(0, 14))

        # From/To: an entry box (read-only, filled by picking a date) plus a
        # 📅 button that opens a real calendar popup — no external library
        # required. Leaving a box empty means that side of the range is unused.
        ctk.CTkLabel(date_filter_panel, text="From", font=("Arial", 11)).pack(side="left", padx=(4, 3))
        self.date_from_var = ctk.StringVar()
        self.date_from_picker = ctk.CTkEntry(date_filter_panel, textvariable=self.date_from_var,
                                              placeholder_text="dd.mm.yyyy", width=100)
        self.date_from_picker.pack(side="left", padx=(0, 3))
        ctk.CTkButton(date_filter_panel, text="📅", width=32, font=("Arial", 12),
                      fg_color="#1f538d", hover_color="#153a63",
                      command=lambda: MiniCalendarPopup(self, self.date_from_picker, on_pick=self.apply_dashboard_filters)
                      ).pack(side="left", padx=(0, 12))

        ctk.CTkLabel(date_filter_panel, text="To", font=("Arial", 11)).pack(side="left", padx=(4, 3))
        self.date_to_var = ctk.StringVar()
        self.date_to_picker = ctk.CTkEntry(date_filter_panel, textvariable=self.date_to_var,
                                            placeholder_text="dd.mm.yyyy", width=100)
        self.date_to_picker.pack(side="left", padx=(0, 3))
        ctk.CTkButton(date_filter_panel, text="📅", width=32, font=("Arial", 12),
                      fg_color="#1f538d", hover_color="#153a63",
                      command=lambda: MiniCalendarPopup(self, self.date_to_picker, on_pick=self.apply_dashboard_filters)
                      ).pack(side="left", padx=(0, 14))

        ctk.CTkButton(date_filter_panel, text="Apply Filter", width=100, font=("Arial", 11, "bold"),
                      fg_color="#228B22", hover_color="#006400",
                      command=self.apply_dashboard_filters).pack(side="left", padx=(0, 6), pady=8)
        ctk.CTkButton(date_filter_panel, text="Clear Filters", width=100, font=("Arial", 11, "bold"),
                      fg_color="#94a3b8", hover_color="#64748b",
                      command=self.clear_dashboard_filters).pack(side="left", padx=(0, 10), pady=8)

        self.history_filters = ctk.CTkTabview(self.tab_dashboard, height=400)
        self.history_filters.pack(fill="both", expand=True, padx=15, pady=5)
        
        tab_q_log = self.history_filters.add("📋 Quotation Records Log")
        tab_p_log = self.history_filters.add("📜 Proforma Invoice Log")

        style = ttk.Style()
        style.theme_use("clam")
        style.configure("Treeview", background="white", fieldbackground="white", rowheight=32, font=("Arial", 12))
        style.configure("Treeview.Heading", background="#f2f5f9", font=("Arial", 12, "bold"), anchor="w")
        style.map("Treeview", background=[("selected", "#b3d1ff")], foreground=[("selected", "black")])

        self.tree_q = ttk.Treeview(tab_q_log, columns=("Ref", "Client", "Project", "Date", "Month", "Amount"), show="headings")
        self.setup_tree_columns(self.tree_q)
        self.tree_q.pack(fill="both", expand=True, padx=5, pady=5)
        # NAYA: Double-click ek record par → usay load karke turant uska
        # PDF preview khol deta hai. Ab dashboard se seedha dekhna ho to
        # alag se "Load" + "Preview" button click karne ki zaroorat nahi.
        self.tree_q.bind("<Double-1>", self.on_dashboard_record_double_click)

        self.tree_p = ttk.Treeview(tab_p_log, columns=("Ref", "Client", "Project", "Date", "Month", "Amount"), show="headings")
        self.setup_tree_columns(self.tree_p)
        self.tree_p.pack(fill="both", expand=True, padx=5, pady=5)
        self.tree_p.bind("<Double-1>", self.on_dashboard_record_double_click)

        frame_actions = ctk.CTkFrame(self.tab_dashboard, fg_color="transparent", height=55)
        frame_actions.pack(fill="x", padx=15, pady=10)
        frame_actions.pack_propagate(False)  # keep this row's height fixed so the buttons can never get squeezed/stretched

        ctk.CTkButton(frame_actions, text="🔄 Refresh Logs", font=("Arial", 12, "bold"), width=140, height=38, command=self.refresh_dashboard_data).pack(side="left", padx=5, pady=5)
        ctk.CTkButton(frame_actions, text="📂 Load Selected Revision Record", font=("Arial", 13, "bold"), height=38, fg_color="#ffcc00", text_color="black", hover_color="#e6b800", command=self.load_selected_record).pack(side="left", padx=5, pady=5)
        ctk.CTkButton(frame_actions, text="🗑️ Delete Selected Record", font=("Arial", 13, "bold"), height=38, fg_color="#cc3333", text_color="white", hover_color="#992222", command=self.delete_selected_record).pack(side="left", padx=5, pady=5)

        self.cached_q_logs = []
        self.cached_p_logs = []
        self.refresh_dashboard_data()

    def setup_tree_columns(self, tree):
        tree.heading("Ref", text="  Reference / Revision Code", anchor="w")
        tree.heading("Client", text="  Client Factory Name", anchor="w")
        tree.heading("Project", text="  Project Module", anchor="w")
        tree.heading("Date", text="  Timestamp", anchor="w")
        tree.heading("Month", text="  Tracking Month", anchor="w")
        tree.heading("Amount", text="  Total Value", anchor="w")

        tree.column("Ref", width=180, anchor="w")
        tree.column("Client", width=260, anchor="w")
        tree.column("Project", width=160, anchor="w")
        tree.column("Date", width=120, anchor="w")
        tree.column("Month", width=130, anchor="w")
        tree.column("Amount", width=130, anchor="w")

    def trigger_system_logout(self):
        ans = messagebox.askyesno("Confirm Logout", "Are you sure you want to log out and lock the active workspace?")
        if ans:
            self.app_loaded = False 
            if self.sidebar:
                self.sidebar.grid_forget()
                self.sidebar.destroy()
            if self.main_workspace:
                self.main_workspace.grid_forget()
                self.main_workspace.destroy()
            self.sidebar = None
            self.main_workspace = None
            self.login_view.pack(expand=True, fill="both")
            self.login_view.clear_fields()
            self.login_view.resize_background()
            messagebox.showinfo("Session Terminated", "Workspace safely locked.")

    def get_parsed_data(self):
        grouped = {}
        has_prices = False
        grand_total = 0.0

        # 1. Pehle active line ki current UI state ko save karein
        if hasattr(self, 'active_heading_card') and self.active_heading_card:
            try:
                if hasattr(self, 'save_current_line_state'):
                    self.save_current_line_state(self.active_heading_card)
            except Exception:
                pass

        # 2. Saare Line Cards Collect Karein (UI Scan + Memory)
        heading_cards = []

        def recursive_cards_search(widget):
            cards = []
            if getattr(widget, 'is_section_heading', False) or hasattr(widget, 'raw_title') or hasattr(widget, 'line_no'):
                cards.append(widget)
            if hasattr(widget, 'winfo_children'):
                for child in widget.winfo_children():
                    cards.extend(recursive_cards_search(child))
            return cards

        # Sabhi sub_scroll_frames se direct scan karein
        if hasattr(self, 'sub_scroll_frames'):
            for frame in self.sub_scroll_frames.values():
                heading_cards.extend(recursive_cards_search(frame))

        # Permanent registry (self.all_heading_cards) ko HAMESHA merge karein.
        # Live widget-scan kabhi kabhi koi line miss kar sakta hai, is liye
        # yeh registry guarantee deta hai ke jitni bhi lines banayi gayi hain
        # sab yahan shamil hon (sirf woh cards jo abhi tak destroy nahi hui).
        for attr in ['all_heading_cards', 'line_cards', 'heading_cards', 'all_line_cards']:
            cards_list = getattr(self, attr, None)
            if isinstance(cards_list, list) and len(cards_list) > 0:
                for c in cards_list:
                    try:
                        if c.winfo_exists():
                            heading_cards.append(c)
                    except Exception:
                        pass

        # Duplicate cards filter karein
        unique_cards = []
        for c in heading_cards:
            if c not in unique_cards:
                unique_cards.append(c)
        heading_cards = unique_cards

        # 3. Agar Line Headings mil jayein (Line 1.1, Line 1.2...)
        if heading_cards:
            def get_line_sort_key(card):
                line_str = str(getattr(card, 'line_no', '')).strip()
                digits = ''.join(filter(str.isdigit, line_str))
                if digits:
                    return (0, int(digits))
                try:
                    return (1, card.winfo_y())
                except Exception:
                    return (2, 0)

            heading_cards.sort(key=get_line_sort_key)

            for card in heading_cards:
                line_no = getattr(card, 'line_no', '').strip()
                title = getattr(card, 'raw_title', '').strip()
                
                full_title = f"{line_no} {title}".strip() if line_no else title
                if not full_title:
                    full_title = "Section"

                # Check selections dictionary OR Force Sync
                saved_selections = getattr(card, 'saved_selections', {})
                
                # Agar kisi line ki saved_selections empty ho, lekin active grid inputs mein data ho
                if not saved_selections and card == getattr(self, 'active_heading_card', None):
                    if hasattr(self, 'save_current_line_state'):
                        self.save_current_line_state(card)
                        saved_selections = getattr(card, 'saved_selections', {})

                if not saved_selections:
                    continue

                selected_items = []
                for item_id, item_data in saved_selections.items():
                    # Multi-type Selection check (Boolean, Int, String)
                    is_sel = item_data.get("selected", False) or item_data.get("checked", False)
                    if is_sel:
                        prod_info = None
                        if hasattr(self, 'inventory_db') and isinstance(self.inventory_db, dict):
                            for cat_items in self.inventory_db.values():
                                for p in cat_items:
                                    if str(p.get("id")) == str(item_id):
                                        prod_info = p
                                        break
                                if prod_info:
                                    break

                        if not prod_info:
                            continue

                        is_code_vis = item_data.get("code_show", True)
                        is_price_vis = item_data.get("price_show", True)
                        code_val = prod_info.get("code", "") if is_code_vis else ""

                        try:
                            qty_val = int(item_data.get("qty", 1))
                        except (ValueError, TypeError):
                            qty_val = 1

                        try:
                            raw_price = item_data.get("price", prod_info.get("price", 0))
                            price_val = float(str(raw_price).replace(",", "").replace("€", ""))
                        except (ValueError, TypeError):
                            price_val = 0.0

                        if is_price_vis and price_val > 0:
                            has_prices = True
                            tot = price_val * qty_val
                            grand_total += tot
                        else:
                            tot = 0.0
                            price_val = "" if not is_price_vis else price_val
                            tot = "" if not is_price_vis else tot

                        selected_items.append({
                            "id": item_id,
                            "name": prod_info.get("name", ""),
                            "code": code_val,
                            "code_visible": is_code_vis,
                            "specs": prod_info.get("specs", ""),
                            "qty": qty_val if is_price_vis else "",
                            "price": price_val,
                            "total": tot,
                            "price_visible": is_price_vis,
                            "select_order": item_data.get("order", 0)
                        })

                if selected_items:
                    selected_items.sort(key=lambda x: x["select_order"])
                    grouped[full_title] = selected_items

            if grouped:
                return grouped, has_prices, grand_total

        # 4. Fallback (Jab Line Headings na bani hon)
        if hasattr(self, 'grid_inputs'):
            for cat, items in self.grid_inputs.items():
                selected = []
                for item in items:
                    cb_val = item["checkbox"].get() if hasattr(item["checkbox"], 'get') else item["checkbox"]
                    if cb_val == 1:
                        is_code_vis = item["code_show"].get() == 1 if hasattr(item["code_show"], 'get') else item["code_show"]
                        is_price_vis = item["price_show"].get() == 1 if hasattr(item["price_show"], 'get') else item["price_show"]
                        code_val = item["code_entry"].get().strip() if is_code_vis and hasattr(item["code_entry"], 'get') else ""

                        try:
                            qty_val = int(item["qty"].get().strip()) if hasattr(item["qty"], 'get') else int(item["qty"])
                        except (ValueError, AttributeError, TypeError):
                            qty_val = 1

                        try:
                            p_str = item["price"].get().strip() if hasattr(item["price"], 'get') else str(item["price"])
                            price_val = float(p_str.replace(",", "").replace("€", ""))
                        except (ValueError, AttributeError, TypeError):
                            price_val = 0.0

                        if is_price_vis and price_val > 0:
                            has_prices = True
                            tot = price_val * qty_val
                            grand_total += tot
                        else:
                            tot = 0.0
                            if not is_price_vis:
                                price_val = ""
                                tot = ""

                        order_val = item["select_order"].get() if hasattr(item["select_order"], 'get') else item.get("select_order", 0)

                        selected.append({
                            "id": item.get("id", ""),
                            "name": item.get("name", ""),
                            "code": code_val,
                            "code_visible": is_code_vis,
                            "specs": item.get("specs", ""),
                            "qty": qty_val if is_price_vis else "",
                            "price": price_val,
                            "total": tot,
                            "price_visible": is_price_vis,
                            "select_order": order_val
                        })

                if selected:
                    selected.sort(key=lambda x: x["select_order"])
                    grouped[cat] = selected

        return grouped, has_prices, grand_total

        # 4. Universal Fallback
        if hasattr(self, 'grid_inputs'):
            for cat, items in self.grid_inputs.items():
                selected = []
                for item in items:
                    cb_val = item["checkbox"].get() if hasattr(item["checkbox"], 'get') else item["checkbox"]
                    if cb_val == 1:
                        is_code_vis = item["code_show"].get() == 1 if hasattr(item["code_show"], 'get') else item["code_show"]
                        is_price_vis = item["price_show"].get() == 1 if hasattr(item["price_show"], 'get') else item["price_show"]
                        code_val = item["code_entry"].get().strip() if is_code_vis and hasattr(item["code_entry"], 'get') else ""

                        try:
                            qty_val = int(item["qty"].get().strip()) if hasattr(item["qty"], 'get') else int(item["qty"])
                        except (ValueError, AttributeError, TypeError):
                            qty_val = 1

                        try:
                            p_str = item["price"].get().strip() if hasattr(item["price"], 'get') else str(item["price"])
                            price_val = float(p_str.replace(",", "").replace("€", ""))
                        except (ValueError, AttributeError, TypeError):
                            price_val = 0.0

                        if is_price_vis and price_val > 0:
                            has_prices = True
                            tot = price_val * qty_val
                            grand_total += tot
                        else:
                            tot = 0.0
                            if not is_price_vis:
                                price_val = ""
                                tot = ""

                        order_val = item["select_order"].get() if hasattr(item["select_order"], 'get') else item.get("select_order", 0)

                        selected.append({
                            "id": item.get("id", ""),
                            "name": item.get("name", ""),
                            "code": code_val,
                            "code_visible": is_code_vis,
                            "specs": item.get("specs", ""),
                            "qty": qty_val if is_price_vis else "",
                            "price": price_val,
                            "total": tot,
                            "price_visible": is_price_vis,
                            "select_order": order_val
                        })

                if selected:
                    selected.sort(key=lambda x: x["select_order"])
                    grouped[cat] = selected

        return grouped, has_prices, grand_total
    
        # 4. Fallback: Agar Line Headings na banai gayi hon
        if hasattr(self, 'grid_inputs'):
            for cat, items in self.grid_inputs.items():
                selected = []
                for item in items:
                    cb_val = item["checkbox"].get() if hasattr(item["checkbox"], 'get') else item["checkbox"]
                    if cb_val == 1:
                        is_code_vis = item["code_show"].get() == 1 if hasattr(item["code_show"], 'get') else item["code_show"]
                        is_price_vis = item["price_show"].get() == 1 if hasattr(item["price_show"], 'get') else item["price_show"]
                        
                        code_val = item["code_entry"].get().strip() if is_code_vis and hasattr(item["code_entry"], 'get') else ""

                        try:
                            qty_val = int(item["qty"].get().strip()) if hasattr(item["qty"], 'get') else int(item["qty"])
                        except (ValueError, AttributeError, TypeError):
                            qty_val = 1

                        try:
                            p_str = item["price"].get().strip() if hasattr(item["price"], 'get') else str(item["price"])
                            price_val = float(p_str.replace(",", "").replace("€", ""))
                        except (ValueError, AttributeError, TypeError):
                            price_val = 0.0

                        if is_price_vis and price_val > 0:
                            has_prices = True
                            tot = price_val * qty_val
                            grand_total += tot
                        else:
                            tot = 0.0
                            if not is_price_vis:
                                price_val = ""
                                tot = ""

                        order_val = item["select_order"].get() if hasattr(item["select_order"], 'get') else item.get("select_order", 0)

                        selected.append({
                            "id": item.get("id", ""),
                            "name": item.get("name", ""),
                            "code": code_val,
                            "code_visible": is_code_vis,
                            "specs": item.get("specs", ""),
                            "qty": qty_val if is_price_vis else "",
                            "price": price_val,
                            "total": tot,
                            "price_visible": is_price_vis,
                            "select_order": order_val
                        })

                if selected:
                    selected.sort(key=lambda x: x["select_order"])
                    grouped[cat] = selected

        return grouped, has_prices, grand_total

        # 4. Fallback: Jab bilkul koi line card create na hua ho
        if hasattr(self, 'grid_inputs'):
            for cat, items in self.grid_inputs.items():
                selected = []
                for item in items:
                    cb_val = item["checkbox"].get() if hasattr(item["checkbox"], 'get') else item["checkbox"]
                    if cb_val == 1:
                        is_code_vis = item["code_show"].get() == 1 if hasattr(item["code_show"], 'get') else item["code_show"]
                        is_price_vis = item["price_show"].get() == 1 if hasattr(item["price_show"], 'get') else item["price_show"]
                        
                        code_val = item["code_entry"].get().strip() if is_code_vis and hasattr(item["code_entry"], 'get') else ""

                        try:
                            qty_val = int(item["qty"].get().strip()) if hasattr(item["qty"], 'get') else int(item["qty"])
                        except (ValueError, AttributeError, TypeError):
                            qty_val = 1

                        try:
                            p_str = item["price"].get().strip() if hasattr(item["price"], 'get') else str(item["price"])
                            price_val = float(p_str.replace(",", "").replace("€", ""))
                        except (ValueError, AttributeError, TypeError):
                            price_val = 0.0

                        if is_price_vis and price_val > 0:
                            has_prices = True
                            tot = price_val * qty_val
                            grand_total += tot
                        else:
                            tot = 0.0
                            if not is_price_vis:
                                price_val = ""
                                tot = ""

                        order_val = item["select_order"].get() if hasattr(item["select_order"], 'get') else item.get("select_order", 0)

                        selected.append({
                            "id": item.get("id", ""),
                            "name": item.get("name", ""),
                            "code": code_val,
                            "code_visible": is_code_vis,
                            "specs": item.get("specs", ""),
                            "qty": qty_val if is_price_vis else "",
                            "price": price_val,
                            "total": tot,
                            "price_visible": is_price_vis,
                            "select_order": order_val
                        })

                if selected:
                    selected.sort(key=lambda x: x["select_order"])
                    grouped[cat] = selected

        return grouped, has_prices, grand_total

        # 4. Fallback: Agar Line Headings na hon toh General Inventory Grid Scan Karein
        if hasattr(self, 'grid_inputs'):
            for cat, items in self.grid_inputs.items():
                selected = []
                for item in items:
                    cb_val = item["checkbox"].get() if hasattr(item["checkbox"], 'get') else item["checkbox"]
                    if cb_val == 1:
                        is_code_vis = item["code_show"].get() == 1 if hasattr(item["code_show"], 'get') else item["code_show"]
                        is_price_vis = item["price_show"].get() == 1 if hasattr(item["price_show"], 'get') else item["price_show"]
                        
                        code_val = item["code_entry"].get().strip() if is_code_vis and hasattr(item["code_entry"], 'get') else ""

                        try:
                            qty_val = int(item["qty"].get().strip()) if hasattr(item["qty"], 'get') else int(item["qty"])
                        except (ValueError, AttributeError, TypeError):
                            qty_val = 1

                        try:
                            p_str = item["price"].get().strip() if hasattr(item["price"], 'get') else str(item["price"])
                            price_val = float(p_str.replace(",", "").replace("€", ""))
                        except (ValueError, AttributeError, TypeError):
                            price_val = 0.0

                        if is_price_vis and price_val > 0:
                            has_prices = True
                            tot = price_val * qty_val
                            grand_total += tot
                        else:
                            tot = 0.0
                            if not is_price_vis:
                                price_val = ""
                                tot = ""

                        order_val = item["select_order"].get() if hasattr(item["select_order"], 'get') else item.get("select_order", 0)

                        selected.append({
                            "id": item.get("id", ""),
                            "name": item.get("name", ""),
                            "code": code_val,
                            "code_visible": is_code_vis,
                            "specs": item.get("specs", ""),
                            "qty": qty_val if is_price_vis else "",
                            "price": price_val,
                            "total": tot,
                            "price_visible": is_price_vis,
                            "select_order": order_val
                        })

                if selected:
                    selected.sort(key=lambda x: x["select_order"])
                    grouped[cat] = selected

        return grouped, has_prices, grand_total

        # AGAR KOI LINE HEADING NA HO, TO DEFAULT CATEGORIES SE DATA LEIN
        for cat, items in getattr(self, 'grid_inputs', {}).items():
            selected = []
            for item in items:
                if item["checkbox"].get() == 1:
                    is_code_vis = item["code_show"].get() == 1
                    is_price_vis = item["price_show"].get() == 1
                    code_val = item["code_entry"].get().strip() if is_code_vis else ""

                    try:
                        qty_val = int(item["qty"].get().strip())
                    except (ValueError, AttributeError):
                        qty_val = 1

                    try:
                        price_val = float(item["price"].get().strip().replace(",", "").replace("€", ""))
                    except (ValueError, AttributeError):
                        price_val = 0.0

                    if is_price_vis and price_val > 0:
                        has_prices = True
                        tot = price_val * qty_val
                        grand_total += tot
                    else:
                        tot = 0.0
                        if not is_price_vis:
                            price_val = ""
                            tot = ""

                    selected.append({
                        "id": item.get("id", ""),
                        "name": item["name"],
                        "code": code_val,
                        "code_visible": is_code_vis,
                        "specs": item.get("specs", ""),
                        "qty": qty_val if is_price_vis else "",
                        "price": price_val,
                        "total": tot,
                        "price_visible": is_price_vis,
                        "select_order": item["select_order"].get()
                    })

            if selected:
                selected.sort(key=lambda x: x["select_order"])
                grouped[cat] = selected

        return grouped, has_prices, grand_total

        # AGAR KOI LINE HEADING NA HO, TO DEFAULT CATEGORIES SE DATA LEIN
        for cat, items in getattr(self, 'grid_inputs', {}).items():
            selected = []
            for item in items:
                if item["checkbox"].get() == 1:
                    is_code_vis = item["code_show"].get() == 1
                    is_price_vis = item["price_show"].get() == 1
                    code_val = item["code_entry"].get().strip() if is_code_vis else ""

                    try:
                        qty_val = int(item["qty"].get().strip())
                    except (ValueError, AttributeError):
                        qty_val = 1

                    try:
                        price_val = float(item["price"].get().strip().replace(",", "").replace("€", ""))
                    except (ValueError, AttributeError):
                        price_val = 0.0

                    if is_price_vis and price_val > 0:
                        has_prices = True
                        tot = price_val * qty_val
                        grand_total += tot
                    else:
                        tot = 0.0
                        if not is_price_vis:
                            price_val = ""
                            tot = ""

                    selected.append({
                        "id": item.get("id", ""),
                        "name": item["name"],
                        "code": code_val,
                        "code_visible": is_code_vis,
                        "specs": item.get("specs", ""),
                        "qty": qty_val if is_price_vis else "",
                        "price": price_val,
                        "total": tot,
                        "price_visible": is_price_vis,
                        "select_order": item["select_order"].get()
                    })

            if selected:
                selected.sort(key=lambda x: x["select_order"])
                grouped[cat] = selected

        return grouped, has_prices, grand_total

        # AGAR KOI LINE HEADING NA HO, TO DEFAULT CATEGORIES SE DATA LEIN
        for cat, items in getattr(self, 'grid_inputs', {}).items():
            selected = []
            for item in items:
                if item["checkbox"].get() == 1:
                    is_code_vis = item["code_show"].get() == 1
                    is_price_vis = item["price_show"].get() == 1
                    code_val = item["code_entry"].get().strip() if is_code_vis else ""

                    try:
                        qty_val = int(item["qty"].get().strip())
                    except (ValueError, AttributeError):
                        qty_val = 1

                    try:
                        price_val = float(item["price"].get().strip().replace(",", "").replace("€", ""))
                    except (ValueError, AttributeError):
                        price_val = 0.0

                    if is_price_vis and price_val > 0:
                        has_prices = True
                        tot = price_val * qty_val
                        grand_total += tot
                    else:
                        tot = 0.0
                        if not is_price_vis:
                            price_val = ""
                            tot = ""

                    selected.append({
                        "id": item.get("id", ""),
                        "name": item["name"],
                        "code": code_val,
                        "code_visible": is_code_vis,
                        "specs": item.get("specs", ""),
                        "qty": qty_val if is_price_vis else "",
                        "price": price_val,
                        "total": tot,
                        "price_visible": is_price_vis,
                        "select_order": item["select_order"].get()
                    })

            if selected:
                selected.sort(key=lambda x: x["select_order"])
                grouped[cat] = selected

        return grouped, has_prices, grand_total

        # AGAR KOI LINE HEADING NA HO, TO DEFAULT CATEGORIES SE DATA LEIN
        for cat, items in getattr(self, 'grid_inputs', {}).items():
            selected = []
            for item in items:
                if item["checkbox"].get() == 1:
                    is_code_vis = item["code_show"].get() == 1
                    is_price_vis = item["price_show"].get() == 1
                    code_val = item["code_entry"].get().strip() if is_code_vis else ""

                    try:
                        qty_val = int(item["qty"].get().strip())
                    except (ValueError, AttributeError):
                        qty_val = 1

                    try:
                        price_val = float(item["price"].get().strip().replace(",", "").replace("€", ""))
                    except (ValueError, AttributeError):
                        price_val = 0.0

                    if is_price_vis and price_val > 0:
                        has_prices = True
                        tot = price_val * qty_val
                        grand_total += tot
                    else:
                        tot = 0.0
                        if not is_price_vis:
                            price_val = ""
                            tot = ""

                    selected.append({
                        "id": item.get("id", ""),
                        "name": item["name"],
                        "code": code_val,
                        "code_visible": is_code_vis,
                        "specs": item.get("specs", ""),
                        "qty": qty_val if is_price_vis else "",
                        "price": price_val,
                        "total": tot,
                        "price_visible": is_price_vis,
                        "select_order": item["select_order"].get()
                    })

            if selected:
                selected.sort(key=lambda x: x["select_order"])
                grouped[cat] = selected

        return grouped, has_prices, grand_total

    def auto_log_state_to_history(self, grand_total):
        ref_num = self.ent_offer.get().strip() or "UNTITLED_REF"
        doc_type = self.doc_type_var.get()
        current_month_folder = datetime.now().strftime("%B_%Y")
        sub_folder_name = "Quotations" if doc_type == "Quotation" else "Proformas"
        target_directory_path = os.path.join(BASE_HISTORY_DIR, current_month_folder, sub_folder_name)
        if not os.path.exists(target_directory_path):
            os.makedirs(target_directory_path)

        # Screen par jo bhi line abhi active hai, usay save karne se pehle
        # sync kar lein — warna uski latest ticks save hone se reh jaati hain
        if hasattr(self, 'active_heading_card') and self.active_heading_card:
            try:
                self.save_current_line_state(self.active_heading_card)
            except Exception:
                pass

        state_payload = {
            "type": doc_type, "factory": self.ent_factory.get().strip(),
            "address": self.ent_address.get().strip() if doc_type == "Proforma" else "",
            "offer": ref_num, "project": self.ent_project.get().strip() if doc_type == "Proforma" else "",
            "date": datetime.now().strftime("%d.%m.%Y"), "month": current_month_folder,
            "total_value": grand_total, "selections": {}, "lines": []
        }

        # Purana flat snapshot backward-compatibility ke liye rakha hai (agar
        # koi record bina headings ke bhi ho)
        for cat, items in self.grid_inputs.items():
            state_payload["selections"][cat] = [{
                "checked": x["checkbox"].get(), 
                "code": x["code_entry"].get(), 
                "code_show": x["code_show"].get(),
                "qty": x["qty"].get(), 
            
                "price": x["price"].get(),
                "price_show": x["price_show"].get()
            } for x in items]

        # --- ASLI FIX: HAR CUSTOM LINE/HEADING KI APNI SAVED SELECTIONS ---
        # Pehle sirf currently-visible checkbox state save hoti thi, isliye
        # load par sirf woh ek line wapis aati thi. Ab har heading card (jo
        # bhi screen par ho ya kisi aur category tab par) alag se collect
        # karke save karte hain, taake load par POORI structure wapis mile.
        def _collect_heading_cards(widget):
            found = []
            if getattr(widget, 'is_section_heading', False):
                found.append(widget)
            if hasattr(widget, 'winfo_children'):
                for child in widget.winfo_children():
                    found.extend(_collect_heading_cards(child))
            return found

        seen_cards = []
        all_cards = []
        if hasattr(self, 'sub_scroll_frames'):
            for frame in self.sub_scroll_frames.values():
                for card in _collect_heading_cards(frame):
                    if card not in seen_cards:
                        seen_cards.append(card)
                        all_cards.append(card)

        if hasattr(self, 'all_heading_cards'):
            for card in self.all_heading_cards:
                try:
                    if card.winfo_exists() and card not in seen_cards:
                        seen_cards.append(card)
                        all_cards.append(card)
                except Exception:
                    pass

        def _line_sort_key(card):
            line_str = str(getattr(card, 'line_no', '')).strip()
            digits = ''.join(filter(str.isdigit, line_str))
            return int(digits) if digits else 0

        all_cards.sort(key=_line_sort_key)

        for card in all_cards:
            state_payload["lines"].append({
                "line_no": getattr(card, 'line_no', ''),
                "title": getattr(card, 'raw_title', ''),
                "category": getattr(card, 'owner_category', ''),
                "selections": getattr(card, 'saved_selections', {})
            })

        # --- SAVE TARGET DECISION: same existing record, or a brand new one? ---
        # By default, treat this as a NEW record saved into today's month folder.
        filename_prefix = f"{doc_type}_{ref_num}"
        new_file_path = os.path.join(target_directory_path, f"{filename_prefix}.json")
        target_file = new_file_path

        # If this workspace was opened by loading a record from the Dashboard,
        # and the reference number still matches that same record, ask the
        # user whether to overwrite that SAME existing file or save a new one.
        was_loaded_record = (
            self.loaded_record_path
            and self.loaded_record_ref == ref_num
            and os.path.exists(self.loaded_record_path)
        )

        if was_loaded_record:
            choice = messagebox.askyesnocancel(
                "Save Changes",
                f"You are editing an existing record ('{ref_num}').\n\n"
                "Do you want to save these changes to the EXISTING file?\n\n"
                "Yes = Update the existing file\n"
                "No = Save as a new, separate record\n"
                "Cancel = Don't save right now"
            )
            if choice is None:
                # User cancelled the save entirely.
                return None
            elif choice:
                # Update the exact same file this session was loaded from.
                target_file = self.loaded_record_path
                # Keep the record's month/folder consistent with where it
                # actually lives on disk, so the Dashboard still finds it
                # under the correct month tab.
                existing_month_folder = os.path.basename(os.path.dirname(os.path.dirname(target_file)))
                state_payload["month"] = existing_month_folder
            else:
                # User chose to save as a brand new record.
                target_file = new_file_path
                self.loaded_record_path = None
                self.loaded_record_ref = None

        with open(target_file, "w") as f: json.dump(state_payload, f, indent=4)

        # This save now becomes the "currently loaded" record for this session,
        # so saving again without reloading will correctly ask about updating it.
        self.loaded_record_path = target_file
        self.loaded_record_ref = ref_num

        self.refresh_dashboard_data()
        return target_file

    def _parse_record_date(self, date_str):
        """Turns the stored 'dd.mm.yyyy' string into a real date object for
        range comparisons. Returns None if it can't be parsed."""
        try:
            return datetime.strptime(str(date_str).strip(), "%d.%m.%Y").date()
        except (ValueError, TypeError):
            return None

    def refresh_dashboard_data(self):
        for row in self.tree_q.get_children(): self.tree_q.delete(row)
        for row in self.tree_p.get_children(): self.tree_p.delete(row)
        
        self.cached_q_logs.clear()
        self.cached_p_logs.clear()
        
        q_count, p_count = 0, 0
        if os.path.exists(BASE_HISTORY_DIR):
            for month_dir in os.listdir(BASE_HISTORY_DIR):
                full_m_path = os.path.join(BASE_HISTORY_DIR, month_dir)
                if os.path.isdir(full_m_path):
                    q_folder = os.path.join(full_m_path, "Quotations")
                    if os.path.exists(q_folder):
                        for file in os.listdir(q_folder):
                            if file.endswith(".json"):
                                try:
                                    with open(os.path.join(q_folder, file), "r") as f: d = json.load(f)
                                    tokens = f"{d.get('offer','')} {d.get('factory','')} N/A (Quotation) {d.get('date','')} {d.get('month','')}".lower()
                                    vals = (d.get("offer"), d.get("factory"), "N/A (Quotation)", d.get("date"), d.get("month"), f"€ {d.get('total_value',0):,.2f}")
                                    rec_month = d.get("month", "")
                                    m_name, _, m_year = rec_month.partition("_")
                                    self.cached_q_logs.append((vals, tokens, self._parse_record_date(d.get("date", "")), m_name, m_year))
                                    q_count += 1
                                except: pass
                    p_folder = os.path.join(full_m_path, "Proformas")
                    if os.path.exists(p_folder):
                        for file in os.listdir(p_folder):
                            if file.endswith(".json"):
                                try:
                                    with open(os.path.join(p_folder, file), "r") as f: d = json.load(f)
                                    tokens = f"{d.get('offer','')} {d.get('factory','')} {d.get('project','N/A')} {d.get('date','')} {d.get('month','')}".lower()
                                    vals = (d.get("offer"), d.get("factory"), d.get("project", "N/A"), d.get("date"), d.get("month"), f"€ {d.get('total_value',0):,.2f}")
                                    rec_month = d.get("month", "")
                                    m_name, _, m_year = rec_month.partition("_")
                                    self.cached_p_logs.append((vals, tokens, self._parse_record_date(d.get("date", "")), m_name, m_year))
                                    p_count += 1
                                except: pass

        self.apply_dashboard_tree_renders(self.cached_q_logs, self.cached_p_logs)
        self.card_quotations.configure(text=f"Quotations Mapped\n{q_count} Documents")
        self.card_proformas.configure(text=f"Proformas Executed\n{p_count} Documents")

        # Keep the Year dropdown populated with whatever years actually
        # exist in the saved history, so it's never stale.
        if hasattr(self, "combo_filter_year"):
            years_found = sorted({y for *_, y in self.cached_q_logs if y} | {y for *_, y in self.cached_p_logs if y})
            self.combo_filter_year.configure(values=["All Years"] + years_found)

        # Re-apply whatever filters are currently set (search text, year,
        # month, date range) instead of showing the unfiltered full list.
        if hasattr(self, "apply_dashboard_filters"):
            self.apply_dashboard_filters()

    def apply_dashboard_tree_renders(self, q_list, p_list):
        for row_data, *_ in q_list: self.tree_q.insert("", "end", values=row_data)
        for row_data, *_ in p_list: self.tree_p.insert("", "end", values=row_data)

    def filter_dashboard_logs_realtime(self, *args):
        # Kept for the search box's live trace callback; now simply routes
        # into the combined filter engine below.
        self.apply_dashboard_filters()

    def apply_dashboard_filters(self, *args):
        """Combined Dashboard filter: free-text search + Year + Month +
        an optional From/To calendar date range — any mix of these can be
        used together (e.g. Year + Month, or just a date range, or all)."""
        if not hasattr(self, "tree_q"):
            return

        query = self.db_search_var.get().strip().lower()
        query_tokens = query.split()

        year_choice = self.filter_year_var.get() if hasattr(self, "filter_year_var") else "All Years"
        month_choice = self.filter_month_var.get() if hasattr(self, "filter_month_var") else "All Months"

        range_from, range_to = None, None
        if hasattr(self, "date_from_var"):
            from_text = self.date_from_var.get().strip()
            to_text = self.date_to_var.get().strip()
            if from_text:
                range_from = self._parse_record_date(from_text)
            if to_text:
                range_to = self._parse_record_date(to_text)

        def record_matches(record):
            _vals, tokens, dt, rec_month_name, rec_year = record
            if query_tokens and not all(t in tokens for t in query_tokens):
                return False
            if year_choice != "All Years" and rec_year != year_choice:
                return False
            if month_choice != "All Months" and rec_month_name != month_choice:
                return False
            if range_from and (dt is None or dt < range_from):
                return False
            if range_to and (dt is None or dt > range_to):
                return False
            return True

        for item in self.tree_q.get_children(): self.tree_q.delete(item)
        for item in self.tree_p.get_children(): self.tree_p.delete(item)

        filtered_q = [r for r in self.cached_q_logs if record_matches(r)]
        filtered_p = [r for r in self.cached_p_logs if record_matches(r)]
        self.apply_dashboard_tree_renders(filtered_q, filtered_p)

    def clear_dashboard_filters(self):
        """Resets every Dashboard filter (search text, Year, Month, and the
        From/To date range) back to 'show everything'."""
        self.db_search_var.set("")
        if hasattr(self, "filter_year_var"):
            self.filter_year_var.set("All Years")
        if hasattr(self, "filter_month_var"):
            self.filter_month_var.set("All Months")
        if hasattr(self, "date_from_var"):
            self.date_from_var.set("")
            self.date_to_var.set("")
        self.apply_dashboard_filters()

    def get_grouped_data_from_record(self, record_data):
        """Saved history record (JSON) se seedha PDF-ready 'grouped' data
        banata hai — bilkul get_parsed_data() jaisa output, lekin sirf
        record ki apni saved lines se, LIVE workspace widgets ko chuye
        baghair. Yehi cheez Dashboard 'double-click preview' ko current
        kaam disturb kiye baghair kaam karne deti hai."""
        grouped = {}
        has_prices = False
        grand_total = 0.0

        for line in record_data.get("lines", []):
            line_no = str(line.get("line_no", "")).strip()
            title = str(line.get("title", "")).strip()
            full_title = f"{line_no} {title}".strip() or "Section"
            selections = line.get("selections", {}) or {}

            selected_items = []
            for item_id, item_data in selections.items():
                is_sel = item_data.get("selected", False) or item_data.get("checked", False)
                if not is_sel:
                    continue

                prod_info = None
                if isinstance(self.inventory_db, dict):
                    for cat_items in self.inventory_db.values():
                        for p in cat_items:
                            if str(p.get("id")) == str(item_id):
                                prod_info = p
                                break
                        if prod_info:
                            break
                if not prod_info:
                    continue

                is_code_vis = item_data.get("code_show", True)
                is_price_vis = item_data.get("price_show", True)
                code_val = prod_info.get("code", "") if is_code_vis else ""

                try:
                    qty_val = int(item_data.get("qty", 1))
                except (ValueError, TypeError):
                    qty_val = 1

                try:
                    raw_price = item_data.get("price", prod_info.get("price", 0))
                    price_val = float(str(raw_price).replace(",", "").replace("€", ""))
                except (ValueError, TypeError):
                    price_val = 0.0

                if is_price_vis and price_val > 0:
                    has_prices = True
                    tot = price_val * qty_val
                    grand_total += tot
                else:
                    tot = 0.0
                    price_val = "" if not is_price_vis else price_val
                    tot = "" if not is_price_vis else tot

                selected_items.append({
                    "id": item_id, "name": prod_info.get("name", ""),
                    "code": code_val, "code_visible": is_code_vis,
                    "specs": prod_info.get("specs", ""),
                    "qty": qty_val if is_price_vis else "",
                    "price": price_val, "total": tot,
                    "price_visible": is_price_vis,
                    "select_order": item_data.get("order", 0)
                })

            if selected_items:
                selected_items.sort(key=lambda x: x["select_order"])
                grouped[full_title] = selected_items

        # FALLBACK: agar "lines" (heading-based) format me data na mile, to
        # purane FLAT format ("selections": {category: [...]}) se try
        # karein — kaafi purane records isi tareeqe se save hue thay, jab
        # tak heading/line feature use nahi hui thi. Bilkul load_selected_record()
        # ke old-format restore jaisa hi index-matching logic.
        if not grouped:
            flat_selections = record_data.get("selections", {}) or {}
            for cat, saved_list in flat_selections.items():
                cat_items = self.inventory_db.get(cat, []) if isinstance(self.inventory_db, dict) else []
                selected = []
                for prod_info, sv in zip(cat_items, saved_list):
                    if not isinstance(sv, dict):
                        continue
                    checked = sv.get("checked", 0)
                    if checked != 1 and checked is not True:
                        continue

                    is_code_vis = sv.get("code_show", 1) == 1 or sv.get("code_show") is True
                    is_price_vis = sv.get("price_show", 1) == 1 or sv.get("price_show") is True
                    code_val = (sv.get("code") or prod_info.get("code", "")) if is_code_vis else ""

                    try:
                        qty_val = int(str(sv.get("qty", 1)).strip())
                    except (ValueError, TypeError):
                        qty_val = 1

                    try:
                        raw_price = sv.get("price", prod_info.get("price", 0))
                        price_val = float(str(raw_price).replace(",", "").replace("€", ""))
                    except (ValueError, TypeError):
                        price_val = 0.0

                    if is_price_vis and price_val > 0:
                        has_prices = True
                        tot = price_val * qty_val
                        grand_total += tot
                    else:
                        tot = 0.0
                        if not is_price_vis:
                            price_val = ""
                            tot = ""

                    selected.append({
                        "id": prod_info.get("id", ""),
                        "name": prod_info.get("name", ""),
                        "code": code_val,
                        "code_visible": is_code_vis,
                        "specs": prod_info.get("specs", ""),
                        "qty": qty_val if is_price_vis else "",
                        "price": price_val,
                        "total": tot,
                        "price_visible": is_price_vis,
                        "select_order": 0
                    })

                if selected:
                    grouped[cat] = selected

        return grouped, has_prices, grand_total

    def on_dashboard_record_double_click(self, event):
        """Dashboard log (Quotation ya Proforma tab) me kisi record par
        double-click karne se SIRF uska PDF preview khulta hai — current
        workspace (jo abhi edit ho rahi ho) bilkul touch nahi hoti. Record
        ko actually edit/load karne ke liye neeche wala separate 'Load'
        button use karein."""
        tree = event.widget
        row_id = tree.identify_row(event.y)
        if not row_id:
            return
        tree.selection_set(row_id)
        tree.focus(row_id)

        active_tab = self.history_filters.get()
        if "Quotation" in active_tab:
            sub_dir_name, doc_prefix = "Quotations", "Quotation"
        else:
            sub_dir_name, doc_prefix = "Proformas", "Proforma"

        item_vals = tree.item(row_id, "values")
        if not item_vals:
            return
        target_ref = item_vals[0]
        target_month = item_vals[4]
        record_path = os.path.join(BASE_HISTORY_DIR, target_month, sub_dir_name, f"{doc_prefix}_{target_ref}.json")

        if not os.path.exists(record_path):
            messagebox.showwarning("Not Found", "This record's saved file could not be located.")
            return

        try:
            with open(record_path, "r", encoding="utf-8") as f:
                record_data = json.load(f)

            grouped, has_prices, grand_total = self.get_grouped_data_from_record(record_data)
            if not grouped:
                messagebox.showwarning("Empty Record", "This record has no selected items to preview.")
                return

            override_data = {
                "factory": record_data.get("factory", ""),
                "doc_type": record_data.get("type", "Quotation"),
                "address": record_data.get("address", ""),
                "project": record_data.get("project", ""),
                "offer": record_data.get("offer", target_ref),
                "grouped": grouped, "has_prices": has_prices, "grand_total": grand_total,
            }

            temp_preview_path = os.path.join(tempfile.gettempdir(), "trutzschler_temp_preview.pdf")
            self.build_pdf_document_engine(temp_preview_path, override_data=override_data)
            open_file_with_default_app(temp_preview_path)
        except Exception as e:
            log_error("Dashboard double-click preview failed", e)
            messagebox.showerror("Preview Error", f"Could not open this record's PDF:\n{str(e)}")

    def load_selected_record(self):
        active_tab = self.history_filters.get()
        if "Quotation" in active_tab:
            sel = self.tree_q.selection(); target_tree = self.tree_q; sub_dir_name = "Quotations"; doc_prefix = "Quotation"
        else:
            sel = self.tree_p.selection(); target_tree = self.tree_p; sub_dir_name = "Proformas"; doc_prefix = "Proforma"
        if not sel:
            messagebox.showwarning("Selection Missing", "Please select a specific history record line profile to load.")
            return
        item_vals = target_tree.item(sel[0], "values")
        target_ref = item_vals[0]; target_month = item_vals[4]
        target_path = os.path.join(BASE_HISTORY_DIR, target_month, sub_dir_name, f"{doc_prefix}_{target_ref}.json")
        if os.path.exists(target_path):
            with open(target_path, "r") as f: data = json.load(f)

            # Remember exactly which file this session was loaded from, so a
            # later save can ask whether to update THIS file or save new.
            self.loaded_record_path = target_path
            self.loaded_record_ref = target_ref

            self.doc_type_var.set(data.get("type", "Quotation"))
            self.toggle_dynamic_fields()
            self.ent_factory.delete(0, "end"); self.ent_factory.insert(0, data.get("factory", ""))
            self.ent_offer.delete(0, "end"); self.ent_offer.insert(0, data.get("offer", ""))
            if data.get("type") == "Proforma":
                self.ent_address.delete(0, "end"); self.ent_address.insert(0, data.get("address", ""))
                self.ent_project.delete(0, "end"); self.ent_project.insert(0, data.get("project", ""))

            # --- ASLI FIX: LOAD PAR PURANI WORKSPACE STATE POORI TARAH SAAF KAREIN ---
            # Warna record load hote hi nayi lines purani lines/ticks ke sath
            # mix ho jaati thi.
            if hasattr(self, 'all_heading_cards'):
                for card in list(self.all_heading_cards):
                    try:
                        if card.winfo_exists():
                            card.destroy()
                    except Exception:
                        pass
                self.all_heading_cards.clear()
            self.active_heading_card = None
            self._switching_lines = False
            if hasattr(self, 'inventory_db') and self.inventory_db:
                for cat in self.inventory_db.keys():
                    try:
                        self.populate_category_rows(cat)
                    except Exception as e:
                        print(f"Error resetting category {cat} on load: {e}")

            saved_lines = data.get("lines", [])
            if saved_lines:
                # --- NEW-FORMAT RECORD: har saved line apni category tab me,
                # apni isolated selections ke sath, poori tarah rebuild karein ---
                first_card = None
                for line in saved_lines:
                    cat = line.get("category")
                    if not cat or cat not in self.inventory_db:
                        continue
                    card = self.create_heading_card(
                        cat,
                        line.get("line_no", ""),
                        line.get("title", "Section"),
                        saved_selections=line.get("selections", {}),
                        auto_activate=False
                    )
                    if first_card is None:
                        first_card = card

                if first_card is not None:
                    self.tabview_sub.set(first_card.owner_category)
                    self.activate_heading_line(first_card, first_card.master, clear_all=False)
            else:
                # --- OLD-FORMAT RECORD (koi lines save nahi thin, sirf flat
                # selections): backward-compatibility ke liye wohi purana raasta ---
                for cat, items in self.grid_inputs.items():
                    saved = data.get("selections", {}).get(cat, [])
                    for s, sv in zip(items, saved):
                        s["checkbox"].deselect()
                        if sv["checked"] == 1: s["checkbox"].select()
                        s["code_entry"].configure(state="normal"); s["code_entry"].delete(0, "end"); s["code_entry"].insert(0, sv.get("code", "")); s["code_entry"].configure(state="disabled")
                        s["code_show"].deselect(); 
                        if sv.get("code_show", 1) == 1: s["code_show"].select()
                        
                        s["price_show"].deselect()
                        if sv.get("price_show", 1) == 1: s["price_show"].select()
                        
                        s["price"].configure(state="normal")
                        s["price"].delete(0, "end")
                        s["price"].insert(0, sv.get("price", ""))
                        s["price"].configure(state="disabled")
                        
                        s["qty"].delete(0, "end"); s["qty"].insert(0, sv["qty"])

            self.tabs_container.set("Matrix Selection")
            self.show_toast(f"Loaded record '{target_ref}'", kind="info")

    def delete_selected_record(self):
        active_tab = self.history_filters.get()
        if "Quotation" in active_tab:
            sel = self.tree_q.selection(); target_tree = self.tree_q; sub_dir_name = "Quotations"; doc_prefix = "Quotation"
        else:
            sel = self.tree_p.selection(); target_tree = self.tree_p; sub_dir_name = "Proformas"; doc_prefix = "Proforma"
        if not sel:
            messagebox.showwarning("Selection Missing", "Please select a specific history record line profile to delete.")
            return
        item_vals = target_tree.item(sel[0], "values")
        target_ref = item_vals[0]; target_month = item_vals[4]
        target_path = os.path.join(BASE_HISTORY_DIR, target_month, sub_dir_name, f"{doc_prefix}_{target_ref}.json")
        if os.path.exists(target_path): os.remove(target_path); messagebox.showinfo("Deleted", "Record profile successfully removed."); self.refresh_dashboard_data()

    def trigger_pdf_preview(self):
        try:
            grouped, _, _ = self.get_parsed_data() 
            if not grouped: 
                messagebox.showwarning("Selection Missing", "No items selected.")
                return
            temp_preview_path = os.path.join(tempfile.gettempdir(), "trutzschler_temp_preview.pdf")
            self.build_pdf_document_engine(temp_preview_path)
            open_file_with_default_app(temp_preview_path)
        except Exception as e:
            messagebox.showerror("Preview Error", f"A problem occurred while generating the PDF preview:\n{str(e)}")

    def trigger_excel_preview(self):
        try:
            grouped, _, _ = self.get_parsed_data() 
            if not grouped: 
                messagebox.showwarning("Selection Missing", "No items selected.")
                return
            temp_xl_path = os.path.join(tempfile.gettempdir(), "trutzschler_temp_preview.xlsx")
            self.build_excel_sheet_engine(temp_xl_path)
            open_file_with_default_app(temp_xl_path)
        except Exception as e:
            messagebox.showerror("Preview Error", f"A problem occurred while generating the Excel preview:\n{str(e)}")

    def generate_pdf_proforma(self):
        try:
            grouped, _, grand_total = self.get_parsed_data() 
            if not grouped: 
                messagebox.showwarning("Empty Selection", "No components selected.")
                return
            save_path = filedialog.asksaveasfilename(defaultextension=".pdf", filetypes=[("PDF Document Replica", "*.pdf")])
            if not save_path: 
                return
            # NAYA: Export ke dauran busy cursor + turant UI refresh — bada
            # offer generate hone me thoda waqt lag sakta hai, is se user
            # ko lagta hai software "atka" nahi, kaam ho raha hai.
            self.configure(cursor="watch")
            self.update_idletasks()
            try:
                self.build_pdf_document_engine(save_path)
                self.auto_log_state_to_history(grand_total)
            finally:
                self.configure(cursor="")
            self.show_toast("PDF exported successfully!", kind="success")
        except Exception as e:
            self.configure(cursor="")
            log_error("PDF export failed", e)
            messagebox.showerror("Export Error", f"PDF export failed:\n{str(e)}")

    def generate_excel_format(self):
        try:
            grouped, _, grand_total = self.get_parsed_data() 
            if not grouped: 
                messagebox.showwarning("Empty Structure", "No machinery configurations mapped.")
                return
            save_p = filedialog.asksaveasfilename(defaultextension=".xlsx", filetypes=[("Excel Sheet", "*.xlsx")])
            if not save_p: 
                return
            self.configure(cursor="watch")
            self.update_idletasks()
            try:
                self.build_excel_sheet_engine(save_p)
                self.auto_log_state_to_history(grand_total)
            finally:
                self.configure(cursor="")
            self.show_toast("Excel spreadsheet generated!", kind="success")
        except Exception as e:
            self.configure(cursor="")
            log_error("Excel export failed", e)
            messagebox.showerror("Export Error", f"Excel export failed:\n{str(e)}")

    def refresh_entire_workspace(self):
            # 1. Left panel ki document configuration fields ko clear karein
            if hasattr(self, 'ent_factory') and self.ent_factory: 
                self.ent_factory.delete(0, 'end')
            if hasattr(self, 'ent_address') and self.ent_address: 
                self.ent_address.delete(0, 'end')
            if hasattr(self, 'ent_offer') and self.ent_offer: 
                self.ent_offer.delete(0, 'end')
            if hasattr(self, 'ent_project') and self.ent_project: 
                self.ent_project.delete(0, 'end')
            
            # Doc type dropdown ko default (Quotation) ya Proforma par reset karein
            if hasattr(self, 'doc_type_var') and self.doc_type_var: 
                self.doc_type_var.set("Quotation")
                self.toggle_dynamic_fields() # Left panel ke visibility fields ko sync karne ke liye
                
            # Search bar ko bhi khali kar dein agar usme kuch likha ho
            if hasattr(self, 'matrix_search_var') and self.matrix_search_var:
                self.matrix_search_var.set("")

            # 2. Har custom Line/Heading card ko FORCE-DESTROY karein. Permanent
            #    registry (all_heading_cards) ko source of truth maan kar chalte
            #    hain kyunke ek live widget-scan kisi inactive tab ki line miss
            #    kar sakta hai. Yehi woh cheez hai jo pehle refresh par "khari"
            #    reh jaati thi.
            if hasattr(self, 'all_heading_cards'):
                for card in list(self.all_heading_cards):
                    try:
                        if card.winfo_exists():
                            card.destroy()
                    except Exception:
                        pass
                self.all_heading_cards.clear()

            self.active_heading_card = None
            self._switching_lines = False

            # A fresh/blank workspace means we are no longer editing any
            # previously loaded history record — the next save should be
            # treated as a brand new record, not an update.
            self.loaded_record_path = None
            self.loaded_record_ref = None

            # 3. Saari categories ke items ko database se bilkul fresh state me repopulate (reset) karein
            if hasattr(self, 'inventory_db') and self.inventory_db:
                for cat in self.inventory_db.keys():
                    try:
                        self.populate_category_rows(cat)
                        if hasattr(self, '_populated_cats'):
                            self._populated_cats.add(cat)
                    except Exception as e:
                        print(f"Error resetting category {cat}: {e}")
                        
            # Success message box taake pata chale system refresh ho gya h
            messagebox.showinfo("System Refreshed", "Successfully — full system refreshed, all lines cleared.")


    # --- FIRST PAGE FULLY CUSTOMIZED PROFORMA INVOICE RE-ENGINEERING ---
    # --- FIRST PAGE FULLY CUSTOMIZED PROFORMA INVOICE RE-ENGINEERING ---
    def build_pdf_document_engine(self, target_path, override_data=None):
        # NAYA: override_data (optional) — jab diya jaye to PDF is diye
        # gaye data se banti hai, live workspace widgets se nahi. Isi
        # tareeqe se Dashboard 'double-click preview' current kaam ko
        # disturb kiye baghair kisi bhi purane record ka PDF dikha sakta
        # hai. Normal Save/Preview flow bilkul pehle jaisa hi chalta hai
        # (override_data=None hota hai).
        if override_data:
            factory = override_data.get("factory", "").strip() or "MAHMOOD TEXTILE"
            doc_type = override_data.get("doc_type", "Quotation")
            if doc_type == "Quotation":
                address = ""
                proj_name = ""
            else:
                address = override_data.get("address", "").strip() or "SURVEY # 31,32,33,34, KARACHI, PAKISTAN"
                proj_name = override_data.get("project", "").strip() or "Draw Frames"
            offer_num = override_data.get("offer", "").strip() or "02-013106-00"
            grouped = override_data.get("grouped", {})
            has_prices = override_data.get("has_prices", False)
            grand_total = override_data.get("grand_total", 0.0)
        else:
            factory = self.ent_factory.get().strip() or "MAHMOOD TEXTILE"
            doc_type = self.doc_type_var.get()

            if doc_type == "Quotation": 
                address = ""
                proj_name = ""
            else:
                address = self.ent_address.get().strip() or "SURVEY # 31,32,33,34, KARACHI, PAKISTAN"
                proj_name = self.ent_project.get().strip() or "Draw Frames"
                
            offer_num = self.ent_offer.get().strip() or "02-013106-00"
            grouped, has_prices, grand_total = self.get_parsed_data()
        
        doc = SimpleDocTemplate(target_path, pagesize=letter, leftMargin=54, rightMargin=54, topMargin=90, bottomMargin=110)
        
        s_normal = ParagraphStyle('Norm', fontName='Helvetica', fontSize=9, leading=14, textColor=colors.HexColor("#222222"))
        s_bold = ParagraphStyle('Bld', fontName='Helvetica-Bold', fontSize=9, leading=14)
        s_factory = ParagraphStyle('Fac', fontName='Helvetica-Bold', fontSize=12, leading=16, textColor=colors.HexColor("#111111"))
        s_title = ParagraphStyle('Ttl', fontName='Helvetica-Bold', fontSize=22, leading=26, spaceBefore=15, spaceAfter=5)
        s_sec = ParagraphStyle('Sec', fontName='Helvetica-Bold', fontSize=12, leading=16, spaceBefore=15, spaceAfter=10, textColor=colors.HexColor("#003399"))
        
        story = [Spacer(1, 15)]
        
        left_block = [Paragraph(factory.upper(), s_factory)]
        if address: 
            left_block.append(Paragraph(address, s_normal))
        else: 
            left_block.append(Spacer(1, 1))
            
        right_block = [
            Paragraph("<b>Official in charge:</b> Waleed Tahir", s_normal),
            Paragraph("<b>Department:</b> Sales", s_normal),
            Paragraph("<b>Fon:</b> +92-324-5600114", s_normal),
            Paragraph("<b>Fax:</b> +92-423-5713161", s_normal),
            Paragraph("<b>E-mail:</b> waleed@machpart.com", s_normal),
            Paragraph("<b>Website:</b> www.truetzschler.com", s_normal)
        ]
        
        header_table = Table([[left_block, right_block]], colWidths=[270, 230])
        header_table.setStyle(TableStyle([('VALIGN', (0,0), (-1,-1), 'TOP'), ('PADDING', (0,0), (-1,-1), 0)]))
        story.append(header_table)
        story.append(Spacer(1, 25))
        
        story.append(Paragraph(doc_type, s_title))
        story.append(Paragraph(f"<b>No.</b> {offer_num}", ParagraphStyle('OffNum', fontName='Helvetica', fontSize=12, leading=16, textColor=colors.HexColor("#444444"))))
        story.append(Spacer(1, 15))
        
        meta_left = []
        if proj_name:
            meta_left.append(Paragraph("<b>for your Project:</b>", s_normal))
            meta_left.append(Paragraph(proj_name, s_bold))
        else: 
            meta_left.append(Spacer(1, 1))
            
        meta_meta = [
            [meta_left, Paragraph(f"<b>Date:</b> {datetime.now().strftime('%d.%m.%Y')}", s_normal)],
            [Spacer(1, 12), ""],
            [Paragraph("<b>Representative:</b><br/>MACHPART <br/>195-p, Gulberg-III<br/>Lahore<br/>Pakistan", s_normal), ""]
        ]
        
        t_meta = Table(meta_meta, colWidths=[270, 230])
        t_meta.setStyle(TableStyle([('VALIGN', (0,0), (-1,-1), 'TOP'), ('PADDING', (0,0), (-1,-1), 1)]))
        story.append(t_meta)
        
        if has_prices:
            story.append(PageBreak())
            story.append(Paragraph("Commercial Price Summary Breakdown", s_sec))
            sum_data = [["Section Description Details Nomenclature", "Net Price (EUR)"]]
            for cat, items in grouped.items():
                cat_subtotal = sum(x["total"] for x in items if isinstance(x["total"], float))
                if cat_subtotal > 0:
                    sum_data.append([f"Subtotal System Machinery Section — {cat}", f"€ {cat_subtotal:,.2f}"])
            sum_data.append(["TOTAL OVERALL SUMMARY VALUE OFFER", f"€ {grand_total:,.2f}"])
            t_sum = Table(sum_data, colWidths=[370, 130])
            t_sum.setStyle(TableStyle([
                ('BACKGROUND', (0,0), (-1,0), colors.HexColor("#f8f9fa")), ('GRID', (0,0), (-1,-1), 0.5, colors.HexColor("#cccccc")),
                ('ALIGN', (1,0), (1,-1), 'RIGHT'), ('FONTNAME', (0,-1), (-1,-1), 'Helvetica-Bold'), ('PADDING', (0,0), (-1,-1), 8)
            ]))
            story.append(t_sum)
            
        for cat, items in grouped.items():
            story.append(PageBreak())
            # Ab 'cat' me seedha GUI ki bani hui Heading Line ka naam aayega (e.g. Line 1 / Custom Name)
            story.append(Paragraph(f"{cat}", s_sec))
            
            cat_has_prices = any(x["price_visible"] and isinstance(x["price"], float) for x in items)
            any_code_visible = any(it["code_visible"] for it in items)
            
            if cat_has_prices:
                if any_code_visible:
                    headers = ["Code", "Subassembly Specification Description Details", "Qty", "Price", "Total"]
                    widths = [65, 230, 25, 90, 90]
                else:
                    headers = ["Subassembly Specification Description Details", "Qty", "Price", "Total"]
                    widths = [295, 25, 90, 90]
                
                table_rows = [headers]
                for it in items:
                    desc_p = Paragraph(f"<b>{it['name']}</b><br/><font color='#555555'>{it['specs']}</font>", s_normal)
                    q_text = str(it["qty"]) if it["qty"] != "" else ""
                    
                    p_text = f"€ {it['price']:,.2f}" if (it.get("price_visible") and isinstance(it.get("price"), float)) else ""
                    t_text = f"€ {it['total']:,.2f}" if (it.get("price_visible") and isinstance(it.get("total"), float)) else ""
                    
                    if any_code_visible:
                        table_rows.append([str(it["code"]), desc_p, q_text, p_text, t_text])
                    else:
                        table_rows.append([desc_p, q_text, p_text, t_text])
                        
                # --- ASLI FIX: HAR HEADING KE PRODUCTS KE AAKHIR ME "Line total"
                # row add karein, bilkul waise jaisa Excel engine already karta
                # hai. Pehle yeh row PDF me sirey se missing thi. ---
                cat_total = sum(x["total"] for x in items if isinstance(x["total"], float))
                # NOTE: jab cells SPAN (merge) ki jaati hain to ReportLab sirf
                # us merge ke SABSE PEHLE (top-left) cell ka content dikhata
                # hai — isliye label hamesha index 0 par hi hona chahiye,
                # warna woh merge hote hi ghayab ho jata hai (pichla bug).
                if any_code_visible:
                    total_row = [Paragraph(f"<b>Line total {cat}</b>", s_bold), "", "", "", f"€ {cat_total:,.2f}"]
                else:
                    total_row = [Paragraph(f"<b>Line total {cat}</b>", s_bold), "", "", f"€ {cat_total:,.2f}"]
                table_rows.append(total_row)
                total_row_idx = len(table_rows) - 1

                t_spec = Table(table_rows, colWidths=widths)
                t_spec.setStyle(TableStyle([
                    ('GRID', (0,0), (-1,-1), 0.5, colors.HexColor("#dddddd")), ('BACKGROUND', (0,0), (-1,0), colors.HexColor("#f2f5f9")),
                    ('VALIGN', (0,0), (-1,-1), 'TOP'), ('PADDING', (0,0), (-1,-1), 7), 
                    ('ALIGN', (-4,1), (-1,-1), 'CENTER'),
                    ('SPAN', (0, total_row_idx), (-2, total_row_idx)),
                    ('BACKGROUND', (0, total_row_idx), (-1, total_row_idx), colors.HexColor("#e8edf5")),
                    ('FONTNAME', (0, total_row_idx), (-1, total_row_idx), 'Helvetica-Bold'),
                    ('ALIGN', (-1, total_row_idx), (-1, total_row_idx), 'RIGHT'),
                    ('LINEABOVE', (0, total_row_idx), (-1, total_row_idx), 1, colors.HexColor("#003399")),
                ]))
            else:
                if any_code_visible:
                    headers = ["Code", "Subassembly Specification Description Details"]
                    widths = [80, 420]
                else:
                    headers = ["Subassembly Specification Description Details"]
                    widths = [500]
                    
                table_rows = [headers]
                for it in items:
                    desc_p = Paragraph(f"<b>{it['name']}</b><br/><font color='#555555'>{it['specs']}</font>", s_normal)
                    if any_code_visible:
                        table_rows.append([str(it["code"]), desc_p])
                    else:
                        table_rows.append([desc_p])
                        
                t_spec = Table(table_rows, colWidths=widths)
                t_spec.setStyle(TableStyle([
                    ('GRID', (0,0), (-1,-1), 0.5, colors.HexColor("#dddddd")), ('BACKGROUND', (0,0), (-1,0), colors.HexColor("#f2f5f9")),
                    ('VALIGN', (0,0), (-1,-1), 'TOP'), ('PADDING', (0,0), (-1,-1), 8)
                ]))
                
            story.append(t_spec)
            
        doc.build(story, canvasmaker=NumberedCanvas)

          # --- FIRST PAGE FULLY DYNAMIC EXCEL ENGINE ---
    def build_excel_sheet_engine(self, target_path):
        factory = self.ent_factory.get().strip() or "MAHMOOD TEXTILE"
        offer_num = self.ent_offer.get().strip()
        doc_type = self.doc_type_var.get()
        
        # EXACT DATA LAYER THAT MAKES PDF PERFECT
        grouped, has_prices, grand_total = self.get_parsed_data()
        
        wb = Workbook()
        ws = wb.active
        ws.title = "Proposal"
        ws.views.sheetView[0].showGridLines = True
        
        f_head = XLFont(name="Arial", size=15, bold=True)
        f_bold = XLFont(name="Arial", size=10, bold=True)
        f_normal = XLFont(name="Arial", size=10)
        fill_y = PatternFill(start_color="FFFF00", end_color="FFFF00", fill_type="solid")
        border_thin = Border(left=Side(style='thin', color='000000'), right=Side(style='thin', color='000000'), top=Side(style='thin', color='000000'), bottom=Side(style='thin', color='000000'))
        
        ws.merge_cells("A1:G1")
        ws["A1"] = f"{factory.upper()} - {doc_type.upper()}"
        ws["A1"].font = f_head
        ws["A1"].fill = fill_y
        ws["A1"].alignment = Alignment(horizontal="center")
        
        ws["A3"] = f"{doc_type} Pricing Framework Working"
        ws["D3"] = f"Ref No {offer_num}" if offer_num else ""
        ws["F2"] = "Based On New Price Matrix"
        ws["F2"].fill = fill_y
        ws["F3"] = datetime.now().strftime("%m/%d/%Y")
        
        current_row = 5
        global_any_code_visible = False

        for cat, items in grouped.items():
            if not items:
                continue
                
            # Category level mapping exactly matching PDF workflow
            cat_has_prices = any(x["price_visible"] and isinstance(x["price"], float) for x in items)
            any_code_visible = any(it["code_visible"] for it in items)
            
            if any_code_visible:
                global_any_code_visible = True
            
            if cat_has_prices:
                # --- MODE 1: PRICING PRESENT LAYOUT ---
                ws.cell(row=current_row, column=2, value="Quantity").font = f_bold
                c_sec = ws.cell(row=current_row, column=3, value=f"Extension of the {cat} Installation")
                c_sec.font = f_bold
                c_sec.fill = fill_y
                
                if any_code_visible:
                    ws.cell(row=current_row, column=4, value="CODE").font = f_bold
                    ws.cell(row=current_row, column=5, value="Unit Price").font = f_bold
                    ws.cell(row=current_row, column=6, value="Total Price").font = f_bold
                    max_col = 7
                else:
                    ws.cell(row=current_row, column=4, value="Unit Price").font = f_bold
                    ws.cell(row=current_row, column=5, value="Total Price").font = f_bold
                    max_col = 6
                        
                for c in range(2, max_col): 
                    ws.cell(row=current_row, column=c).border = border_thin
                current_row += 1
                
                cat_total = 0
                for item in items:
                    ws.cell(row=current_row, column=2, value=item["qty"])
                    ws.cell(row=current_row, column=3, value=item["name"])
                    
                    if any_code_visible:
                        ws.cell(row=current_row, column=4, value=str(item["code"]) if item.get("code_visible") else None)
                        p_col, t_col = 5, 6
                    else:
                        p_col, t_col = 4, 5
                        
                    if item.get("price_visible") and isinstance(item.get("price"), float):
                        ws.cell(row=current_row, column=p_col, value=item["price"]).number_format = '#,##0.00'
                        ws.cell(row=current_row, column=t_col, value=item["total"]).number_format = '#,##0.00'
                        cat_total += item["total"]
                    else:
                        ws.cell(row=current_row, column=p_col, value=None)
                        ws.cell(row=current_row, column=t_col, value=None)
                        
                    for c in range(2, max_col): 
                        ws.cell(row=current_row, column=c).font = f_normal
                        ws.cell(row=current_row, column=c).border = border_thin
                    current_row += 1
                    
                ws.cell(row=current_row, column=3, value=f"Line total {c_sec.value}").font = f_bold
                ws.cell(row=current_row, column=t_col, value=cat_total).font = f_bold
                ws.cell(row=current_row, column=t_col).number_format = '#,##0.00'
            else:
                # --- MODE 2: PURE SPECIFICATIONS (PRICES UNTICKED) ---
                c_sec = ws.cell(row=current_row, column=3, value=f"{cat} Specification Details")
                c_sec.font = f_bold
                c_sec.fill = fill_y
                
                # Header Border Fix (range ko 3 se 5 tak chalana hai agar code visible ho)
                end_border_col = 5 if any_code_visible else 4
                for c in range(3, end_border_col):
                    ws.cell(row=current_row, column=c).border = border_thin
                
                if any_code_visible:
                    ws.cell(row=current_row, column=4, value="CODE").font = f_bold
                    ws.cell(row=current_row, column=4).border = border_thin
                    
                current_row += 1
                for item in items:
                    ws.cell(row=current_row, column=3, value=item["name"]).font = f_normal
                    ws.cell(row=current_row, column=3).border = border_thin
                    
                    if any_code_visible:
                        ws.cell(row=current_row, column=4, value=str(item["code"]) if item.get("code_visible") else None).font = f_normal
                        ws.cell(row=current_row, column=4).border = border_thin
                    current_row += 1
                    
                ws.cell(row=current_row, column=3, value=f"{cat} Specification Mapped Successfully").font = f_bold
            
            current_row += 2
            
        # GRAND TOTAL: Strict numerical and boolean filter matching PDF structure
        if has_prices and isinstance(grand_total, (int, float)) and grand_total > 0: 
            final_col_idx = 6 if global_any_code_visible else 5
            ws.cell(row=current_row, column=3, value="GRAND TOTAL").font = f_bold
            ws.cell(row=current_row, column=final_col_idx, value=grand_total).font = f_bold
            ws.cell(row=current_row, column=final_col_idx).number_format = '#,##0.00'

        
            
        
            
       
            
        # --- COLUMNS AUTO-FIT & WIDTH SETTING ENGINE ---
            
            ws.column_dimensions['C'].width = 42  # Description / Item Name
            ws.column_dimensions['D'].width = 15  # Code column
            ws.column_dimensions['E'].width = 16  # Qty column
            ws.column_dimensions['F'].width = 18  # Unit Price column
            ws.column_dimensions['G'].width = 19  # Total Price / Grand Total column
            
            wb.save(target_path)

if __name__ == "__main__":
    app = MasterSystemApp()
    app.mainloop()
