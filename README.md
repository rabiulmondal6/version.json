from datetime import datetime
import sqlite3
import tkinter as tk
from tkinter import filedialog, messagebox, ttk
import pandas as pd
import warnings
from reportlab.lib.pagesizes import letter
from reportlab.lib import colors
from reportlab.platypus import SimpleDocTemplate, Paragraph, Spacer, Table, TableStyle
from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
from PIL import Image, ImageTk
import os
import sys
import subprocess
import threading
import urllib.request
import json

# --- Update Configuration ---
CURRENT_VERSION = "2.0"
# নিচে আপনার সার্ভারের বা গিটহাবের version.json লিংকের Raw URL বসাতে হবে। 
# (আপাতত আমি একটি ডেমো লিংক দিয়ে রাখলাম, এটি পরে চেঞ্জ করে নেবেন)
UPDATE_URL = "https://github.com/rabiulmondal6/version.json/blob/main/version.json"


# Database Setup with Remarks and Color Tag support
def init_db():
    conn = sqlite3.connect("data_store.db")
    cursor = conn.cursor()
    
    # Records Table
    cursor.execute("""
        CREATE TABLE IF NOT EXISTS records (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            date TEXT,
            amount REAL,
            service_type TEXT,
            grn_no TEXT,
            application_no TEXT,
            khatiyan_plot TEXT,
            transaction_no TEXT,
            citizen_name TEXT,
            citizen_address TEXT,
            phone_number TEXT,
            deo_name TEXT,
            remarks TEXT,
            color_tag TEXT
        )
    """)
    
    # Profile Table
    cursor.execute("""
        CREATE TABLE IF NOT EXISTS profile (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            bsk_name TEXT,
            bsk_code TEXT,
            deo_code TEXT,
            deo_name TEXT,
            mobile_no TEXT,
            bsk_account TEXT
        )
    """)
    
    cursor.execute("SELECT COUNT(*) FROM profile")
    if cursor.fetchone()[0] == 0:
        cursor.execute("""
            INSERT INTO profile (bsk_name, bsk_code, deo_code, deo_name, mobile_no, bsk_account)
            VALUES (?, ?, ?, ?, ?, ?)
        """, ("Default BSK Unit", "BSK-001", "DEO-001", "Admin DEO", "01700000000", "ACC-123456"))
    
    conn.commit()
    conn.close()

def get_db_connection():
    """Helper to guarantee tables exist before any database operation."""
    init_db()
    return sqlite3.connect("data_store.db")

init_db()

class ExcelSpreadsheet(tk.Frame):
    def __init__(self, parent, headers, data):
        super().__init__(parent, bg="#ffffff")
        self.headers = headers
        self.raw_data = data 

        self.tree = ttk.Treeview(self, columns=[str(i) for i in range(len(headers))], show="headings", selectmode="extended")
        
        for idx, h_text in enumerate(headers):
            self.tree.heading(str(idx), text=h_text)
            self.tree.column(str(idx), width=110, anchor="center")

        v_scroll = ttk.Scrollbar(self, orient=tk.VERTICAL, command=self.tree.yview)
        h_scroll = ttk.Scrollbar(self, orient=tk.HORIZONTAL, command=self.tree.xview)
        self.tree.configure(yscrollcommand=v_scroll.set, xscrollcommand=h_scroll.set)

        self.tree.pack(side=tk.TOP, fill=tk.BOTH, expand=True)
        v_scroll.pack(side=tk.RIGHT, fill=tk.Y)
        h_scroll.pack(side=tk.BOTTOM, fill=tk.X)

        self.load_grid_data()
        
        self.tree.bind("<ButtonPress-1>", self.track_cell_click, add="+")
        self.tree.bind("<Control-c>", self.copy_selection)

    def track_cell_click(self, event):
        self.clicked_col = self.tree.identify_column(event.x)

    def load_grid_data(self):
        for row_data in self.raw_data:
            self.tree.insert("", tk.END, values=[str(val) for val in row_data])

    def copy_selection(self, event=None):
        selected_items = self.tree.selection()
        if not selected_items:
            return
        
        if len(selected_items) == 1 and hasattr(self, 'clicked_col') and self.clicked_col:
            try:
                col_idx = int(self.clicked_col.replace('#', '')) - 1
                item_values = self.tree.item(selected_items[0], "values")
                if 0 <= col_idx < len(item_values):
                    cell_text = str(item_values[col_idx])
                    self.winfo_toplevel().clipboard_clear()
                    self.winfo_toplevel().clipboard_append(cell_text)
                    return
            except Exception:
                pass
        
        copied_rows = []
        for item in selected_items:
            vals = self.tree.item(item, "values")
            copied_rows.append("\t".join(str(v) for v in vals))
            
        clipboard_text = "\n".join(copied_rows)
        try:
            self.winfo_toplevel().clipboard_clear()
            self.winfo_toplevel().clipboard_append(clipboard_text)
        except Exception:
            pass


class DataEntryApp:
    def __init__(self, root):
        self.root = root
        self.root.title(f"BSK Data Management & Reporting Portal - Pro Edition (v{CURRENT_VERSION})")
        self.root.geometry("1650x900")
        self.root.configure(bg="#0f172a")

        try:
            self.root.iconbitmap("icon.ico")
        except Exception:
            pass

        style = ttk.Style()
        style.theme_use("clam")
        style.configure("Treeview.Heading", background="#1e293b", foreground="#38bdf8", font=("Segoe UI", 10, "bold"), relief="flat")
        style.configure("Treeview", rowheight=32, font=("Segoe UI", 10), background="#ffffff", fieldbackground="#ffffff", bordercolor="#cbd5e1", borderwidth=1)
        style.map("Treeview", background=[('selected', '#3b82f6')], foreground=[('selected', 'white')])

        header_frame = tk.Frame(self.root, bg="#1e293b", height=85)
        header_frame.pack(fill=tk.X, side=tk.TOP)
        header_frame.pack_propagate(False)
        
        try:
            if os.path.exists("logo.png"):
                img = Image.open("logo.png")
                img = img.resize((50, 50), Image.Resampling.LANCZOS)
                self.logo_img = ImageTk.PhotoImage(img)
                logo_lbl = tk.Label(header_frame, image=self.logo_img, bg="#1e293b")
                logo_lbl.pack(side=tk.LEFT, padx=(20, 10), pady=15)
        except Exception as e:
            pass

        title_container = tk.Frame(header_frame, bg="#1e293b")
        title_container.pack(side=tk.LEFT, pady=12)

        tk.Label(title_container, text="BSK DATA MANAGEMENT & REPORTING PORTAL", fg="#f8fafc", bg="#1e293b", font=("Segoe UI", 14, "bold")).pack(anchor="w")
        tk.Label(title_container, text="Bangla Sahayata Kendra | Government of West Bengal", fg="#38bdf8", bg="#1e293b", font=("Segoe UI", 9)).pack(anchor="w")

        btn_container = tk.Frame(header_frame, bg="#1e293b")
        btn_container.pack(side=tk.RIGHT, padx=25)

        new_entry_btn = tk.Button(btn_container, text="➕ New Data Entry", bg="#10b981", fg="white", font=("Segoe UI", 10, "bold"), bd=0, padx=16, pady=8, cursor="hand2", command=self.open_advanced_entry_window, activebackground="#059669", activeforeground="white")
        new_entry_btn.pack(side=tk.LEFT, padx=10)

        profile_btn = tk.Button(btn_container, text="⚙ BSK PROFILE SETUP", bg="#3b82f6", fg="white", font=("Segoe UI", 10, "bold"), bd=0, padx=16, pady=8, cursor="hand2", command=self.open_profile_window, activebackground="#2563eb", activeforeground="white")
        profile_btn.pack(side=tk.LEFT)

        profile_display_frame = tk.LabelFrame(self.root, text=" Active BSK Unit Status ", font=("Segoe UI", 9, "bold"), bg="#1e293b", fg="#38bdf8", bd=1, relief="solid", padx=15, pady=8)
        profile_display_frame.pack(fill=tk.X, padx=20, pady=(10, 0))

        self.p_labels = {}
        p_titles = [
            ("BSK Name:", "bsk_name"), 
            ("BSK Code:", "bsk_code"), 
            ("DEO Code:", "deo_code"), 
            ("DEO Name:", "deo_name"), 
            ("Mobile No:", "mobile_no"), 
            ("BSK Account:", "bsk_account")
        ]

        for idx, (title, key) in enumerate(p_titles):
            tk.Label(profile_display_frame, text=title, bg="#1e293b", fg="#94a3b8", font=("Segoe UI", 9, "bold")).grid(row=0, column=idx*2, sticky="w", padx=(15, 4))
            val_lbl = tk.Label(profile_display_frame, text="", bg="#0f172a", fg="#f8fafc", font=("Segoe UI", 9, "bold"), width=15, relief="solid", bd=1, anchor="w")
            val_lbl.grid(row=0, column=idx*2+1, sticky="w", padx=4)
            self.p_labels[key] = val_lbl

        self.update_main_profile_labels()

        main_frame = tk.Frame(self.root, bg="#0f172a")
        main_frame.pack(fill=tk.BOTH, expand=True, padx=20, pady=10)

        left_frame = tk.Frame(main_frame, bg="#0f172a")
        left_frame.pack(side=tk.LEFT, fill=tk.BOTH, expand=True, padx=(0, 0))

        control_panel = tk.Frame(left_frame, bg="#1e293b", padx=10, pady=10, relief="solid", bd=1)
        control_panel.pack(fill=tk.X, pady=(0, 8))

        search_row = tk.Frame(control_panel, bg="#1e293b")
        search_row.pack(fill=tk.X, pady=(0, 6))

        tk.Label(search_row, text="🔍 Global Search:", bg="#1e293b", fg="#cbd5e1", font=("Segoe UI", 9, "bold")).pack(side=tk.LEFT, padx=2)
        self.search_entry = tk.Entry(search_row, font=("Segoe UI", 9), width=18, relief="solid", bd=1, bg="#0f172a", fg="#f8fafc", insertbackground="white")
        self.search_entry.pack(side=tk.LEFT, padx=2)
        self.search_entry.bind("<KeyRelease>", self.filter_data)

        tk.Label(search_row, text="Date:", bg="#1e293b", fg="#cbd5e1", font=("Segoe UI", 9, "bold")).pack(side=tk.LEFT, padx=(8, 2))
        self.filter_date_entry = tk.Entry(search_row, font=("Segoe UI", 9), width=11, relief="solid", bd=1, bg="#0f172a", fg="#f8fafc", insertbackground="white")
        self.filter_date_entry.pack(side=tk.LEFT, padx=2)
        self.filter_date_entry.bind("<KeyRelease>", self.filter_data)

        tk.Button(search_row, text="❌ Clear", bg="#64748b", fg="white", font=("Segoe UI", 8, "bold"), bd=0, padx=6, pady=3, cursor="hand2", command=self.clear_filter).pack(side=tk.LEFT, padx=6)

        btn_row = tk.Frame(control_panel, bg="#1e293b")
        btn_row.pack(fill=tk.X)

        tk.Button(btn_row, text="📈 View Reports & Filter", bg="#7c3aed", fg="white", font=("Segoe UI", 9, "bold"), bd=0, padx=8, pady=5, cursor="hand2", command=self.open_report_window, activebackground="#6d28d9", activeforeground="white").pack(side=tk.LEFT, padx=3)
        tk.Button(btn_row, text="📥 Import Excel", bg="#d97706", fg="white", font=("Segoe UI", 9, "bold"), bd=0, padx=8, pady=5, cursor="hand2", command=self.import_excel, activebackground="#b45309", activeforeground="white").pack(side=tk.RIGHT, padx=3)
        tk.Button(btn_row, text="📋 Sample Format", bg="#0284c7", fg="white", font=("Segoe UI", 9, "bold"), bd=0, padx=8, pady=5, cursor="hand2", command=self.download_sample_format, activebackground="#0369a1", activeforeground="white").pack(side=tk.RIGHT, padx=3)
        tk.Button(btn_row, text="📊 Export Excel", bg="#059669", fg="white", font=("Segoe UI", 9, "bold"), bd=0, padx=8, pady=5, cursor="hand2", command=self.export_excel, activebackground="#047857", activeforeground="white").pack(side=tk.RIGHT, padx=3)

        table_frame = tk.Frame(left_frame, bg="#1e293b", bd=1, relief="solid")
        table_frame.pack(fill=tk.BOTH, expand=True)

        columns = ("Sl", "Date", "Name", "Amount", "Address", "Phone", "Service Type", "GRN", "App No", "Khatiyan", "Trans No", "Remarks", "Color", "DEO")
        self.tree = ttk.Treeview(table_frame, columns=columns, show="headings", selectmode="extended")
        
        col_widths = [40, 85, 115, 75, 110, 90, 90, 75, 75, 75, 80, 90, 75, 75]
        for col, w in zip(columns, col_widths):
            self.tree.heading(col, text=col)
            self.tree.column(col, width=w, anchor="center")

        self.tree.tag_configure('green', background='#dcfce7')
        self.tree.tag_configure('yellow', background='#fef9c3')
        self.tree.tag_configure('red', background='#fee2e2')
        self.tree.tag_configure('blue', background='#e0f2fe')
        self.tree.tag_configure('purple', background='#f3e8ff')
        self.tree.tag_configure('evenrow', background='#ffffff')
        self.tree.tag_configure('oddrow', background='#f8fafc')

        scrollbar_y = ttk.Scrollbar(table_frame, orient=tk.VERTICAL, command=self.tree.yview)
        scrollbar_x = ttk.Scrollbar(table_frame, orient=tk.HORIZONTAL, command=self.tree.xview)
        self.tree.configure(yscrollcommand=scrollbar_y.set, xscrollcommand=scrollbar_x.set)
        
        self.tree.pack(side=tk.TOP, fill=tk.BOTH, expand=True)
        scrollbar_y.pack(side=tk.RIGHT, fill=tk.Y)
        scrollbar_x.pack(side=tk.BOTTOM, fill=tk.X)

        def setup_drag_scrolling(tree_widget):
            def on_press(event):
                tree_widget._drag_start_x = event.x
                tree_widget._drag_start_y = event.y
                self.clicked_col = tree_widget.identify_column(event.x)

            def on_drag(event):
                if not hasattr(tree_widget, '_drag_start_x') or not hasattr(tree_widget, '_drag_start_y'):
                    tree_widget._drag_start_x = event.x
                    tree_widget._drag_start_y = event.y
                    return
                
                dx = tree_widget._drag_start_x - event.x
                dy = tree_widget._drag_start_y - event.y
                
                if abs(dy) > 5:
                    direction = 1 if dy > 0 else -1
                    tree_widget.yview_scroll(direction, "units")
                    tree_widget._drag_start_y = event.y
                    
                if abs(dx) > 5:
                    direction = 1 if dx > 0 else -1
                    tree_widget.xview_scroll(direction, "units")
                    tree_widget._drag_start_x = event.x

            tree_widget.bind("<ButtonPress-1>", on_press, add="+")
            tree_widget.bind("<B1-Motion>", on_drag, add="+")

        setup_drag_scrolling(self.tree)

        self.tree.bind("<Double-1>", self.open_edit_window)
        self.tree.bind("<Button-3>", self.show_context_menu)
        self.tree.bind("<Control-c>", self.copy_row_event)

        self.context_menu = tk.Menu(self.root, tearoff=0, bg="#1e293b", fg="#f8fafc", activebackground="#3b82f6", activeforeground="white")
        self.context_menu.add_command(label="📋 Copy Selected Cell", command=self.copy_cell)
        self.context_menu.add_command(label="📋 Copy Selected Row(s)", command=self.copy_row)
        self.context_menu.add_command(label="✏ Edit Selected Record", command=self.edit_selected_from_menu)
        self.context_menu.add_command(label="🗑 Delete Selected Record", command=self.delete_selected_record)

        self.load_data()
        
        # --- Start Auto Update Check in Background ---
        threading.Thread(target=self.perform_update_check, daemon=True).start()

    # --- AUTO UPDATER LOGIC START ---
    def perform_update_check(self):
        try:
            req = urllib.request.Request(UPDATE_URL, headers={'User-Agent': 'Mozilla/5.0'})
            with urllib.request.urlopen(req, timeout=5) as response:
                data = json.loads(response.read().decode())
                
            latest_version = data.get("latest_version")
            download_url = data.get("download_url")
            
            if latest_version and float(latest_version) > float(CURRENT_VERSION):
                self.root.after(2000, self.prompt_update, latest_version, download_url)
        except Exception as e:
            print("Update check failed:", e)

    def prompt_update(self, latest_version, download_url):
        ans = messagebox.askyesno("Software Update Available!", f"Great news! Version {latest_version} is available.\n\nDo you want to download and install it automatically right now?", parent=self.root)
        if ans:
            self.download_and_apply_update(download_url)

    def download_and_apply_update(self, download_url):
        update_win = tk.Toplevel(self.root)
        update_win.title("Updating Software...")
        update_win.geometry("350x120")
        update_win.configure(bg="#1e293b")
        update_win.grab_set()
        
        tk.Label(update_win, text="Downloading new version...", bg="#1e293b", fg="#38bdf8", font=("Segoe UI", 11, "bold")).pack(pady=(20, 10))
        tk.Label(update_win, text="Please wait, the application will restart automatically.", bg="#1e293b", fg="#cbd5e1", font=("Segoe UI", 9)).pack()
        update_win.update()

        try:
            new_file_name = "update_new.tmp"
            urllib.request.urlretrieve(download_url, new_file_name)
            
            current_file = sys.argv[0]
            is_exe = current_file.lower().endswith('.exe')
            
            if is_exe:
                restart_cmd = f'start "" "{os.path.basename(current_file)}"'
            else:
                restart_cmd = f'start "" "{sys.executable}" "{os.path.basename(current_file)}"'

            bat_content = f"""@echo off
echo Updating Software, Please wait...
timeout /t 3 /nobreak > NUL
del "{current_file}"
ren "{new_file_name}" "{os.path.basename(current_file)}"
{restart_cmd}
del "%~f0"
"""
            with open("updater.bat", "w") as f:
                f.write(bat_content)
            
            subprocess.Popen("updater.bat", shell=True)
            self.root.quit()
            sys.exit()

        except Exception as e:
            update_win.destroy()
            messagebox.showerror("Update Error", f"Failed to apply update automatically. Please check your internet connection.\nError: {e}", parent=self.root)
    # --- AUTO UPDATER LOGIC END ---

    def validate_phone(self, p):
        if p == "":
            return True
        if p.isdigit() and len(p) <= 10:
            return True
        return False

    def validate_name(self, p):
        if p == "":
            return True
        if all(char.isalpha() or char.isspace() for char in p):
            return True
        return False

    def clean_phone_number(self, phone_val):
        if pd.isna(phone_val):
            return ""
        digits = "".join(filter(str.isdigit, str(phone_val)))
        if len(digits) > 10:
            digits = digits[-10:]
        return digits

    def get_profile_data(self):
        conn = get_db_connection()
        cursor = conn.cursor()
        cursor.execute("SELECT bsk_name, bsk_code, deo_code, deo_name, mobile_no, bsk_account FROM profile LIMIT 1")
        data = cursor.fetchone()
        conn.close()
        return data

    def update_main_profile_labels(self):
        data = self.get_profile_data()
        if data:
            keys = ["bsk_name", "bsk_code", "deo_code", "deo_name", "mobile_no", "bsk_account"]
            for idx, key in enumerate(keys):
                self.p_labels[key].config(text=f" {data[idx]}")

    def open_profile_window(self):
        profile_win = tk.Toplevel(self.root)
        profile_win.title("BSK Profile Setup")
        profile_win.configure(bg="#1e293b")
        profile_win.grab_set()

        try:
            profile_win.iconbitmap("icon.ico")
        except Exception:
            pass

        tk.Label(profile_win, text="BSK Profile Setup", font=("Segoe UI", 14, "bold"), bg="#1e293b", fg="#f8fafc").pack(pady=(18, 10))

        form_frame = tk.Frame(profile_win, bg="#1e293b")
        form_frame.pack(padx=25, pady=10, fill=tk.BOTH, expand=True)

        p_fields = [
            ("BSK Name:", "bsk_name"),
            ("BSK Code:", "bsk_code"),
            ("DEO Code:", "deo_code"),
            ("DEO Name:", "deo_name"),
            ("Mobile No:", "mobile_no"),
            ("BSK Account Number:", "bsk_account")
        ]

        current_data = self.get_profile_data()
        p_entries = {}

        for i, (lbl_text, key) in enumerate(p_fields):
            tk.Label(form_frame, text=lbl_text, bg="#1e293b", fg="#cbd5e1", font=("Segoe UI", 9, "bold")).grid(row=i, column=0, sticky="w", pady=6)
            ent = tk.Entry(form_frame, font=("Segoe UI", 10), width=26, relief="solid", bd=1, bg="#0f172a", fg="#f8fafc", insertbackground="white")
            ent.grid(row=i, column=1, pady=6, padx=10)
            if current_data:
                ent.insert(0, current_data[i])
            p_entries[key] = ent

        def save_profile():
            vals = [ent.get() for ent in p_entries.values()]
            conn = get_db_connection()
            cursor = conn.cursor()
            cursor.execute("DELETE FROM profile")
            cursor.execute("""
                INSERT INTO profile (bsk_name, bsk_code, deo_code, deo_name, mobile_no, bsk_account)
                VALUES (?, ?, ?, ?, ?, ?)
            """, tuple(vals))
            conn.commit()
            conn.close()
            
            self.update_main_profile_labels()
            messagebox.showinfo("Success", "BSK Profile Setup Updated Successfully!")
            profile_win.destroy()

        tk.Button(profile_win, text="Save Profile", bg="#059669", fg="white", font=("Segoe UI", 10, "bold"), bd=0, padx=15, pady=8, cursor="hand2", command=save_profile, activebackground="#047857", activeforeground="white").pack(pady=18)
        profile_win.update_idletasks()
        profile_win.geometry(f"{profile_win.winfo_reqwidth() + 50}x{profile_win.winfo_reqheight() + 20}")

    def open_advanced_entry_window(self):
        entry_win = tk.Toplevel(self.root)
        entry_win.title("Advanced Commercial Data Entry Terminal")
        entry_win.configure(bg="#0f172a")
        entry_win.grab_set()

        try:
            entry_win.iconbitmap("icon.ico")
        except Exception:
            pass

        banner = tk.Frame(entry_win, bg="#1e293b", padx=20, pady=12)
        banner.pack(fill=tk.X)
        tk.Label(banner, text="BSK DATA MANAGEMENT & REPORTING PORTAL", font=("Segoe UI", 12, "bold"), fg="#38bdf8", bg="#1e293b").pack(side=tk.LEFT)

        form_card = tk.Frame(entry_win, bg="#1e293b", padx=20, pady=15, bd=1, relief="solid")
        form_card.pack(fill=tk.BOTH, expand=True, padx=20, pady=15)

        vcmd_phone = (entry_win.register(self.validate_phone), '%P')
        vcmd_name = (entry_win.register(self.validate_name), '%P')

        fields = [
            ("Date (DD-MM-YYYY):", "date_entry"), 
            ("Citizen Name:", "name_entry"),
            ("Phone Number:", "phone_number"),
            ("Citizen Address:", "address_entry"),
            ("Amount (BDT):", "amount_entry"),
            ("Service Type:", "service_entry"), 
            ("GRN No:", "grn_entry"),
            ("Application No:", "app_entry"), 
            ("Khatiyan/Plot:", "khatiyan_entry"),
            ("Transaction No:", "trans_entry"), 
            ("Remarks / Notes:", "remarks_entry"),
            ("Color Tag:", "color_entry"),
            ("DEO Name (Profile):", "deo_name")
        ]

        entries = {}
        for i, (label_text, key) in enumerate(fields):
            tk.Label(form_card, text=label_text, bg="#1e293b", fg="#cbd5e1", font=("Segoe UI", 9, "bold")).grid(row=i, column=0, sticky="w", pady=4)
            
            if key == "phone_number":
                ent = tk.Entry(form_card, font=("Segoe UI", 10), width=32, relief="solid", bd=1, bg="#0f172a", fg="#f8fafc", insertbackground="white", validate="key", validatecommand=vcmd_phone)
                ent.grid(row=i, column=1, pady=4, padx=5)
            elif key == "name_entry":
                ent = tk.Entry(form_card, font=("Segoe UI", 10), width=32, relief="solid", bd=1, bg="#0f172a", fg="#f8fafc", insertbackground="white", validate="key", validatecommand=vcmd_name)
                ent.grid(row=i, column=1, pady=4, padx=5)
            elif key == "color_entry":
                color_options = ["None", "Green (#dcfce7)", "Yellow (#fef9c3)", "Red (#fee2e2)", "Blue (#e0f2fe)", "Purple (#f3e8ff)"]
                ent = ttk.Combobox(form_card, values=color_options, font=("Segoe UI", 10), width=30, state="readonly")
                ent.set("None")
                ent.grid(row=i, column=1, pady=4, padx=5)
            else:
                ent = tk.Entry(form_card, font=("Segoe UI", 10), width=32, relief="solid", bd=1, bg="#0f172a", fg="#f8fafc", insertbackground="white")
                ent.grid(row=i, column=1, pady=4, padx=5)
            
            if key == "date_entry": 
                ent.insert(0, datetime.today().strftime('%d-%m-%Y'))
            
            if key == "deo_name":
                profile_data = self.get_profile_data()
                ent.insert(0, profile_data[3] if profile_data else "")
                ent.config(state="disabled", disabledbackground="#0f172a", disabledforeground="#f8fafc")

            entries[key] = ent

        def save_advanced_data():
            profile_data = self.get_profile_data()
            deo_val = profile_data[3] if profile_data else ""

            date_val = entries["date_entry"].get().strip()
            if not date_val:
                messagebox.showerror("Validation Error", "Date is required!", parent=entry_win)
                return

            phone_val = self.clean_phone_number(entries["phone_number"].get())
            if len(phone_val) != 10 and phone_val != "":
                messagebox.showerror("Validation Error", "Mobile number must be exactly 10 digits!", parent=entry_win)
                return

            raw_amount = entries["amount_entry"].get().strip()
            amount_val = float(raw_amount) if raw_amount else 0.0

            color_selection = entries["color_entry"].get()
            color_code = "None"
            if "Green" in color_selection: color_code = "green"
            elif "Yellow" in color_selection: color_code = "yellow"
            elif "Red" in color_selection: color_code = "red"
            elif "Blue" in color_selection: color_code = "blue"
            elif "Purple" in color_selection: color_code = "purple"

            data = [
                date_val,
                amount_val,
                entries["service_entry"].get().strip(),
                entries["grn_entry"].get().strip(),
                entries["app_entry"].get().strip(),
                entries["khatiyan_entry"].get().strip(),
                entries["trans_entry"].get().strip(),
                entries["name_entry"].get().strip(),
                entries["address_entry"].get().strip(),
                phone_val,
                deo_val,
                entries["remarks_entry"].get().strip(),
                color_code
            ]

            conn = get_db_connection()
            cursor = conn.cursor()
            cursor.execute("""
                INSERT INTO records (date, amount, service_type, grn_no, application_no, khatiyan_plot, transaction_no, citizen_name, citizen_address, phone_number, deo_name, remarks, color_tag)
                VALUES (?,?,?,?,?,?,?,?,?,?,?,?,?)
            """, tuple(data))
            conn.commit()
            conn.close()
            
            messagebox.showinfo("Success", "Record Saved Successfully!", parent=entry_win)
            entry_win.destroy()
            self.load_data()

        btn_box = tk.Frame(entry_win, bg="#0f172a")
        btn_box.pack(pady=12)
        
        tk.Button(btn_box, text="💾 Save Record", bg="#2563eb", fg="white", font=("Segoe UI", 10, "bold"), bd=0, padx=16, pady=8, cursor="hand2", command=save_advanced_data, activebackground="#1d4ed8", activeforeground="white").pack(side=tk.LEFT, padx=6)
        tk.Button(btn_box, text="❌ Cancel", bg="#475569", fg="white", font=("Segoe UI", 10, "bold"), bd=0, padx=16, pady=8, cursor="hand2", command=entry_win.destroy, activebackground="#334155", activeforeground="white").pack(side=tk.LEFT, padx=6)

        entry_win.update_idletasks()
        entry_win.geometry(f"{entry_win.winfo_reqwidth() + 50}x{entry_win.winfo_reqheight() + 20}")

    def show_context_menu(self, event):
        item = self.tree.identify_row(event.y)
        self.clicked_col = self.tree.identify_column(event.x)
        if item:
            if item not in self.tree.selection():
                self.tree.selection_set(item)
            self.context_menu.tk_popup(event.x_root, event.y_root)

    def copy_cell(self):
        selected_item = self.tree.selection()
        if selected_item and hasattr(self, 'clicked_col') and self.clicked_col:
            try:
                col_idx = int(self.clicked_col.replace('#', '')) - 1
                item_values = self.tree.item(selected_item[0], "values")
                if 0 <= col_idx < len(item_values):
                    cell_text = str(item_values[col_idx])
                    self.root.clipboard_clear()
                    self.root.clipboard_append(cell_text)
            except ValueError:
                pass

    def copy_row_event(self, event=None):
        selected_items = self.tree.selection()
        if not selected_items:
            return
        
        if len(selected_items) == 1 and hasattr(self, 'clicked_col') and self.clicked_col:
            try:
                col_idx = int(self.clicked_col.replace('#', '')) - 1
                item_values = self.tree.item(selected_items[0], "values")
                if 0 <= col_idx < len(item_values):
                    cell_text = str(item_values[col_idx])
                    self.root.clipboard_clear()
                    self.root.clipboard_append(cell_text)
                    return
            except ValueError:
                pass
        self.copy_row(show_msg=False)

    def copy_row(self, show_msg=True):
        selected_items = self.tree.selection()
        if selected_items:
            copied_data = []
            for item in selected_items:
                item_values = self.tree.item(item, "values")
                row_text = "\t".join(str(val) for val in item_values)
                copied_data.append(row_text)
            
            self.root.clipboard_clear()
            self.root.clipboard_append("\n".join(copied_data))
            if show_msg:
                messagebox.showinfo("Copied", f"{len(selected_items)} Row(s) data copied to clipboard successfully!")

    def edit_selected_from_menu(self):
        selected_items = self.tree.selection()
        if selected_items:
            self.tree.selection_set(selected_items[0])
            self.open_edit_window(None)

    def delete_selected_record(self):
        selected_items = self.tree.selection()
        if not selected_items:
            return
        
        if len(selected_items) > 1:
            if messagebox.askyesno("Confirm Delete", f"Are you sure you want to delete {len(selected_items)} selected records?", parent=self.root):
                conn = get_db_connection()
                cursor = conn.cursor()
                for item in selected_items:
                    record_id = self.tree.item(item, "values")[0]
                    cursor.execute("DELETE FROM records WHERE id=?", (record_id,))
                conn.commit()
                conn.close()
                messagebox.showinfo("Success", f"{len(selected_items)} Records Deleted Successfully!")
                self.load_data()
        else:
            item_values = self.tree.item(selected_items[0], "values")
            if not item_values:
                return
            record_id = item_values[0]

            if messagebox.askyesno("Confirm Delete", f"Are you sure you want to delete Record ID: {record_id}?", parent=self.root):
                conn = get_db_connection()
                cursor = conn.cursor()
                cursor.execute("DELETE FROM records WHERE id=?", (record_id,))
                conn.commit()
                conn.close()
                messagebox.showinfo("Success", "Record Deleted Successfully!")
                self.load_data()

    def open_edit_window(self, event):
        selected_items = self.tree.selection()
        if not selected_items:
            return
        
        item_values = self.tree.item(selected_items[0], "values")
        if not item_values:
            return

        record_id = item_values[0]
        
        edit_win = tk.Toplevel(self.root)
        edit_win.title(f"Modify Record - ID: {record_id}")
        edit_win.configure(bg="#1e293b")
        edit_win.grab_set()

        try:
            edit_win.iconbitmap("icon.ico")
        except Exception:
            pass

        tk.Label(edit_win, text=f"Edit Record ID: {record_id}", font=("Segoe UI", 13, "bold"), bg="#1e293b", fg="#38bdf8").pack(pady=15)

        form_frame = tk.Frame(edit_win, bg="#1e293b")
        form_frame.pack(padx=20, pady=5, fill=tk.BOTH, expand=True)

        edit_fields = [
            ("Date (DD-MM-YYYY):", "date"), 
            ("Citizen Name:", "name"),
            ("Phone Number:", "phone"),
            ("Citizen Address:", "address"),
            ("Amount:", "amount"),
            ("Service Type:", "service"), 
            ("GRN No:", "grn"),
            ("Application No:", "app"), 
            ("Khatiyan/Plot:", "khatiyan"),
            ("Transaction No:", "trans"),
            ("Remarks / Notes:", "remarks"),
            ("Color Tag:", "color")
        ]

        current_vals = [
            item_values[1], item_values[2], item_values[5], item_values[4], 
            item_values[3], item_values[6], item_values[7], item_values[8], 
            item_values[9], item_values[10], item_values[11], item_values[12]
        ]

        edit_entries = {}
        for i, (lbl_text, key) in enumerate(edit_fields):
            tk.Label(form_frame, text=lbl_text, bg="#1e293b", fg="#cbd5e1", font=("Segoe UI", 9, "bold")).grid(row=i, column=0, sticky="w", pady=4)
            
            if key == "color":
                color_options = ["None", "Green (#dcfce7)", "Yellow (#fef9c3)", "Red (#fee2e2)", "Blue (#e0f2fe)", "Purple (#f3e8ff)"]
                ent = ttk.Combobox(form_frame, values=color_options, font=("Segoe UI", 10), width=26, state="readonly")
                matched_val = "None"
                for opt in color_options:
                    if current_vals[i] in opt.lower():
                        matched_val = opt
                ent.set(matched_val)
                ent.grid(row=i, column=1, pady=4, padx=10)
            else:
                ent = tk.Entry(form_frame, font=("Segoe UI", 10), width=28, relief="solid", bd=1, bg="#0f172a", fg="#f8fafc", insertbackground="white")
                ent.grid(row=i, column=1, pady=4, padx=10)
                ent.insert(0, current_vals[i] if current_vals[i] else "")
            
            edit_entries[key] = ent

        def update_record():
            date_val = edit_entries["date"].get().strip()
            if not date_val:
                messagebox.showerror("Validation Error", "Date is required!", parent=edit_win)
                return

            phone_val = self.clean_phone_number(edit_entries["phone"].get())
            if len(phone_val) != 10 and phone_val != "":
                messagebox.showerror("Validation Error", "Mobile number must be exactly 10 digits!", parent=edit_win)
                return

            raw_amount = edit_entries["amount"].get().strip()
            amount_val = float(raw_amount) if raw_amount else 0.0

            color_selection = edit_entries["color"].get()
            color_code = "None"
            if "Green" in color_selection: color_code = "green"
            elif "Yellow" in color_selection: color_code = "yellow"
            elif "Red" in color_selection: color_code = "red"
            elif "Blue" in color_selection: color_code = "blue"
            elif "Purple" in color_selection: color_code = "purple"

            vals = [
                date_val, 
                amount_val, 
                edit_entries["service"].get().strip(), 
                edit_entries["grn"].get().strip(), 
                edit_entries["app"].get().strip(), 
                edit_entries["khatiyan"].get().strip(), 
                edit_entries["trans"].get().strip(), 
                edit_entries["name"].get().strip(), 
                edit_entries["address"].get().strip(), 
                phone_val, 
                edit_entries["remarks"].get().strip(), 
                color_code
            ]

            conn = get_db_connection()
            cursor = conn.cursor()
            cursor.execute("""
                UPDATE records SET date=?, amount=?, service_type=?, grn_no=?, application_no=?, khatiyan_plot=?, transaction_no=?, citizen_name=?, citizen_address=?, phone_number=?, remarks=?, color_tag=?
                WHERE id=?
            """, tuple(vals) + (record_id,))
            conn.commit()
            conn.close()

            messagebox.showinfo("Success", "Record Updated Successfully!", parent=edit_win)
            edit_win.destroy()
            self.load_data()

        def delete_from_edit():
            if messagebox.askyesno("Confirm Delete", f"Are you sure you want to delete Record ID: {record_id}?", parent=edit_win):
                conn = get_db_connection()
                cursor = conn.cursor()
                cursor.execute("DELETE FROM records WHERE id=?", (record_id,))
                conn.commit()
                conn.close()
                messagebox.showinfo("Success", "Record Deleted Successfully!", parent=edit_win)
                edit_win.destroy()
                self.load_data()

        btn_box = tk.Frame(edit_win, bg="#1e293b")
        btn_box.pack(pady=15)
        tk.Button(btn_box, text="🔄 Update Record", bg="#d97706", fg="white", font=("Segoe UI", 10, "bold"), bd=0, padx=12, pady=8, cursor="hand2", command=update_record, activebackground="#b45309", activeforeground="white").pack(side=tk.LEFT, padx=6)
        tk.Button(btn_box, text="🗑 Delete Record", bg="#dc2626", fg="white", font=("Segoe UI", 10, "bold"), bd=0, padx=12, pady=8, cursor="hand2", command=delete_from_edit, activebackground="#b91c1c", activeforeground="white").pack(side=tk.LEFT, padx=6)

        edit_win.update_idletasks()
        edit_win.geometry(f"{edit_win.winfo_reqwidth() + 50}x{edit_win.winfo_reqheight() + 20}")

    def open_report_window(self):
        report_win = tk.Toplevel(self.root)
        report_win.title("Professional Enterprise Report View & Dynamic Filters - Excel Style")
        report_win.geometry("1480x850")
        report_win.configure(bg="#f1f5f9")
        report_win.grab_set()
        
        try:
            report_win.iconbitmap("icon.ico")
        except Exception:
            pass
        
        current_filtered_records = []
        current_filter_text = "All Records"

        profile = self.get_profile_data()
        bsk_name = profile[0] if profile else "BSK Enterprise Unit"
        bsk_code = profile[1] if profile else "N/A"
        deo_name = profile[3] if profile else "N/A"

        banner = tk.Frame(report_win, bg="#1e293b", padx=20, pady=12)
        banner.pack(fill=tk.X)
        
        report_title_frame = tk.Frame(banner, bg="#1e293b")
        report_title_frame.pack(side=tk.LEFT)
        tk.Label(report_title_frame, text="BSK DATA MANAGEMENT & REPORTING PORTAL", font=("Segoe UI", 12, "bold"), fg="#38bdf8", bg="#1e293b").pack(anchor="w")
        tk.Label(report_title_frame, text="Bangla Sahayata Kendra | Government of West Bengal", font=("Segoe UI", 9), fg="#94a3b8", bg="#1e293b").pack(anchor="w")

        tk.Label(banner, text=f"DEO: {deo_name} | Date: {datetime.today().strftime('%d-%m-%Y')}", font=("Segoe UI", 9), fg="#94a3b8", bg="#1e293b").pack(side=tk.RIGHT, pady=3)

        summary_frame = tk.Frame(report_win, bg="#f1f5f9", padx=15, pady=10)
        summary_frame.pack(fill=tk.X)

        self.today_amt_lbl = tk.Label(summary_frame, text="Today Total: 0.00 BDT", font=("Segoe UI", 10, "bold"), bg="#0284c7", fg="white", padx=15, pady=10, relief="solid", bd=0)
        self.today_amt_lbl.pack(side=tk.LEFT, padx=5, expand=True, fill=tk.X)

        self.month_amt_lbl = tk.Label(summary_frame, text="Current Month Total: 0.00 BDT", font=("Segoe UI", 10, "bold"), bg="#7c3aed", fg="white", padx=15, pady=10, relief="solid", bd=0)
        self.month_amt_lbl.pack(side=tk.LEFT, padx=5, expand=True, fill=tk.X)

        self.filter_total_amt_lbl = tk.Label(summary_frame, text="Filtered/Selected Total: 0.00 BDT", font=("Segoe UI", 10, "bold"), bg="#d97706", fg="white", padx=15, pady=10, relief="solid", bd=0)
        self.filter_total_amt_lbl.pack(side=tk.LEFT, padx=5, expand=True, fill=tk.X)

        self.all_amt_lbl = tk.Label(summary_frame, text="All Total Amount: 0.00 BDT", font=("Segoe UI", 10, "bold"), bg="#059669", fg="white", padx=15, pady=10, relief="solid", bd=0)
        self.all_amt_lbl.pack(side=tk.LEFT, padx=5, expand=True, fill=tk.X)

        control_bar = tk.Frame(report_win, bg="#e2e8f0", padx=15, pady=8)
        control_bar.pack(fill=tk.X)

        tk.Label(control_bar, text="Filter Type:", bg="#e2e8f0", font=("Segoe UI", 9, "bold")).pack(side=tk.LEFT, padx=2)
        filter_mode_combo = ttk.Combobox(control_bar, values=["All Total", "Day Wise", "Month Wise", "Year Wise"], font=("Segoe UI", 9), width=14, state="readonly")
        filter_mode_combo.set("All Total")
        filter_mode_combo.pack(side=tk.LEFT, padx=4)

        tk.Label(control_bar, text="Search:", bg="#e2e8f0", font=("Segoe UI", 9, "bold")).pack(side=tk.LEFT, padx=(10, 2))
        rep_search_ent = tk.Entry(control_bar, font=("Segoe UI", 9), width=16)
        rep_search_ent.pack(side=tk.LEFT, padx=2)

        def apply_report_filters():
            mode = filter_mode_combo.get()
            query = rep_search_ent.get().strip().lower()
            load_report_data(mode, query)

        filter_mode_combo.bind("<<ComboboxSelected>>", lambda e: apply_report_filters())
        rep_search_ent.bind("<KeyRelease>", lambda e: apply_report_filters())

        def reset_report_filters():
            filter_mode_combo.set("All Total")
            rep_search_ent.delete(0, tk.END)
            load_report_data("All Total", "")

        tk.Button(control_bar, text="🔍 Apply Main Filter", bg="#2563eb", fg="white", font=("Segoe UI", 8, "bold"), bd=0, padx=8, pady=4, cursor="hand2", command=apply_report_filters).pack(side=tk.LEFT, padx=6)
        tk.Button(control_bar, text="🔄 Reset All", bg="#64748b", fg="white", font=("Segoe UI", 8, "bold"), bd=0, padx=8, pady=4, cursor="hand2", command=reset_report_filters).pack(side=tk.LEFT, padx=2)

        def export_pdf_report():
            file_path = filedialog.asksaveasfilename(defaultextension=".pdf", filetypes=[("PDF Files", "*.pdf")], initialfile="bsk_professional_filtered_report.pdf", parent=report_win)
            if file_path:
                try:
                    doc = SimpleDocTemplate(file_path, pagesize=letter, rightMargin=30, leftMargin=30, topMargin=30, bottomMargin=30)
                    elements = []
                    styles = getSampleStyleSheet()
                    
                    title_style = ParagraphStyle('TitleStyle', parent=styles['Heading1'], fontName='Helvetica-Bold', fontSize=12, textColor=colors.HexColor("#1e293b"), alignment=1)
                    sub_style = ParagraphStyle('SubStyle', parent=styles['Normal'], fontName='Helvetica', fontSize=9, textColor=colors.HexColor("#64748b"), alignment=1)

                    elements.append(Paragraph("<b>BSK DATA MANAGEMENT & REPORTING PORTAL</b>", title_style))
                    elements.append(Paragraph("Bangla Sahayata Kendra | Government of West Bengal", sub_style))
                    elements.append(Paragraph(f"<b>Unit:</b> {bsk_name} ({bsk_code}) | <b>Generated By:</b> {deo_name}", sub_style))
                    elements.append(Paragraph(f"<b>Data Filter Range:</b> {current_filter_text}", sub_style))
                    elements.append(Spacer(1, 15))

                    table_data = [["ID", "Date", "Citizen Name", "Amount (BDT)", "Service Type", "DEO Name", "Remarks"]]
                    for r in current_filtered_records:
                        table_data.append([str(r[0]), str(r[1]), str(r[8]), str(r[2]), str(r[3]), str(r[11]), str(r[12])])

                    pdf_table = Table(table_data, colWidths=[35, 70, 115, 75, 100, 90, 85])
                    pdf_table.setStyle(TableStyle([
                        ('BACKGROUND', (0, 0), (-1, 0), colors.HexColor("#1e293b")),
                        ('TEXTCOLOR', (0, 0), (-1, 0), colors.whitesmoke),
                        ('ALIGN', (0, 0), (-1, -1), 'CENTER'),
                        ('FONTNAME', (0, 0), (-1, 0), 'Helvetica-Bold'),
                        ('FONTSIZE', (0, 0), (-1, 0), 9),
                        ('BOTTOMPADDING', (0, 0), (-1, 0), 8),
                        ('BACKGROUND', (0, 1), (-1, -1), colors.HexColor("#f8fafc")),
                        ('GRID', (0, 0), (-1, -1), 0.5, colors.HexColor("#cbd5e1")),
                        ('FONTNAME', (0, 1), (-1, -1), 'Helvetica'),
                        ('FONTSIZE', (0, 1), (-1, -1), 8),
                    ]))
                    elements.append(pdf_table)
                    doc.build(elements)
                    messagebox.showinfo("Success", f"Professional PDF Report Generated Successfully at:\n{file_path}", parent=report_win)
                except Exception as e:
                    messagebox.showerror("Error", f"Failed to generate PDF:\n{str(e)}", parent=report_win)

        def export_excel_report():
            file_path = filedialog.asksaveasfilename(defaultextension=".xlsx", filetypes=[("Excel Files", "*.xlsx")], initialfile="bsk_professional_filtered_report.xlsx", parent=report_win)
            if file_path:
                try:
                    columns = ["ID", "Date", "Amount", "Service Type", "GRN", "App No", "Khatiyan", "Trans No", "Name", "Address", "Phone", "DEO", "Remarks", "Color"]
                    data = []
                    for r in current_filtered_records:
                        data.append(list(r))
                    df = pd.DataFrame(data, columns=columns)
                    df['Data Filter Range Applied'] = current_filter_text
                    df.to_excel(file_path, index=False)
                    messagebox.showinfo("Success", f"Excel Report Generated Successfully at:\n{file_path}", parent=report_win)
                except Exception as e:
                    messagebox.showerror("Error", f"Failed to generate Excel:\n{str(e)}", parent=report_win)

        tk.Button(control_bar, text="📄 Export PDF Report", bg="#dc2626", fg="white", font=("Segoe UI", 9, "bold"), bd=0, padx=10, pady=5, cursor="hand2", command=export_pdf_report).pack(side=tk.RIGHT, padx=4)
        tk.Button(control_bar, text="📊 Export Excel Report", bg="#10b981", fg="white", font=("Segoe UI", 9, "bold"), bd=0, padx=10, pady=5, cursor="hand2", command=export_excel_report).pack(side=tk.RIGHT, padx=4)

        notebook = ttk.Notebook(report_win)
        notebook.pack(fill=tk.BOTH, expand=True, padx=15, pady=10)

        tab_all = tk.Frame(notebook, bg="white")
        tab_service = tk.Frame(notebook, bg="white")
        tab_deo = tk.Frame(notebook, bg="white")

        notebook.add(tab_all, text=" 📊 All Records View (Excel Spreadsheet Style) ")
        notebook.add(tab_service, text=" 📁 Category: Service Type Wise ")
        notebook.add(tab_deo, text=" 👤 Category: DEO Name Wise ")

        columns_all = ("Sl", "Date", "Name", "Amount", "Address", "Phone", "Service Type", "GRN", "App No", "Khatiyan", "Trans No", "Remarks", "Color", "DEO")
        columns_service = ("Service Type", "Total Transactions", "Total Collection Amount (BDT)")
        columns_deo = ("DEO Name", "Total Transactions", "Total Collection Amount (BDT)")

        spread_all_container = tk.Frame(tab_all, bg="white")
        spread_all_container.pack(fill=tk.BOTH, expand=True, padx=5, pady=5)

        spread_service_container = tk.Frame(tab_service, bg="white")
        spread_service_container.pack(fill=tk.BOTH, expand=True, padx=5, pady=5)

        spread_deo_container = tk.Frame(tab_deo, bg="white")
        spread_deo_container.pack(fill=tk.BOTH, expand=True, padx=5, pady=5)

        def open_category_details(cat_type, cat_value):
            detail_win = tk.Toplevel(report_win)
            detail_win.title(f"Detailed Records - {cat_type}: {cat_value} (Excel Style)")
            detail_win.geometry("1400x700")
            detail_win.configure(bg="#0f172a")
            
            try:
                detail_win.iconbitmap("icon.ico")
            except Exception:
                pass

            tk.Label(detail_win, text=f"Detailed View for {cat_type}: {cat_value}", font=("Segoe UI", 14, "bold"), bg="#0f172a", fg="#38bdf8").pack(pady=10)

            filtered_details = []
            for r in current_filtered_records:
                r_service = r[3] if r[3] else "General Service"
                r_deo = r[11] if r[11] else "Default DEO"
                if (cat_type == "Service Type" and r_service == cat_value) or \
                   (cat_type == "DEO Name" and r_deo == cat_value):
                    filtered_details.append((r[0], r[1], r[8], r[2], r[9], r[10], r[3], r[4], r[5], r[6], r[7], r[12], r[13], r[11]))

            sheet_frame = tk.Frame(detail_win, bg="white")
            sheet_frame.pack(fill=tk.BOTH, expand=True, padx=15, pady=10)
            
            ExcelSpreadsheet(sheet_frame, columns_all, filtered_details).pack(fill=tk.BOTH, expand=True)

            tk.Label(detail_win, text="💡 Tip: Click and drag across multiple cells to select them, then press Ctrl+C to copy to Excel.", bg="#0f172a", fg="#94a3b8", font=("Segoe UI", 9)).pack(pady=5)

        def load_report_data(mode="All Total", query=""):
            nonlocal current_filtered_records, current_filter_text

            conn = get_db_connection()
            cursor = conn.cursor()
            cursor.execute("SELECT id, date, amount, service_type, grn_no, application_no, khatiyan_plot, transaction_no, citizen_name, citizen_address, phone_number, deo_name, remarks, color_tag FROM records ORDER BY id DESC")
            all_records = cursor.fetchall()
            conn.close()

            today_str = datetime.today().strftime('%d-%m-%Y')
            current_month_str = datetime.today().strftime('%m-%Y')

            today_total = 0.0
            month_total = 0.0
            all_total = 0.0
            filter_total = 0.0

            service_summary = {}
            deo_summary = {}
            filtered_records = []

            for r in all_records:
                r_id, r_date, r_amount, r_service, r_grn, r_app, r_khatiyan, r_trans, r_name, r_address, r_phone, r_deo, r_remarks, r_color = r
                
                try:
                    amt_val = float(r_amount)
                except:
                    amt_val = 0.0

                all_total += amt_val

                if r_date.strip() == today_str:
                    today_total += amt_val

                if current_month_str in r_date.strip():
                    month_total += amt_val

                include = True
                r_date_clean = r_date.strip()

                if mode == "Day Wise":
                    if r_date_clean != today_str:
                        include = False
                elif mode == "Month Wise":
                    if current_month_str not in r_date_clean:
                        include = False
                elif mode == "Year Wise":
                    current_year_str = datetime.today().strftime('%Y')
                    if not r_date_clean.endswith(current_year_str):
                        include = False

                row_str = f"{r_id} {r_date} {r_amount} {r_service} {r_grn} {r_app} {r_khatiyan} {r_trans} {r_name} {r_address} {r_phone} {r_deo} {r_remarks}".lower()
                if query and query not in row_str:
                    include = False

                if include:
                    filtered_records.append(r)
                    filter_total += amt_val

                    s_key = r_service if r_service else "General Service"
                    if s_key not in service_summary:
                        service_summary[s_key] = [0, 0.0]
                    service_summary[s_key][0] += 1
                    service_summary[s_key][1] += amt_val

                    d_key = r_deo if r_deo else "Default DEO"
                    if d_key not in deo_summary:
                        deo_summary[d_key] = [0, 0.0]
                    deo_summary[d_key][0] += 1
                    deo_summary[d_key][1] += amt_val
            
            current_filtered_records = filtered_records
            
            filter_str = mode
            if query:
                filter_str += f" | Search: '{query}'"
            current_filter_text = filter_str

            self.today_amt_lbl.config(text=f"Today Total: {today_total:.2f} BDT")
            self.month_amt_lbl.config(text=f"Current Month Total: {month_total:.2f} BDT")
            self.filter_total_amt_lbl.config(text=f"Filtered Total: {filter_total:.2f} BDT")
            self.all_amt_lbl.config(text=f"All Total Amount: {all_total:.2f} BDT")

            for widget in spread_all_container.winfo_children():
                widget.destroy()
            formatted_all_data = []
            for r in filtered_records:
                formatted_all_data.append((r[0], r[1], r[8], r[2], r[9], r[10], r[3], r[4], r[5], r[6], r[7], r[12], r[13], r[11]))
            ExcelSpreadsheet(spread_all_container, columns_all, formatted_all_data).pack(fill=tk.BOTH, expand=True)

            for widget in spread_service_container.winfo_children():
                widget.destroy()
            service_data_list = []
            for srv, data in service_summary.items():
                service_data_list.append([srv, data[0], f"{data[1]:.2f}"])
            
            srv_sheet = ExcelSpreadsheet(spread_service_container, columns_service, service_data_list)
            srv_sheet.pack(fill=tk.BOTH, expand=True)

            for widget in spread_deo_container.winfo_children():
                widget.destroy()
            deo_data_list = []
            for deo, data in deo_summary.items():
                deo_data_list.append([deo, data[0], f"{data[1]:.2f}"])
            
            deo_sheet = ExcelSpreadsheet(spread_deo_container, columns_deo, deo_data_list)
            deo_sheet.pack(fill=tk.BOTH, expand=True)

        load_report_data()

    def load_data(self):
        for row in self.tree.get_children(): 
            self.tree.delete(row)
        conn = get_db_connection()
        cursor = conn.cursor()
        cursor.execute("SELECT id, date, amount, service_type, grn_no, application_no, khatiyan_plot, transaction_no, citizen_name, citizen_address, phone_number, deo_name, remarks, color_tag FROM records ORDER BY id DESC")
        
        for index, row in enumerate(cursor.fetchall()):
            r_id, r_date, r_amount, r_service, r_grn, r_app, r_khatiyan, r_trans, r_name, r_address, r_phone, r_deo, r_remarks, r_color = row
            display_row = (r_id, r_date, r_name, r_amount, r_address, r_phone, r_service, r_grn, r_app, r_khatiyan, r_trans, r_remarks, r_color, r_deo)
            
            c_tag = r_color if r_color in ['green', 'yellow', 'red', 'blue', 'purple'] else ('evenrow' if index % 2 == 0 else 'oddrow')
            self.tree.insert("", tk.END, values=display_row, tags=(c_tag,))
        conn.close()

    def filter_data(self, event=None):
        search_query = self.search_entry.get().strip().lower()
        date_query = self.filter_date_entry.get().strip()

        for row in self.tree.get_children(): 
            self.tree.delete(row)

        conn = get_db_connection()
        cursor = conn.cursor()
        cursor.execute("SELECT id, date, amount, service_type, grn_no, application_no, khatiyan_plot, transaction_no, citizen_name, citizen_address, phone_number, deo_name, remarks, color_tag FROM records ORDER BY id DESC")
        rows = cursor.fetchall()
        conn.close()

        for index, row in enumerate(rows):
            r_id, r_date, r_amount, r_service, r_grn, r_app, r_khatiyan, r_trans, r_name, r_address, r_phone, r_deo, r_remarks, r_color = map(str, row)
            
            row_combined = f"{r_id} {r_date} {r_amount} {r_service} {r_grn} {r_app} {r_khatiyan} {r_trans} {r_name} {r_address} {r_phone} {r_deo} {r_remarks}".lower()
            text_match = (search_query in row_combined)
            date_match = (date_query in r_date) or (date_query == "")

            if text_match and date_match:
                display_row = (r_id, r_date, r_name, r_amount, r_address, r_phone, r_service, r_grn, r_app, r_khatiyan, r_trans, r_remarks, r_color, r_deo)
                c_tag = r_color if r_color in ['green', 'yellow', 'red', 'blue', 'purple'] else ('evenrow' if index % 2 == 0 else 'oddrow')
                self.tree.insert("", tk.END, values=display_row, tags=(c_tag,))

    def clear_filter(self):
        self.search_entry.delete(0, tk.END)
        self.filter_date_entry.delete(0, tk.END)
        self.load_data()

    def export_excel(self):
        conn = get_db_connection()
        df = pd.read_sql_query("SELECT * FROM records", conn)
        conn.close()
        file_name = "bsk_enterprise_report.xlsx"
        df.to_excel(file_name, index=False)
        messagebox.showinfo("Export Successful", f"Excel Report Saved as {file_name}")

    def download_sample_format(self):
        columns = [
            'date', 'amount', 'service_type', 'grn_no', 'application_no', 
            'khatiyan_plot', 'transaction_no', 'citizen_name', 'citizen_address', 
            'phone_number', 'deo_name', 'remarks', 'color_tag'
        ]
        df = pd.DataFrame(columns=columns)
        file_path = filedialog.asksaveasfilename(defaultextension=".xlsx", filetypes=[("Excel Files", "*.xlsx")], initialfile="sample_template.xlsx")
        if file_path:
            df.to_excel(file_path, index=False)
            messagebox.showinfo("Success", f"Sample Template Saved Successfully at:\n{file_path}")

    def import_excel(self):
        file_path = filedialog.askopenfilename(filetypes=[("Excel Files", "*.xlsx *.xls")])
        if file_path:
            try:
                df = pd.read_excel(file_path)
                df.dropna(how='all', inplace=True)
                
                expected_columns = [
                    'date', 'amount', 'service_type', 'grn_no', 'application_no', 
                    'khatiyan_plot', 'transaction_no', 'citizen_name', 'citizen_address', 
                    'phone_number', 'deo_name', 'remarks', 'color_tag'
                ]
                
                for col in expected_columns:
                    if col not in df.columns:
                        df[col] = ""

                df = df[expected_columns]

                if 'amount' in df.columns:
                    df['amount'] = pd.to_numeric(df['amount'], errors='coerce').fillna(0.0)

                text_columns = ['date', 'service_type', 'grn_no', 'application_no', 
                                'khatiyan_plot', 'transaction_no', 'citizen_name', 'citizen_address', 
                                'phone_number', 'deo_name', 'remarks', 'color_tag']
                for col in text_columns:
                    if col in df.columns:
                        df[col] = df[col].fillna("").astype(str)
                        df[col] = df[col].replace('nan', '')

                if 'phone_number' in df.columns:
                    df['phone_number'] = df['phone_number'].apply(self.clean_phone_number)

                if 'date' in df.columns:
                    def parse_date(val):
                        if pd.isna(val) or str(val).strip() == "" or str(val).strip().lower() == "nan":
                            return datetime.today().strftime('%d-%m-%Y')
                        if isinstance(val, pd.Timestamp) or isinstance(val, datetime):
                            return val.strftime('%d-%m-%Y')
                        val_str = str(val).strip()
                        try:
                            with warnings.catch_warnings():
                                warnings.simplefilter("ignore", category=UserWarning)
                                parsed_dt = pd.to_datetime(val_str, errors='raise', dayfirst=True)
                            return parsed_dt.strftime('%d-%m-%Y')
                        except:
                            return val_str

                    df['date'] = df['date'].apply(parse_date)

                conn = get_db_connection()
                df.to_sql('records', conn, if_exists='append', index=False)
                conn.close()
                
                self.load_data()
                messagebox.showinfo("Success", "Excel-এ ডেটা সফলভাবে ইমপোর্ট হয়েছে!")
            except Exception as e:
                messagebox.showerror("Import Error", f"Failed to import data due to format issue:\n{str(e)}")

if __name__ == "__main__":
    root = tk.Tk()
    app = DataEntryApp(root)
    root.mainloop()
