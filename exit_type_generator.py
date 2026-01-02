import tkinter as tk
from tkinter import messagebox, filedialog
from tkinter import ttk
from reportlab.lib.pagesizes import letter
from reportlab.pdfgen import canvas
from reportlab.lib.units import inch
from reportlab.lib import colors
from reportlab.pdfbase import pdfmetrics
from reportlab.pdfbase.ttfonts import TTFont
import os
import json
from datetime import datetime
import platform
import subprocess
import sys
from PIL import Image, ImageTk



ALIGNMENT_FILE = "alignment_settings.json"

# ------------------- TOOLTIP CLASS -------------------
class ToolTip(object):
    def __init__(self, widget, text="widget info"):
        self.widget = widget
        self.text = text
        self.tipwindow = None
        self.id = None
        self.x = self.y = 0
        widget.bind("<Enter>", self.show_tip)
        widget.bind("<Leave>", self.hide_tip)

    def show_tip(self, event=None):
        "Display text in tooltip window"
        if self.tipwindow or not self.text:
            return
        x = self.widget.winfo_rootx() + 40
        y = self.widget.winfo_rooty() + 20
        self.tipwindow = tw = tk.Toplevel(self.widget)
        tw.wm_overrideredirect(True)
        tw.wm_geometry("+%d+%d" % (x, y))
        label = tk.Label(
            tw,
            text=self.text,
            justify=tk.LEFT,
            background="#000",
            foreground="white",
            relief=tk.SOLID,
            borderwidth=1,
            font=("Arial", 10)
        )
        label.pack(ipadx=6, ipady=3)

    def hide_tip(self, event=None):
        tw = self.tipwindow
        self.tipwindow = None
        if tw:
            tw.destroy()
# -------------------------------------------------------
# def resource_path(relative_path):
#     """Return absolute path for PyInstaller."""
#     try:
#         base_path = sys._MEIPASS
#     except Exception:
#         base_path = os.path.abspath(".")
# 
#     return os.path.join(base_path, relative_path)

def resource_path(relative_path):
    """
    Return the absolute path to a resource that works for both
    PyInstaller (one-file & one-folder) and normal execution.
    """

    # If running as a PyInstaller bundle
    if hasattr(sys, '_MEIPASS'):
        base_path = sys._MEIPASS
    else:
        # Running in development mode
        base_path = os.path.dirname(os.path.abspath(__file__))

    return os.path.join(base_path, relative_path)

class ExitTypeGeneratorApp:

    

    def __init__(self, root):
        # Load messages or state
        self.root = root
        self.root.title("Exit Type Generator")
        self.root.geometry("1100x880")
        
        # Configure style
        style = ttk.Style()
        style.theme_use('clam')
        
        # Professional Colors
        bg_main = "#eaeff2"       # Soft bluish-gray background
        bg_card = "#ffffff"       # White card
        accent = "#3498db"        # Primary Blue
        accent_hover = "#2980b9"
        text_primary = "#2c3e50"  # Dark Slate
        text_secondary = "#7f8c8d" # Gray
        
        self.root.configure(bg=bg_main)
        
        # Style Definitions
        style.configure("Main.TFrame", background=bg_main)
        style.configure("Card.TFrame", background=bg_card, relief="flat")
        
        # Labels
        style.configure("CardLabel.TLabel", background=bg_card, foreground=text_secondary, font=("Segoe UI Semibold", 9))
        style.configure("Header.TLabel", background=bg_main, foreground=text_primary, font=("Segoe UI", 18, "bold"))
        
        # Entries
        style.configure("Card.TEntry", fieldbackground="#f8f9fa", padding=8, borderwidth=1, relief="solid", bordercolor="#dcdde1")
        style.map("Card.TEntry", bordercolor=[("focus", accent)])
        
        # Buttons
        style.configure("Primary.TButton", font=("Segoe UI", 10, "bold"), background=accent, foreground="white", padding=8)
        style.map("Primary.TButton", background=[("active", accent_hover)])
        
        style.configure("Secondary.TButton", font=("Segoe UI", 10), background="#ecf0f1", foreground=text_primary, padding=6)

        # === Scrollable Container ===
        # 1. Main Canvas
        self.canvas = tk.Canvas(root, borderwidth=0, background=bg_main, highlightthickness=0)
        self.scrollbar = ttk.Scrollbar(root, orient="vertical", command=self.canvas.yview)
        
        # 2. Scrollable Frame
        self.scrollable_frame = ttk.Frame(self.canvas, style="Main.TFrame")
        
        # 3. Configure scrolling
        self.scrollable_frame.bind(
            "<Configure>",
            lambda e: self.canvas.configure(scrollregion=self.canvas.bbox("all"))
        )
        self.canvas.create_window((0, 0), window=self.scrollable_frame, anchor="n") # Anchor North for centering
        self.canvas.configure(yscrollcommand=self.scrollbar.set)
        
        # 4. Pack
        self.scrollbar.pack(side="right", fill="y")
        self.canvas.pack(side="left", fill="both", expand=True)
        self.root.bind_all("<MouseWheel>", self._on_mousewheel)

        # Make the scrollable frame fill the canvas width
        self.frame_id = self.canvas.create_window((0, 0), window=self.scrollable_frame, anchor="nw")
        
        def _configure_frame_width(event):
            # Update the width of the frame to match the canvas
            canvas_width = event.width
            self.canvas.itemconfig(self.frame_id, width=canvas_width)
            
        self.canvas.bind("<Configure>", _configure_frame_width)

        # Load Icons
        try:
            self.icon_save = ImageTk.PhotoImage(Image.open(resource_path("icons/save.png")).resize((24, 24), Image.LANCZOS))
            self.icon_saveas = ImageTk.PhotoImage(Image.open(resource_path("icons/saveas.png")).resize((24, 24), Image.LANCZOS))
            self.icon_clear = ImageTk.PhotoImage(Image.open(resource_path("icons/clear.png")).resize((24, 24), Image.LANCZOS))
            self.icon_load = ImageTk.PhotoImage(Image.open(resource_path("icons/add.png")).resize((24, 24), Image.LANCZOS))
            self.icon_print = ImageTk.PhotoImage(Image.open(resource_path("icons/print.png")).resize((24, 24), Image.LANCZOS))
        except Exception:
            # Fallback if icons missing
            self.icon_save = None
            self.icon_saveas = None
            self.icon_clear = None
            self.icon_load = None
            self.icon_print = None

        # Track last created PDF path (for printing)
        self.last_pdf_path = None

        # Load settings
        (
            self.offset_x,
            self.offset_y,
            self.default_folder,
            self.filename_base,
        ) = self.load_settings()

        # Register Fonts
        try:
            verdana_path = "C:\\Windows\\Fonts\\verdana.ttf"
            verdana_bold_path = "C:\\Windows\\Fonts\\verdanab.ttf"
            pdfmetrics.registerFont(TTFont("Verdana", verdana_path))
            pdfmetrics.registerFont(TTFont("Verdana-Bold", verdana_bold_path))
        except Exception:
            pass

        # Try setting window icon
        try:
            icon_path = resource_path("exittype.ico")
            self.root.iconbitmap(icon_path)
        except Exception:
            pass

        # === Header ===
        # Removed to save space per user request
        # header_frame = ttk.Frame(self.scrollable_frame, style="Main.TFrame")
        # header_frame.pack(pady=(30, 20))
        # ttk.Label(header_frame, text="Exit Type Generator", style="Header.TLabel").pack()

        # === THE CARD (Main Content) ===
        # A centered white card that fills most of the width
        self.card = tk.Frame(self.scrollable_frame, bg="#ffffff", padx=30, pady=30, relief="solid", bd=1)
        self.card.configure(highlightbackground="#dcdde1", highlightthickness=1)
        self.card.pack(padx=40, pady=(0, 30), ipadx=10, ipady=10, fill=tk.X, expand=True)

        # 1. Controls (Settings)
        # Use a grid layout for the controls
        controls_frame = tk.Frame(self.card, bg="#ffffff")
        controls_frame.pack(fill=tk.X, pady=(0, 20))
        
        # Alignment
        align_btn = ttk.Button(controls_frame, text="⚙ Alignment", command=self.open_alignment_window, style="Secondary.TButton")
        align_btn.pack(side=tk.LEFT)
        
        # Folder
        folder_frame = tk.Frame(controls_frame, bg="#ffffff")
        folder_frame.pack(side=tk.RIGHT)
        ttk.Label(folder_frame, text="Save Folder:", style="CardLabel.TLabel").pack(side=tk.LEFT, padx=(10, 5))
        self.folder_var = tk.StringVar(value=self.default_folder or "")
        ttk.Entry(folder_frame, textvariable=self.folder_var, width=25, style="Card.TEntry").pack(side=tk.LEFT)
        ttk.Button(folder_frame, text="Browse", command=self.choose_folder, style="Secondary.TButton").pack(side=tk.LEFT, padx=5)

        # Filename
        fname_frame = tk.Frame(self.card, bg="#ffffff")
        fname_frame.pack(fill=tk.X, pady=(0, 20))
        ttk.Label(fname_frame, text="Filename Base:", style="CardLabel.TLabel").pack(side=tk.LEFT, padx=(0, 5))
        self.filename_base_var = tk.StringVar(value=self.filename_base or "exit")
        ttk.Entry(fname_frame, textvariable=self.filename_base_var, width=20, style="Card.TEntry").pack(side=tk.LEFT)
        ttk.Button(fname_frame, text="Save Settings", command=self.save_all_settings, style="Secondary.TButton").pack(side=tk.LEFT, padx=10)

        ttk.Separator(self.card, orient="horizontal").pack(fill=tk.X, pady=10)

        # 2. Form Fields (Grid Layout - 6 Columns)
        self.form_frame = tk.Frame(self.card, bg="#ffffff")
        self.form_frame.pack(fill=tk.X, pady=10)
        
        # Configure columns to expand (6 possible slots for flexibility)
        for c in range(6):
            self.form_frame.columnconfigure(c, weight=1)

        # Define Layout: List of Rows, where each Row is a list of (Label, ColumnSpan)
        layout_rows = [
            # Row 1: Inspection details (2 items -> span 3 each)
            [("Inspection No", 3), ("Inspection Date", 3)],
            
            # Row 2: Exporter (Wide -> span 6)
            [("Exporter", 6)],
            
            # Row 3: Bill No, BOE Date, AWB (User Req: 3 items -> span 2 each)
            [("Bill No", 2), ("BOE Date", 2), ("Air way Bill NO", 2)],
            
            # Row 4: COO, Point of Exit, Destination (User Req: 3 items -> span 2 each)
            [("Country of Origin", 2), ("Point of Exit", 2), ("Destination", 2)],
            
            # Row 5: Quantity (5%) + Description (95%)
            [("Quantity|Description", 6)],
            
            # Row 6: All Logistics (Total Q, Weight, Container, Seal)
            [("Final_Row_Group", 6)],
        ]
        
        self.entries = {}
        
        current_row = 0
        
        for row_items in layout_rows:
            current_col = 0
            for label_text, span in row_items:
                
                # Special Case 1: Quantity + Description (5% / 95%)
                if label_text == "Quantity|Description":
                    # Create a nested frame
                    group_frame = tk.Frame(self.form_frame, bg="#ffffff")
                    group_frame.grid(row=current_row, column=0, columnspan=6, sticky="ew", padx=10, pady=8)
                    
                    group_frame.columnconfigure(0, weight=1)  # 5%
                    group_frame.columnconfigure(1, weight=20) # 95%
                    
                    # Qty
                    qty_container = tk.Frame(group_frame, bg="#ffffff")
                    qty_container.grid(row=0, column=0, sticky="ew", padx=(0, 10))
                    ttk.Label(qty_container, text="Quantity", style="CardLabel.TLabel").pack(anchor="w", pady=(0, 4))
                    qty_entry = tk.Text(
                        qty_container, height=4, width=5, 
                        font=("Segoe UI", 10), 
                        relief="flat", bd=0,
                        highlightthickness=1, highlightcolor="#3498db", highlightbackground="#dcdde1",
                        bg="#f8f9fa",
                        undo=True, maxundo=-1
                    )
                    qty_entry.bind("<Tab>", self.focus_next_window)
                    qty_entry.pack(fill=tk.X)
                    self.entries["Quantity"] = qty_entry

                    # Desc
                    desc_container = tk.Frame(group_frame, bg="#ffffff")
                    desc_container.grid(row=0, column=1, sticky="ew")
                    ttk.Label(desc_container, text="Description", style="CardLabel.TLabel").pack(anchor="w", pady=(0, 4))
                    desc_entry = tk.Text(
                        desc_container, height=4, width=30, 
                        font=("Segoe UI", 10), 
                        relief="flat", bd=0,
                        highlightthickness=1, highlightcolor="#3498db", highlightbackground="#dcdde1",
                        bg="#f8f9fa",
                        undo=True, maxundo=-1
                    )
                    desc_entry.bind("<Tab>", self.focus_next_window)
                    desc_entry.pack(fill=tk.X)
                    self.entries["Description"] = desc_entry
                    self.desc_text = desc_entry 
                    
                    continue

                # Special Case 2: Final Row (Total Q | Weight | Container | Seal)
                if label_text == "Final_Row_Group":
                    # Nested frame
                    fin_frame = tk.Frame(self.form_frame, bg="#ffffff")
                    fin_frame.grid(row=current_row, column=0, columnspan=6, sticky="ew", padx=10, pady=8)
                    
                    # Distribution: Qty(2), Weight(1), Cont(4), Seal(2)
                    # Total 9 shares.
                    # Weight ~ 11%
                    fin_frame.columnconfigure(0, weight=2)
                    fin_frame.columnconfigure(1, weight=1)
                    fin_frame.columnconfigure(2, weight=4)
                    fin_frame.columnconfigure(3, weight=2)
                    
                    # 1. Total Quantity
                    tq_cont = tk.Frame(fin_frame, bg="#ffffff")
                    tq_cont.grid(row=0, column=0, sticky="ew", padx=(0, 10))
                    ttk.Label(tq_cont, text="Total Quantity", style="CardLabel.TLabel").pack(anchor="w", pady=(0, 4))
                    tq_entry = tk.Text(tq_cont, height=2, width=15, font=("Segoe UI", 10), 
                        relief="flat", bd=0, 
                        highlightthickness=1, highlightcolor="#3498db", highlightbackground="#dcdde1", 
                        bg="#f8f9fa",
                        undo=True, maxundo=-1)
                    tq_entry.bind("<Tab>", self.focus_next_window)
                    tq_entry.pack(fill=tk.X)
                    self.entries["Total Quantity"] = tq_entry

                    # 2. Total Weight
                    tw_cont = tk.Frame(fin_frame, bg="#ffffff")
                    tw_cont.grid(row=0, column=1, sticky="ew", padx=(0, 10))
                    ttk.Label(tw_cont, text="Total Weight", style="CardLabel.TLabel").pack(anchor="w", pady=(0, 4))
                    tw_entry = ttk.Entry(tw_cont, style="Card.TEntry", width=10)
                    tw_entry.pack(fill=tk.X)
                    self.entries["Total Weight"] = tw_entry
                    
                    # 3. Container
                    c_cont = tk.Frame(fin_frame, bg="#ffffff")
                    c_cont.grid(row=0, column=2, sticky="ew", padx=(0, 10))
                    ttk.Label(c_cont, text="Container/Vehicle No", style="CardLabel.TLabel").pack(anchor="w", pady=(0, 4))
                    c_entry = ttk.Entry(c_cont, style="Card.TEntry", width=20)
                    c_entry.pack(fill=tk.X)
                    self.entries["Container/Vehicle No"] = c_entry
                    
                    # 4. Seal
                    s_cont = tk.Frame(fin_frame, bg="#ffffff")
                    s_cont.grid(row=0, column=3, sticky="ew")
                    ttk.Label(s_cont, text="Custom Seal NO", style="CardLabel.TLabel").pack(anchor="w", pady=(0, 4))
                    s_entry = ttk.Entry(s_cont, style="Card.TEntry", width=15)
                    s_entry.pack(fill=tk.X)
                    self.entries["Custom Seal NO"] = s_entry
                    
                    continue

                # Standard Processing for other fields
                # Container for the field
                field_container = tk.Frame(self.form_frame, bg="#ffffff")
                field_container.grid(row=current_row, column=current_col, columnspan=span, sticky="ew", padx=10, pady=8)
                
                # Update column pointer
                current_col += span

                # Determine Widget Type
                is_multiline = label_text in ["Exporter", "Bill No", "Quantity", "Total Quantity", "Country of Origin", "Air way Bill NO"]
                
                # Label (Top)
                ttk.Label(field_container, text=label_text, style="CardLabel.TLabel").pack(anchor="w", pady=(0, 4))
                
                # Input (Bottom)
                if is_multiline:
                    entry = tk.Text(
                        field_container, 
                        height=2, 
                        width=30, 
                        font=("Segoe UI", 10), 
                        relief="flat", bd=0,
                        highlightthickness=1,
                        highlightcolor="#3498db",
                        highlightbackground="#dcdde1",
                        bg="#f8f9fa",
                        undo=True, maxundo=-1
                    )
                    entry.bind("<Tab>", self.focus_next_window)
                    entry.pack(fill=tk.X)
                else:
                    entry = ttk.Entry(field_container, style="Card.TEntry", width=15)
                    entry.pack(fill=tk.X)
                
                self.entries[label_text] = entry
            
            # Move to next row
            current_row += 1
        
        # Description was moved into the loop above.

        # 3. Footer Actions (Floating or Sticky)
        # We put them inside the card at bottom
        ttk.Separator(self.card, orient="horizontal").pack(fill=tk.X, pady=20)
        
        action_frame = tk.Frame(self.card, bg="#ffffff")
        action_frame.pack(fill=tk.X)
        
        # Primary Action
        gen_btn = ttk.Button(action_frame, text="GENERATE PDF", command=self.generate_pdf_label, style="Primary.TButton")
        gen_btn.pack(side=tk.RIGHT, padx=5)
        
        # Secondary Actions
        ttk.Button(action_frame, text="Save As...", command=self.generate_pdf_label_saveas, style="Secondary.TButton").pack(side=tk.RIGHT, padx=5)
        ttk.Button(action_frame, text="Open PDF", command=self.open_last_pdf_file, style="Secondary.TButton").pack(side=tk.RIGHT, padx=5)
        ttk.Button(action_frame, text="Print", command=self.print_last_pdf, style="Secondary.TButton").pack(side=tk.RIGHT, padx=5)
        ttk.Button(action_frame, text="Load", command=self.load_old_document, style="Secondary.TButton").pack(side=tk.RIGHT, padx=5)
        
        # Clear (Left side)
        ttk.Button(action_frame, text="Clear Form", command=self.clear_fields, style="Secondary.TButton").pack(side=tk.LEFT, padx=0)

        # === Preview ===
        ttk.Label(self.scrollable_frame, text="Preview Data:", style="CardLabel.TLabel").pack(pady=(20, 5))
        self.preview_text = tk.Text(
            self.scrollable_frame, 
            height=6, 
            width=85, 
            font=("Consolas", 8), 
            relief="flat", bd=0,
            highlightthickness=1,
            highlightcolor="#3498db",
            highlightbackground="#dcdde1",
            bg="#f8f9fa", 
            fg="#2c3e50"
        )
        self.preview_text.pack(padx=20, pady=5)


    def _on_mousewheel(self, event):
        """Enable mousewheel scrolling."""
        self.canvas.yview_scroll(int(-1*(event.delta/120)), "units")

    def focus_next_window(self, event):
        """Handle Tab key to move focus to next widget."""
        event.widget.tk_focusNext().focus()
        return "break"

    def show_toast(self, message, duration=3000):
        """Show a temporary notification (toast)."""
        toast = tk.Toplevel(self.root)
        toast.wm_overrideredirect(True)
        
        # Position
        x = self.root.winfo_x() + (self.root.winfo_width() // 2) - 100
        y = self.root.winfo_y() + self.root.winfo_height() - 100
        toast.geometry(f"200x50+{x}+{y}")
        
        # Style
        toast.configure(bg="#2d3436", relief="solid", bd=1)
        label = tk.Label(toast, text=message, bg="#2d3436", fg="white", font=("Segoe UI", 10))
        label.pack(expand=True, fill="both")
        
        # Close after duration
        toast.after(duration, toast.destroy)

    def open_last_pdf_file(self):
        """Open the last generated PDF."""
        if not self.last_pdf_path or not os.path.exists(self.last_pdf_path):
            messagebox.showwarning("Open PDF", "No PDF file found to open.")
            return
        
        try:
            os.startfile(self.last_pdf_path)
        except Exception as e:
            messagebox.showerror("Error", f"Could not open file: {e}")



    # -------------------------
    # Settings (load/save)
    # -------------------------
    def load_settings(self):
        """Load alignment offsets, default folder, and filename base from JSON."""
        if not os.path.exists(ALIGNMENT_FILE):
            # Defaults: offsets 0, folder = user's Documents/ExitLabels, filename base 'exit'
            default_folder = os.path.join(os.path.expanduser("~"), "Documents", "ExitLabels")
            return 0.0, 0.0, default_folder, "exit"
        try:
            with open(ALIGNMENT_FILE, "r") as f:
                data = json.load(f)
            ox = float(data.get("offset_x", 0.0))
            oy = float(data.get("offset_y", 0.0))
            folder = data.get("default_folder", os.path.join(os.path.expanduser("~"), "Documents", "ExitLabels"))
            fname = data.get("filename_base", "exit")
            return ox, oy, folder, fname
        except Exception:
            return 0.0, 0.0, os.path.join(os.path.expanduser("~"), "Documents", "ExitLabels"), "exit"

    def save_settings_to_file(self):
        """Persist current settings to JSON."""
        data = {
            "offset_x": self.offset_x,
            "offset_y": self.offset_y,
            "default_folder": self.folder_var.get() or self.default_folder,
            "filename_base": self.filename_base_var.get() or "exit",
        }
        with open(ALIGNMENT_FILE, "w") as f:
            json.dump(data, f, indent=4)

    def save_all_settings(self):
        """Save offsets, folder and filename base from UI into settings file."""
        try:
            # offsets already in self.offset_x / offset_y
            self.default_folder = self.folder_var.get() or self.default_folder
            self.filename_base = self.filename_base_var.get() or "exit"
            self.save_settings_to_file()
            messagebox.showinfo("Saved", "Settings saved to alignment_settings.json")
        except Exception as e:
            messagebox.showerror("Error", f"Could not save settings: {e}")

    # -------------------------
    # Folder chooser
    # -------------------------
    def choose_folder(self):
        folder = filedialog.askdirectory(title="Choose default folder to save PDFs")
        if folder:
            self.folder_var.set(folder)

    # -------------------------
    # Alignment Settings Window
    # -------------------------
    def open_alignment_window(self):
        win = tk.Toplevel(self.root)
        win.title("Adjust PDF Alignment")
        win.geometry("360x230")
        win.config(bg="#f1f2f6")

        tk.Label(win, text="X Offset (Left/Right) — inches", font=("Arial", 11)).pack(pady=5)
        x_entry = tk.Entry(win, width=18)
        x_entry.insert(0, str(self.offset_x))
        x_entry.pack()

        tk.Label(win, text="Y Offset (Up/Down) — inches", font=("Arial", 11)).pack(pady=5)
        y_entry = tk.Entry(win, width=18)
        y_entry.insert(0, str(self.offset_y))
        y_entry.pack()

        def save():
            try:
                new_x = float(x_entry.get())
                new_y = float(y_entry.get())
                self.offset_x, self.offset_y = new_x, new_y
                self.save_settings_to_file()
                messagebox.showinfo("Saved", "Alignment settings saved!")
                win.destroy()
            except ValueError:
                messagebox.showerror("Error", "Offsets must be numbers (can be decimals).")

        def reset():
            x_entry.delete(0, tk.END)
            y_entry.delete(0, tk.END)
            x_entry.insert(0, "0")
            y_entry.insert(0, "0")

        tk.Button(win, text="Save", bg="#27ae60", fg="white", font=("Arial", 11, "bold"), command=save).pack(pady=10)
        tk.Button(win, text="Reset to Default", bg="#c0392b", fg="white", font=("Arial", 11, "bold"), command=reset).pack()

    # -------------------------
    # Main functions
    # -------------------------
    def clear_fields(self):
        for entry in self.entries.values():
            if isinstance(entry, tk.Text):
                entry.delete("1.0", tk.END)
            else:
                entry.delete(0, tk.END)
        self.desc_text.delete("1.0", tk.END)
        self.preview_text.delete("1.0", tk.END)

    def generate_auto_filename(self):
        """Generate a unique filename using the base and timestamp."""
        base = (self.filename_base_var.get() or self.filename_base or "exit").strip()
        ts = datetime.now().strftime("%Y%m%d_%H%M%S")
        fname = f"{base}_{ts}.pdf"
        return fname

    def ensure_folder_exists(self, folder_path):
        """Create folder if missing."""
        if not folder_path:
            # fallback to Documents/ExitLabels
            folder_path = os.path.join(os.path.expanduser("~"), "Documents", "ExitLabels")
        if not os.path.exists(folder_path):
            try:
                os.makedirs(folder_path, exist_ok=True)
            except Exception as e:
                raise IOError(f"Could not create folder '{folder_path}': {e}")
        return folder_path

    def generate_pdf_label(self):
        """Create PDF and auto-save to default folder (no dialog)."""
        data = {}
        for label, widget in self.entries.items():
            if isinstance(widget, tk.Text):
                data[label] = widget.get("1.0", tk.END).strip()
            else:
                data[label] = widget.get().strip()
        data["Description"] = self.desc_text.get("1.0", tk.END).strip()

        # Show preview in UI
        preview = "\n".join([f"{k}: {v}" for k, v in data.items()])
        self.preview_text.delete("1.0", tk.END)
        self.preview_text.insert(tk.END, preview)

        # Determine folder and ensure it exists
        folder = self.folder_var.get() or self.default_folder
        try:
            folder = self.ensure_folder_exists(folder)
        except IOError as e:
            messagebox.showerror("Folder Error", str(e))
            return

        # Generate filename and full path
        filename = self.generate_auto_filename()
        fullpath = os.path.join(folder, filename)

        # Create PDF
        try:
            self.create_pdf_label(data, fullpath)

            # Save JSON copy (Option A)
            json_path = fullpath.replace(".pdf", ".json")
            try:
                with open(json_path, "w", encoding="utf-8") as jf:
                    json.dump(data, jf, indent=4, ensure_ascii=False)
            except Exception as je:
                # Not fatal — show a warning but keep going
                messagebox.showwarning("JSON Save Warning", f"PDF created but failed to save JSON: {je}")

            # Save current folder/filename base to settings
            self.default_folder = folder
            self.filename_base = self.filename_base_var.get() or self.filename_base
            self.save_settings_to_file()

            # store last pdf for printing/opening
            self.last_pdf_path = fullpath

            self.show_toast("PDF Generated Successfully!")
            
        except Exception as e:
            messagebox.showerror("Error", f"Failed to create PDF: {e}")

    def generate_pdf_label_saveas(self):
        """Generate PDF but open Save As dialog — user chooses exact filename."""
        data = {}
        for label, widget in self.entries.items():
            if isinstance(widget, tk.Text):
                data[label] = widget.get("1.0", tk.END).strip()
            else:
                data[label] = widget.get().strip()
        data["Description"] = self.desc_text.get("1.0", tk.END).strip()

        # Show preview
        preview = "\n".join([f"{k}: {v}" for k, v in data.items()])
        self.preview_text.delete("1.0", tk.END)
        self.preview_text.insert(tk.END, preview)

        # Ask user where to save (fallback to default folder as initialdir)
        initial_dir = self.folder_var.get() or self.default_folder or os.path.expanduser("~")
        save_path = filedialog.asksaveasfilename(
            defaultextension=".pdf",
            filetypes=[("PDF Files", "*.pdf")],
            title="Save Exit Type PDF As",
            initialdir=initial_dir,
            initialfile=f"{self.filename_base_var.get() or self.filename_base}_",
        )
        if not save_path:
            return
        try:
            # ensure folder exists for chosen path
            folder = os.path.dirname(save_path)
            self.ensure_folder_exists(folder)
            self.create_pdf_label(data, save_path)

            # Save JSON copy (Option A)
            json_path = save_path.replace(".pdf", ".json")
            try:
                with open(json_path, "w", encoding="utf-8") as jf:
                    json.dump(data, jf, indent=4, ensure_ascii=False)
            except Exception as je:
                messagebox.showwarning("JSON Save Warning", f"PDF created but failed to save JSON: {je}")

            # update default folder to chosen folder
            self.folder_var.set(folder)
            self.default_folder = folder
            self.filename_base = self.filename_base_var.get() or self.filename_base
            self.save_settings_to_file()

            # store last pdf for printing/opening
            self.last_pdf_path = save_path

            self.show_toast("PDF Saved Successfully!")
            
        except Exception as e:
            messagebox.showerror("Error", f"Failed to create PDF: {e}")

    def create_pdf_label(self, data, filename):
        """Generate the label PDF with global offset applied (offset are in inches)."""
        c = canvas.Canvas(filename, pagesize=letter)
        width, height = letter
        c.setFillColor(colors.black)

        # Convert offset inches to points (ReportLab uses points)
        ox_in = float(self.offset_x)
        oy_in = float(self.offset_y)
        ox = ox_in * inch
        oy = oy_in * inch

        # Use registered Verdana if available
        try:
            c.setFont("Verdana", 10)
        except Exception:
            c.setFont("Helvetica", 10)

        field_positions = {
            "Inspection No": (4.7 * inch, height - 1.6 * inch),
            "Exporter": (-0.9 * inch, height - 2.16 * inch),
            "Inspection Date": (4.55 * inch, height - 2.2 * inch),
            "Bill No": (-1.3 * inch, height - 3.13 * inch),
            "BOE Date": (1.65 * inch, height - 3.2 * inch),
            "Air way Bill NO": (4.05 * inch, height - 3.2 * inch),
            "Country of Origin": (-1.3 * inch, height - 3.83 * inch),
            "Point of Exit": (1.45 * inch, height - 3.9 * inch),
            "Destination": (4.35 * inch, height - 3.9 * inch),
            "Quantity": (-1.55 * inch, height - 5.5 * inch),
            "Total Quantity": (-0.9 * inch, height - 7.35 * inch),
            "Total Weight": (1.7 * inch, height - 7.38 * inch),
            "Container/Vehicle No": (-1.2 * inch, height - 8.23 * inch),
            "Custom Seal NO": (1.7 * inch, height - 8.23 * inch),
        }

        for key, (x, y) in field_positions.items():
            value = data.get(key, "")
            x += ox
            y += oy
            if key in ["Exporter", "Bill No", "Quantity", "Total Quantity", "Country of Origin", "Air way Bill NO"]:
                lines = value.splitlines()
                y_off = 0
                for line in lines:
                    c.drawString(x + 1.8 * inch, y - y_off, line)
                    y_off += 0.15 * inch
            else:
                c.drawString(x + 1.8 * inch, y, value)

        # Description block
        desc_x = 1 * inch + ox
        desc_y = height - 4.9 * inch + oy
        desc_lines = data.get("Description", "").splitlines()
        y_offset = desc_y - 0.3 * inch
        for line in desc_lines:
            c.drawString(desc_x + 1 * inch, y_offset, line)
            y_offset -= 0.25 * inch

        c.save()

    # -------------------------
    # New: Load saved JSON and fill form
    # -------------------------
    def load_old_document(self):
        """Load JSON file and auto-fill all fields."""
        json_file = filedialog.askopenfilename(
            title="Select Saved Document (JSON)",
            filetypes=[("JSON Files", "*.json")],
            initialdir=self.folder_var.get() or self.default_folder or os.path.expanduser("~"),
        )
        if not json_file:
            return

        try:
            with open(json_file, "r", encoding="utf-8") as f:
                data = json.load(f)

            # Fill all fields
            for key, widget in self.entries.items():
                if key in data:
                    if isinstance(widget, tk.Text):
                        widget.delete("1.0", tk.END)
                        widget.insert("1.0", data[key])
                    else:
                        widget.delete(0, tk.END)
                        widget.insert(0, data[key])

            # Description
            if "Description" in data:
                self.desc_text.delete("1.0", tk.END)
                self.desc_text.insert("1.0", data["Description"])

            # Set preview
            preview = "\n".join([f"{k}: {v}" for k, v in data.items()])
            self.preview_text.delete("1.0", tk.END)
            self.preview_text.insert(tk.END, preview)

            # Set last_pdf_path if JSON and PDF same base exists in same dir
            possible_pdf = json_file.replace(".json", ".pdf")
            if os.path.exists(possible_pdf):
                self.last_pdf_path = possible_pdf

            messagebox.showinfo("Loaded", "Saved document loaded successfully.")
        except Exception as e:
            messagebox.showerror("Error", f"Failed to load document:\n{e}")

    # -------------------------
    # New: Print PDF helper
    # -------------------------
    def print_last_pdf(self):
        """Print the most recently created PDF directly to default printer."""
        try:
            if not self.last_pdf_path:
                messagebox.showerror("Error", "No PDF available to print. Create or load a document first.")
                return
            
            if not os.path.exists(self.last_pdf_path):
                messagebox.showerror("Error", "PDF file not found. Please regenerate or load a document.")
                return

            system = platform.system()
            if system == "Windows":
                # This sends to the default printer without dialog
                try:
                    os.startfile(self.last_pdf_path, "print")
                    messagebox.showinfo("Printing", "Document sent to default printer.")
                except Exception as e:
                    messagebox.showerror("Printing Error", f"Failed to print (Windows):\n{e}")
            elif system in ("Linux", "Darwin"):
                # Try using lpr (common on UNIX). This may require lpr to be installed.
                try:
                    subprocess.run(["lpr", self.last_pdf_path], check=True)
                    messagebox.showinfo("Printing", "Document sent to printer via lpr.")
                except FileNotFoundError:
                    messagebox.showerror("Printing Error", "Printing not supported on this OS in this app (lpr not found).")
                except subprocess.CalledProcessError as e:
                    messagebox.showerror("Printing Error", f"lpr failed:\n{e}")
                except Exception as e:
                    messagebox.showerror("Printing Error", f"Unexpected printing error:\n{e}")
            else:
                messagebox.showerror("Printing Error", f"Unsupported platform for direct print: {system}")

        except Exception as e:
            messagebox.showerror("Error", f"Printing failed:\n{e}")


if __name__ == "__main__":
    root = tk.Tk()
    app = ExitTypeGeneratorApp(root)
    root.mainloop()
