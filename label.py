import tkinter as tk
from tkinter import messagebox, filedialog
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
            # Load icon once (important!)
        self.icon_save = ImageTk.PhotoImage(
            Image.open(resource_path("icons/save.png")).resize((28, 28), Image.LANCZOS)
        )
        self.icon_saveas = ImageTk.PhotoImage(
            Image.open(resource_path("icons/saveas.png")).resize((28, 28), Image.LANCZOS)
        )
        self.icon_clear = ImageTk.PhotoImage(
            Image.open(resource_path("icons/clear.png")).resize((28, 28), Image.LANCZOS)
        )
        self.icon_load = ImageTk.PhotoImage(
            Image.open(resource_path("icons/add.png")).resize((28, 28), Image.LANCZOS)
        )
        self.icon_print = ImageTk.PhotoImage(
            Image.open(resource_path("icons/print.png")).resize((28, 28), Image.LANCZOS)
        )
        self.root = root
        self.root.title("Exit Type Generator")
        self.root.geometry("800x880")
        self.root.configure(bg="#f4f6f9")

        # Track last created PDF path (for printing)
        self.last_pdf_path = None

        # Load settings (offsets + folder + filename base)
        (
            self.offset_x,
            self.offset_y,
            self.default_folder,
            self.filename_base,
        ) = self.load_settings()

        # Register Verdana fonts (use fallback if not present)
        try:
            verdana_path = "C:\\Windows\\Fonts\\verdana.ttf"
            verdana_bold_path = "C:\\Windows\\Fonts\\verdanab.ttf"
            pdfmetrics.registerFont(TTFont("Verdana", verdana_path))
            pdfmetrics.registerFont(TTFont("Verdana-Bold", verdana_bold_path))
        except Exception:
            # If Verdana not found, register a standard font (ReportLab default)
            pass

        # Try setting icon
        try:
            # self.root.iconbitmap("exittype.ico")
            icon_path = resource_path("exittype.ico")
            self.root.iconbitmap(icon_path)

            # base_path = sys._MEIPASS  # PyInstaller temp folder
        except tk.TclError:
            print("Icon not found, using default")

        # Title
        tk.Label(
            root,
            text="Exit Type Generator",
            font=("Arial", 18, "bold"),
            bg="#2e86de",
            fg="white",
            pady=10,
        ).pack(fill=tk.X)

        # Top controls: Alignment / Folder
        top_frame = tk.Frame(root, bg="#f4f6f9")
        top_frame.pack(fill=tk.X, pady=6)

        tk.Button(
            top_frame,
            text="Adjust Alignment (X/Y)",
            command=self.open_alignment_window,
            bg="#8e44ad",
            fg="white",
            font=("Arial", 11, "bold"),
            padx=8,
            pady=5,
        ).pack(side=tk.LEFT, padx=8)

        # Folder display + choose
        folder_frame = tk.Frame(top_frame, bg="#f4f6f9")
        folder_frame.pack(side=tk.RIGHT, padx=8)

        tk.Label(folder_frame, text="Save Folder:", bg="#f4f6f9").grid(row=0, column=0, sticky="e")
        self.folder_var = tk.StringVar(value=self.default_folder or "")
        self.folder_entry = tk.Entry(folder_frame, textvariable=self.folder_var, width=48)
        self.folder_entry.grid(row=0, column=1, padx=6)
        tk.Button(folder_frame, text="Choose Folder", command=self.choose_folder).grid(row=0, column=2, padx=6)

        # Filename base (auto 'exit' by your choice)
        fname_frame = tk.Frame(root, bg="#f4f6f9")
        fname_frame.pack(fill=tk.X, pady=(0, 6))
        tk.Label(fname_frame, text="Filename base:", bg="#f4f6f9").pack(side=tk.LEFT, padx=(12, 4))
        self.filename_base_var = tk.StringVar(value=self.filename_base or "exit")
        self.filename_base_entry = tk.Entry(fname_frame, textvariable=self.filename_base_var, width=20)
        self.filename_base_entry.pack(side=tk.LEFT)

        # Save settings button (saves offsets + folder + filename base)
        tk.Button(
            fname_frame,
            text="Save Settings",
            command=self.save_all_settings,
            bg="#27ae60",
            fg="white",
            font=("Arial", 10, "bold"),
            padx=8,
            pady=3,
        ).pack(side=tk.LEFT, padx=12)
        

        # Form
        self.form_frame = tk.Frame(root, bg="#f4f6f9")
        self.form_frame.pack(pady=8)

        labels = [
            "Inspection No",
            "Exporter",
            "Inspection Date",
            "Bill No",
            "BOE Date",
            "Air way Bill NO",
            "Country of Origin",
            "Point of Exit",
            "Destination",
            "Quantity",
            "Total Quantity",
            "Total Weight",
            "Container/Vehicle No",
            "Custom Seal NO",
        ]

        self.entries = {}

        for i, label_text in enumerate(labels):
            tk.Label(
                self.form_frame,
                text=label_text + ":",
                font=("Arial", 10, "bold"),
                bg="#f4f6f9",
            ).grid(row=i, column=0, sticky="e", padx=10, pady=5)

            if label_text in ["Exporter", "Bill No", "Quantity", "Total Quantity", "Country of Origin", "Air way Bill NO"]:
                entry = tk.Text(self.form_frame, width=50, height=2, font=("Arial", 11, "bold"))
            else:
                entry = tk.Entry(self.form_frame, width=50, font=("Arial", 11, "bold"))

            entry.grid(row=i, column=1, padx=10, pady=5)
            self.entries[label_text] = entry

        # Description
        tk.Label(
            self.form_frame,
            text="Description:",
            font=("Arial", 10, "bold"),
            bg="#f4f6f9",
        ).grid(row=len(labels), column=0, sticky="ne", padx=10, pady=5)
        self.desc_text = tk.Text(self.form_frame, width=50, height=3, font=("Arial", 11, "bold"))
        self.desc_text.grid(row=len(labels), column=1, padx=10, pady=5)

        # Buttons
        btn_frame = tk.Frame(root, bg="#f4f6f9")
        btn_frame.pack(pady=10)

        # tk.Button(
        #     btn_frame,
        #     text="Generate PDF (Auto-save)",
        #     command=self.generate_pdf_label,
        #     bg="#27ae60",
        #     fg="white",
        #     font=("Arial", 11, "bold"),
        #     padx=15,
        #     pady=5,
        # ).grid(row=0, column=0, padx=10)
        # tk.Button(
        #     btn_frame,
        #     image=self.icon_generate,
        #     command=self.generate_pdf_label,
        #     bg="#27ae60",
        #     bd=0,
        #     activebackground="#27ae60",
        #     padx=10,
        #     pady=10
        # ).grid(row=0, column=0, padx=10)
        generate_btn = tk.Button(
            btn_frame,
            image=self.icon_save,
            command=self.generate_pdf_label,
            bg="#e2e9e5",
            activebackground="#27ae60",
            bd=0,
            padx=10,
            pady=10,
        )
        generate_btn.grid(row=0, column=0, padx=10)

        # Add tooltip here
        ToolTip(generate_btn, "Generate PDF")



        # tk.Button(
        #     btn_frame,
        #     text="Generate PDF (Save As...)",
        #     command=self.generate_pdf_label_saveas,
        #     bg="#2980b9",
        #     fg="white",
        #     font=("Arial", 11, "bold"),
        #     padx=15,
        #     pady=5,
        # ).grid(row=0, column=1, padx=10)
        saveas_btn = tk.Button(
            btn_frame,
            image=self.icon_saveas,
            command=self.generate_pdf_label_saveas,
            bg="#e0e0e0",
            activebackground="#27ae60",
            bd=0,
            padx=10,
            pady=10,
        )
        saveas_btn.grid(row=0, column=1, padx=10)

        # Add tooltip here
        ToolTip(saveas_btn, "Save As")
        

        # tk.Button(
        #     btn_frame,
        #     text="Clear All",
        #     command=self.clear_fields,
        #     bg="#c0392b",
        #     fg="white",
        #     font=("Arial", 11, "bold"),
        #     padx=15,
        #     pady=5,
        # ).grid(row=0, column=2, padx=10)

        clear_btn = tk.Button(
            btn_frame,
            image=self.icon_clear,
            command=self.clear_fields,
            bg="#e0e0e0",
            activebackground="#27ae60",
            bd=0,
            padx=10,
            pady=10,
        )
        clear_btn.grid(row=0, column=2, padx=10)

        # Add tooltip here
        ToolTip(clear_btn, "Clear ALL")

        # New: Load old doc and print buttons
        # tk.Button(
        #     btn_frame,
        #     text="Load Old Document",
        #     command=self.load_old_document,
        #     bg="#8e44ad",
        #     fg="white",
        #     font=("Arial", 11, "bold"),
        #     padx=15,
        #     pady=5,
        # ).grid(row=0, column=3, padx=10)

        load_btn = tk.Button(
            btn_frame,
            image=self.icon_load,
            command=self.load_old_document,
            bg="#e0e0e0",
            activebackground="#27ae60",
            bd=0,
            padx=10,
            pady=10,
        )
        load_btn.grid(row=0, column=3, padx=10)

        # Add tooltip here
        ToolTip(load_btn, "Load Old Document")


        # tk.Button(
        #     btn_frame,
        #     text="Print PDF",
        #     command=self.print_last_pdf,
        #     bg="#16a085",
        #     fg="white",
        #     font=("Arial", 11, "bold"),
        #     padx=15,
        #     pady=5,
        # ).grid(row=0, column=4, padx=10)

        print_btn = tk.Button(
            btn_frame,
            image=self.icon_print,
            command=self.print_last_pdf,
            bg="#e0e0e0",
            activebackground="#27ae60",
            bd=0,
            padx=10,
            pady=10,
        )
        print_btn.grid(row=0, column=4, padx=10)

        # Add tooltip here
        ToolTip(print_btn, "print_last_pdf")

        # Preview
        tk.Label(root, text="Exit Label Preview:", font=("Arial", 11, "bold"), bg="#f4f6f9").pack(pady=(10, 0))
        self.preview_text = tk.Text(root, height=10, width=90, font=("Arial", 10))
        self.preview_text.pack(padx=10, pady=5)

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

            # store last pdf for printing
            self.last_pdf_path = fullpath

            messagebox.showinfo("Success", f"PDF saved to:\n{fullpath}")
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

            # store last pdf for printing
            self.last_pdf_path = save_path

            messagebox.showinfo("Success", f"PDF saved to:\n{save_path}")
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
