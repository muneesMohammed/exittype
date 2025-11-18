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

ALIGNMENT_FILE = "alignment_settings.json"


class ExitTypeGeneratorApp:
    def __init__(self, root):
        self.root = root
        self.root.title("Exit Type Generator")
        self.root.geometry("780x820")
        self.root.configure(bg="#f4f6f9")

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
            self.root.iconbitmap("labelicon.ico")
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

            if label_text in ["Exporter", "Bill No", "Quantity", "Total Quantity"]:
                entry = tk.Text(self.form_frame, width=46, height=2, font=("Arial", 10))
            else:
                entry = tk.Entry(self.form_frame, width=56)

            entry.grid(row=i, column=1, padx=10, pady=5)
            self.entries[label_text] = entry

        # Description
        tk.Label(
            self.form_frame,
            text="Description:",
            font=("Arial", 10, "bold"),
            bg="#f4f6f9",
        ).grid(row=len(labels), column=0, sticky="ne", padx=10, pady=5)
        self.desc_text = tk.Text(self.form_frame, width=46, height=3, font=("Arial", 10))
        self.desc_text.grid(row=len(labels), column=1, padx=10, pady=5)

        # Buttons
        btn_frame = tk.Frame(root, bg="#f4f6f9")
        btn_frame.pack(pady=10)

        tk.Button(
            btn_frame,
            text="Generate PDF (Auto-save)",
            command=self.generate_pdf_label,
            bg="#27ae60",
            fg="white",
            font=("Arial", 11, "bold"),
            padx=15,
            pady=5,
        ).grid(row=0, column=0, padx=10)

        tk.Button(
            btn_frame,
            text="Generate PDF (Save As...)",
            command=self.generate_pdf_label_saveas,
            bg="#2980b9",
            fg="white",
            font=("Arial", 11, "bold"),
            padx=15,
            pady=5,
        ).grid(row=0, column=1, padx=10)

        tk.Button(
            btn_frame,
            text="Clear All",
            command=self.clear_fields,
            bg="#c0392b",
            fg="white",
            font=("Arial", 11, "bold"),
            padx=15,
            pady=5,
        ).grid(row=0, column=2, padx=10)

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
            # Save current folder/filename base to settings
            self.default_folder = folder
            self.filename_base = self.filename_base_var.get() or self.filename_base
            self.save_settings_to_file()
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

            # update default folder to chosen folder
            self.folder_var.set(folder)
            self.default_folder = folder
            self.filename_base = self.filename_base_var.get() or self.filename_base
            self.save_settings_to_file()

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
            "Inspection Date": (4.6 * inch, height - 2.2 * inch),
            "Bill No": (-1.3 * inch, height - 3.13 * inch),
            "BOE Date": (1.7 * inch, height - 3.2 * inch),
            "Air way Bill NO": (4.05 * inch, height - 3.25 * inch),
            "Country of Origin": (-1.2 * inch, height - 3.9 * inch),
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
            if key in ["Exporter", "Bill No", "Quantity", "Total Quantity"]:
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


if __name__ == "__main__":
    root = tk.Tk()
    app = ExitTypeGeneratorApp(root)
    root.mainloop()
