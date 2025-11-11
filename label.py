import tkinter as tk
from tkinter import messagebox, filedialog
from reportlab.lib.pagesizes import letter
from reportlab.pdfgen import canvas
from reportlab.lib.units import cm, inch
from reportlab.lib import colors
from reportlab.pdfbase import pdfmetrics
from reportlab.pdfbase.ttfonts import TTFont
import os


class ExitTypeGeneratorApp:
    def __init__(self, root):
        self.root = root
        self.root.title("Exit Type Generator")
        self.root.geometry("600x750")
        self.root.configure(bg="#f4f6f9")

        # Register Verdana fonts
        verdana_path = "C:\\Windows\\Fonts\\verdana.ttf"
        verdana_bold_path = "C:\\Windows\\Fonts\\verdanab.ttf"
        pdfmetrics.registerFont(TTFont("Verdana", verdana_path))
        pdfmetrics.registerFont(TTFont("Verdana-Bold", verdana_bold_path))

        # Try setting icon
        try:
            self.root.iconbitmap("labelicon.ico")
        except tk.TclError:
            print("Icon not found, using default")

        # Title label
        title_label = tk.Label(
            root, text="Exit Type Generator",
            font=("Arial", 18, "bold"), bg="#2e86de", fg="white", pady=10
        )
        title_label.pack(fill=tk.X)

        # Form frame
        self.form_frame = tk.Frame(root, bg="#f4f6f9")
        self.form_frame.pack(pady=15)

        # Labels list
        labels = [
            "Inspection No", "Exporter", "Inspection Date", "Bill No", "BOE Date",
            "Manifest", "Country of Origin", "Point of Exit", "Destination",
            "Quantity", "Total Quantity", "Total Weight", "Container/Vehicle No"
        ]

        self.entries = {}

        # Create form entries
        for i, label_text in enumerate(labels):
            tk.Label(self.form_frame, text=label_text + ":", font=("Arial", 10), bg="#f4f6f9") \
                .grid(row=i, column=0, sticky="e", padx=10, pady=5)

            # Exporter and Bill No are multi-line text boxes
            if label_text in ["Exporter", "Bill No"]:
                entry = tk.Text(self.form_frame, width=45, height=3, font=("Arial", 10))
            else:
                entry = tk.Entry(self.form_frame, width=45)

            entry.grid(row=i, column=1, padx=10, pady=5)
            self.entries[label_text] = entry

        # Description (multi-line)
        desc_label = tk.Label(self.form_frame, text="Description:", font=("Arial", 10), bg="#f4f6f9")
        desc_label.grid(row=len(labels), column=0, sticky="ne", padx=10, pady=5)
        self.desc_text = tk.Text(self.form_frame, width=45, height=5, font=("Arial", 10))
        self.desc_text.grid(row=len(labels), column=1, padx=10, pady=5)

        # Buttons
        btn_frame = tk.Frame(root, bg="#f4f6f9")
        btn_frame.pack(pady=10)

        tk.Button(btn_frame, text="Generate PDF", command=self.generate_pdf_label,
                  bg="#27ae60", fg="white", font=("Arial", 11, "bold"), padx=15, pady=5).grid(row=0, column=0, padx=10)
        tk.Button(btn_frame, text="Clear All", command=self.clear_fields,
                  bg="#c0392b", fg="white", font=("Arial", 11, "bold"), padx=15, pady=5).grid(row=0, column=1, padx=10)

        # Preview area
        tk.Label(root, text="Exit Label Preview:", font=("Arial", 11, "bold"), bg="#f4f6f9").pack(pady=(10, 0))
        self.preview_text = tk.Text(root, height=10, width=70, font=("Courier New", 10))
        self.preview_text.pack(padx=10, pady=5)

    def clear_fields(self):
        """Clear all input fields and preview"""
        for entry in self.entries.values():
            if isinstance(entry, tk.Text):
                entry.delete("1.0", tk.END)
            else:
                entry.delete(0, tk.END)
        self.desc_text.delete('1.0', tk.END)
        self.preview_text.delete('1.0', tk.END)

    def generate_pdf_label(self):
        """Collect data, display preview, and create PDF"""
        data = {}
        for label, widget in self.entries.items():
            if isinstance(widget, tk.Text):
                data[label] = widget.get("1.0", tk.END).strip()
            else:
                data[label] = widget.get().strip()
        data["Description"] = self.desc_text.get("1.0", tk.END).strip()

        # Show preview
        preview_lines = []
        for label, value in data.items():
            preview_lines.append(f"{label}: {value}")
        preview_text = "\n".join(preview_lines)

        self.preview_text.delete('1.0', tk.END)
        self.preview_text.insert(tk.END, preview_text)

        # Ask where to save
        pdf_filename = filedialog.asksaveasfilename(
            defaultextension=".pdf", filetypes=[("PDF Files", "*.pdf")],
            title="Save Exit Type PDF As"
        )
        if not pdf_filename:
            return

        # Generate PDF
        self.create_pdf_label(data, pdf_filename)
        messagebox.showinfo("Success", f"Exit Type PDF saved as:\n{pdf_filename}")

    def create_pdf_label(self, data, filename):
        """Generate the label PDF with precise field positioning"""
        c = canvas.Canvas(filename, pagesize=letter)
        width, height = letter
        c.setFillColor(colors.black)
        c.setFont("Verdana", 10)

        field_positions = {
            "Inspection No": (4.7 * inch, height - 1.4 * inch),
            "Exporter": (-0.88 * inch, height - 2 * inch),
            "Inspection Date": (4.65 * inch, height - 2.1 * inch),
            "Bill No": (-1.3 * inch, height - 3 * inch),
            "BOE Date": (1.7 * inch, height - 3 * inch),
            "Manifest": (4.05 * inch, height - 3 * inch),
            "Country of Origin": (-1.2 * inch, height - 3.75 * inch),
            "Point of Exit": (1.5 * inch, height - 3.75 * inch),
            "Destination": (4.35 * inch, height - 3.75 * inch),
            "Quantity": (-1.5 * inch, height - 5.5 * inch),
            "Total Quantity": (-0.86 * inch, height - 7.2 * inch),
            "Total Weight": (1.7 * inch, height - 7.2 * inch),
            "Container/Vehicle No": (-1.2 * inch, height - 8 * inch),
        }

        # Draw fields
        for key, (x, y) in field_positions.items():
            value = data.get(key, "")
            c.setFont("Verdana-Bold", 10)

            if key in ["Exporter", "Bill No"]:
                lines = value.splitlines()
                y_offset = 0
                for line in lines:
                    c.drawString(x + 1.8 * inch, y - y_offset, line)
                    y_offset += 0.25 * inch
            else:
                c.drawString(x + 1.8 * inch, y, value)

        # Description block
        desc_x = 1 * inch
        desc_y = height - 4.8 * inch
        # c.drawString(desc_x, desc_y, "Description:")
        # c.setFont("Verdana-Bold", 10)

        desc_text = data.get("Description", "")
        desc_lines = desc_text.splitlines()

        y_offset = desc_y - 0.3 * inch
        for line in desc_lines:
            c.drawString(desc_x + 1.1 * inch, y_offset, line)
            y_offset -= 0.25 * inch

        c.save()


if __name__ == "__main__":
    root = tk.Tk()
    app = ExitTypeGeneratorApp(root)
    root.mainloop()
